VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmBatalBilling 
   Caption         =   "Retur Invoice"
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
   Begin VB.CommandButton TblKeluar 
      Caption         =   "&Kembali ke Menu"
      Height          =   390
      Left            =   9870
      TabIndex        =   3
      Top             =   7110
      Width           =   1890
   End
   Begin VB.CommandButton TblSimpan 
      Caption         =   "&Retur"
      Height          =   390
      Left            =   7980
      TabIndex        =   2
      Top             =   7110
      Width           =   1890
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2310
      TabIndex        =   33
      Top             =   90
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
      Height          =   5070
      Left            =   405
      TabIndex        =   30
      Top             =   1980
      Width           =   11460
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   38
         Left            =   1545
         TabIndex        =   67
         Top             =   4125
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
         Left            =   1545
         MultiLine       =   -1  'True
         TabIndex        =   66
         Top             =   3750
         Width           =   9825
      End
      Begin VB.ComboBox cboStatus 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         ItemData        =   "frmBatalBilling.frx":0000
         Left            =   5280
         List            =   "frmBatalBilling.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
      Begin VB.Frame Frame4 
         Caption         =   "Biaya"
         Height          =   3150
         Left            =   7065
         TabIndex        =   40
         Top             =   540
         Width           =   4320
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   39
            Left            =   930
            TabIndex        =   41
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
            TabIndex        =   8
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
            Format          =   "#,##0.00##"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   41
            Left            =   2355
            TabIndex        =   42
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
            TabIndex        =   9
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
            Format          =   "#,##0.00##"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   930
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   60
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
            TabIndex        =   12
            Top             =   2310
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
            TabIndex        =   64
            Top             =   2535
            Width           =   180
         End
         Begin VB.Label Label12 
            Caption         =   "+"
            Height          =   210
            Left            =   3990
            TabIndex        =   63
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
            TabIndex        =   62
            Top             =   2340
            Width           =   945
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Total"
            Height          =   255
            Left            =   1230
            TabIndex        =   61
            Top             =   1980
            Width           =   945
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "SubBlet"
            Height          =   255
            Left            =   -165
            TabIndex        =   59
            Top             =   1575
            Width           =   945
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bahan"
            Height          =   255
            Left            =   -165
            TabIndex        =   58
            Top             =   1260
            Width           =   945
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Labour"
            Height          =   255
            Left            =   -165
            TabIndex        =   46
            Top             =   315
            Width           =   945
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Part"
            Height          =   255
            Left            =   -165
            TabIndex        =   45
            Top             =   630
            Width           =   945
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   255
            Left            =   1215
            TabIndex        =   44
            Top             =   2730
            Width           =   945
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Oli"
            Height          =   255
            Left            =   -165
            TabIndex        =   43
            Top             =   945
            Width           =   945
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3150
         Left            =   45
         TabIndex        =   32
         Top             =   540
         Width           =   7005
         Begin MSFlexGridLib.MSFlexGrid DGrid 
            Height          =   2025
            Left            =   75
            TabIndex        =   7
            Top             =   615
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
         TabIndex        =   31
         Top             =   4380
         Width           =   11310
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   37
            Left            =   1950
            TabIndex        =   13
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
            Left            =   10155
            TabIndex        =   14
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
            Left            =   9060
            TabIndex        =   65
            Top             =   270
            Width           =   945
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Suplay Slip/Keterangan"
            Height          =   255
            Left            =   -45
            TabIndex        =   57
            Top             =   285
            Width           =   1875
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   31
         Left            =   6900
         TabIndex        =   5
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
         Left            =   7500
         TabIndex        =   47
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
         Left            =   3375
         TabIndex        =   50
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
         TabIndex        =   51
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
         Left            =   10125
         TabIndex        =   54
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
         Left            =   8505
         TabIndex        =   6
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
         Left            =   450
         TabIndex        =   68
         Top             =   4170
         Width           =   945
      End
      Begin VB.Label Label6 
         Caption         =   "Terbilang (Rupiah)"
         Height          =   270
         Left            =   105
         TabIndex        =   69
         Top             =   3810
         Width           =   1395
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. No."
         Height          =   255
         Left            =   9045
         TabIndex        =   56
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "T SA."
         Height          =   255
         Left            =   7425
         TabIndex        =   55
         Top             =   285
         Width           =   945
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. WO"
         Height          =   255
         Left            =   2295
         TabIndex        =   52
         Top             =   300
         Width           =   945
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "MCH"
         Height          =   255
         Left            =   5805
         TabIndex        =   49
         Top             =   285
         Width           =   945
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         Height          =   255
         Left            =   4200
         TabIndex        =   48
         Top             =   315
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   405
      TabIndex        =   23
      Top             =   420
      Width           =   11460
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   900
         TabIndex        =   15
         Top             =   525
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
         Left            =   900
         TabIndex        =   16
         Top             =   840
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
         Left            =   900
         TabIndex        =   17
         Top             =   1155
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
         Left            =   6240
         TabIndex        =   18
         Top             =   195
         Width           =   1410
         _ExtentX        =   2487
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
         Left            =   6240
         TabIndex        =   19
         Top             =   510
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
         Index           =   25
         Left            =   6240
         TabIndex        =   21
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
         Left            =   10830
         TabIndex        =   20
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
         Left            =   7965
         TabIndex        =   34
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
         Left            =   8055
         TabIndex        =   35
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
         Left            =   9000
         TabIndex        =   36
         Top             =   195
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
         Left            =   6240
         TabIndex        =   38
         Top             =   1140
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
         Index           =   57
         Left            =   900
         TabIndex        =   1
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2430
         TabIndex        =   70
         Top             =   225
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
         Left            =   4350
         TabIndex        =   72
         Top             =   210
         Width           =   765
         _ExtentX        =   1349
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
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   5130
         TabIndex        =   73
         Top             =   225
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
         Caption         =   "Syarat Bayar"
         Height          =   255
         Left            =   3150
         TabIndex        =   74
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Lgn"
         Height          =   255
         Left            =   -300
         TabIndex        =   71
         Top             =   255
         Width           =   1065
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E/G No."
         Height          =   255
         Left            =   5145
         TabIndex        =   39
         Top             =   1185
         Width           =   945
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
         Height          =   255
         Left            =   7920
         TabIndex        =   37
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         Height          =   255
         Left            =   5145
         TabIndex        =   29
         Top             =   870
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   5145
         TabIndex        =   28
         Top             =   540
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         Height          =   255
         Left            =   5145
         TabIndex        =   27
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   255
         Left            =   -195
         TabIndex        =   26
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   255
         Left            =   -195
         TabIndex        =   25
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   -195
         TabIndex        =   24
         Top             =   570
         Width           =   945
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   75
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
      FormatString    =   "Retur Invoice"
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
   Begin VB.Label No 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Inv"
      Height          =   255
      Left            =   195
      TabIndex        =   22
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmBatalBilling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Baris As Integer
Dim RateCPI As Double
Dim NoInvoice As String
Dim vNoRetur As String
Dim vKodeRetur As String
Dim vJamRetur As String
Dim vTglRetur As String
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
            Case "PPN", "DISC"
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
            Case "PPN", "DISC"
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
            Case 9, 10, 11, 15, 18, 20, 22, 23, 27, 29, 30, 32, 33, 34, 35, 36, 43, 44
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
    Text(57) = ""
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
                .TextMatrix(i, 3) = Format(.TextMatrix(i + 1, 3), "#,##0.00")
                .TextMatrix(i, 4) = Format(.TextMatrix(i + 1, 4), "#,##0.00")
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
                .TextMatrix(Baris, 3) = Koma2Titik(Text(34))
                .TextMatrix(Baris, 4) = Koma2Titik(Text(34)) * RateCPI
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

Private Sub TblSimpan_Click()
    If Trim(Text(57).Text) = "" Then
        MsgBox "Masukkan kode pelanggan !"
    Else
        If MsgBox("Yakin Retur Invoice ?", vbYesNo, "Konfirmasi") = vbYes Then
            gNoInvoice = Text(0)
            frmBtlinvoice.Show (vbModal)
        End If
    End If
End Sub

Private Sub CetakFormKecil()
    Dim rpt As New RptInvRetur
    
    rpt.Label48 = "No Urut Retur : " & vNoRetur
    rpt.Label39 = gSecmanager
    rpt.Label40 = "KEPALA BENGKEL"
    rpt.Label2 = "NPWP : " & gNpwp
    rpt.Label3 = "PKP  : " & gPkp
    rpt.Label44 = "TELP : " & gTelp
    rpt.Label18 = "BIAYA JASA  : " & AddSpace(Format(Text(39), "#,##0"), 13, 3, " ")
    rpt.Label19 = "SUBLET      : " & AddSpace(Format(Text(5), "#,##0"), 13, 3, " ")
    rpt.Label20 = "PARTS       : " & AddSpace(Format(Text(40), "#,##0"), 13, 3, " ")
    rpt.Label21 = "OIL/GRERASE : " & AddSpace(Format(Text(45), "#,##0"), 13, 3, " ")
    rpt.Label22 = "MATERIAL    : " & AddSpace(Format(Text(4), "#,##0"), 13, 3, " ")
    rpt.Label23 = "SUBTOTAL    : " & AddSpace(Format(Text(6), "#,##0"), 13, 3, " ")
    rpt.Label24 = "METERAI     : " & AddSpace(Format(Text(7), "#,##0"), 13, 3, " ")
    rpt.Label25 = "GRAND TOTAL : " & AddSpace(Format(Text(41), "#,##0"), 13, 3, " ")
    
    rpt.Label26 = AddSpace("URAIAN BIAYA JASA", 25, 1, " ") + _
                AddSpace("MKN", 4, 3, " ") + _
                AddSpace("RUPIAH", 12, 3, " ")
    rpt.Label27 = AddSpace(DGrid.TextMatrix(1, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(1, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(1, 4), "#,##0"), 12, 3, " ")
    rpt.Label28 = AddSpace(DGrid.TextMatrix(2, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(2, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(2, 4), "#,##0"), 12, 3, " ")
    rpt.Label29 = AddSpace(DGrid.TextMatrix(3, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(3, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(3, 4), "#,##0"), 12, 3, " ")
    rpt.Label30 = AddSpace(DGrid.TextMatrix(4, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(4, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(4, 4), "#,##0"), 12, 3, " ")
    rpt.Label31 = AddSpace(DGrid.TextMatrix(5, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(5, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(5, 4), "#,##0"), 12, 3, " ")
    rpt.Label32 = AddSpace(DGrid.TextMatrix(6, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(6, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(6, 4), "#,##0"), 12, 3, " ")
    rpt.Label33 = AddSpace(DGrid.TextMatrix(7, 1), 25, 1, " ") + _
                AddSpace(DGrid.TextMatrix(7, 5), 4, 3, " ") + _
                AddSpace(Format(DGrid.TextMatrix(7, 4), "#,##0"), 12, 3, " ")
    
    rpt.Label8 = "NO. WO. : " & Text(26)
    rpt.Label9 = "FRAME : " & Text(25)
    rpt.Label10 = "E/G : " & Text(21)
    rpt.Label11 = "NAMA : " & Text(1)
    rpt.Label12 = "ALAMAT : " & Text(2)
    rpt.Label13 = "KOTA : " & Text(3)
    rpt.Label14 = Text(19) & "(" & Text(24) & ")"
    rpt.Label15 = "SA : " & Text(13)
    rpt.Label16 = "OPT : " & gSA
    
    rpt.Label47 = "Kode Plgn  : " & Text(57)
    rpt.Label4 = "No. Retur  : " & vKodeRetur & "/" & cboStatus.Text
    rpt.Label5 = "Tanggal    : " & vTglRetur
    rpt.Label6 = "Jam        : " & vJamRetur
    rpt.Label7 = "No. Polisi : " & Text(16)
    
    rpt.Label36 = Text1.Text
    rpt.Label37 = Text(37)
    rpt.Label38 = Text(38)
    
    rpt.Label45 = "Bank     : " & gBank
    rpt.Label46 = "Acc. No. : " & gAcc

    rpt.Printer.PaperWidth = 1440 * 8.5
    rpt.Printer.PaperHeight = 1440 * 5.5
    rpt.Printer.Orientation = 1
        
    If gPreviewCetak = True Then
        rpt.PrintReport False
        Unload rpt
    Else
        rpt.Show vbModal
    End If
End Sub

Private Sub Cetak()
    Dim rpt As New RptFromRetur
    
    gKodePlg = Text(57)
    gInvNo = Text(42) & "/" & cboStatus.Text
    gInvTgl = vTglRetur
    gInvCust = Text(1)
    gInvAlmt = Text(2)
    gInvKota = Text(3) & "-" & vNoTelp
    gInvNPWP = Text(14)
    gInvNoPolisi = Text(16)
    gInvModel = Text(19) & " (" & Text(24) & ")"
    gInvNoMesin = Text(21)
    gInvNoChas = Text(25)
    gInvNoOrder = Text(26)
    gInvSA = Text(13)
    gInvStatus = cboStatus.Text
    gInvMech = Text(31)
    gInvPrint = vJamRetur
    gInvBilling = gSA
    gInvUraian1 = DGrid.TextMatrix(1, 1)
    gInvUraian2 = DGrid.TextMatrix(2, 1)
    gInvUraian3 = DGrid.TextMatrix(3, 1)
    gInvUraian4 = DGrid.TextMatrix(4, 1)
    gInvUraian5 = DGrid.TextMatrix(5, 1)
    gInvUraian6 = DGrid.TextMatrix(6, 1)
    gInvUraian7 = DGrid.TextMatrix(7, 1)
    gInvRate1 = ""
    gInvRate2 = ""
    gInvRate3 = ""
    gInvRate4 = ""
    gInvRate5 = ""
    gInvRate6 = ""
    gInvRate7 = ""
    gInvRp1 = Format(DGrid.TextMatrix(1, 4), "#,##0")
    gInvRp2 = Format(DGrid.TextMatrix(2, 4), "#,##0")
    gInvRp3 = Format(DGrid.TextMatrix(3, 4), "#,##0")
    gInvRp4 = Format(DGrid.TextMatrix(4, 4), "#,##0")
    gInvRp5 = Format(DGrid.TextMatrix(5, 4), "#,##0")
    gInvRp6 = Format(DGrid.TextMatrix(6, 4), "#,##0")
    gInvRp7 = Format(DGrid.TextMatrix(7, 4), "#,##0")
    gGInvRp = Format(Text(39), "#,##0")
    gInvTRp1 = Format(Text(39), "#,##0")
    gInvTRp2 = Format(Text(5), "#,##0")
    gInvTRp3 = Format(Text(40), "#,##0")
    gInvTRp4 = Format(Text(45), "#,##0")
    gInvTRp5 = Format(Text(4), "#,##0")
    gInvTRp6 = Format(Text(6), "#,##0")
    gInvTRp7 = Format(Text(7), "#,##0")
    gGInvTRp = Format(Text(41), "#,##0")
    gInvTerbilang = Text1.Text
    gInvSS = Text(37)
    gInvCatatan = Text(38)
    
    gInvSM = gSecmanager
    rpt.Field67.DataValue = vNoRetur
    rpt.Field66.DataValue = vKodeRetur & "/" & cboStatus.Text
    rpt.Field65.DataValue = gKodePlg
    rpt.Field39.DataValue = gInvNo
    rpt.Field40.DataValue = vTglRetur

    rpt.Field1.DataValue = gInvCust
    rpt.Field4.DataValue = gInvAlmt
    rpt.Field28.DataValue = gInvKota
    rpt.Field7.DataValue = gInvNPWP

    rpt.Field2.DataValue = gInvNoPolisi
    rpt.Field5.DataValue = gInvModel
    rpt.Field12.DataValue = gInvNoMesin
    rpt.Field14.DataValue = gInvNoChas
    
    rpt.Field29.DataValue = gInvNoOrder
    rpt.Field30.DataValue = gInvSA

    rpt.Field31.DataValue = gInvStatus
    rpt.Field32.DataValue = gInvMech

    rpt.Field33.DataValue = gInvPrint
    rpt.Field38.DataValue = gInvBilling
    
    rpt.Field16.DataValue = gInvUraian1
    rpt.Field17.DataValue = gInvUraian2
    rpt.Field18.DataValue = gInvUraian3
    rpt.Field19.DataValue = gInvUraian4
    rpt.Field20.DataValue = gInvUraian5
    rpt.Field21.DataValue = gInvUraian6
    rpt.Field22.DataValue = gInvUraian7
    
    rpt.Field41.DataValue = gInvRate1
    rpt.Field42.DataValue = gInvRate2
    rpt.Field43.DataValue = gInvRate3
    rpt.Field44.DataValue = gInvRate4
    rpt.Field45.DataValue = gInvRate5
    rpt.Field46.DataValue = gInvRate6
    rpt.Field47.DataValue = gInvRate7
    
    rpt.Field48.DataValue = gInvRp1
    rpt.Field49.DataValue = gInvRp2
    rpt.Field50.DataValue = gInvRp3
    rpt.Field51.DataValue = gInvRp4
    rpt.Field52.DataValue = gInvRp5
    rpt.Field53.DataValue = gInvRp6
    rpt.Field54.DataValue = gInvRp7
    
    rpt.Field55.DataValue = gInvTRp1
    rpt.Field56.DataValue = gInvTRp2
    rpt.Field57.DataValue = gInvTRp3
    rpt.Field58.DataValue = gInvTRp4
    rpt.Field59.DataValue = gInvTRp5
    rpt.Field60.DataValue = gInvTRp6
    rpt.Field61.DataValue = gInvTRp7
    
    rpt.Field23.DataValue = gGInvRp
    rpt.Field27.DataValue = gGInvTRp
    rpt.Field25.DataValue = gInvTerbilang
    
    rpt.Field62.DataValue = gInvSS
    rpt.Field63.DataValue = gInvCatatan
    rpt.Field64.DataValue = gInvSM
    '9=A4
    'kwarto

'    rpt.Printer.PaperSize = 9
    rpt.Printer.PaperSize = 256
    rpt.Printer.PaperHeight = 1440 * 11
    rpt.Printer.PaperWidth = 1440 * 8.5
    
    rpt.Printer.Orientation = 1
    
    If gPreviewCetak = True Then
        rpt.PrintReport False
        Unload rpt
    Else
        rpt.Show vbModal
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

Private Sub PanggilDataOrder(NomorInv As String)
    sSQL = "Select *, datepart(hh,FD_RETUR) as JAM, datepart(mi,FD_RETUR) as MENIT from T_INVOICE where NOINV ='" & NomorInv & "'"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        'Data Customer & Order
        vKodeRetur = oRs1("FS_NO_RETUR")
        vNoRetur = oRs1("FS_KD_RETUR")
        vJamRetur = Format(oRs1("JAM"), "00") & ":" & Format(oRs1("MENIT"), "00")
        vTglRetur = TglDDMMYYYY(oRs1("FD_RETUR"))
        
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
        Text(40) = 0
        Text(41) = oRs1("TOTAL")
        Text(42) = oRs1("NOINV")
        Text(45) = 0
        cboStatus.Text = oRs1("STATUS")
        Text(57) = oRs1("FS_KD_PLG")
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

'        If oRs("RPR_COND") = "K" Then
'            MsgBox "Kendaraan telah keluar, batalkan registrasi dahulu !", vbOKOnly, "Perhatian"
'            TblSimpan.Enabled = False
'            TblKeluar.SetFocus
'        Else
'            TblSimpan.Enabled = True
'        End If
    End If
    oRs1.Close
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "INVBATAL"
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
            Case 57
                If gConnecAAS = True Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "KODEPLG"
                    frmCariAcess.Show (vbModal)
                End If
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
    Dim NoInv As String
    
    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                TblSimpan.Enabled = True
                NoInv = Text(0)
                sSQL = "Select * from T_INVOICE where NOINV ='" & NoInv & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "INVBATAL"
                    frmCari.Show (vbModal)
                Else
                    Call PanggilDataOrder(NoInv)
                    If Trim(Text(57).Text) = "" And gConnecAAS Then
                        Text(57).BackColor = &HFFFFFF
                    Else
                        Text(57).BackColor = &HC0E0FF
                    End If
                    gMDL = Text(17)
                    Text(39) = TotalJasa
                    Call HitungTotal
                    If Trim(oRs("FS_NO_RETUR")) <> "" Then
                        If MsgBox("Mencetak Invoice Retur ?", vbYesNo, "Cetak Ulang") = vbYes Then
                            If gUkuranRetur Then
                                Call Cetak
                            Else
                                Call CetakFormKecil
                            End If
                        End If
                        TblSimpan.Enabled = False
                    End If
                End If
                oRs.Close
    
                If gConnecAAS Then
                    'Panggil data pelanggan Syarat Bayar AAS
                    'LOGONSTR2 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=auser;Data Source=" & LetakAcces & "ardata" & Trim(gGenerasi) & ".MDB" & ";Persist Security Info=False;Jet OLEDB:System database=" & SysAcces
                    LOGONSTR2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
                
                    oConnAcces.ConnectionString = LOGONSTR2
                    oConnAcces.CommandTimeout = 120
                    oConnAcces.ConnectionTimeout = 120
                    oConnAcces.Open
                    
                    sSQL = "select KodeLgn, SyaratBayar from Armst where KodeLgn = '" & Trim(Text(57)) & "'"
                    oRs1.Open sSQL, oConnAcces, adOpenStatic
                    If Not oRs1.BOF Then
                        Text(12) = oRs1("SyaratBayar")
                    End If
                    oRs1.Close
                    oConnAcces.Close
                End If
            Else
                Call HapusForm
            End If
        Case 9
            If CekTgl(Text(Index)) = True Then
                Text(Index) = TglDDMMYYYY(Text(Index))
            Else
                Text(Index) = TglDDMMYYYY(Date)
            End If
'            gTglInv = Text(Index)
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
        Case 57
            If Trim(Text(Index)) <> "" And gConnecAAS Then
                'LOGONSTR2 = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=auser;Data Source=" & LetakAcces & "ardata" & Trim(gGenerasi) & ".MDB" & ";Persist Security Info=False;Jet OLEDB:System database=" & SysAcces
                LOGONSTR2 = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & NisUSER & ";Password='" & NisPass & "';Initial Catalog= " & NisDb & ";Data Source= " & NisKomp & ""
                
                oConnAcces.ConnectionString = LOGONSTR2
                oConnAcces.CommandTimeout = 120
                oConnAcces.ConnectionTimeout = 120
                oConnAcces.Open

                sSQL = "select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp, SyaratBayar from Armst where KodeLgn = '" & Trim(Text(Index)) & "'"
                oRs.Open sSQL, oConnAcces, adOpenStatic
                If oRs.BOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "KODEPLG"
                    frmCariAcess.Show (vbModal)
                Else
                    If IsNull(oRs("NamaLgn")) Then
                        Text(1) = ""
                    Else
                        Text(1) = oRs("NamaLgn")
                    End If
                    If IsNull(oRs("Alamat1")) Then
                        Text(2) = ""
                    Else
                        Text(2) = oRs("Alamat1")
                    End If
                    If IsNull(oRs("Alamat2")) Then
                        Text(3) = ""
                    Else
                        Text(3) = oRs("Alamat2")
                    End If
                    If IsNull(oRs("Npwp")) Then
                        Text(14) = ""
                    Else
                        Text(14) = oRs("Npwp")
                    End If
                    Text(12) = oRs("SyaratBayar")
                End If
                oRs.Close
                oConnAcces.Close
            End If
    End Select
End Sub

Private Sub HitungTotal()
    Dim Nilai As Double
    
    Text(41) = Koma2Titik(Text(39)) + Koma2Titik(Text(40)) _
             + Koma2Titik(Text(45)) + Koma2Titik(Text(4)) _
             + Koma2Titik(Text(5)) + Koma2Titik(Text(7))
    Text(6) = Koma2Titik(Text(39)) + Koma2Titik(Text(40)) _
             + Koma2Titik(Text(45)) + Koma2Titik(Text(4)) _
             + Koma2Titik(Text(5))
    
    Nilai = Val(SimpanKoma2Titik(Text(41)))
    
    Text1.Text = Angka2Huruf(Nilai, 1)
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
