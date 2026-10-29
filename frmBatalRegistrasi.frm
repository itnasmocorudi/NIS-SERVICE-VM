VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmBatalRegistrasi 
   Caption         =   "Pembatalan Registrasi & Close Order"
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2460
      TabIndex        =   53
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
      TabIndex        =   35
      Top             =   2895
      Width           =   11430
      Begin VB.CommandButton TblSimpan 
         Caption         =   "&Batal"
         Height          =   390
         Left            =   7560
         TabIndex        =   8
         Top             =   4200
         Width           =   1890
      End
      Begin VB.CommandButton TblKeluar 
         Caption         =   "&Kembali ke Menu"
         Height          =   390
         Left            =   9450
         TabIndex        =   9
         Top             =   4200
         Width           =   1890
      End
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         Height          =   3930
         Left            =   4560
         TabIndex        =   50
         Top             =   210
         Width           =   6780
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   37
            Left            =   1590
            TabIndex        =   6
            Top             =   2970
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
            Left            =   1590
            TabIndex        =   7
            Top             =   3285
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
            Left            =   390
            TabIndex        =   5
            Top             =   645
            Width           =   6045
            _ExtentX        =   10663
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
            Left            =   -15
            TabIndex        =   52
            Top             =   3030
            Width           =   1485
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saran"
            Height          =   255
            Left            =   525
            TabIndex        =   51
            Top             =   3330
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   3930
         Left            =   105
         TabIndex        =   36
         Top             =   210
         Width           =   4395
         Begin VB.Frame Frame4 
            Caption         =   "Estimasi Labor, Parts & Material"
            Height          =   1650
            Left            =   90
            TabIndex        =   37
            Top             =   2175
            Width           =   4200
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   39
               Left            =   1215
               TabIndex        =   38
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
               TabIndex        =   22
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
               TabIndex        =   39
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
               TabIndex        =   59
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
               TabIndex        =   60
               Top             =   975
               Width           =   945
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Total"
               Height          =   255
               Left            =   120
               TabIndex        =   42
               Top             =   1290
               Width           =   945
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Part"
               Height          =   255
               Left            =   120
               TabIndex        =   41
               Top             =   660
               Width           =   945
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Labour"
               Height          =   255
               Left            =   120
               TabIndex        =   40
               Top             =   345
               Width           =   945
            End
         End
         Begin VB.ComboBox cboStatus 
            BackColor       =   &H00C0E0FF&
            Height          =   315
            ItemData        =   "frmBatalRegistrasi.frx":0000
            Left            =   1260
            List            =   "frmBatalRegistrasi.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   210
            Width           =   735
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   42
            Left            =   1245
            TabIndex        =   43
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
            TabIndex        =   44
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
            TabIndex        =   45
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
            TabIndex        =   4
            Top             =   1185
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
            Left            =   1260
            TabIndex        =   3
            Top             =   870
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   1260
            TabIndex        =   2
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
            TabIndex        =   21
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
            TabIndex        =   58
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
            TabIndex        =   55
            Top             =   1545
            Width           =   945
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "GRP"
            Height          =   255
            Left            =   165
            TabIndex        =   54
            Top             =   600
            Width           =   945
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inv. No."
            Height          =   255
            Left            =   165
            TabIndex        =   49
            Top             =   1845
            Width           =   945
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Status"
            Height          =   255
            Left            =   180
            TabIndex        =   48
            Top             =   285
            Width           =   945
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "K. Meter"
            Height          =   255
            Left            =   165
            TabIndex        =   47
            Top             =   1245
            Width           =   945
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "MCH"
            Height          =   255
            Left            =   165
            TabIndex        =   46
            Top             =   915
            Width           =   945
         End
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2415
      Left            =   420
      TabIndex        =   24
      Top             =   480
      Width           =   11430
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   10
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
         Left            =   1440
         TabIndex        =   11
         Top             =   630
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
         Left            =   1440
         TabIndex        =   12
         Top             =   945
         Width           =   3285
         _ExtentX        =   5794
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
         Left            =   1440
         TabIndex        =   13
         Top             =   1260
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
         Index           =   10
         Left            =   1440
         TabIndex        =   14
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
         Left            =   1440
         TabIndex        =   15
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
         Left            =   7440
         TabIndex        =   16
         Top             =   285
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
         Left            =   7440
         TabIndex        =   17
         Top             =   615
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
         Index           =   18
         Left            =   7440
         TabIndex        =   19
         Top             =   1245
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
         Left            =   7440
         TabIndex        =   20
         Top             =   1560
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
         Left            =   9315
         TabIndex        =   18
         Top             =   615
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
         Left            =   9195
         TabIndex        =   56
         Top             =   630
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
         Left            =   7440
         TabIndex        =   57
         Top             =   930
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
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         Height          =   255
         Left            =   6345
         TabIndex        =   34
         Top             =   1605
         Width           =   945
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         Height          =   255
         Left            =   6345
         TabIndex        =   33
         Top             =   1290
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   6345
         TabIndex        =   32
         Top             =   645
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         Height          =   255
         Left            =   6345
         TabIndex        =   31
         Top             =   330
         Width           =   945
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disc."
         Height          =   255
         Left            =   900
         TabIndex        =   30
         Top             =   1920
         Width           =   390
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         Height          =   255
         Left            =   915
         TabIndex        =   29
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telp. No."
         Height          =   255
         Left            =   345
         TabIndex        =   28
         Top             =   1305
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   255
         Left            =   345
         TabIndex        =   27
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   255
         Left            =   345
         TabIndex        =   26
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   345
         TabIndex        =   25
         Top             =   360
         Width           =   945
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1455
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
      FormatString    =   "Pembatalan Registrasi & Close Order"
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
      Left            =   345
      TabIndex        =   23
      Top             =   165
      Width           =   945
   End
End
Attribute VB_Name = "frmBatalRegistrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Baris As Integer
Dim RateCPI As Double
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
            Case 5, 6, 7, 8, 9, 14, 15, 21, 22, 23, 20, 26, 27, 28, 29, 32, 33, 34, 35, 36
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

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call DGrid_DblClick
    End If
End Sub

Private Sub Form_Load()
    Baris = 1
    DGrid.Rows = 11
    DGrid.Cols = 5
    DGrid.ColWidth(0) = 0
    DGrid.ColWidth(1) = 2670
    DGrid.ColWidth(2) = 800
    DGrid.ColWidth(3) = 1050
    DGrid.ColWidth(4) = 1445
    DGrid.ColAlignment(1) = flexAlignLeftCenter
    DGrid.ColAlignment(2) = flexAlignLeftCenter
    DGrid.ColAlignment(3) = flexAlignLeftCenter
    DGrid.ColAlignment(4) = flexAlignRightCenter
    DGrid.TextMatrix(0, 1) = "Keterangan"
    DGrid.TextMatrix(0, 2) = "JP"
    DGrid.TextMatrix(0, 3) = "Rate"
    DGrid.TextMatrix(0, 4) = "Kondisi"
    Call HapusForm
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub SaveOrderPekerjaan(NomorWO As String, NomorPolisi As String)
    sSQL = "Select * from JOBDTA where NO_WO ='" & NomorWO & "' and NO_POLISI ='" & NomorPolisi & "' and DELIV<>'BTL' AND RPR_COND = 'K'"
    oRs.Open sSQL, oConn, adOpenStatic
    Call ClearsSql
    gsSql1(0) = "RPR_COND"
    gsSql2(0) = "S"
    If Not oRs.BOF Then
        sSQL = GensSql("JOBDTA", True) + " where NO_WO ='" & NomorWO & "' and NO_POLISI ='" & NomorPolisi & "' and DELIV<>'BTL' AND RPR_COND = 'K'"
'    Else
'        sSQL = GensSql("JOBDTA", False)
    End If
    oConn.Execute sSQL
    oRs.Close

    sSQL = "Select * from T_INVOICE where NO_WO ='" & NomorWO & "' and NO_POLISI ='" & NomorPolisi & "' and DELIV<>'BTL' AND RPR_COND = 'K'"
    oRs.Open sSQL, oConn, adOpenStatic
    Call ClearsSql
    gsSql1(0) = "RPR_COND"
    gsSql2(0) = "S"
    If Not oRs.BOF Then
        sSQL = GensSql("T_INVOICE", True) + " where NO_WO ='" & NomorWO & "' and NO_POLISI ='" & NomorPolisi & "' and DELIV<>'BTL' AND RPR_COND = 'K'"
'    Else
'        sSQL = GensSql("T_INVOICE", False)
    End If
    oConn.Execute sSQL
    oRs.Close

End Sub

Private Sub TblSimpan_Click()
    Call SaveOrderPekerjaan(Text(0), Text(16))
    Call HapusForm
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

Private Sub PanggilDataOrder(NomorWO As String)
    sSQL = "Select * from JOBDTA where NO_WO ='" & NomorWO & "' and DELIV<>'BTL' AND RPR_COND = 'K'"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        'Data Customer & Order
        Text(1) = oRs1("CUSTOMER")
        Text(2) = oRs1("CUST_ADRES")
        Text(3) = oRs1("KOTA")
        Text(4) = oRs1("TELPON")
        Text(16) = oRs1("NO_POLISI")
        Text(17) = oRs1("MDL")
        Text(24) = oRs1("MODEL")
        Text(18) = oRs1("COLOR")
        Text(13) = oRs1("FRONT")
        Text(37) = oRs1("REMARK1")
        Text(10) = oRs1("PPN")
        Text(11) = oRs1("DISC")
        Text(12) = oRs1("GROUP")
        Text(19) = oRs1("FMDL")
        Text(45) = oRs1("MTRL")
        
        Text(42) = oRs1("NOINV")
        Text(31) = oRs1("MECH")
        Text(30) = oRs1("KM")
        Text(38) = oRs1("SUGGEST")
        Text(39) = oRs1("LABOR")
        Text(40) = oRs1("PART")
        Text(41) = oRs1("TOTAL")
        Text(44) = oRs1("PAID")
        Text(43) = oRs1("DELIV")
        cboStatus.Text = oRs1("STATUS")
'        Text(32) = oRs1("JOB")
'        Text(5) = TglDDMMYYYY(oRs1("DATE"))
'        Text(6) = oRs1("TIME")
        
'        If oRs1("TIME_M") <> "" Then
'            Text(7) = oRs1("TIME_M")
'        End If
'        If oRs1("TIME_S") <> "" Then
'            Text(8) = oRs1("TIME_S")
'        End If
'        If oRs1("TIME_A") <> "" Then
'            Text(9) = oRs1("TIME_A")
'        End If
        
        
''Format(Time, "HH:NN")

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

    End If
    oRs1.Close
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "REGISTBATAL"
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
    Dim NoWO As String
    
    Select Case Index
        Case 0
            If Trim(Text(Index)) <> "" Then
                NoWO = Text(0)
                sSQL = "Select * from JOBDTA where NO_WO ='" & NoWO & "' and DELIV<>'BTL' AND RPR_COND = 'K'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "REGISTBATAL"
                    frmCari.Show (vbModal)
                Else
                    Call PanggilDataOrder(NoWO)
                End If
                oRs.Close
                gMDL = Text(17)
                Call HitungNilaiKaliRate
                Call HitungTotal
            Else
                Call HapusForm
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
            RateCPI = 0
    End Select
End Sub

