VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmBatalBooking 
   Caption         =   "Pembatalan Booking"
   ClientHeight    =   7755
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2865
      TabIndex        =   53
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
      Caption         =   "Input Booking"
      Height          =   4830
      Left            =   405
      TabIndex        =   49
      Top             =   2640
      Width           =   11415
      Begin VB.CommandButton TblBatal 
         Caption         =   "&Batal"
         Height          =   390
         Left            =   7560
         TabIndex        =   29
         Top             =   4350
         Width           =   1890
      End
      Begin VB.CommandButton TblKeluar 
         Caption         =   "&Kembali ke Menu"
         Height          =   390
         Left            =   9450
         TabIndex        =   28
         Top             =   4350
         Width           =   1890
      End
      Begin VB.Frame Frame5 
         Enabled         =   0   'False
         Height          =   4080
         Left            =   6090
         TabIndex        =   52
         Top             =   195
         Width           =   5250
         Begin MSFlexGridLib.MSFlexGrid DGrid 
            Height          =   1545
            Left            =   240
            TabIndex        =   25
            Top             =   975
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   2725
            _Version        =   393216
            FixedCols       =   0
            ScrollBars      =   0
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   37
            Left            =   1455
            TabIndex        =   26
            Top             =   2700
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
            Left            =   1455
            TabIndex        =   27
            Top             =   3015
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
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Pesan Khusus"
            Height          =   255
            Left            =   195
            TabIndex        =   67
            Top             =   2760
            Width           =   1125
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Saran"
            Height          =   255
            Left            =   360
            TabIndex        =   66
            Top             =   3060
            Width           =   945
         End
      End
      Begin VB.Frame Frame3 
         Enabled         =   0   'False
         Height          =   4065
         Left            =   60
         TabIndex        =   50
         Top             =   195
         Width           =   6000
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   28
            Left            =   1650
            TabIndex        =   16
            Top             =   195
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
            Left            =   1650
            TabIndex        =   17
            Top             =   510
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
            Left            =   1650
            TabIndex        =   18
            Top             =   825
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
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
            Index           =   9
            Left            =   1650
            TabIndex        =   56
            Top             =   1140
            Width           =   360
            _ExtentX        =   635
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
            Index           =   14
            Left            =   1650
            TabIndex        =   19
            Top             =   1455
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
            Index           =   15
            Left            =   1650
            TabIndex        =   59
            Top             =   2970
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
            Index           =   22
            Left            =   1650
            TabIndex        =   62
            Top             =   3285
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
            Index           =   23
            Left            =   1650
            TabIndex        =   64
            Top             =   3600
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
            Height          =   240
            Index           =   20
            Left            =   1650
            TabIndex        =   20
            Top             =   1785
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
            Index           =   26
            Left            =   1650
            TabIndex        =   21
            Top             =   2010
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
            Left            =   1650
            TabIndex        =   22
            Top             =   2235
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
            Index           =   29
            Left            =   1650
            TabIndex        =   23
            Top             =   2460
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
            Index           =   30
            Left            =   1650
            TabIndex        =   24
            Top             =   2685
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
         Begin VB.Label Label20 
            Caption         =   "(HH:MM)"
            Height          =   240
            Left            =   2415
            TabIndex        =   71
            Top             =   570
            Width           =   1140
         End
         Begin VB.Label Label30 
            Caption         =   "(DD-MM-YYYY)"
            Height          =   240
            Left            =   2940
            TabIndex        =   70
            Top             =   255
            Width           =   1140
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Keluhan && Permintaan Customer"
            Height          =   630
            Left            =   555
            TabIndex        =   68
            Top             =   1815
            Width           =   945
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inv. No."
            Height          =   255
            Left            =   -210
            TabIndex        =   65
            Top             =   3645
            Width           =   1710
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
            Height          =   255
            Left            =   -210
            TabIndex        =   63
            Top             =   3330
            Width           =   1710
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "D [ Datang ] / T [Telepon]"
            Height          =   255
            Left            =   2100
            TabIndex        =   61
            Top             =   870
            Width           =   2850
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Labour"
            Height          =   255
            Left            =   -210
            TabIndex        =   60
            Top             =   3000
            Width           =   1710
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "T. SA"
            Height          =   255
            Left            =   -210
            TabIndex        =   58
            Top             =   1500
            Width           =   1710
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tech"
            Height          =   255
            Left            =   -210
            TabIndex        =   57
            Top             =   1185
            Width           =   1710
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Via"
            Height          =   255
            Left            =   -210
            TabIndex        =   55
            Top             =   870
            Width           =   1710
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Booking"
            Height          =   255
            Left            =   255
            TabIndex        =   54
            Top             =   555
            Width           =   1245
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Booking"
            Height          =   255
            Left            =   255
            TabIndex        =   51
            Top             =   240
            Width           =   1245
         End
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   90
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
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
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2190
      Left            =   405
      TabIndex        =   33
      Top             =   405
      Width           =   11415
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1080
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
         Left            =   1080
         TabIndex        =   2
         Top             =   525
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
         Left            =   1080
         TabIndex        =   3
         Top             =   840
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
         Left            =   1080
         TabIndex        =   4
         Top             =   1155
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
         Index           =   5
         Left            =   1080
         TabIndex        =   5
         Top             =   1470
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
         Left            =   1080
         TabIndex        =   6
         Top             =   1785
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
         Index           =   10
         Left            =   6330
         TabIndex        =   7
         Top             =   210
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
         Left            =   6330
         TabIndex        =   8
         Top             =   525
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
         Left            =   6330
         TabIndex        =   9
         Top             =   840
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
         Left            =   6330
         TabIndex        =   10
         Top             =   1155
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
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   8550
         TabIndex        =   11
         Top             =   195
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
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   8550
         TabIndex        =   12
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
         Index           =   18
         Left            =   8550
         TabIndex        =   13
         Top             =   1140
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
         Index           =   19
         Left            =   8550
         TabIndex        =   31
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
         Index           =   21
         Left            =   8550
         TabIndex        =   15
         Top             =   1770
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
         Index           =   24
         Left            =   10410
         TabIndex        =   30
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   25
         Left            =   8550
         TabIndex        =   14
         Top             =   1455
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   10290
         TabIndex        =   69
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
         Index           =   53
         Left            =   3045
         TabIndex        =   72
         Top             =   1785
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   2925
         TabIndex        =   73
         Top             =   1785
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "E/G No."
         Height          =   255
         Left            =   7455
         TabIndex        =   48
         Top             =   1815
         Width           =   945
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
         Height          =   255
         Left            =   7455
         TabIndex        =   47
         Top             =   1500
         Width           =   945
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         Height          =   255
         Left            =   7455
         TabIndex        =   46
         Top             =   1185
         Width           =   945
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         Height          =   255
         Left            =   7455
         TabIndex        =   45
         Top             =   555
         Width           =   945
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Th. Mobil"
         Height          =   255
         Left            =   7455
         TabIndex        =   44
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DDTE."
         Height          =   255
         Left            =   5610
         TabIndex        =   43
         Top             =   1200
         Width           =   555
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         Height          =   255
         Left            =   5745
         TabIndex        =   42
         Top             =   885
         Width           =   420
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disc."
         Height          =   255
         Left            =   5790
         TabIndex        =   41
         Top             =   555
         Width           =   390
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         Height          =   255
         Left            =   5805
         TabIndex        =   40
         Top             =   255
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
         Height          =   255
         Left            =   -15
         TabIndex        =   39
         Top             =   1830
         Width           =   945
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
         Height          =   255
         Left            =   -15
         TabIndex        =   38
         Top             =   1515
         Width           =   945
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Telp. No."
         Height          =   255
         Left            =   -15
         TabIndex        =   37
         Top             =   1200
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   255
         Left            =   -15
         TabIndex        =   36
         Top             =   885
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   255
         Left            =   -15
         TabIndex        =   35
         Top             =   570
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         Height          =   255
         Left            =   -15
         TabIndex        =   34
         Top             =   255
         Width           =   945
      End
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
      FormatString    =   "Pembatalan Booking"
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
      Caption         =   "No. Polisi"
      Height          =   255
      Left            =   315
      TabIndex        =   32
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmBatalBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Baris As Integer
Dim Enter As Integer
Const MaxBrs = 5

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
    For i = 1 To 38
        Select Case i
            Case 14
                Text(i) = gSA
            Case 10
                Text(i) = "0,1"
            Case 31, 32, 33, 34, 35, 36
            Case 13, 28
                Text(i) = ">>->>->>>>"
            Case 8
                Text(i) = ">>:>>"
            Case Else
                Text(i) = ""
        End Select
    Next
    Text(53) = ""
    DGrid.TextMatrix(1, 1) = ""
    DGrid.TextMatrix(2, 1) = ""
    DGrid.TextMatrix(3, 1) = ""
    DGrid.TextMatrix(4, 1) = ""
    DGrid.TextMatrix(5, 1) = ""
    

'    sSQL = "Select count(*) as Jml_Data from BOOK where DELIV='BOK'"
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
End Sub

Private Sub DGrid_DblClick()
    If Baris > 1 And DGrid.TextMatrix(DGrid.Row, 1) <> "" Then
        A = DGrid.Row
        'AMBIL
        With DGrid
            Text(33) = .TextMatrix(A, 1)
        End With
        'HAPUS
        For i = A To Baris
            With DGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
            End With
        Next
        Baris = Baris - 1
        Text(33).SetFocus
    End If
End Sub

Private Sub dGrid_GotFocus()
    If Trim(Text(33)) <> "" Then
        If MaxBrs >= Baris Then
            With DGrid
                .TextMatrix(Baris, 1) = Text(33)
            End With
            Baris = Baris + 1
            Text(33) = ""
            Text(33).SetFocus
        Else
            MsgBox "Maksimal " & Str(MaxBrs) & " !"
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
'    Call CenterForm(Me, True)
    Baris = 1
    DGrid.Rows = 11
    DGrid.Cols = 3
    DGrid.ColWidth(0) = 0
    DGrid.ColWidth(1) = 3500
    DGrid.ColWidth(2) = 1200
    DGrid.ColAlignment(1) = flexAlignLeftCenter
    DGrid.TextMatrix(0, 1) = "Keluhan & Order Booking"
    Call HapusForm
End Sub

Private Sub TblBatal_Click()
    If MsgBox("Yakin Membatalkan booking ?", vbYesNo, "Konfirmasi") = vbYes Then
        sSQL = "UPDATE BOOK set DELIV='BTL', PAID='D', FS_LOG_DEL='" & gSA & "', FD_TGL_DEL='" & TglYYYYMMDD(gTanggal) & "' " _
            & " where NO_POLISI ='" & Text(0) & "' and DELIV='BOK'"
        oConn.Execute sSQL
        Call HapusForm
        Text(0) = ""
        Text(0).SetFocus
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub TblSimpan_Click()
    Dim Boleh As Boolean
    Dim Pesan As String
    Boleh = True
    
    If CekTgl(Text(28)) = False Then
        Pesan = Pesan + "Tgl Booking. "
        Boleh = False
    End If
    If Text(8) = ">>:>>" Then
        Pesan = Pesan + "Jam Booking. "
        Boleh = False
    End If
    If Trim(Text(0)) = "" Then
        Pesan = Pesan + "No. Polisi "
        Boleh = False
    End If
    
    If Boleh = True Then
        Call SaveDataBooking(Text(0))
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

Private Sub SaveDataBooking(NoPolisi As String)
    
    sSQL = "Select * from BOOK where NO_POLISI ='" & NoPolisi & "' and DELIV='BOK'"
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
    gsSql2(4) = Text(0)
    gsSql1(5) = "NO_ENGINE"
    gsSql2(5) = Text(21)
'    gsSql1(6) = "VDS"
'    gsSql2(6) = Left(Text(25), 6)
    gsSql1(6) = "DELIV"
    gsSql2(6) = "BOK"
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
    gsSql1(18) = "DATE"
    gsSql2(18) = TglYYYYMMDD(Text(28))
    gsSql1(19) = "TIME"
    gsSql2(19) = Text(8)
    gsSql1(20) = "PAID"
    gsSql2(20) = Text(7)
    gsSql1(21) = "FRONT"
    gsSql2(21) = Text(14)
    gsSql1(22) = "REQ1"
    gsSql2(22) = Text(20)
    gsSql1(23) = "REQ2"
    gsSql2(23) = Text(26)
    gsSql1(24) = "REQ3"
    gsSql2(24) = Text(27)
    gsSql1(25) = "REQ4"
    gsSql2(25) = Text(29)
    gsSql1(26) = "REQ5"
    gsSql2(26) = Text(30)
    gsSql1(27) = "REMARK1"
    gsSql2(27) = Text(37)
    gsSql1(28) = "SUGGEST"
    gsSql2(28) = Text(38)
    gsSql1(28) = "NOFAK"
    gsSql2(28) = "BOOKED"

    gsSql1(29) = "REP1"
    gsSql2(29) = DGrid.TextMatrix(1, 1)
    gsSql1(30) = "REP2"
    gsSql2(30) = DGrid.TextMatrix(2, 1)
    gsSql1(31) = "REP3"
    gsSql2(31) = DGrid.TextMatrix(3, 1)
    gsSql1(32) = "REP4"
    gsSql2(32) = DGrid.TextMatrix(4, 1)
    gsSql1(33) = "REP5"
    gsSql2(33) = DGrid.TextMatrix(5, 1)

    If Not oRs.BOF Then
        sSQL = GensSql("BOOK", True) + " where NO_POLISI ='" & NoPolisi & "' and DELIV='BOK'"
    Else
        sSQL = GensSql("BOOK", False)
    End If
    oConn.Execute sSQL
    oRs.Close
End Sub

Private Sub PanggilDataBooking(NoPolisi As String)
    Dim j As Integer
    
    sSQL = "Select * from BOOK where NO_POLISI ='" & NoPolisi & "' and DELIV='BOK'"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        Text(1) = oRs1("CUSTOMER")
        Text(2) = oRs1("CUST_ADRES")
        Text(3) = oRs1("KOTA")
        Text(4) = oRs1("TELPON")
        Text(0) = oRs1("NO_POLISI")
        Text(21) = oRs1("NO_ENGINE")
        Text(25) = oRs1("NO_CHASSIS")
        Text(16) = oRs1("FS_TH_MOBIL")
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
        Text(28) = TglDDMMYYYY(oRs1("DATE"))
        Text(7) = oRs1("PAID")
        Text(8) = oRs1("TIME")
        Text(14) = oRs1("FRONT")
        Text(20) = oRs1("REQ1")
        Text(26) = oRs1("REQ2")
        Text(27) = oRs1("REQ3")
        Text(29) = oRs1("REQ4")
        Text(30) = oRs1("REQ5")
        Text(37) = oRs1("REMARK1")
        Text(38) = oRs1("SUGGEST")
        
        DGrid.TextMatrix(1, 1) = oRs1("REP1")
        DGrid.TextMatrix(2, 1) = oRs1("REP2")
        DGrid.TextMatrix(3, 1) = oRs1("REP3")
        DGrid.TextMatrix(4, 1) = oRs1("REP4")
        DGrid.TextMatrix(5, 1) = oRs1("REP5")
        Baris = 1
        For j = 1 To 5
            If Trim(DGrid.TextMatrix(j, 1)) <> "" Then
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
                gSearchOption = "BOOKBATAL"
                frmCari.Show (vbModal)
            Case 17
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
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
        Case 10, 11
            KeyAscii = CekNumerik(KeyAscii)
    End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                sSQL = "Select * from BOOK where NO_POLISI ='" & Text(0) & "' AND DELIV='BOK'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    Call HapusForm
                    MsgBox "No. Polisi tersebut tidak ada !"
                Else
                    Call PanggilDataBooking(Text(0))
                End If
                oRs.Close
            Else
                Call HapusForm
            End If
        Case 10, 11
            If Koma2Titik(Text(Index)) > 1 Then
                Text(Index) = 0
            End If
        Case 28
            If CekTgl(Text(Index)) = False Then
                Text(Index) = gTanggal
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
            End If
    End Select
End Sub
