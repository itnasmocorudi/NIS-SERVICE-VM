VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmtCustomer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Customer"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14550
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   14550
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9060
      Left            =   480
      TabIndex        =   57
      Top             =   0
      Width           =   13380
      Begin VB.CommandButton cmdExit 
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   11295
         Picture         =   "frmtCustomer.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   8520
         Width           =   2010
      End
      Begin VB.CommandButton cmdHapus 
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   9300
         Picture         =   "frmtCustomer.frx":06CF
         Style           =   1  'Graphical
         TabIndex        =   137
         Top             =   8520
         Width           =   2010
      End
      Begin VB.CommandButton cmdSimpan 
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   7320
         Picture         =   "frmtCustomer.frx":0C7E
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   8520
         Width           =   2010
      End
      Begin VB.CommandButton cmdPelanggan 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4640
         Picture         =   "frmtCustomer.frx":1252
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   430
         Width           =   2320
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   4055
         Left            =   7320
         TabIndex        =   86
         Top             =   4440
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CheckBox chAsuransi 
            BackColor       =   &H00FFFFFF&
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
            Left            =   5580
            TabIndex        =   134
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdPlgAsuransi 
            BackColor       =   &H00DCFCFB&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            Picture         =   "frmtCustomer.frx":1B33
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   3480
            Width           =   2320
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   38
            Left            =   1050
            TabIndex        =   87
            Top             =   1170
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   39
            Left            =   1050
            TabIndex        =   88
            Top             =   1470
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Index           =   37
            Left            =   1050
            TabIndex        =   43
            Top             =   840
            Width           =   4305
            _ExtentX        =   7594
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   40
            Left            =   1050
            TabIndex        =   44
            Top             =   2475
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   41
            Left            =   3915
            TabIndex        =   45
            Top             =   2475
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   42
            Left            =   1050
            TabIndex        =   46
            Top             =   2790
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   43
            Left            =   1050
            TabIndex        =   47
            Top             =   3120
            Visible         =   0   'False
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Index           =   52
            Left            =   1050
            TabIndex        =   42
            Top             =   480
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   53
            Left            =   1050
            TabIndex        =   125
            Top             =   1800
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   54
            Left            =   1050
            TabIndex        =   126
            Top             =   2100
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   55
            Left            =   2865
            TabIndex        =   127
            Top             =   2100
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox2 
            Height          =   270
            Left            =   2640
            TabIndex        =   135
            Top             =   480
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            X1              =   0
            X2              =   6000
            Y1              =   3960
            Y2              =   3960
         End
         Begin VB.Image Image3 
            Height          =   390
            Left            =   0
            Picture         =   "frmtCustomer.frx":238A
            Top             =   0
            Width           =   6000
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PSTN No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   130
            Top             =   1845
            Width           =   675
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hp.  No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   129
            Top             =   2145
            Width           =   600
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ";"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2755
            TabIndex        =   128
            Top             =   2145
            Width           =   75
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Lgn"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   270
            TabIndex        =   124
            Top             =   525
            Width           =   660
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   375
            TabIndex        =   95
            Top             =   3165
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "s/d"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3120
            TabIndex        =   94
            Top             =   2520
            Width           =   225
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Limit"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   93
            Top             =   2835
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   360
            TabIndex        =   92
            Top             =   2520
            Width           =   570
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   525
            TabIndex        =   91
            Top             =   885
            Width           =   405
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Kota"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   90
            Top             =   1515
            Width           =   330
         End
         Begin VB.Label Label32 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   435
            TabIndex        =   89
            Top             =   1215
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "&H000000FF&"
         Height          =   4695
         Left            =   120
         TabIndex        =   80
         Top             =   3720
         Width           =   7095
         Begin VB.CheckBox chLog 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Simpan ke Log"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   5040
            TabIndex        =   27
            Top             =   4215
            Value           =   1  'Checked
            Width           =   1845
         End
         Begin VB.CheckBox chSda 
            BackColor       =   &H00FFFFFF&
            Caption         =   "sda"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   6120
            TabIndex        =   11
            Top             =   40
            Width           =   735
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   31
            Left            =   930
            TabIndex        =   13
            Top             =   810
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   32
            Left            =   930
            TabIndex        =   16
            Top             =   1710
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Height          =   270
            Index           =   34
            Left            =   930
            TabIndex        =   18
            Top             =   2325
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   33
            Left            =   930
            TabIndex        =   17
            Top             =   2010
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptInclude   =   0   'False
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
            Mask            =   ">CCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   30
            Left            =   930
            TabIndex        =   12
            Top             =   480
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   35
            Left            =   930
            TabIndex        =   19
            Top             =   2625
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   36
            Left            =   3330
            TabIndex        =   20
            Top             =   2625
            Width           =   1710
            _ExtentX        =   3016
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   "9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   4410
            TabIndex        =   24
            Top             =   3570
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox MaskEdBox7 
            Height          =   270
            Left            =   6830
            TabIndex        =   113
            Top             =   3570
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   13
            Left            =   930
            TabIndex        =   23
            Top             =   3600
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   270
            Left            =   2760
            TabIndex        =   114
            Top             =   3600
            Width           =   60
            _ExtentX        =   106
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   930
            TabIndex        =   22
            Top             =   3270
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   48
            Left            =   930
            TabIndex        =   25
            Top             =   3900
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Left            =   930
            TabIndex        =   26
            Top             =   4200
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   57
            Left            =   930
            TabIndex        =   14
            Top             =   1100
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   58
            Left            =   930
            TabIndex        =   15
            Top             =   1400
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   60
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   59
            Left            =   930
            TabIndex        =   21
            Top             =   2950
            Width           =   5985
            _ExtentX        =   10557
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
         Begin VB.Line Line2 
            BorderColor     =   &H00FF0000&
            X1              =   0
            X2              =   7080
            Y1              =   4680
            Y2              =   4680
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   0
            Picture         =   "frmtCustomer.frx":30C8
            Top             =   0
            Width           =   6990
         End
         Begin VB.Label Label59 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   450
            TabIndex        =   133
            Top             =   2985
            Width           =   360
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "GSM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   495
            TabIndex        =   132
            Top             =   2670
            Width           =   315
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CDMA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2760
            TabIndex        =   131
            Top             =   2670
            Width           =   435
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Tempat / Tgl Lahir"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   119
            Top             =   3960
            Width           =   690
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   118
            Top             =   3645
            Width           =   495
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Hobi"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3960
            TabIndex        =   117
            Top             =   3600
            Width           =   315
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Disc."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   165
            TabIndex        =   116
            Top             =   3315
            Width           =   645
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   3600
            TabIndex        =   115
            Top             =   3322
            Width           =   120
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Pos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   85
            Top             =   2055
            Width           =   660
         End
         Begin VB.Label Label27 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   315
            TabIndex        =   84
            Top             =   855
            Width           =   495
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Kota"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   83
            Top             =   1755
            Width           =   330
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PSTN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   435
            TabIndex        =   82
            Top             =   2370
            Width           =   375
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   405
            TabIndex        =   81
            Top             =   510
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   " Tipe Pelanggan "
         Height          =   855
         Left            =   5520
         TabIndex        =   78
         Top             =   2355
         Visible         =   0   'False
         Width           =   1455
         Begin VB.OptionButton OpIn 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Internal"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton OpExt 
            BackColor       =   &H00FFFFFF&
            Caption         =   "External"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   4
         Left            =   990
         TabIndex        =   53
         Top             =   1065
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   5
         Left            =   990
         TabIndex        =   5
         Top             =   1365
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   6
         Left            =   990
         TabIndex        =   54
         Top             =   2025
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Index           =   11
         Left            =   990
         TabIndex        =   55
         Top             =   3255
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   8430
         TabIndex        =   28
         Top             =   435
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   44
         Left            =   1035
         TabIndex        =   49
         Top             =   8505
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   28
         Left            =   12285
         TabIndex        =   56
         Top             =   3705
         Visible         =   0   'False
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   11880
         TabIndex        =   41
         Top             =   1065
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   18
         Left            =   8430
         TabIndex        =   29
         Top             =   1065
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   20
         Left            =   8430
         TabIndex        =   31
         Top             =   1860
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   21
         Left            =   8430
         TabIndex        =   32
         Top             =   2175
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   24
         Left            =   8430
         TabIndex        =   33
         Top             =   2775
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   23
         Left            =   8430
         TabIndex        =   59
         Top             =   2475
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   26
         Left            =   8430
         TabIndex        =   38
         Top             =   3705
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   27
         Left            =   8430
         TabIndex        =   39
         Top             =   4035
         Width           =   1935
         _ExtentX        =   3413
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
         Index           =   22
         Left            =   12150
         TabIndex        =   60
         Top             =   2160
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   2565
         TabIndex        =   2
         Top             =   435
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   11520
         TabIndex        =   61
         Top             =   2175
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   8
         Left            =   990
         TabIndex        =   8
         Top             =   2640
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   8430
         TabIndex        =   72
         Top             =   765
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox6 
         Height          =   270
         Left            =   11520
         TabIndex        =   73
         Top             =   435
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   29
         Left            =   12285
         TabIndex        =   74
         Top             =   4035
         Visible         =   0   'False
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   3
         Left            =   3500
         TabIndex        =   52
         Top             =   765
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2700
         TabIndex        =   76
         Top             =   765
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   9
         Left            =   990
         TabIndex        =   9
         Top             =   2940
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   10
         Left            =   3510
         TabIndex        =   10
         Top             =   2940
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   990
         TabIndex        =   7
         Top             =   2325
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   19
         Left            =   8430
         TabIndex        =   30
         Top             =   1365
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         Top             =   435
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Index           =   45
         Left            =   8430
         TabIndex        =   34
         Top             =   3090
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   270
         Left            =   3900
         TabIndex        =   97
         Top             =   435
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   25
         Left            =   8430
         TabIndex        =   35
         Top             =   3390
         Width           =   495
         _ExtentX        =   873
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
         Index           =   46
         Left            =   9045
         TabIndex        =   36
         Top             =   3390
         Width           =   690
         _ExtentX        =   1217
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
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   47
         Left            =   9870
         TabIndex        =   37
         Top             =   3390
         Width           =   495
         _ExtentX        =   873
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   11520
         TabIndex        =   121
         Top             =   2775
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox8 
         Height          =   270
         Left            =   1600
         TabIndex        =   122
         Top             =   435
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   990
         TabIndex        =   4
         Top             =   765
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   50
         Left            =   990
         TabIndex        =   0
         Top             =   435
         Width           =   585
         _ExtentX        =   1032
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
         Index           =   51
         Left            =   870
         TabIndex        =   40
         Top             =   4260
         Visible         =   0   'False
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Height          =   270
         Index           =   56
         Left            =   990
         TabIndex        =   6
         Top             =   1680
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   0
         Picture         =   "frmtCustomer.frx":422E
         Top             =   0
         Width           =   13395
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   11640
         TabIndex        =   120
         Top             =   450
         Width           =   120
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   9720
         TabIndex        =   112
         Top             =   1080
         Width           =   120
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   10440
         TabIndex        =   111
         Top             =   4080
         Width           =   120
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   10440
         TabIndex        =   110
         Top             =   3720
         Width           =   120
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2370
         TabIndex        =   109
         Top             =   2370
         Width           =   120
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(YYYY)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   9120
         TabIndex        =   108
         Top             =   1905
         Width           =   480
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2955
         TabIndex        =   107
         Top             =   2978
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   106
         Top             =   1140
         Width           =   405
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   375
         TabIndex        =   105
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   104
         Top             =   2100
         Width           =   330
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rangka"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   330
         TabIndex        =   103
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   435
         TabIndex        =   102
         Top             =   3300
         Width           =   435
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   495
         TabIndex        =   101
         Top             =   2685
         Width           =   375
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Lgn"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   100
         Top             =   803
         Width           =   660
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   555
         TabIndex        =   99
         Top             =   2985
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Pos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   98
         Top             =   2370
         Width           =   660
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Km Terakhir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7440
         TabIndex        =   96
         Top             =   3135
         Width           =   840
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO TAM"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7710
         TabIndex        =   79
         Top             =   1410
         Width           =   585
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Wilayah"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         TabIndex        =   77
         Top             =   810
         Width           =   570
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Disc."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11805
         TabIndex        =   75
         Top             =   4080
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7830
         TabIndex        =   71
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   70
         Top             =   8535
         Width           =   585
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PPN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11865
         TabIndex        =   69
         Top             =   3750
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11415
         TabIndex        =   68
         Top             =   1110
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7320
         TabIndex        =   67
         Top             =   1110
         Width           =   975
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Th. Mobil"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7635
         TabIndex        =   66
         Top             =   1905
         Width           =   645
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7860
         TabIndex        =   65
         Top             =   2220
         Width           =   420
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7800
         TabIndex        =   64
         Top             =   2820
         Width           =   480
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Mesin"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7605
         TabIndex        =   63
         Top             =   3750
         Width           =   705
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Kunci"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7635
         TabIndex        =   62
         Top             =   4080
         Width           =   675
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7650
         TabIndex        =   58
         Top             =   3435
         Width           =   645
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PKP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   123
         Top             =   4305
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9375
      Left            =   0
      Picture         =   "frmtCustomer.frx":566F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   450
   End
End
Attribute VB_Name = "frmtCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim pds As Boolean

Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
          Dim aa As String
          Dim xHasil As String
1        On Error GoTo GensSql_Error

2         xHasil = ""
3         If xUpdate Then
4             xHasil = "UPDATE " + xTabel + " SET "
5             i = 0
6             Do While gsSql1(i) <> ""
7                 If i <> 0 Then
8                     xHasil = xHasil + ","
9                 End If
                  
10                aa = UCase(gsSql1(i))
11                Select Case aa
                  Case "PPN", "DISC"
12                    xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
13                Case Else
14                    xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
15                End Select
16                i = i + 1
17            Loop
18        Else
19            xHasil = "INSERT INTO " + xTabel + " ("
20            i = 0
21            Do While gsSql1(i) <> ""
22                If i <> 0 Then
23                    xHasil = xHasil + ","
24                End If
25                xHasil = xHasil + gsSql1(i)
26                i = i + 1
27            Loop
28            xHasil = xHasil + ") VALUES ("
              
29            i = 0
30            Do While gsSql1(i) <> ""
31                If i <> 0 Then
32                    xHasil = xHasil + ","
33                End If
                  
34                aa = UCase(gsSql1(i))
35                Select Case aa
                  Case "PPN", "DISC"
36                    xHasil = xHasil + gsSql2(i)
37                Case Else
38                    xHasil = xHasil + "'" + gsSql2(i) + "'"
39                End Select
40                i = i + 1
41            Loop
42            xHasil = xHasil + ")"
43        End If
44        GensSql = xHasil

45       On Error GoTo 0
46       Exit Function

GensSql_Error:

47        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GensSql of Form frmtCustomer")
End Function

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub chAsuransi_Click()
    If chAsuransi.value = Checked Then
        Text(52).Enabled = True
        Text(37).Enabled = True
        Text(40).Enabled = True
        Text(41).Enabled = True
        Text(42).Enabled = True
        Text(43).Enabled = True
        Text(52).SetFocus
    ElseIf chAsuransi.value = Unchecked Then
        Text(52).Enabled = False
        Text(37).Enabled = False
        Text(40).Enabled = False
        Text(41).Enabled = False
        Text(42).Enabled = False
        Text(43).Enabled = False
        Text(37) = ""
        n = Text(40).Mask
        Text(40).Mask = ""
        Text(40) = ""
        Text(40).Mask = n
        n = Text(41).Mask
        Text(41).Mask = ""
        Text(41) = ""
        Text(41).Mask = n
        Text(42) = "0"
        Text(43) = ""
    End If
End Sub

Private Sub chAsuransi_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub chLog_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub chSda_Click()
    If chSda.value = Checked Then
        Text(30) = Text(4)
        Text(31) = Text(5)
        Text(57) = Text(56)
        Text(32) = Text(6)
        Text(33) = Text(7)
        Text(34) = Text(8)
        Text(35) = Text(9)
        Text(36) = Text(10)
        Text(12).SetFocus
    ElseIf chSda.value = Unchecked Then
        Text(30) = ""
        Text(31) = ""
        Text(32) = ""
        Text(33) = ""
        Text(34) = ""
        Text(35) = ""
        Text(36) = ""
        Text(12) = ""
        Text(13) = ""
        Text(14) = ""
        Text(48) = ""
        Text(49) = ">>->>->>>>"
        Text(30).SetFocus
    End If
End Sub

Private Sub chSda_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    Set rsTemp = objO.FetchData("SELECT NO_CHASSIS FROM JOBDTA WITH(NOLOCK) WHERE DELIV <> 'BTL' AND FS_LOG_DEL = '' AND NO_CHASSIS = '" & Trim(Trim(Text(50)) & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")) & "'")
    If Not rsTemp.EOF Then
        MsgBox "Kendaraan ini tidak boleh dihapus, karena sudah memiliki transaksi. "
    Else
        If MsgBox("Ingin menghapus data ini ?", vbYesNo, "Konfirmasi") = vbYes Then
            sSQL = "delete CUSTDAT where NO_CHASSIS ='" & Trim(Text(50)) & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'"
            Call objO.UpdateData(sSQL)
            Call hapusForm
        End If
    End If
End Sub

Private Sub cmdPelanggan_Click()
    With frmtPlgn
        .Text(57) = Text(2)
        .Text(1) = Text(4)
        .Text(2) = Left(Text(5), 40)
        .Text(3) = Text(6)
        .Text(4) = Text(8) + "," + Trim(Text(9)) + "," + Trim(Text(10))
        .Text(6) = Trim(Text(11))
        gBantuan2 = "kosong"
        frmtPlgn.show vbModal, Me
    End With
End Sub

Private Sub cmdPlgAsuransi_Click()
    With frmtPlgn
        .Text(57) = Trim(Text(52))
        .Text(1) = Trim(Text(37))
        .Text(2) = Trim(Text(38))
        .Text(3) = Trim(Text(39))
        .Text(4) = Text(8) + "," + Trim(Text(9)) + "," + Trim(Text(10))
        .Text(6) = Trim(Text(11))
        gBantuan2 = "kosong"
        frmtPlgn.show vbModal, Me
    End With
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
      Dim rsSearch As ADODB.Recordset

1         On Error GoTo vbError
2         If Trim(Text(0)) <> "" Or Trim(Text(1)) <> "" Then
3             boleh = True
              
              ' Cek Field mandatory yg kosong
4             If Trim(Text(50)) = "" And Text(1) = "" And Text(0) = "" Then
5                 boleh = False
6                 pesan = pesan + "No Rangka. "
7             End If
8             If Text(2) = "" Then
9                 boleh = False
10                pesan = pesan + "Kode Pelanggan. "
11            End If
12            If Text(8) = "" And Text(9) = "" And Text(10) = "" Then
13                boleh = False
14                pesan = pesan + "No Telpon. "
15            End If
16            If Text(34) = "" And Text(35) = "" And Text(36) = "" Then
17                boleh = False
18                pesan = pesan + "No Telpon Surat Menyurat. "
19            End If
20            If Text(12) = "" Then
21                boleh = False
22                pesan = pesan + "Job Description. "
23            End If
24            If Text(15) = "" Then
25                boleh = False
26                pesan = pesan + "Dealer Penjual. "
27            Else
28                Set rsSearch = objO.FetchData("SELECT FS_KD_DEALER FROM T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER = '" & Trim(Text(15)) & "'")
29                If rsSearch.EOF Then
30                    boleh = False
31                    pesan = pesan & vbCr & "Dealer Penjual tidak sesuai master. " & vbCr
32                End If
33            End If
34            If Text(18) = ">>->>->>>>" Then
35                boleh = False
36                pesan = pesan + "Tanggal Delivery. "
37            Else
                  ''2016-05-27 || penambahan validasi delivery date
38                If Not CekTgl(Text(18)) Then
39                   boleh = False
40                   pesan = pesan + "Tanggal Delivery. "
41                ElseIf Trim(Text(18)) = "01/01/1900" Then
42                   boleh = False
43                   pesan = pesan + "Tanggal Delivery tidak boleh 01/01/1900."
44                ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Trim(Text(18)), "dd/mm/yyyy")) > 0 Then
45                   boleh = False
46                   pesan = pesan + "Tanggal Delivery tidak boleh melebihi tanggal hari ini."
47                Else
48                   Text(18) = TglDDMMYYYY(Text(18))
49                End If
50            End If
51            If Text(19) = "" Then
52                boleh = False
53                pesan = pesan + "DOTAM. "
54            End If
55            If Text(20) = "" Then
56                boleh = False
57                pesan = pesan + "Tahun Kendaraan. "
58            ElseIf IsNumeric(Trim(Text(20))) = False Then
59                boleh = False
60                pesan = pesan + "Tahun Kendaraan harus angka. "
61            ElseIf year(isnow) - CInt(Trim(Text(20))) > 50 Then
62                boleh = False
63                pesan = pesan + "Tahun Kendaraan minimal tahun " & year(isnow) - 50 & ". "
64            ElseIf CInt(Trim(Text(20))) - year(isnow) > 0 Then
65                boleh = False
66                pesan = pesan + "Tahun Kendaraan maksimal tahun " & year(isnow) & ". "
67            End If
68            If Text(21) = "" Then
69                boleh = False
70                pesan = pesan + "Model. "
71            Else
72                Set rsSearch = objO.FetchData("SELECT MDL FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Trim(Text(21)) & "'")
73                If rsSearch.EOF Then
74                    boleh = False
75                    pesan = pesan & vbCr & "Model tidak terdapat pada master. " & vbCr
76                End If
77            End If
78            If Text(22) = "" Then
79                boleh = False
80                pesan = pesan & vbCr & "Ada kesalahan pada master model. " & vbCr
81            Else
82                Set rsSearch = objO.FetchData("SELECT CLASSCODE FROM CARCLASS WITH(NOLOCK) WHERE CLASSCODE = '" & Trim(Text(22)) & "'")
83                If rsSearch.EOF Then
84                    boleh = False
85                    pesan = pesan & vbCr & "Ada kesalahan pada master model. " & vbCr
86                End If
87            End If
88            If Text(24) = "" Then
89                boleh = False
90                pesan = pesan + "Warna. "
91            End If
92            If Text(26) = "" Then
93                boleh = False
94                pesan = pesan + "No Mesin. "
95            End If
96            If Text(27) = "" Then
97                boleh = False
98                pesan = pesan + "No Kunci. "
99            End If
100           If Trim(Text(45)) = "" Then
101               Text(45) = 0
102           End If
103           If Text(59) = "" Then
104               MsgBox "Alamat Email tidak boleh kosong" & Chr(13) & "Input konfirmasi untuk melanjutkan proses. (Jika Customer memang tidak punya alamat Email)", vbCritical, "Kelangkapan Data Email"
105               If Trim(InputBox("Masukkan kata konfirmasi " & Chr(13) & "Konfirmasi : Pelanggan belum memiliki email")) <> "Pelanggan belum memiliki email" Then
106                   boleh = False
107                   pesan = pesan + "Email. "
108               End If
109           End If
              
              '20/12/2012  Cust Non Axtif
              Dim rsNon As ADODB.Recordset
110           Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH(NOLOCK) WHERE kodeLgn = '" & Text(2) & "' and nonaktifflag = 0")
111           If rsNon.RecordCount = 0 Then
112               boleh = False
113               pesan = "Kode Langgan : " & Text(2) & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Hubungi Admin Untuk Mengaktifkan Kode Langgan Tersebut"
114           End If
              
115           If boleh Then
116               Call saveLogMenu(gSA, ComptName, gVersi, "frmCustomer", "cmdSimpan", 0)
117               erTranAll = False
118               Call objHT.OpenTransactionAll
                  
119               Call ClearsSql
120               gsSql1(0) = "NO_CHASSIS"
121               gsSql2(0) = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
122               gsSql1(1) = "FS_KD_PLG"
123               gsSql2(1) = Text(2)
124               gsSql1(2) = "FS_KD_WIL"
125               gsSql2(2) = Text(3)
126               gsSql1(3) = "CUSTOMER"
127               gsSql2(3) = Text(4)
128               gsSql1(4) = "CUST_ADRES"
129               gsSql2(4) = Text(5) + Text(56)
130               gsSql1(5) = "KOTA"
131               gsSql2(5) = Text(6)
132               gsSql1(6) = "KODEPOS"
133               gsSql2(6) = Text(7)
134               gsSql1(7) = "TELPON"
135               gsSql2(7) = Text(8)
136               gsSql1(8) = "GSM"
137               gsSql2(8) = Text(9)
138               gsSql1(9) = "CDMA"
139               gsSql2(9) = Text(10)
140               gsSql1(10) = "NPWP"
141               gsSql2(10) = Text(11)
142               gsSql1(11) = "USEJOBDSC"
143               gsSql2(11) = Text(12)
144               gsSql1(12) = "USEAGAMA"
145               gsSql2(12) = Text(13)
146               gsSql1(13) = "USEFS_NM_HOBI"
147               gsSql2(13) = Text(14)
148               gsSql1(14) = "FS_KD_DEALER"
149               gsSql2(14) = Text(15)
150               gsSql1(15) = "FS_NM_DEALER"
151               gsSql2(15) = Text(16)
152               gsSql1(16) = "CODE"
153               gsSql2(16) = Text(17)
154               gsSql1(17) = "DDATE"
155               gsSql2(17) = Format(IIf(Replace(Replace(Text(18), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(18), ">", ""), "-", "/")), "dd/mm/yyyy")
156               gsSql1(18) = "DOTAM"
157               gsSql2(18) = Format(IIf(Replace(Replace(Text(19), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(19), ">", ""), "-", "/")), "mm/dd/yyyy")
158               gsSql1(19) = "FS_TH_MOBIL"
159               gsSql2(19) = Text(20)
160               gsSql1(20) = "MDL"
161               gsSql2(20) = Text(21)
162               gsSql1(21) = "FMDL"
163               gsSql2(21) = Text(23)
164               gsSql1(22) = "MODEL"
165               gsSql2(22) = Text(22)
166               gsSql1(23) = "COLOR"
167               gsSql2(23) = Text(24)
168               gsSql1(24) = "NO_POLISI"
169               gsSql2(24) = Trim(Text(25)) & IIf(Trim(Text(25)) <> "", "-", "") & Trim(Text(46)) & IIf(Trim(Text(47)) <> "", "-", "") & Text(47)
170               gsSql1(25) = "NO_ENGINE"
171               gsSql2(25) = Text(26)
172               gsSql1(26) = "KEYNO"
173               gsSql2(26) = Text(27)
174               gsSql1(27) = "PPN"
175               gsSql2(27) = SimpanKoma2Titik(Text(28))
176               gsSql1(28) = "DISC"
177               gsSql2(28) = Val(Text(29))
178               gsSql1(29) = "USENAME"
179               gsSql2(29) = Text(30)
180               gsSql1(30) = "USEADDR"
181               gsSql2(30) = Text(31) + Text(57) + Text(58)
182               gsSql1(31) = "USECITY"
183               gsSql2(31) = Text(32)
184               gsSql1(32) = "USEPOS"
185               gsSql2(32) = Text(33)
186               gsSql1(33) = "USEPSTN"
187               gsSql2(33) = Text(34)
188               gsSql1(34) = "USEGSM"
189               gsSql2(34) = Text(35)
190               gsSql1(35) = "USECDMA"
191               gsSql2(35) = Text(36)
192               gsSql1(36) = "ASURANSI"
193               gsSql2(36) = Text(37)
194               gsSql1(37) = "ASR_ADDR"
195               gsSql2(37) = Text(38)
196               gsSql1(38) = "ASR_CITY"
197               gsSql2(38) = Text(39)
198               gsSql1(39) = "ASR_TGLM"
199               gsSql2(39) = Format(IIf(Replace(Replace(Text(40), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(40), ">", ""), "-", "/")), "mm/dd/yyyy")
200               gsSql1(40) = "ASR_TGLS"
201               gsSql2(40) = Format(IIf(Replace(Replace(Text(41), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(41), ">", ""), "-", "/")), "mm/dd/yyyy")
202               gsSql1(41) = "ASR_LIMIT"
203               gsSql2(41) = Text(42)
204               gsSql1(42) = "ASR_NOTE"
205               gsSql2(42) = Text(43)
206               gsSql1(43) = "NOTE"
207               gsSql2(43) = Text(44)
208               gsSql1(44) = "FS_KM_AKHIR"
209               gsSql2(44) = CDbl(Text(45))
210               gsSql1(45) = "USERADD"
211               gsSql2(45) = gSA
212               gsSql1(46) = "USET4LAHIR"
213               gsSql2(46) = Text(48)
214               gsSql1(47) = "USETGLLAHIR"
215               gsSql2(47) = Format(IIf(Replace(Replace(Text(49), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(49), ">", ""), "-", "/")), "mm/dd/yyyy")
216               gsSql1(48) = "PKP"
217               gsSql2(48) = Trim(Text(51))
218               gsSql1(49) = "KDPLG_ASR"
219               gsSql2(49) = Trim(Text(52))
220               gsSql1(50) = "USEEMAIL"
221               gsSql2(50) = Trim(Text(59))
                  
222               If Log.loyalityCustomer = 1 Then
223                   Set rsSearch = objHT.FetchDataBengkel("select * froM RplNasApp..UnitCustomer where NoRangka = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
224                   If rsSearch.EOF Then
225                       Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM RplNasApp..UnitCustomer WHERE IDCUSTSERVICE = '" & Trim(Text(2)) & "'")
226                       If Not rsSearch.EOF Then
227                           sSQL = "INSERT INTO [RplNasApp]..UnitCustomer(ddate,fmdl,noEngine,IDCustService,NoRangka,KdDlrPenjual,nmDlrPenjual,noPolisi,mdl,thmobil,userentry,tglentry)"
228                           sSQL = sSQL & " VALUES ('" & Format(IIf(Replace(Replace(Text(18), ">", ""), "-", "/") = "//", "", Replace(Replace(Text(18), ">", ""), "-", "/")), "dd/mm/yyyy") & "','" & Nul(Text(23)) & "','" & Nul(Text(26)) & "','" & Nul(Text(2)) & "','" & Nul(Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")) & "','" & Nul(Text(15)) & "','" & Nul(Text(16)) & "','" & Nul(Trim(Text(25)) & IIf(Trim(Text(25)) <> "", "-", "") & Trim(Text(46)) & IIf(Trim(Text(47)) <> "", "-", "") & Text(47)) & "','" & Nul(Text(21)) & "','" & Nul(Text(20)) & "','" & gSA & "',getdate())"
229                           Call objHT.InsertData_Bengkel(sSQL)
230                       End If
                          
231                   End If
232               End If
                  
233               Set rsSearch = objO.FetchData("Select * from CUSTDAT WITH(NOLOCK) where NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' ")
234               If Not rsSearch.BOF Then
235                   Set rsTemp = objO.FetchData("SELECT * FROM HIST_PEMILIK WITH(NOLOCK) WHERE NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' AND FS_KD_PLG = '" & Text(2) & "'")
236                   If rsTemp.EOF Then
237                       Call objHT.UpdateData_Bengkel(GensSql2("HIST_PEMILIK", False))
238                   Else
239                       Call objHT.UpdateData_Bengkel(GensSql2("HIST_PEMILIK", True) + " WHERE NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' AND FS_KD_PLG = '" & Text(2) & "'")
240                   End If
241                   rsSearch.MoveFirst
                      
242                   sSQL = GensSql("CUSTDAT", True) + " where NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'"
243                   If rsSearch("ASURANSI") <> Trim(Text(37)) Then
244                       Set rsTemp = objO.FetchData("SELECT * FROM HIST_ASURANSI WITH(NOLOCK) WHERE fs_kd_plg = '" & Text(2) & "' AND no_Chassis = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' ORDER BY urut DESC")
245                       If Not rsTemp.EOF Then
246                           If rsTemp.RecordCount > 1 Then
247                               Call objHT.UpdateData_Bengkel("UPDATE HIST_ASURANSI SET urut = '" & CDbl(rsTemp("urut")) + 1 & "' WHERE fs_kd_plg = '" & Text(2) & "' AND no_Chassis = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
248                           ElseIf rsTemp.RecordCount = 1 Then
249                               Call objHT.UpdateData_Bengkel("UPDATE HIST_ASURANSI SET urut = '1' WHERE fs_kd_plg = '" & Text(2) & "' AND no_Chassis = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
250                           End If
251                       End If
252                       If Trim(Text(37)) <> "" Then
253                           Call objHT.InsertData_Bengkel("INSERT INTO HIST_ASURANSI (NO_CHASSIS,FS_KD_PLG,CUSTOMER,ASURANSI,ASR_ADDR,ASR_CITY,ASR_TGLM,ASR_TGLS,ASR_LIMIT,ASR_NOTE,KDPLG_ASR) VALUES ('" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "','" & Text(2) & "','" & Text(4) & "','" & Text(37) & "','" & Text(38) & "','" & Text(39) & "','" & Format(gsSql2(39), "MM/DD/YYYY") & "','" & Format(gsSql2(40), "MM/DD/YYYY") & "'," & CDbl(Text(42)) & ",'" & Text(43) & "','" & Trim(Text(52)) & "')")
254                       End If
255                   End If
256               Else
257                   sSQL = GensSql("CUSTDAT", False)
258                   Set rsTemp = objO.FetchData("SELECT * FROM HIST_PEMILIK WITH(NOLOCK) WHERE NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' AND FS_KD_PLG = '" & Text(2) & "'")
259                   If rsTemp.EOF Then
260                       Call objHT.UpdateData_Bengkel(GensSql2("HIST_PEMILIK", False))
261                   Else
262                       Call objHT.UpdateData_Bengkel(GensSql2("HIST_PEMILIK", True) + "WHERE NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "' AND FS_KD_PLG = '" & Text(2) & "'")
263                   End If
264                   If Text(37) <> "" Then
265                       Call objHT.InsertData_Bengkel("INSERT INTO HIST_ASURANSI (NO_CHASSIS,FS_KD_PLG,CUSTOMER,ASURANSI,ASR_ADDR,ASR_CITY,ASR_TGLM,ASR_TGLS,ASR_LIMIT,ASR_NOTE,KDPLG_ASR) VALUES ('" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "','" & Text(2) & "','" & Text(4) & "','" & Text(37) & "','" & Text(38) & "','" & Text(39) & "','" & Format(gsSql2(39), "MM/DD/YYYY") & "','" & Format(gsSql2(40), "MM/DD/YYYY") & "'," & CDbl(Text(42)) & ",'" & Text(43) & "','" & Trim(Text(52)) & "')")
266                   End If
267               End If
268               Call objHT.UpdateData_Bengkel(sSQL)
                  
                  ' UPDATE DATA CUSTOMER KE SPKCABANG
269               If Trim(Text(30)) <> "" Then
                      ' KLO BELUM ADA PDS TANYA MO PDS APA GAK, KLO PDS TIDAK UPDATE
270                   If Not pds Then
271                       If MsgBox("Akan melakukan transaksi PDS ?", vbDefaultButton2 + vbYesNo, "PDS") = vbNo Then
272                               If chLog = 1 Then
273                                   Call objHT.InsertData_SPK("INSERT INTO SCUST VALUES ('" & Trim(gsSql2(0)) & "','" & Trim(Text(30)) & "','" & Trim(Text(31)) + Trim(Text(57)) + Trim(Text(58)) & "','" & Trim(Text(33)) & "','" & Trim(Text(32)) & "','','" & Trim(Text(34)) & "','" & Trim(Text(35)) & "','" & Trim(Text(36)) & "', '" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN") & "','" & Trim(Text(59)) & "')")
                                      'Call objHT.InsertData_SPK("INSERT INTO SCUST VALUES ('" & Trim(gsSql2(0)) & "','" & Trim(Text(30)) & "','" & Trim(Text(31)) + Trim(Text(57)) + Trim(Text(58)) & "','" & Trim(Text(33)) & "','" & Trim(Text(32)) & "','','" & Trim(Text(34)) & "','" & Trim(Text(35)) & "','" & Trim(Text(36)) & "', '" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN") & "')")
274                               End If
275                               Call objHT.UpdateData_SPK("UPDATE SPK SET SNAMA = '" & Trim(Text(30)) & "', SALAMAT = '" & Trim(Text(31)) + Trim(Text(57)) + Trim(Text(58)) & "', SKODEPOS = '" & Trim(Text(33)) & "', SKOTA = '" & Trim(Text(32)) & "',  " _
                                                  & "STELP = '" & Trim(Text(34)) & "', SGSM = '" & Trim(Text(35)) & "', SCDMA = '" & Trim(Text(36)) & "', EMAIL = '" & Trim(Text(59)) & "' WHERE NORANGKA = '" & Trim(gsSql2(0)) & "' AND SPK_STATUS <> '1'")
                                  'Call objHT.UpdateData_SPK("UPDATE SPK SET SNAMA = '" & Trim(Text(30)) & "', SALAMAT = '" & Trim(Text(31)) + Trim(Text(57)) + Trim(Text(58)) & "', SKODEPOS = '" & Trim(Text(33)) & "', SKOTA = '" & Trim(Text(32)) & "',  " _
                                  '                & "STELP = '" & Trim(Text(34)) & "', SGSM = '" & Trim(Text(35)) & "', SCDMA = '" & Trim(Text(36)) & "' WHERE NORANGKA = '" & Trim(gsSql2(0)) & "' AND SPK_STATUS <> '1'")
276                       End If
277                   End If
278               End If
                  
                  ' UPDATE NO RANGKA DI BOOK
279               If gBantuan = "booking" Then
280                   Call objHT.UpdateData_Bengkel("UPDATE book SET no_chassis = '" & gsSql2(0) & "' WHERE nofak = '" & gBantuan2 & "'")
281               End If
                  
                  ' TRANSFER HISTORY DARI BENGKEL LAMA
               '   Dim noRgk As String
               '   Dim rsHlp As ADodb.Recordset
                  
               '      noRgk = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
               '           Set rsHlp = objO.FetchDataAAS("SELECT * FROM CUSTDAT WHERE NO_CHASSIS = '" & noRgk & "'")
               '           If Not rsHlp.EOF Then
               '               If MsgBox("Akan transfer history Bengkel Lama ?", vbDefaultButton2 + vbYesNo, "PDS") = vbYes Then
               '                   Call transferHistoryLama(noRgk, rsHlp)
               '                  updateHistory = False
               '               Else
               '                   updateHistory = True
               '               End If
               '           End If
                      
                              
282               If erTranAll = False Then
283                   Call objHT.CloseTransactionAll
284                   Call saveLogMenu(gSA, ComptName, gVersi, "frmCustomer", "cmdSimpan", 1)
285               Else
286                   Call objHT.CloseTransactionAll_Err
287                   Call saveLogMenu(gSA, ComptName, gVersi, "frmCustomer", "cmdSimpan", 2)
288                   MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
289               End If
290               With frmOrder
291                   If .Visible Then
292                       .Text(1) = Text(2)
293                       .Text(14) = gsSql2(24)
294                       Unload Me
295                   End If
296               End With
                  
297               Text(25) = ""
298               Call hapusForm
299               Unload Me
300           Else
301               MsgBox "Lengkapi field berikut ..." & vbCr & pesan, vbCritical, "Field Mandatory"
302           End If
303       End If
304       Exit Sub
          
vbError:
305       If erTranAll = True Then
306           Call objHT.CloseTransactionAll_Err
307           Call saveLogMenu(gSA, ComptName, gVersi, "frmCustomer", "cmdSimpan", 2)
308           MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmTCustomer"), vbInformation, "Gagal Simpan"
309       End If
310       Exit Sub
End Sub

Private Sub Form_Load()
    Call hapusForm
    pds = False
    chAsuransi.value = Unchecked
    chAsuransi_Click
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub OpExt_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub OpIn_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
        
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            If Index = 0 Then
                Text(Index) = Right(gMainKeyCari, 8)
                Text(1) = Mid(gMainKeyCari, 4, Len(gMainKeyCari) - 11)
                Text(50) = Left(gMainKeyCari, 3)
            Else
                Text(Index) = gMainKeyCari
            End If
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
    
    Select Case Index
        Case 50
            If Text(Index) <> "" And Text(1) <> "" Then
                PanggilCustomer
            End If
    End Select
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
                gSearchOption = "CUSTOMER2"
                frmCari.show (vbModal)
            Case 2
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.show (vbModal)
            Case 13
                gMainKeyCari = Text(Index)
                gSearchOption = "AGAMA"
                frmCari.show (vbModal)
            Case 14
                gMainKeyCari = Text(Index)
                gSearchOption = "HOBI"
                frmCari.show (vbModal)
            Case 15
                gMainKeyCari = Text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.show (vbModal)
            Case 21
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 24
                If Text(21) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "WARNA"
                    gBantuan = Trim(Text(21))
                    frmCariSPK.show (vbModal)
                Else
                    MsgBox "Isi Model Kendaraan"
                    Text(21).SelLength = Len(Text(21))
                    Text(21).SetFocus
                    Exit Sub
                End If
            Case 37
                gMainKeyCari = Text(Index)
                gSearchOption = "ASURANSI"
                frmCari.show (vbModal)
            Case 50
                gMainKeyCari = Text(Index)
                gSearchOption = "COUNTRY"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
    KeyAscii = UpCase(KeyAscii)
    Select Case Index
        Case 7, 33, 46
            KeyAscii = CekNumerik(KeyAscii)
        Case 47, 50, 1, 0
            KeyAscii = IIf(CekNumerik(KeyAscii) = 0, CekAlfabet(KeyAscii), CekNumerik(KeyAscii))
    End Select
End Sub

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 20
            If Trim(Text(Index)) = "" Then
                Text(Index) = 0
                Text(Index).SetFocus
                Exit Sub
            End If
            If CDbl(Trim(Text(Index))) > (year(isnow) + 1) Or CDbl(Trim(Text(Index))) < (year(isnow) - 60) Then
                MsgBox "Tahun Kendaraan salah"
                Text(Index) = ""
                Text(Index).SetFocus
            End If
        Case 42
            If Trim(Text(Index)) = "" Then
                Text(Index) = 0
                Exit Sub
            End If
        Case 45
            If Text(Index) = "" Then
                Text(Index) = 0
                Exit Sub
            End If
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As ADODB.Recordset

1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 0
3                 If Text(Index) <> "" Then
4                     Set rsSearch = objO.FetchData("select * from custdat WITH(NOLOCK) where no_chassis = '" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'")
5                     If rsSearch.EOF Then
6                         If Text(Index) = "" Then
7                             Cancel = True
8                             gMainKeyCari = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
9                             gSearchOption = "CUSTOMER2"
10                            frmCari.show (vbModal)
11                        End If
12                        PanggilCustomer
13                    Else
14                        PanggilCustomer
15                    End If
16                End If
17            Case 2
18                If Trim(Text(Index)) <> "" Then
19                    Text(3) = Left(Text(Index), 3)
                          
                          'Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp from Armst where KodeLgn = '" & Trim(text(Index)) & "'")
                          '20/12/2012 Cust Non Axtif
20                        Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp from Armst WITH(NOLOCK) where KodeLgn = '" & Trim(Text(Index)) & "' and nonaktifflag  = 0")
21                        If rsSearch.BOF Then
22                            temp = MsgBox("Tambah Kode Pelanggan baru ?", vbYesNo, "KODE PELANGGAN")
23                            If temp = vbYes Then
24                                cmdPelanggan_Click
25                            Else
26                                Cancel = True
27                                gMainKeyCari = Text(Index)
28                                gSearchOption = "KODEPLG"
29                                frmCariAcess.show (vbModal)
30                            End If
31                        Else
32                            If IsNull(rsSearch("NamaLgn")) Then
33                                Text(4) = ""
34                            Else
35                                Text(4) = Left(Replace(Trim(rsSearch("NamaLgn")), vbNewLine, ""), 60)
36                            End If
37                            Text(5) = Left(Trim(Nul(rsSearch("Alamat1"))) & " " & Trim(Nul(rsSearch("Alamat2"))), 60)
                              
38                                If Len(Trim(Nul(rsSearch("Alamat1"))) & " " & Trim(Nul(rsSearch("Alamat2")))) > 60 Then
39                                    Text(56) = Mid(Trim(Nul(rsSearch("Alamat1"))) & " " & Trim(Nul(rsSearch("Alamat2"))), 61, Len(Trim(Nul(rsSearch("Alamat1"))) & " " & Trim(Nul(rsSearch("Alamat2")))) - 60)
40                                Else
41                                    Text(56) = ""
42                                End If
                          
43                            If IsNull(rsSearch("Alamat3")) Then
44                                Text(6) = ""
45                            Else
46                                Text(6) = rsSearch("Alamat3")
47                            End If
48                            If IsNull(rsSearch("Telepon")) Then
49                                Text(8) = ""
50                                Text(9) = ""
51                                Text(10) = ""
52                            Else
                                  Dim nDpn1 As String
                                  Dim nTgh1 As String
                                  Dim nBlk1 As String
53                                nDpn1 = ""
54                                nTgh1 = ""
55                                nBlk1 = ""
56                                Call pecahNoTelp(Trim(rsSearch("Telepon")), nDpn1, nTgh1, nBlk1, ",")
57                                Text(8) = Trim(nDpn1)
58                                Text(9) = Trim(nTgh1)
59                                Text(10) = Trim(nBlk1)
60                            End If
61                            If IsNull(rsSearch("Npwp")) Then
62                                Text(11) = ""
63                            Else
64                                Text(11) = rsSearch("Npwp")
65                            End If
66                            If IsNull(rsSearch("KodeWil")) Then
67                                Text(3) = ""
68                            Else
69                                Text(3) = rsSearch("KodeWil")
70                            End If
71                        End If
72                    If Log.loyalityCustomer = 1 Then
73                        Set rsTemp = objO.FetchData("SELECT a.alamatemail as email FROM RPLNASAPP..MASTERUSER a INNER JOIN RPLNASAPP..NewCustomer b ON a.alamatemail = b.email WHERE a.verifikasi = 1 and b.idcust = '" & Trim(Text(2)) & "' ")
74                        If Not rsTemp.EOF Then
75                            Text(59).Enabled = False
76                            Text(59) = Nul(rsTemp!email)
77                        End If
                           
78                    End If
79                    Set rsSearch = objO.FetchDataACC("select KodeWil,NamaWil from Arwilayah WITH(NOLOCK) where KodeWil = '" & Left(Text(3), 3) & "'")
80                    If Not rsSearch.BOF Then
81                        Text(6) = rsSearch("NamaWil")
82                    End If
83                End If
84            Case 13
85                If Text(Index) <> "" Then
86                    Set rsSearch = objO.FetchData("select * from t_agama WITH(NOLOCK) where agama = '" & Text(Index) & "'")
87                    If rsSearch.BOF Then
88                        Cancel = True
89                        gMainKeyCari = Text(Index)
90                        gSearchOption = "AGAMA"
91                        frmCari.show (vbModal)
92                    End If
93                End If
94            Case 14
95                If Text(Index) <> "" Then
96                    Set rsSearch = objO.FetchData("select * from t_hobi WITH(NOLOCK) where fs_nm_hobi = '" & Text(Index) & "'")
97                    If rsSearch.BOF Then
98                        Cancel = True
99                        gMainKeyCari = Text(Index)
100                       gSearchOption = "HOBI"
101                       frmCari.show (vbModal)
102                   End If
103               End If
104           Case 15
105               If Text(Index) <> "" Then
106                   Set rsSearch = objO.FetchData("select * from t_dealer WITH(NOLOCK) where fs_kd_dealer = '" & Text(Index) & "' and FS_LOG_DEL = ''")
107                   If rsSearch.BOF Then
108                       Cancel = True
109                       gMainKeyCari = Text(Index)
110                       gBantuan = "ALL"
111                       gSearchOption = "DEALER"
112                       frmCari.show (vbModal)
113                   Else
114                       Text(16) = rsSearch("FS_NM_DEALER")
115                   End If
116               Else
117                   Text(16) = ""
118               End If
119           Case 18, 19, 49
120               If Text(Index) <> ">>->>->>>>" Then
121                   If Not CekTgl(Text(Index)) Then
122                       Text(Index) = ">>->>->>>>"
123                       Sendkeys ("+{TAB}")
124                       Exit Sub
125                   End If
126               End If
127           Case 21
128               If Text(Index) <> "" Then
129                   Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(Index) & "'")
130                   If rsSearch.BOF Then
131                       Cancel = True
132                       gMainKeyCari = Text(Index)
133                       gSearchOption = "MDL"
134                       frmCari.show (vbModal)
135                       gGroupModel = ""
136                   Else
137                       Text(21) = rsSearch("MDL")
138                       Text(23) = rsSearch("Desc")
139                       Text(22) = rsSearch("MODEL")
140                       gGroupModel = rsSearch("FS_GRP_MDL")
141                       If Text(1) <> Trim(rsSearch("NO_CHASSIS")) And Text(1) <> "" Then
142                           If MsgBox("No Rangka Berubah !" & vbCr & "Mengikuti perubahan ?", vbYesNo, "UBAH KODE RANGKA") = vbYes Then
143                               Text(1) = rsSearch("NO_CHASSIS")
144                           End If
145                       ElseIf Text(1) = "" Then
146                           Text(1) = Nul(rsSearch("NO_CHASSIS"))
147                       End If
148                   End If
149               End If
150           Case 37
151               If Text(Index) <> "" Then
152                   Set rsSearch = objO.FetchData("select * from t_asuransi WITH(NOLOCK) where asuransi = '" & Trim(Text(Index)) & "'")
153                   If rsSearch.BOF Then
154                       Cancel = True
155                       gMainKeyCari = Text(Index)
156                       gSearchOption = "ASURANSI"
157                       frmCari.show (vbModal)
158                       Text(38) = ""
159                       Text(39) = ""
160                   Else
161                       Text(38) = Trim(rsSearch("Alamat"))
162                       Text(39) = Trim(rsSearch("Kota"))
163                       Text(52) = Trim(rsSearch("KDPLG_ASR"))
164                       Text(53) = Trim(rsSearch("PSTN"))
165                       Text(54) = Trim(rsSearch("GSM"))
166                       Text(55) = Trim(rsSearch("CDMA"))
167                   End If
168               End If
169           Case 45
170               If Text(Index) <> "" Or Text(Index) <> "0" Then
171                   Set rsSearch = objO.FetchData("SELECT * FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Replace(Trim(Text(50)), ">", "") & Trim(Replace(Text(1), ">", "")) & Trim(Replace(Text(0), ">", "")) & "' ORDER BY KM")
172                   If Not rsSearch.EOF Then
173                       If CDbl(Text(Index)) < CDbl(rsSearch("KM")) Then
174                           MsgBox "KM tidak boleh lebih kecil dari KM Terakhir"
175                           Text(Index) = CDbl(rsSearch("KM"))
176                       End If
177                   End If
178               End If
179           Case 47
180               If Text(Index) <> "" Then
181                   nPol = Trim(Replace(Text(25), ">", "")) & Trim(Replace(Text(46), ">", "")) & Trim(Replace(Text(47), ">", ""))
182                   Set rsSearch = objO.FetchData("SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_POLISI = '" & Trim(nPol) & "' AND NO_CHASSIS <> '" & Replace(Trim(Text(50)), ">", "") & Trim(Replace(Text(1), ">", "")) & Trim(Replace(Text(0), ">", "")) & "'")
183                   If Not rsSearch.EOF Then
184                       MsgBox "No. Polisi digunakan unit kendaraan lain !!!" & vbCr & "Pastikan No. Polisi yang anda masukkan benar !!"
185                       Text(25) = ""
186                       Text(46) = ""
187                       Text(47) = ""
188                       Text(25).SetFocus
189                       Exit Sub
190                   End If
191               End If
192           Case 50
193               If Text(Index) <> "" Then
194                   If Log.RPLHistori = 0 Then
195                       Set rsSearch = objO.FetchData("SELECT * FROM T_COUNTRY WITH(NOLOCK) WHERE KODE = '" & Trim(Text(Index)) & "'")
196                   Else
197                       Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_COUNTRY WITH(NOLOCK) WHERE KODE = '" & Trim(Text(Index)) & "'")
198                   End If
199                   If rsSearch.EOF Then
200                      If MsgBox("Unit bukan Kendaraan Toyota ?", vbYesNo, "Konfirmasi No Rangka") = vbNo Then
201                           gMainKeyCari = Text(Index)
202                           gSearchOption = "COUNTRY"
203                           frmCari.show (vbModal)
204                      End If
205                   End If
206               End If
207       End Select

208      On Error GoTo 0
209      Exit Sub

text_Validate_Error:

210       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmtCustomer")
End Sub

Private Sub hapusForm()
          Dim i As Integer
          
1        On Error GoTo hapusForm_Error

2         For i = 0 To 59
3             Select Case i
                  Case 8, 9, 10, 34, 35, 36
4                     Text(i) = ">>>>>>>>>>>>>>"
5                 Case 25
6                     Text(i) = ""
7                     Text(i).Mask = "??"
8                 Case 46
9                     Text(i) = ""
10                    Text(i).Mask = "9999"
11                Case 47
12                    Text(i) = ""
13                    Text(i).Mask = "CCC"
14                Case 18, 19, 40, 41, 49
15                    n = Text(i).Mask
16                    Text(i).Mask = ""
17                    Text(i) = ""
18                    Text(i).Mask = n
19                Case 20, 28, 29, 42, 45
20                    Text(i) = 0
21                Case Else
22                    Text(i) = ""
23            End Select
24        Next

25       On Error GoTo 0
26       Exit Sub

hapusForm_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmtCustomer")
          
End Sub

Private Function cekHistoryHO(noRgk As String) As Boolean
      Dim rsHlp As ADODB.Recordset
          '  Dim rsTemp As ADODB.Recordset

1         On Error GoTo vbError
          
2     If Log.LogHOConn = 1 Then
3         If ConHO = True Then
4             Set rsHlp = objO.FetchDataHO("SELECT TOP 1 CUSTOMER,CUST_ADRES,KOTA, '' AS kodepos, TELPON, NO_POLISI, " & _
                                  "NO_ENGINE, '' AS wmi, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, FS_KD_DEALER, " & _
                                  "FS_NM_DEALER, DDATE, '' AS remark, CODE, NPWP , PPN, 0 AS DISC, '' AS keyno, " & _
                                  "front AS [USER], '' AS SPK, invDATE AS TGL_REP, JOB, '' AS JOBCD, km AS FS_KM_AKHIR, " & _
                                  "'' AS USEJOBDSC, invDATE AS RCDATE, '' AS KD_K, '' AS CR, '' AS CDK, " & _
                                  "'1900-01-01' AS TGL_D, '' ASCDD, '' AS CRDTCD, FS_TH_MOBIL, '' AS USEFS_NM_HOBI, " & _
                                  "FS_KD_PLG, FS_KD_WIL, '1900-01-01' AS dotam, CUSTOMER AS USENAME, CUST_ADRES AS USEADDR, " & _
                                  "KOTA AS USECITY, '' AS USEPOS, '' AS USEPSTN, '' AS USEGSM, '' AS USECDMA, '' AS USET4LAHIR, " & _
                                  "'1900-01-01' AS USETGLLAHIR, '' AS ASURANSI, '' AS ASR_ADDR, '' AS ASR_CITY, " & _
                                  "'1900-01-01' AS ASR_TGLM, '1900-01-01' AS ASR_TGLS,'0' AS  ASR_LIMIT, '' AS ASR_NOTE, " & _
                                  "'' AS NOTE, '' AS GSM, '' AS CDMA, '' AS USEAGAMA, FRONT AS USERADD, 0 AS FLAGKM, " & _
                                  "0 AS HEAVYUSE, '' AS pkp, '' AS KDPLG_ASR, '' AS USEEMAIL  " & _
                                  "from T_invoiceNew WITH(NOLOCK) WHERE TPSVC <> 'W' and THS = 0 and PDS = 0 and NO_CHASSIS = '" & noRgk & "' and Right(NOINV, 1) <> 'B' Order By InvDate Desc ")

5             If Not rsHlp.EOF Then
                  
6                 Text(2) = rsHlp("FS_KD_PLG")
7                 Text(4) = Left(rsHlp("CUSTOMER"), 60)
8                 Text(5) = Left(Trim(rsHlp("CUST_ADRES")), 60)
9                 If Len(Trim(rsHlp("CUST_ADRES"))) > 60 Then
10                    Text(56) = Mid(Nul(rsHlp("CUST_ADRES")), 61, Len(rsHlp("CUST_ADRES")) - 60)
11                End If
                  
12                Text(6) = rsHlp("KOTA")
13                Text(7) = rsHlp("KODEPOS")
14                Text(8) = rsHlp("TELPON")
15                Text(9) = rsHlp("GSM")
16                Text(10) = rsHlp("CDMA")
                  
17                Text(11) = rsHlp("NPWP")
18                Text(12) = rsHlp("USEJOBDSC")
19                Text(13) = rsHlp("USEAGAMA")
20                Text(14) = rsHlp("USEFS_NM_HOBI")
                  
                  '' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  '' Text(15) = rsHlp("FS_KD_DEALER")
                  '' Text(16) = rsHlp("FS_NM_DEALER")
                  '' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  
                  
                  ''-------------------------------------------------
                  ''2016-02-01 | diganti dengan nilai dari salesunit
                  ''-------------------------------------------------
                  
21                If Log.RPLHistori = 0 Then
22                    Set rsTemp = objO.FetchData("SELECT * FROM SALESUNIT WHERE NORANGKA = '" & Trim(noRgk) & "' ORDER BY tglentry desc")
23                Else
24                    Set rsTemp = objO.FetchData("select ISNULL(b.fs_kd_dealer,'') as FS_KD_DEALER,ISNULL(b.fs_nm_dealer,'') as FS_NM_DEALER from [" & dbReplikasiHO & "]..Summary a left join [" & dbReplikasiHO & "]..t_dealer b on a.KodeDealer = b.kodecabang where a.NoRangka = '" & Trim(noRgk) & "'")
25                End If
26                Text(15) = ""
27                Text(16) = ""
28                If Not rsTemp.EOF Then
29                    rsTemp.MoveFirst
30                    Text(15) = rsTemp("FS_KD_DEALER")
31                    Text(16) = rsTemp("FS_NM_DEALER")
32                End If
33                Text(17) = rsHlp("CODE")
34                Text(18) = TglDDMMYYYY(NulTgl(rsHlp("DDATE")))
35                Text(19) = TglDDMMYYYY(NulTgl(rsHlp("DOTAM")))
36                Text(20) = rsHlp("FS_TH_MOBIL")
37                Text(21) = rsHlp("MDL")
38                Text(22) = rsHlp("MODEL")
39                Text(23) = rsHlp("FMDL")
40                Text(24) = rsHlp("COLOR")
41                If Len(rsHlp("NO_POLISI")) > 0 Then
                      Dim nPol1 As String
42                    nPol1 = rsHlp("NO_POLISI")
                      Dim nDpn1 As String
                      Dim nTgh1 As String
                      Dim nBlk1 As String
43                    nDpn1 = ""
44                    nTgh1 = ""
45                    nBlk1 = ""
46                    Call pecahNoPol(nPol1, nDpn1, nTgh1, nBlk1)
47                    Text(25) = Trim(nDpn1)
48                    Text(46) = Trim(nTgh1)
49                    Text(47) = Trim(nBlk1)
50                Else
51                    Text(25) = ""
52                    Text(46) = ""
53                    Text(47) = ""
54                End If
                  
55                Text(26) = rsHlp("NO_ENGINE")
56                Text(27) = rsHlp("KEYNO")
57                Text(28) = rsHlp("PPN")
58                Text(29) = rsHlp("DISC")
59                Text(30) = rsHlp("USENAME")
60                Text(31) = Left(Nul(rsHlp("USEADDR")), 60)
61                Text(57) = ""
62                Text(58) = ""
63                If Len(Nul(rsHlp("USEADDR"))) > 60 Then
64                    Text(57) = Mid(Nul(rsHlp("USEADDR")), 61, Len(Nul(rsHlp("USEADDR"))) - 60)
65                    If Len(Nul(rsHlp("USEADDR"))) > 120 Then
66                        Text(58) = Mid(Nul(rsHlp("USEADDR")), 121, Len(Nul(rsHlp("USEADDR"))) - 120)
67                    End If
68                End If
69                Text(32) = rsHlp("USECITY")
70                Text(33) = rsHlp("USEPOS")
71                Text(34) = rsHlp("USEPSTN")
72                Text(35) = rsHlp("USEGSM")
73                Text(36) = rsHlp("USECDMA")
74                If rsHlp("ASURANSI") = "" Then
75                    chAsuransi.value = Unchecked
76                Else
77                    chAsuransi.value = Checked
78                End If
79                Text(37) = rsHlp("ASURANSI")
80                Text(38) = rsHlp("ASR_ADDR")
81                Text(39) = rsHlp("ASR_CITY")
82                Text(40) = TglDDMMYYYY(rsHlp("ASR_TGLM"))
83                Text(41) = TglDDMMYYYY(rsHlp("ASR_TGLS"))
84                Text(42) = rsHlp("ASR_LIMIT")
85                Text(43) = rsHlp("ASR_NOTE")
86                Text(44) = rsHlp("NOTE")
87                Text(45) = rsHlp("FS_KM_AKHIR")
88                Text(51) = rsHlp("PKP")
89                Text(59) = Nul(Trim(rsHlp("USEEMAIL")))
                  
                  ' INPUT HISTORY SERVICE
                  'Call panggilHistory(noRgk, nPol1)
90                cekHistoryHO = True
91            Else
92                cekHistoryHO = False
93            End If
94        End If
95    End If
96    Exit Function
vbError:
97        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekHistoryHO of Form frmtCustomer")
98        cekHistoryHO = False
99        Exit Function
End Function

Private Sub panggilHistory(noRgk As String, noPols As String)
1         On Error GoTo vbError
          Dim rsHist As ADODB.Recordset
              
          ' AMBIL HISTORY DARI HO
          Dim rsHlp2 As ADODB.Recordset
          Dim dBaseHO As String
          
2     If Log.LogHOConn = 1 Then
3         If ConHO = True Then
4             dBaseHO = HODb
5             Set rsHlp = objO.FetchDataHO("SELECT * FROM HISTORYnew WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRgk) & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC")
6             If Not rsHlp.EOF Then
7                 Do While Not rsHlp.EOF
8                     If rsHlp("NO_CHASSIS") <> "" Then
                          ' INPUT HISTORY
9                         Set rsHlp2 = objO.FetchData("SELECT * FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "' AND KM= '" & rsHlp("KM") & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC, NO_WO DESC")
10                        If rsHlp2.EOF Then
11                            Call objO.InsertData("INSERT INTO HISTORY VALUES ('" & rsHlp("CUSTOMER") & "', '" & rsHlp("CUST_ADRES") & "', '" & rsHlp("KOTA") & "','" & rsHlp("NO_POLISI") & "', '" & rsHlp("NO_ENGINE") & "','" & rsHlp("NO_CHASSIS") & "','" & rsHlp("FRONT") & "','" & rsHlp("MECH") & "','" & rsHlp("KM") & "','" & Format(rsHlp("DATE"), "mm/dd/yyyy") & "'," & rsHlp("PART_G") & ", " _
                                           & "" & rsHlp("PART_NG") & "," & rsHlp("OIL_G") & "," & rsHlp("OIL_NG") & "," & rsHlp("MTRL") & "," & rsHlp("LABOR") & "," & rsHlp("BODY") & "," & rsHlp("SBLT") & "," & rsHlp("MET") & "," & rsHlp("TOTAL") & ",'" & rsHlp("JOB") & "','" & rsHlp("REP1") & "','" & rsHlp("REP2") & "','" & rsHlp("REP3") & "','" & rsHlp("REP4") & "'," _
                                           & "'" & rsHlp("REP5") & "','" & rsHlp("REP6") & "','" & rsHlp("REP7") & "','" & rsHlp("MEKANIK1") & "','" & rsHlp("MEKANIK2") & "','" & rsHlp("MEKANIK3") & "'," _
                                           & "'" & rsHlp("MEKANIK4") & "','" & rsHlp("MEKANIK5") & "','" & rsHlp("MEKANIK6") & "','" & rsHlp("MEKANIK7") & "','" & rsHlp("MEKANIK8") & "','" & rsHlp("KLPKERJA1") & "','" & rsHlp("KLPKERJA2") & "','" & rsHlp("KLPKERJA3") & "','" & rsHlp("KLPKERJA4") & "','" & rsHlp("KLPKERJA5") & "','" & rsHlp("KLPKERJA6") & "'," _
                                           & "'" & rsHlp("KLPKERJA7") & "','" & rsHlp("KLPKERJA8") & "','" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "','" & rsHlp("NOFAK") & "','" & rsHlp("STATUS") & "','" & rsHlp("PAID") & "','" & rsHlp("DISC") & "','" & rsHlp("DISCT") & "','" & rsHlp("REMARK1") & "','" & rsHlp("FN_RETURN") & "','" & rsHlp("FS_RC") & "'," _
                                           & "'" & rsHlp("FS_KD_DEALER") & "','" & rsHlp("FS_NM_DEALER") & "','" & rsHlp("TIPEWO") & "','" & rsHlp("KODETIPE") & "', '" & rsHlp("SBE") & "','" & rsHlp("FS_KD_PLG") & "','" & rsHlp("FS_KD_WIL") & "','" & rsHlp("THS") & "','" & rsHlp("MOU") & "','" & rsHlp("PDS") & "','" & rsHlp("FS_LOG_DEL") & "','" & rsHlp("FD_TGL_DEL") & "', " _
                                           & "'" & rsHlp("ALASAN_DEL") & "', '','" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "')")
12                        End If
                          
                          ' INPUT PART
                          Dim stringAwal As String
                          Dim stringPotong1 As String
                          Dim stringPotong2 As String
                          Dim nNext As Integer
                          Dim hasil(15) As String
13                        If rsHlp("MEMOPART") <> "" Then
14                            Set rsHlp2 = objO.FetchData("SELECT * FROM T_PART WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "'")
15                            If rsHlp2.EOF Then
16                                stringAwal = Trim(rsHlp("MEMOPART"))
17                                Do While Len(stringAwal) > 0
18                                    stringPotong1 = Left(stringAwal, InStr(stringAwal, "|") - 1)
19                                    nNext = 0
20                                    Do While Len(stringPotong1) > 0
21                                        stringPotong2 = Left(stringPotong1, IIf(InStr(stringPotong1, "^") = 0, Len(stringPotong1), InStr(stringPotong1, "^") - 1))
22                                        hasil(nNext) = Trim(stringPotong2)
23                                        nNext = nNext + 1
24                                        stringPotong1 = Mid(stringPotong1, InStr(stringPotong1, "^") + 1, Len(stringPotong1) - Len(stringPotong2))
25                                    Loop
26                                    stringAwal = Mid(stringAwal, InStr(stringAwal, "|") + 1, Len(stringAwal) - Len(stringPotong1))
                                      
                                      ' SIMPAN
27                                    Call objO.InsertData("INSERT INTO T_PART VALUES ('" & Trim(hasil(0)) & "','" & Trim(hasil(1)) & "','" & Trim(hasil(2)) & "','" & Trim(hasil(3)) & "','" & Trim(hasil(4)) & "', '" & IIf(Trim(hasil(5)), 1, 0) & "','" & Trim(hasil(6)) & "', " _
                                                   & "'" & Trim(hasil(7)) & "','" & Trim(hasil(8)) & "','" & Trim(hasil(9)) & "')")
28                                Loop
29                            End If
30                        End If
                          
                          ' INPUT SUBLET
31                        If rsHlp("MEMOSUBLET") <> "" Then
32                            Set rsHlp2 = objO.FetchData("SELECT * FROM HIST_SUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsHlp("NO_WO")) & "' AND NO_CHASSIS = '" & noRgk & "'")
33                            If rsHlp2.EOF Then
34                                stringAwal = Trim(rsHlp("MEMOSUBLET"))
35                                Do While Len(stringAwal) > 0
36                                    stringPotong1 = Left(stringAwal, InStr(stringAwal, "|") - 1)
37                                    nNext = 0
38                                    Do While Len(stringPotong1) > 0
39                                        stringPotong2 = Left(stringPotong1, IIf(InStr(stringPotong1, "^") = 0, Len(stringPotong1), InStr(stringPotong1, "^") - 1))
40                                        hasil(nNext) = Trim(stringPotong2)
41                                        nNext = nNext + 1
42                                        stringPotong1 = Mid(stringPotong1, InStr(stringPotong1, "^") + 1, Len(stringPotong1) - Len(stringPotong2))
43                                    Loop
44                                    stringAwal = Mid(stringAwal, InStr(stringAwal, "|") + 1, Len(stringAwal) - Len(stringPotong1))
                                      
                                      ' SIMPAN
45                                    Call objO.InsertData("INSERT INTO HIST_SUBLET VALUES ('" & Trim(hasil(0)) & "','" & Trim(hasil(1)) & "','" & Trim(hasil(2)) & "','" & Trim(hasil(3)) & "','" & Trim(hasil(4)) & "', '" & Trim(hasil(5)) & "','" & Trim(hasil(6)) & "', " _
                                                   & "'" & Trim(hasil(7)) & "','" & Trim(hasil(8)) & "','" & Trim(hasil(9)) & "','" & Trim(hasil(10)) & "','" & Trim(hasil(11)) & "','" & Trim(hasil(12)) & "','" & Trim(hasil(13)) & "','" & Trim(hasil(14)) & "')")
46                                Loop
47                            End If
48                        End If
49                    End If
50                    rsHlp.MoveNext
51                Loop
52            End If
53        End If
54    End If
55    Exit Sub
vbError:
56    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilHistory of Form frmtCustomer")
57        Exit Sub
End Sub

Private Function cekBengkelLama(noRgk As String) As Boolean
      Dim rsHlp As ADODB.Recordset
1        On Error GoTo cekBengkelLama_Error

2             Set rsHlp = objO.FetchDataAAS("SELECT * FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & noRgk & "'")
3             If Not rsHlp.EOF Then
                  
4                 Text(2) = rsHlp("FS_KD_PLG")
5                 Text(4) = Left(rsHlp("CUSTOMER"), 60)
6                 Text(5) = Left(rsHlp("CUST_ADRES"), 60)
7                 If Len(rsHlp("CUST_ADRES")) > 60 Then
8                     Text(56) = Mid(Nul(rsHlp("CUST_ADRES")), 61, Len(rsHlp("CUST_ADRES")) - 60)
9                 End If
                  
10                Text(6) = rsHlp("KOTA")
11                Text(7) = rsHlp("KODEPOS")
12                Text(8) = rsHlp("TELPON")
13                Text(9) = ""
14                Text(10) = ""
                  
15                Text(11) = rsHlp("NPWP")
16                Text(12) = ""
17                Text(13) = ""
18                Text(14) = rsHlp("FS_NM_HOBI")
19                Text(15) = rsHlp("FS_KD_DEALER")
20                Text(16) = rsHlp("FS_NM_DEALER")
21                Text(17) = rsHlp("CODE")
22                Text(18) = TglDDMMYYYY(NulTgl(Replace(Replace(rsHlp("DDATE"), ">", "0"), " ", 0)))
23                Text(19) = "01-01-1900"
24                Text(20) = rsHlp("FS_TH_MOBIL")
25                Text(21) = rsHlp("MDL")
26                Text(22) = rsHlp("MODEL")
27                Text(23) = rsHlp("FMDL")
28                Text(24) = rsHlp("COLOR")
29                If Len(rsHlp("NO_POLISI")) > 0 Then
                      Dim nPol1 As String
30                    nPol1 = rsHlp("NO_POLISI")
                      Dim nDpn1 As String
                      Dim nTgh1 As String
                      Dim nBlk1 As String
31                    nDpn1 = ""
32                    nTgh1 = ""
33                    nBlk1 = ""
34                    Call pecahNoPol(nPol1, nDpn1, nTgh1, nBlk1)
35                    Text(25) = Trim(nDpn1)
36                    Text(46) = Trim(nTgh1)
37                    Text(47) = Trim(nBlk1)
38                Else
39                    Text(25) = ""
40                    Text(46) = ""
41                    Text(47) = ""
42                End If
                  
43                Text(26) = rsHlp("NO_ENGINE")
44                Text(27) = rsHlp("KEYNO")
45                Text(28) = rsHlp("PPN")
46                Text(29) = rsHlp("DISC")
47                Text(30) = rsHlp("CUSTOMER")
48                Text(31) = Left(Nul(rsHlp("CUST_ADRES")), 60)
49                Text(57) = ""
50                Text(58) = ""
51                If Len(Nul(rsHlp("CUST_ADRES"))) > 60 Then
52                    Text(57) = Mid(Nul(rsHlp("CUST_ADRES")), 61, Len(Nul(rsHlp("CUST_ADRES"))) - 60)
53                    If Len(Nul(rsHlp("CUST_ADRES"))) > 120 Then
54                        Text(58) = Mid(Nul(rsHlp("CUST_ADRES")), 121, Len(Nul(rsHlp("CUST_ADRES"))) - 120)
55                    End If
56                End If
                  
57                Text(32) = rsHlp("KOTA")
58                Text(33) = rsHlp("KODEPOS")
59                Text(34) = rsHlp("TELPON")
60                Text(35) = ""
61                Text(36) = ""
62                chAsuransi.value = Unchecked
63                Text(37) = ""
64                Text(38) = ""
65                Text(39) = ""
66                Text(40) = "01-01-1900"
67                Text(41) = "01-01-1900"
68                Text(42) = "0"
69                Text(43) = ""
70                Text(44) = ""
71                Text(45) = rsHlp("FS_KM_AKHIR")
72                Text(51) = ""
                  
73                cekBengkelLama = True
74            Else
75                cekBengkelLama = False
76            End If

77       On Error GoTo 0
78       Exit Function

cekBengkelLama_Error:

79        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekBengkelLama of Form frmtCustomer")
End Function

Private Function cekSPK(noRangka As String) As Boolean
      Dim rsHlp As ADODB.Recordset
      Dim rsHlp2 As ADODB.Recordset
1        On Error GoTo cekSPK_Error

2             Set rsHlp = objO.FetchDataSPK("select (case when a.alm1nama=1 then a.nama_faktur+' '+a.alamat1_faktur else a.nama_faktur end) as NamaFaktur, " _
                                         & "(case when a.alm1nama=1 then a.alamat2_faktur+' '+a.alamat3_faktur+' '+a.alamat4_faktur+' '+a.alamat5_faktur else a.alamat1_faktur+' '+a.alamat2_faktur+' '+a.alamat3_faktur+' '+a.alamat4_faktur+' '+a.alamat5_faktur end) as AlamatFaktur, " _
                                         & "(case when a.kotamadya='LAINNYA' then a.wil_lain else a.kotamadya end) as Kota, " _
                                         & "b.kodepos,b.telp ,b.gsmfaktur,b.cdmafaktur,b.agama,b.hoby,ISNULL(b.tglcr1,'01/01/3000') AS TGLCR1,ISNULL(c.tgldo,'01/01/3000') AS TGLDO,ISNULL(b.tahun,0) AS TAHUN,c.kodekategori,d.namakategory,e.namaind,b.nopol,c.nomesin,c.[key], " _
                                         & "a.snama , a.salamat, a.skota, a.skodepos, a.stelp, a.sgsm, a.scdma, a.kode_lgn from spk a WITH(NOLOCK) left join afi b WITH(NOLOCK) on a.spk_no=b.spk_no " _
                                         & "left join inmst c WITH(NOLOCK) on a.norangka=c.kodeitem left join inkategory d WITH(NOLOCK) on c.kodekategori=d.kodekategory left join inwarnaext e WITH(NOLOCK) on c.kodewarna=e.kodewarna " _
                                         & "where a.norangka = '" & noRangka & "'")
3             If Not rsHlp.EOF Then
              
4                 Text(2) = Nul(rsHlp("KODE_LGN"))
5                 Text(3) = Left(Trim(Nul(rsHlp("KODE_LGN"))), 3)
6                 Text(4) = Left(Replace(Nul(rsHlp("NAMAFAKTUR")), vbNewLine, ""), 60)
7                 Text(5) = Left(Replace(Nul(Trim(rsHlp("ALAMATFAKTUR"))), vbNewLine, ""), 60)
8                 If Len(Replace(Trim(Nul(rsHlp("ALAMATFAKTUR"))), vbNewLine, "")) > 60 Then
9                     Text(56) = Mid(Replace(Nul(Trim(rsHlp("ALAMATFAKTUR"))), vbNewLine, ""), 61, Len(Replace(Nul(Trim(rsHlp("ALAMATFAKTUR"))), vbNewLine, "")) - 60)
10                End If
11                Text(6) = Nul(rsHlp("KOTA"))
12                Text(7) = Nul(rsHlp("KODEPOS"))
13                Text(8) = Nul(rsHlp("TELP"))
14                Text(9) = Nul(rsHlp("GSMFAKTUR"))
15                Text(10) = Nul(rsHlp("CDMAFAKTUR"))
                  
16                Text(15) = gKdCbg
17                Text(16) = gNmCbg
              
18                Text(18) = TglDDMMYYYY(NulTgl(rsHlp("TGLCR1")))
19                Text(19) = TglDDMMYYYY(NulTgl(rsHlp("TGLDO")))
20                Text(20) = CDbl(IIf(rsHlp("TAHUN") = "", 0, rsHlp("TAHUN")))
21                Text(21) = Nul(rsHlp("KODEKATEGORI"))
22                Text(22) = objO.FetchData_String("SELECT MODEL FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Nul(rsHlp("KODEKATEGORI")) & "'")
23                Text(23) = Nul(rsHlp("NAMAKATEGORY"))
24                Text(24) = Nul(rsHlp("NAMAIND"))
25                nPol = Trim(Nul(rsHlp("NOPOL")))
                  Dim nDpn1 As String
                  Dim nTgh1 As String
                  Dim nBlk1 As String
26                nDpn1 = ""
27                nTgh1 = ""
28                nBlk1 = ""
29                If nPol <> "" Then
30                    Call pecahNoPol(Trim(nPol), nDpn1, nTgh1, nBlk1)
31                    Text(25) = Trim(nDpn1)
32                    Text(46) = Trim(nTgh1)
33                    Text(47) = Trim(nBlk1)
34                End If
              
35                Text(26) = Nul(rsHlp("NOMESIN"))
36                Text(27) = Nul(rsHlp("KEY"))
                  
37                Text(30) = Left(Replace(Nul(rsHlp("SNAMA")), vbNewLine, ""), 60)
38                Text(31) = Left(Replace(Nul(Trim(rsHlp("SALAMAT"))), vbNewLine, ""), 60)
39                Text(57) = ""
40                Text(58) = ""
41                If Len(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, "")) > 60 Then
42                    Text(57) = Mid(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, ""), 61, Len(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, "")) - 60)
43                    If Len(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, "")) > 120 Then
44                        Text(58) = Mid(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, ""), 121, Len(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, "")) - 120)
45                    End If
46                End If
                  
47                Text(32) = Nul(rsHlp("SKOTA"))
48                Text(33) = Nul(rsHlp("SKODEPOS"))
49                Text(34) = Nul(rsHlp("STELP"))
50                Text(35) = Nul(rsHlp("SGSM"))
51                Text(36) = Nul(rsHlp("SCDMA"))
                  
52                sNamaSpk = Nul(rsHlp("SNAMA"))
53                sAlamatSpk = Nul(rsHlp("SALAMAT"))
54                sKotaSpk = Nul(rsHlp("SKOTA"))
55                sKodeposSpk = Nul(rsHlp("SKODEPOS"))
56                sTelpSpk = Nul(rsHlp("STELP"))
57                sGsmSpk = Nul(rsHlp("SGSM"))
58                sCdmaSpk = Nul(rsHlp("SCDMA"))
                  
59                chAsuransi.value = Unchecked
60                cekSPK = True
61            Else
62                cekSPK = False
63            End If

64       On Error GoTo 0
65       Exit Function

cekSPK_Error:

66        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekSPK of Form frmtCustomer")
End Function

Private Function cekSPKHO(noRangka As String) As Boolean
1        On Error GoTo cekSPKHO_Error

2     If Log.LogHOConn = 1 Then
      Dim rsHlp As ADODB.Recordset
3             Set rsHlp = objO.FetchDataSPKHO("SELECT NAMAFAKTUR, ALAMATFAKTUR, KOTA, KODEPOS, TELP, GSMFAKTUR, CDMAFAKTUR, AGAMA, HOBY, TGLCR1, TGLDO, TAHUN, " _
                                           & "a.KODEKATEGORI , NAMAKATEGORY, NAMAIND, noPol, A.NOMESIN, [KEYNO], SNAMA, SALAMAT, SKOTA, SKODEPOS, STELP, SGSM, SCDMA, a.dealer " _
                                           & "FROM SPK99..DATACUST A WITH(NOLOCK) INNER JOIN SPK00..INKATEGORY B WITH(NOLOCK) ON A.KODEKATEGORI = B.KODEKATEGORY " _
                                           & "INNER JOIN SPK00..INWARNAEXT C WITH(NOLOCK) ON A.KODEWARNA = C.KODEWARNA INNER JOIN SPK00..INMST D WITH(NOLOCK) ON A.NORANGKA = D.KODEITEM " _
                                           & "WHERE NORANGKA = '" & noRangka & "'")
4             If Not rsHlp.EOF Then
5                 Text(4) = Left(Nul(rsHlp("NAMAFAKTUR")), 60)
6                 Text(5) = Nul(Left(Replace(Trim(Nul(rsHlp("ALAMATFAKTUR"))), vbNewLine, ""), 60))
7                 If Len(Replace(Trim(Nul(rsHlp("ALAMATFAKTUR"))), vbNewLine, "")) > 60 Then
8                     Text(56) = Mid(Replace(Nul(Trim(rsHlp("ALAMATFAKTUR"))), vbNewLine, ""), 61, Len(Replace(Trim(Nul(rsHlp("ALAMATFAKTUR"))), vbNewLine, "")) - 60)
9                 End If
10                Text(6) = Nul(rsHlp("KOTA"))
11                Text(7) = Nul(rsHlp("KODEPOS"))
12                Text(8) = Nul(rsHlp("TELP"))
13                Text(9) = Nul(rsHlp("GSMFAKTUR"))
14                Text(10) = Nul(rsHlp("CDMAFAKTUR"))
15                Text(15) = cariDealer(Trim(rsHlp("DEALER")), "kode")
16                Text(16) = cariDealer(Trim(rsHlp("DEALER")), "nama")
17                Text(18) = TglDDMMYYYY(NulTgl(rsHlp("TGLCR1")))
18                Text(19) = TglDDMMYYYY(NulTgl(rsHlp("TGLDO")))
19                Text(20) = CDbl(NulA(rsHlp("TAHUN")))
20                Text(21) = Nul(rsHlp("KODEKATEGORI"))
21                Text(22) = objO.FetchData_String("SELECT MODEL FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Nul(rsHlp("KODEKATEGORI")) & "'")
22                Text(23) = Nul(rsHlp("NAMAKATEGORY"))
23                Text(24) = Nul(rsHlp("NAMAIND"))
24                nPol = Trim(Nul(rsHlp("NOPOL")))
                  Dim nDpn1 As String
                  Dim nTgh1 As String
                  Dim nBlk1 As String
25                nDpn1 = ""
26                nTgh1 = ""
27                nBlk1 = ""
28                If nPol <> "" Then
29                    Call pecahNoPol(Trim(nPol), nDpn1, nTgh1, nBlk1)
30                    Text(25) = Trim(nDpn1)
31                    Text(46) = Trim(nTgh1)
32                    Text(47) = Trim(nBlk1)
33                End If
                              
34                Text(26) = Nul(rsHlp("NOMESIN"))
35                Text(27) = Nul(rsHlp("KEYNO"))
36                Text(30) = Nul(rsHlp("SNAMA"))
37                Text(31) = Left(Replace(Nul(Trim(rsHlp("SALAMAT"))), vbNewLine, ""), 60)
38                Text(57) = ""
39                Text(58) = ""
40                If Len(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, "")) > 60 Then
41                    Text(57) = Mid(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, ""), 61, Len(Replace(Nul(rsHlp("SALAMAT")), vbNewLine, "")) - 60)
42                    If Len(Replace(Nul(Trim(rsHlp("SALAMAT"))), vbNewLine, "")) > 120 Then
43                        Text(58) = Mid(Replace(Nul(Trim(rsHlp("SALAMAT"))), vbNewLine, ""), 121, Len(Replace(Nul(Trim(rsHlp("SALAMAT"))), vbNewLine, "")) - 120)
44                    End If
45                End If
46                Text(32) = Nul(rsHlp("SKOTA"))
47                Text(33) = Nul(rsHlp("SKODEPOS"))
48                Text(34) = Nul(rsHlp("STELP"))
49                Text(35) = Nul(rsHlp("SGSM"))
50                Text(36) = Nul(rsHlp("SCDMA"))
51                chAsuransi.value = Unchecked
52                cekSPKHO = True
53            Else
54                cekSPKHO = False
55            End If
56        End If

57       On Error GoTo 0
58       Exit Function

cekSPKHO_Error:

59        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekSPKHO of Form frmtCustomer")
End Function

Private Function cariDealer(IniDealer As String, tipe As String) As String
1        On Error GoTo cariDealer_Error

2         Select Case tipe
              Case "kode"
3                 cariDealer = objO.FetchData_String("SELECT FS_KD_DEALER FROM T_DEALER WITH(NOLOCK) WHERE KODECABANG = '" & Trim(IniDealer) & "'")
4             Case "nama"
5                 cariDealer = objO.FetchData_String("SELECT FS_NM_DEALER FROM T_DEALER WITH(NOLOCK) WHERE KODECABANG = '" & Trim(IniDealer) & "'")
6         End Select

7        On Error GoTo 0
8        Exit Function

cariDealer_Error:

9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariDealer of Form frmtCustomer")
End Function

Private Function cekHistoryOnlyPDS(noChs As String) As Boolean
1        On Error GoTo cekHistoryOnlyPDS_Error

2         cekHistoryOnlyPDS = False
3         Set rsSearch = objO.FetchData("SELECT PDS FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noChs) & "'")
4         If Not rsSearch.EOF Then
5             If rsSearch.RecordCount = 1 Then
6                 If Trim(rsSearch("PDS")) = "1" Then
7                     cekHistoryOnlyPDS = True
8                 End If
9             End If
10        End If

11       On Error GoTo 0
12       Exit Function

cekHistoryOnlyPDS_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekHistoryOnlyPDS of Form frmtCustomer")
End Function

Private Function cekHistoryPDS(noChs As String) As Boolean
1        On Error GoTo cekHistoryPDS_Error

2         cekHistoryPDS = objO.FetchData_Boolean("SELECT NO_CHASSIS FROM HISTORY WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(noChs) & "' AND PDS = '1'")

3        On Error GoTo 0
4        Exit Function

cekHistoryPDS_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekHistoryPDS of Form frmtCustomer")
End Function

Private Sub PanggilCustomer(Optional str As String)
      Dim nPol As String
      Dim nChassis As String
      Dim n As String
      Dim newFeed As Boolean
      Dim rsCust As ADODB.Recordset

1        On Error GoTo PanggilCustomer_Error

2         newFeed = False
3         If str <> "" Then
4             sSQL = "Select * from CUSTDAT WITH(NOLOCK) where NO_CHASSIS ='" & str & "'"
5         Else
6             sSQL = "Select * from CUSTDAT WITH(NOLOCK) where NO_CHASSIS ='" & Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "") & "'"
7             str = Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")
8         End If
9         Set rsCust = objO.FetchData(sSQL)
10        If rsCust.BOF Then
11            If cekSPK(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
12                If cekHistoryHO(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
13                    If cekSPKHO(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
                         ' If cekBengkelLama(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) = False Then
                         ' End If
14                    End If
15                End If
16            End If
              
17        Else
18            pds = cekHistoryPDS(rsCust("NO_CHASSIS"))
19            If cekHistoryOnlyPDS(rsCust("NO_CHASSIS")) Then
20                If Not cekSPK(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))) Then
21                    n = cekSPKHO(IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", "")))
22                Else
23                    GoTo LANJUTAN
24                End If
25            Else
26                nChassis = IIf(str <> "", str, Replace(Trim(Text(50)), ">", "") & Replace(Replace(Text(1), ">", ""), " ", "") & Replace(Replace(Text(0), ">", ""), " ", ""))
27                If Nul(rsCust("FS_KD_PLG")) <> Trim(Text(2)) And Trim(Text(2)) <> "" Then
28                    If MsgBox("Data Pelanggan Berbeda. Pakai data Baru ?", vbYesNo, "KODE PELANGGAN") = vbYes Then
29                        newFeed = True
30                        Exit Sub
31                    End If
32                End If
                  
LANJUTAN:
33                If Not newFeed Then
34                    Text(2) = Nul(rsCust("FS_KD_PLG"))
35                    Text(4) = Left(Nul(rsCust("CUSTOMER")), 60)
36                    Text(5) = Left(Nul(rsCust("CUST_ADRES")), 60)
37                    If Len(Trim(Nul(rsCust("CUST_ADRES")))) > 60 Then
38                        Text(56) = Mid(Nul(rsCust("CUST_ADRES")), 61, Len(Nul(rsCust("CUST_ADRES"))) - 60)
39                    End If
40                    Text(6) = Nul(rsCust("KOTA"))
41                    Text(7) = Nul(rsCust("KODEPOS"))
42                    Text(8) = Nul(rsCust("TELPON"))
43                    Text(9) = Nul(rsCust("GSM"))
44                    Text(10) = Nul(rsCust("CDMA"))
                      
45                    Text(11) = Nul(rsCust("NPWP"))
46                    Text(51) = Nul(rsCust("PKP"))
47                    Text(12) = Nul(rsCust("USEJOBDSC"))
48                    Text(13) = Nul(Trim(rsCust("USEAGAMA")))
49                    Text(14) = Nul(rsCust("USEFS_NM_HOBI"))
                      
50                    Text(28) = Nul(rsCust("PPN"))
51                    Text(29) = Nul(rsCust("DISC"))
                      
                      '================= DATA MRS =================
52                    Text(30) = Nul(rsCust("USENAME"))
53                    Text(31) = Left(Trim(Nul(rsCust("USEADDR"))), 60)
54                    If Len(Trim(Nul(rsCust("USEADDR")))) > 60 Then
55                        Text(57) = Mid(Trim(Nul(rsCust("USEADDR"))), 61, Len(Trim(Nul(rsCust("USEADDR")))) - 60)
56                        If Len(Trim(Nul(rsCust("USEADDR")))) > 120 Then
57                            Text(58) = Mid(Trim(Nul(rsCust("USEADDR"))), 121, Len(Trim(Nul(rsCust("USEADDR")))) - 120)
58                        End If
59                    End If
60                    Text(32) = Nul(rsCust("USECITY"))
61                    Text(33) = Nul(rsCust("USEPOS"))
62                    Text(34) = Nul(rsCust("USEPSTN"))
63                    Text(35) = Nul(rsCust("USEGSM"))
64                    Text(36) = Nul(rsCust("USECDMA"))
65                    Text(48) = Nul(rsCust("USET4LAHIR"))
66                    n = Text(49).Mask
67                    Text(49).Mask = ""
68                    Text(49) = Nul(IIf(Replace(rsCust("USETGLLAHIR"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("USETGLLAHIR"), "/", "-") = "01-01-1900", "", Replace(rsCust("USETGLLAHIR"), "/", "-"))))
69                    Text(49).Mask = n
70                    Text(59) = Nul(Trim(rsCust("USEEMAIL")))
                      
                      '================= DATA ASURANSI =================
71                    If Trim(rsCust("ASURANSI")) = "" Then
72                        chAsuransi.value = Unchecked
73                    Else
74                        chAsuransi.value = Checked
75                    End If
76                    Text(52) = Nul(rsCust("KDPLG_ASR"))
77                    Text(37) = Nul(rsCust("ASURANSI"))
78                    Text(38) = Nul(rsCust("ASR_ADDR"))
79                    Text(39) = Nul(rsCust("ASR_CITY"))
80                    n = Text(40).Mask
81                    Text(40).Mask = ""
82                    Text(40) = Nul(IIf(Replace(rsCust("ASR_TGLM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("ASR_TGLM"), "/", "-") = "01-01-1900", "", Replace(rsCust("ASR_TGLM"), "/", "-"))))
83                    Text(40).Mask = n
84                    n = Text(41).Mask
85                    Text(41).Mask = ""
86                    Text(41) = Nul(IIf(Replace(rsCust("ASR_TGLS"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("ASR_TGLS"), "/", "-") = "01-01-1900", "", Replace(rsCust("ASR_TGLS"), "/", "-"))))
87                    Text(41).Mask = n
                      
                      Dim rsAsr As ADODB.Recordset
88                    Set rsAsr = objO.FetchData("SELECT * FROM T_ASURANSI WITH(NOLOCK) WHERE ASURANSI = '" & Trim(Text(37)) & "'")
89                    If Not rsAsr.EOF Then
90                        Text(53) = Nul(rsAsr("PSTN"))
91                        Text(54) = Nul(rsAsr("GSM"))
92                        Text(55) = Nul(rsAsr("CDMA"))
93                    End If
94                    Text(42) = Nul(rsCust("ASR_LIMIT"))
95                    Text(43) = Nul(rsCust("ASR_NOTE"))
96                    Text(44) = Nul(rsCust("NOTE"))
97                End If
                  '================= DATA KENDARAAN =================
98                Text(45) = Nul(rsCust("FS_KM_AKHIR"))
                  ''--------------------------------------
                  ''2016-02-01 | ambil data dari salesunit
                  ''--------------------------------------
99                 Text(15) = Nul(rsCust("FS_KD_DEALER"))
100                Text(16) = Nul(rsCust("FS_NM_DEALER"))
                  ''---------------------------------------
                  
      '            Text(15) = ""
      '            Text(16) = ""
                  '' 2016-04-11 || jika fs_kd_dealer kosong maka dicari dari salesunit
101               If Nul(rsCust!FS_KD_DEALER) = "" Then
                      
                     ''2016-07-29 "" SUMMARY MENYIMPAN SEMUA DATA PENJUALAN SEKALIPUN ada RETUR SEHINGGA YANG DIAKUI PENJUALAN TERAKHIR
102                  If Log.RPLHistori = 0 Then
103                       Set rsTemp = objO.FetchData("SELECT FS_KD_DEALER,FS_NM_DEALER FROM SALESUNIT WHERE FS_KD_DEALER <> '' AND NORANGKA = '" & Trim(str) & "' ORDER BY TGLENTRY DESC ")
104                  Else
105                       Set rsTemp = objO.FetchData("select ISNULL(b.fs_kd_dealer,'') as FS_KD_DEALER,ISNULL(b.fs_nm_dealer,'') as FS_NM_DEALER from [" & dbReplikasiHO & "]..Summary a left join [" & dbReplikasiHO & "]..t_dealer b on a.KodeDealer = b.kodecabang where isnull(b.FS_KD_DEALER,'') <> '' AND a.NoRangka = '" & Trim(str) & "'")
106                  End If
107                  If Not rsTemp.EOF Then
108                      Text(15) = Nul(rsTemp!FS_KD_DEALER)
109                      Text(16) = Nul(rsTemp("FS_NM_DEALER"))
110                  End If
111               End If
                              
112               Text(17) = Nul(rsCust("CODE"))
113               Text(18).Mask = ""
114               Text(18) = IIf(TglDDMMYYYY(NulTgl(rsCust("DDATE"))) = "01-01-1900" Or TglDDMMYYYY(NulTgl(rsCust("DDATE"))) = "01-01-3000", "", TglDDMMYYYY(NulTgl(rsCust("DDATE")))) 'Nul(IIf(Replace(rsCust("DDATE"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("DDATE"), "/", "-") = "01-01-1900", "", Replace(rsCust("DDATE"), "/", "-"))))
115               Text(18).Mask = "99-99-9999"
116               Text(19).Mask = ""
117               Text(19) = IIf(TglDDMMYYYY(NulTgl(rsCust("DOTAM"))) = "01-01-1900" Or TglDDMMYYYY(NulTgl(rsCust("DOTAM"))) = "01-01-3000", "", TglDDMMYYYY(NulTgl(rsCust("DOTAM")))) 'Nul(IIf(Replace(rsCust("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsCust("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsCust("DOTAM"), "/", "-"))))
118               Text(19).Mask = "99-99-9999"
119               Text(20) = Nul(rsCust("FS_TH_MOBIL"))
120               Text(21) = Nul(rsCust("MDL"))
121               Text(22) = Nul(rsCust("MODEL"))
122               Text(23) = Nul(rsCust("FMDL"))
123               Text(24) = Nul(rsCust("COLOR"))
124               If Len(Nul(rsCust("NO_POLISI"))) > 0 Then
125                   nPol = Nul(rsCust("NO_POLISI"))
                      Dim nDpn1 As String
                      Dim nTgh1 As String
                      Dim nBlk1 As String
126                   nDpn1 = ""
127                   nTgh1 = ""
128                   nBlk1 = ""
129                   Call pecahNoPol(Trim(nPol), nDpn1, nTgh1, nBlk1)
130                   Text(25) = Trim(nDpn1)
131                   Text(46) = Trim(nTgh1)
132                   Text(47) = Trim(nBlk1)
133               Else
134                   Text(25) = ""
135                   Text(46) = ""
136                   Text(47) = ""
137               End If
138               Text(26) = Nul(rsCust("NO_ENGINE"))
139               Text(27) = Nul(rsCust("KEYNO"))
                  
140               If Log.loyalityCustomer = 1 Then
141                   Set rsTemp = objO.FetchData("SELECT * FROM [RplNasApp]..masteruser WHERE verifikasi = 1 and alamatemail = '" & Trim(Text(59)) & "' ")
142                   If Not rsTemp.EOF Then
143                       Text(59).Enabled = False
144                   End If
                      
145                   Set rsTemp = objO.FetchData("SELECT b.* FROM RPLNASAPP..MASTERUSER a INNER JOIN RPLNASAPP..NewCustomer b ON a.alamatemail = b.email WHERE a.VERIFIKASI = 1 and isnull(b.isUpdate,0) = 1 AND a.ALAMATEMAIL = '" & Trim(Text(59)) & "' ")
146                   If Not rsTemp.EOF Then
147                       Text(2).Enabled = False
148                       Text(4) = UCase(Trim(Nul(rsTemp!namacust1) & " " & Nul(rsTemp!namacust2) & " " & Nul(rsTemp!namacust3)))
149                       If Len(Trim(Nul(rsTemp("alamat")) & " RT." & Nul(rsTemp("rt")) & " RW." & Nul(rsTemp("rw")) & " RT." & Nul(rsTemp("rt")) & " Kelurahan." & Nul(rsTemp("kelurahan")))) > 60 Then
150                           Text(5) = UCase(Left(Trim(Nul(rsTemp("alamat")) & " RT." & Nul(rsTemp("rt")) & " RW." & Nul(rsTemp("rw")) & " RT." & Nul(rsTemp("rt")) & " Kelurahan." & Nul(rsTemp("kelurahan"))), 60))
151                           Text(56) = UCase(Mid(Trim(Nul(rsTemp("alamat")) & " RT." & Nul(rsTemp("rt")) & " RW." & Nul(rsTemp("rw")) & " RT." & Nul(rsTemp("rt")) & " Kelurahan." & Nul(rsTemp("kelurahan"))), 61, 60))
152                       Else
153                           Text(5) = UCase(Left(Trim(Nul(rsTemp("alamat")) & " RT." & Nul(rsTemp("rt")) & " RW." & Nul(rsTemp("rw")) & " RT." & Nul(rsTemp("rt")) & " Kelurahan." & Nul(rsTemp("kelurahan"))), 60))
154                       End If
155                       Text(60) = UCase(Nul(rsTemp!kecamatan))
156                       Text(6) = UCase(Nul(rsTemp!Kota))
157                       Text(7) = UCase(Nul(rsTemp!KodePos))
158                       Text(14) = UCase(Nul(rsTemp!hobi))
159                       Text(13) = UCase(Nul(rsTemp!Agama))
160                       Text(12) = UCase(Nul(rsTemp!pekerjaan))
161                       Text(8) = UCase(Nul(rsTemp!telp))
162                       Text(9) = UCase(Nul(rsTemp!hp1))
163                       Text(10) = UCase(Nul(rsTemp!hp2))
164                   End If
165               End If
                  
166           End If
167       End If

168      On Error GoTo 0
169      Exit Sub

PanggilCustomer_Error:

170       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PanggilCustomer of Form frmtCustomer")
End Sub

