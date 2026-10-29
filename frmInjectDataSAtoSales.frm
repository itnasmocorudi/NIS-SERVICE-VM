VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmInjectDataSAtoSales 
   Caption         =   "Inject Data SA to Sales"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   13890
   Begin VB.CheckBox ckReferensi 
      Caption         =   "Referensi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   68
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   480
      TabIndex        =   67
      Top             =   4440
      Visible         =   0   'False
      Width           =   9615
      Begin VB.TextBox txtNoTelp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   72
         Top             =   650
         Width           =   2655
      End
      Begin VB.TextBox txtNama 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   70
         Top             =   360
         Width           =   4575
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   26
         Left            =   1680
         TabIndex        =   73
         Top             =   960
         Width           =   1875
         _ExtentX        =   3307
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   3555
         TabIndex        =   74
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
         Index           =   27
         Left            =   3675
         TabIndex        =   75
         Top             =   960
         Width           =   3915
         _ExtentX        =   6906
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
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Sekarang"
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
         TabIndex        =   76
         Top             =   1005
         Width           =   1005
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Telp"
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
         Left            =   990
         TabIndex        =   71
         Top             =   690
         Width           =   540
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Lengkap"
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
         TabIndex        =   69
         Top             =   405
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Customer"
      Height          =   3495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   13335
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   33
         Top             =   1830
         Width           =   2715
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   32
         Top             =   1500
         Width           =   2715
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   31
         Top             =   1170
         Width           =   2715
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1155
         MaxLength       =   100
         TabIndex        =   30
         Top             =   840
         Width           =   2715
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   5010
         MaxLength       =   150
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   2130
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   5010
         MaxLength       =   3
         TabIndex        =   2
         Top             =   1875
         Width           =   570
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   7
         Left            =   5850
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1875
         Width           =   570
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   1155
         TabIndex        =   4
         Top             =   540
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   31
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   1155
         TabIndex        =   5
         Top             =   3090
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
         Height          =   270
         Index           =   6
         Left            =   1155
         TabIndex        =   6
         Top             =   2475
         Width           =   1830
         _ExtentX        =   3228
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
         Height          =   285
         Index           =   1
         Left            =   3660
         TabIndex        =   7
         Top             =   240
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
         Left            =   2865
         TabIndex        =   8
         Top             =   240
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
         Index           =   7
         Left            =   1155
         TabIndex        =   9
         Top             =   2775
         Width           =   1830
         _ExtentX        =   3228
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
         Height          =   285
         Index           =   5
         Left            =   1155
         TabIndex        =   10
         Top             =   2160
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
         Index           =   0
         Left            =   1155
         TabIndex        =   11
         Top             =   240
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
         Height          =   270
         Index           =   3
         Left            =   3165
         TabIndex        =   12
         Top             =   540
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   31
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   4
         Left            =   5010
         TabIndex        =   13
         Top             =   540
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   31
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   8
         Left            =   3600
         TabIndex        =   34
         Top             =   2775
         Width           =   1830
         _ExtentX        =   3228
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
         Height          =   285
         Index           =   10
         Left            =   8310
         TabIndex        =   35
         Top             =   480
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   11
         Left            =   8310
         TabIndex        =   36
         Top             =   795
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   14
         Left            =   8310
         TabIndex        =   37
         Top             =   1395
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   13
         Left            =   8310
         TabIndex        =   38
         Top             =   1095
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
         Index           =   19
         Left            =   8310
         TabIndex        =   39
         Top             =   2325
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   503
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
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   20
         Left            =   8310
         TabIndex        =   40
         Top             =   2640
         Width           =   1935
         _ExtentX        =   3413
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
         Index           =   12
         Left            =   11670
         TabIndex        =   41
         Top             =   780
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
         Index           =   15
         Left            =   8310
         TabIndex        =   42
         Top             =   1710
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   8310
         TabIndex        =   43
         Top             =   2010
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   17
         Left            =   8925
         TabIndex        =   44
         Top             =   2010
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   18
         Left            =   9750
         TabIndex        =   45
         Top             =   2010
         Width           =   495
         _ExtentX        =   873
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   23
         Left            =   8310
         TabIndex        =   61
         Top             =   2950
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka"
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
         Left            =   7410
         TabIndex        =   62
         Top             =   3000
         Width           =   780
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
         Left            =   7530
         TabIndex        =   53
         Top             =   2055
         Width           =   645
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
         Left            =   7515
         TabIndex        =   52
         Top             =   2700
         Width           =   675
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
         Left            =   7485
         TabIndex        =   51
         Top             =   2370
         Width           =   705
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
         Left            =   7680
         TabIndex        =   50
         Top             =   1440
         Width           =   480
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
         Left            =   7740
         TabIndex        =   49
         Top             =   840
         Width           =   420
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
         Left            =   7515
         TabIndex        =   48
         Top             =   525
         Width           =   645
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
         Left            =   7320
         TabIndex        =   47
         Top             =   1755
         Width           =   840
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
         Left            =   9000
         TabIndex        =   46
         Top             =   525
         Width           =   480
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
         Left            =   3045
         TabIndex        =   29
         Top             =   285
         Width           =   570
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
         Left            =   375
         TabIndex        =   28
         Top             =   2205
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
         Left            =   720
         TabIndex        =   27
         Top             =   2820
         Width           =   315
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
         Left            =   375
         TabIndex        =   26
         Top             =   285
         Width           =   660
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
         Left            =   660
         TabIndex        =   25
         Top             =   2520
         Width           =   375
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
         Left            =   600
         TabIndex        =   24
         Top             =   3135
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
         Left            =   630
         TabIndex        =   23
         Top             =   615
         Width           =   405
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
         Left            =   3105
         TabIndex        =   22
         Top             =   2820
         Width           =   435
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RT/RW"
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
         Index           =   23
         Left            =   4245
         TabIndex        =   21
         Top             =   1875
         Width           =   510
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "           /"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   4740
         TabIndex        =   20
         Top             =   1635
         Width           =   780
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jalan /"
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
         Index           =   21
         Left            =   4470
         TabIndex        =   19
         Top             =   1095
         Width           =   480
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kelurahan"
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
         Index           =   19
         Left            =   315
         TabIndex        =   18
         Top             =   1875
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kecamatan"
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
         Index           =   18
         Left            =   240
         TabIndex        =   17
         Top             =   1545
         Width           =   795
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kab/Kota"
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
         Index           =   14
         Left            =   375
         TabIndex        =   16
         Top             =   1245
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Provinsi"
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
         Index           =   13
         Left            =   480
         TabIndex        =   15
         Top             =   930
         Width           =   555
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " Kompleks; No."
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
         Index           =   7
         Left            =   3930
         TabIndex        =   14
         Top             =   1335
         Width           =   1065
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   21
      Left            =   2040
      TabIndex        =   54
      Top             =   3720
      Width           =   1875
      _ExtentX        =   3307
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
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   270
      Left            =   3960
      TabIndex        =   55
      Top             =   3720
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
      Index           =   22
      Left            =   4080
      TabIndex        =   57
      Top             =   3720
      Width           =   3915
      _ExtentX        =   6906
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
   Begin ucXPButton.XPButton cmdKeluar 
      Height          =   495
      Left            =   12600
      TabIndex        =   58
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInjectDataSAtoSales.frx":0000
      PICN            =   "frmInjectDataSAtoSales.frx":001C
      PICH            =   "frmInjectDataSAtoSales.frx":694F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdBatal 
      Height          =   495
      Left            =   11400
      TabIndex        =   59
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInjectDataSAtoSales.frx":D606
      PICN            =   "frmInjectDataSAtoSales.frx":D622
      PICH            =   "frmInjectDataSAtoSales.frx":13C1B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdSimpan 
      Height          =   495
      Left            =   10200
      TabIndex        =   60
      Top             =   3720
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmInjectDataSAtoSales.frx":1A4CA
      PICN            =   "frmInjectDataSAtoSales.frx":1A4E6
      PICH            =   "frmInjectDataSAtoSales.frx":20CE9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   24
      Left            =   2040
      TabIndex        =   63
      Top             =   4080
      Width           =   1875
      _ExtentX        =   3307
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   3960
      TabIndex        =   64
      Top             =   4080
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
      Index           =   25
      Left            =   4080
      TabIndex        =   65
      Top             =   4080
      Width           =   3915
      _ExtentX        =   6906
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
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unit Yang Diinginkan"
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
      TabIndex        =   66
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Sales"
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
      Left            =   1200
      TabIndex        =   56
      Top             =   3765
      Width           =   780
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7575
      Left            =   0
      Picture         =   "frmInjectDataSAtoSales.frx":2788F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmInjectDataSAtoSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmInjectDataSAtoSales
' DateTime  : 05/12/2019 14:53
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

Private rsTemp      As ADODB.Recordset


Private Sub hapusForm()
    For i = 0 To 27
        Text(i) = ""
    Next
    For i = 1 To 7
        txt(i).Text = ""
    Next
    txtNama.Text = ""
    txtNoTelp.Text = ""
    ckReferensi.value = vbUnchecked
    Frame2.Visible = False
    If Log.hansaRenkeiBP = 1 Then
        Label3.Visible = True
        Text(24).Visible = True
        MaskEdBox1.Visible = True
        Text(25).Visible = True
        Frame2.Visible = False
        ckReferensi.Visible = True
    Else
        Label3.Visible = False
        Text(24).Visible = False
        MaskEdBox1.Visible = False
        Text(25).Visible = False
        Frame2.Visible = False
        ckReferensi.Visible = False
    End If
End Sub

Private Sub ckReferensi_Click()
    If ckReferensi.value = vbChecked Then
        Frame2.Visible = True
    Else
        Frame2.Visible = False
    End If
End Sub

Private Sub cmdBatal_Click()
    Call hapusForm
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
          Dim insertDealer    As Boolean
          Dim isBegin         As Boolean
          Dim isReferensi     As Boolean
1        On Error GoTo cmdSimpan_Click_Error
2         isReferensi = False
3         isBegin = False
4         insertDealer = False
5         If NulSID(Text(0)) = "" And NulSID(Text(23)) = "" Then
6             MsgBox "Data customer harus dilengkapi.", vbInformation
7             Exit Sub
8         End If
          
9         If Log.hansaRenkeiBP = 1 Then
10            If NulSID(Text(24)) <> "" Then
11                Set rsTemp = objO.FetchData("select Top 50 MDL, [DESC] as keterangan, MODEL, FS_GRP_MDL from [" & dbReplikasiHO & "]..model WITH (NOLOCK) WHERE FS_LOG_DEL = '' and mdl = '" & NulSID(Text(24)) & "'")
12                If Not rsTemp.EOF Then
13                    Text(25) = NulSID(rsTemp!keterangan)
14                Else
15                    MsgBox "Data Kendaraan tidak ditemukan.", vbInformation
16                    Text(24) = ""
17                    Exit Sub
                      
18                End If
19            End If
20            If ckReferensi.value = vbChecked Then
21                If NulSID(txtNama.Text) = "" Then
22                    MsgBox "Nama Referensi harus diisi.", vbInformation
23                    Exit Sub
24                End If
25                If NulSID(txtNoTelp.Text) = "" Then
26                    MsgBox "No Telp Referensi harus diisi.", vbInformation
27                    Exit Sub
28                End If
29                If NulSID(Text(26)) <> "" Then
30                    Set rsTemp = objO.FetchData("select Top 50 MDL, [DESC] as keterangan, MODEL, FS_GRP_MDL from [" & dbReplikasiHO & "]..model WITH (NOLOCK) WHERE FS_LOG_DEL = '' and mdl = '" & NulSID(Text(26)) & "'")
31                    If Not rsTemp.EOF Then
32                        Text(27) = NulSID(rsTemp!keterangan)
33                    Else
34                        MsgBox "Data Kendaraan tidak ditemukan.", vbInformation
35                        Text(26) = ""
36                        Exit Sub
                          
37                    End If
38                End If
39                isReferensi = True
40            End If
41        End If
42        If NulSID(Text(21)) <> "" Then
          
              ''Set rsTemp = objO.FetchDataSPK("SELECT KodeSales,NamaSales FROM ArSales WITH (nolock) where TdkAktif = 0 and kodesales = '" & NulSID(Text(21)) & "'")
              '' 20211125DEV
43            If Log.FlagSalesHO = 0 Then
44                Set rsTemp = objO.FetchDataSPK("select KodeSales,NamaSales from ArSales WITH (NOLOCK) where TdkAktif = 0 AND KODESALES = '" & NulSID(Text(21)) & "'")
45            Else
46                Set rsTemp = objO.FetchDataHO("select KodeSales,NamaSales from Stars99..ArSales WITH (NOLOCK) where KodeCabang = '" & Log.kodeCabang & "' and TdkAktif = 0 AND KODESALES = '" & NulSID(Text(21)) & "'")
47            End If
                      
48            If Not rsTemp.EOF Then
49                Text(22) = NulSID(rsTemp!namasales)
50            Else
51                MsgBox "Data sales tidak ditemukan.", vbInformation
52                Text(21) = ""
53                Exit Sub
                  
54            End If
55        Else
56            If MsgBox("Apakah yakin untuk inject data ke sales?", vbYesNo, "Konfirmasi Inject") = vbYes Then
57                insertDealer = True
58            Else
59                Exit Sub
60            End If
61        End If
          
          Dim kodepr      As String
          Dim grupCust    As Integer
          Dim namapr      As String
          Dim namaLengkap As String
          Dim kodeProsp   As String
          Dim simpanHO    As Boolean
62        If isReferensi Then
63            namaLengkap = NulSID(txtNama.Text)
64        Else
65            namaLengkap = NulSID(Text(2)) & " " & NulSID(Text(3)) & " " & NulSID(Text(4))
66        End If
67        namapr = Left(Replace(namaLengkap, " ", "") & "-------", 7)
68        kodepr = namapr & year(isnow)
          '' 20200408DEV
69        If Log.sentralisasiStars = 1 Then
70            If Log.flagOffBookingLocal = 1 Then
71                Set rsTemp = objO.FetchDataHO("select top 1 Convert(Numeric(19,0),Right(IsNull(KodePr,0),4)) nomor from [stars" & Right(gDb, 2) & "]..ProspectMst where KodePr like '" & kodepr & "%' order by Convert(Numeric(19,0),Right(IsNull(KodePr,0),4)) desc")
72            Else
73                Set rsTemp = objO.FetchData("select top 1 Convert(Numeric(19,0),Right(IsNull(KodePr,0),4)) nomor from [" & dbReplikasiHO & "]..ProspectMst where kodecabang = '" & Log.kodeCabang & "' and KodePr like '" & kodepr & "%' order by Convert(Numeric(19,0),Right(IsNull(KodePr,0),4)) desc")
74            End If
75        Else
76            Set rsTemp = objO.FetchData("select top 1 Convert(Numeric(19,0),Right(IsNull(KodePr,0),4)) nomor from [stars" & Right(gDb, 2) & "]..ProspectMst where KodePr like '" & kodepr & "%' order by Convert(Numeric(19,0),Right(IsNull(KodePr,0),4)) desc")
77        End If
78        If Not rsTemp.EOF Then
79            kodepr = kodepr & Format(rsTemp!nomor + 1, "000#")
80        Else
81            kodepr = kodepr & "0001"
82        End If
          
83        If IsNumeric(Left(Text(0), 2)) = True And isReferensi = False Then
84            grupCust = CInt(Left(Text(0), 2))
85            grupCust = objO.FetchData_Integer("SELECT IDTypecust FROM bridgingTipeCustomer WITH (NOLOCK) WHERE SIDTIPECUST = '" & grupCust & "'")
86        Else
87            grupCust = 1
88        End If
          
89        kodeProsp = "PR" & Right(year(isnow), 2) & "0001"
90        If Log.flagOffBookingLocal = 1 Then
91            Call objHT.OpenTransactionImportHO
92        Else
93            Call objHT.OpenTransactionBengkel
94        End If
95        erTranAll = False
96        isBegin = True
97        simpanHO = True
          '' 20200408DEV
98        If Log.sentralisasiStars = 1 Then
99            If Log.flagOffBookingLocal = 1 Then
100               sSQL = "insert into [STARS" & Right(gDb, 2) & "].dbo.ProspectMst (KodePr,Nama,Alamat,Kota,KodePos,"
101               sSQL = sSQL & " Telp,GSM,CDMA,KodeUser,TglEntry,SalesPr,Grup,Sumber,sbrdt,kodecabang )"
102               sSQL = sSQL & " values ('" & kodepr & "','" & namaLengkap & "','" & Left(NulSID(txt(5)) & " " & NulSID(txt(4)) & ", " & NulSID(txt(3)) & ", " & NulSID(txt(2)) & ", " & NulSID(txt(1)), 200) & "',"
103               sSQL = sSQL & " '" & NulSID(txt(2)) & "','" & NulSID(Text(5)) & "','" & NulSID(Text(6)) & "','" & NulSID(Text(7)) & "','" & NulSID(Text(8)) & "'"
104               sSQL = sSQL & " ,'" & gSA & "',getdate(),'" & NulSID(Text(21)) & "'," & grupCust & ",17,2,'" & Log.kodeCabang & "')"
              
105           Else
106               sSQL = "insert into [" & dbReplikasiHO & "]..ProspectMst (KodePr,Nama,Alamat,Kota,KodePos,"
107               sSQL = sSQL & " Telp,GSM,CDMA,KodeUser,TglEntry,SalesPr,Grup,Sumber,sbrdt,kodecabang )"
108               sSQL = sSQL & " values ('" & kodepr & "','" & namaLengkap & "','" & Left(NulSID(txt(5)) & " " & NulSID(txt(4)) & ", " & NulSID(txt(3)) & ", " & NulSID(txt(2)) & ", " & NulSID(txt(1)), 200) & "',"
109               sSQL = sSQL & " '" & NulSID(txt(2)) & "','" & NulSID(Text(5)) & "','" & NulSID(Text(6)) & "','" & NulSID(Text(7)) & "','" & NulSID(Text(8)) & "'"
110               sSQL = sSQL & " ,'" & gSA & "',getdate(),'" & NulSID(Text(21)) & "'," & grupCust & ",17,2,'" & Log.kodeCabang & "')"
111           End If
112       Else
113           simpanHO = False
114           sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectMst (KodePr,Nama,Alamat,Kota,KodePos,"
115           sSQL = sSQL & " Telp,GSM,CDMA,KodeUser,TglEntry,SalesPr,Grup,Sumber,sbrdt )"
116           sSQL = sSQL & " values ('" & kodepr & "','" & namaLengkap & "','" & Left(NulSID(txt(5)) & " " & NulSID(txt(4)) & ", " & NulSID(txt(3)) & ", " & NulSID(txt(2)) & ", " & NulSID(txt(1)), 200) & "',"
117           sSQL = sSQL & " '" & NulSID(txt(2)) & "','" & NulSID(Text(5)) & "','" & NulSID(Text(6)) & "','" & NulSID(Text(7)) & "','" & NulSID(Text(8)) & "'"
118           sSQL = sSQL & " ,'" & gSA & "',getdate(),'" & NulSID(Text(21)) & "'," & grupCust & ",17,2)"
119       End If
          
120       If Log.hansaRenkeiBP = 1 And isReferensi Then
121           If Log.flagOffBookingLocal = 1 Then
122               sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectMst (KodePr,Nama,Alamat,Kota,KodePos,"
123               sSQL = sSQL & " Telp,GSM,CDMA,KodeUser,TglEntry,SalesPr,Grup,Sumber,sbrdt,kodecabang )"
124               sSQL = sSQL & " values ('" & kodepr & "','" & namaLengkap & "','',"
125               sSQL = sSQL & " '','','','" & NulSID(txtNoTelp.Text) & "',''"
126               sSQL = sSQL & " ,'" & gSA & "',getdate(),'" & NulSID(Text(21)) & "'," & grupCust & ",17,2,'" & Log.kodeCabang & "')"
              
127           Else
128               sSQL = "insert into [" & dbReplikasiHO & "]..ProspectMst (KodePr,Nama,Alamat,Kota,KodePos,"
129               sSQL = sSQL & " Telp,GSM,CDMA,KodeUser,TglEntry,SalesPr,Grup,Sumber,sbrdt,kodecabang )"
130               sSQL = sSQL & " values ('" & kodepr & "','" & namaLengkap & "','',"
131               sSQL = sSQL & " '','','','" & NulSID(txtNoTelp.Text) & "',''"
132               sSQL = sSQL & " ,'" & gSA & "',getdate(),'" & NulSID(Text(21)) & "'," & grupCust & ",17,2,'" & Log.kodeCabang & "')"
133           End If
134       End If
135       If Log.flagOffBookingLocal = 1 And simpanHO Then
136           Call objHT.InsertData_Ho(sSQL)
137       Else
138           Call objHT.InsertData_Bengkel(sSQL)
139       End If
140       simpanHO = True
141       If isReferensi Then
142               If Log.flagOffBookingLocal = 1 Then
143                   sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectKary(KodeKategory,NamaKategory,Norangka,KodeKaryawan,namaKaryawan,"
144                   sSQL = sSQL & " SA,Kodepr,Nama,Alamat,Kota,Kecamatan,Kelurahan,Telp,"
145                   sSQL = sSQL & " Gsm,Cdma,TglEntry,KodeSales,kodecabang ,KodeKategoryDicari,NamaKategoryDicari,referensi,namaPereferensi,idcustPereferensi) values ('" & Text(26) & "','" & Text(27) & "','','" & gSA & "','" & NulSID(gNama_SA) & "',"
146                   sSQL = sSQL & " '1','" & kodepr & "','" & namaLengkap & "','',"
147                   sSQL = sSQL & " '','','','" & NulSID(txtNoTelp.Text) & "',"
148                   sSQL = sSQL & " '','',getdate(),'" & NulSID(Text(21)) & "','" & Log.kodeCabang & "','" & Text(24) & "','" & Text(25) & "',1,'" & NulSID(NulSID(Text(2)) & " " & NulSID(Text(3)) & " " & NulSID(Text(4))) & "','" & NulSID(Text(0)) & "')"
                  
149               Else
150                   sSQL = "insert into [" & dbReplikasiHO & "]..ProspectKary(KodeKategory,NamaKategory,Norangka,KodeKaryawan,namaKaryawan,"
151                   sSQL = sSQL & " SA,Kodepr,Nama,Alamat,Kota,Kecamatan,Kelurahan,Telp,"
152                   sSQL = sSQL & " Gsm,Cdma,TglEntry,KodeSales,kodecabang ,KodeKategoryDicari,NamaKategoryDicari,referensi,namaPereferensi,idcustPereferensi) values ('" & Text(26) & "','" & Text(27) & "','','" & gSA & "','" & NulSID(gNama_SA) & "',"
153                   sSQL = sSQL & " '1','" & kodepr & "','" & namaLengkap & "','',"
154                   sSQL = sSQL & " '','','','" & NulSID(txtNoTelp.Text) & "',"
155                   sSQL = sSQL & " '','',getdate(),'" & NulSID(Text(21)) & "','" & Log.kodeCabang & "','" & Text(24) & "','" & Text(25) & "',1,'" & NulSID(NulSID(Text(2)) & " " & NulSID(Text(3)) & " " & NulSID(Text(4))) & "','" & NulSID(Text(0)) & "')"
156               End If
157       Else
              '' 20200408DEV
158           If Log.hansaRenkeiBP = 1 Then
159               If Log.flagOffBookingLocal = 1 Then
160                   sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectKary(KodeKategory,NamaKategory,Norangka,KodeKaryawan,namaKaryawan,"
161                   sSQL = sSQL & " SA,Kodepr,Nama,Alamat,Kota,Kecamatan,Kelurahan,Telp,"
162                   sSQL = sSQL & " Gsm,Cdma,TglEntry,KodeSales,kodecabang ,KodeKategoryDicari,NamaKategoryDicari,referensi) values ('" & Text(11) & "','" & Text(13) & "','" & Text(23) & "','" & gSA & "','" & NulSID(gNama_SA) & "',"
163                   sSQL = sSQL & " '1','" & kodepr & "','" & namaLengkap & "','" & NulSID(txt(5)) & "',"
164                   sSQL = sSQL & " '" & NulSID(txt(2)) & "','" & NulSID(txt(3)) & "','" & NulSID(txt(4)) & "','" & NulSID(Text(6)) & "',"
165                   sSQL = sSQL & " '" & NulSID(Text(7)) & "','" & NulSID(Text(8)) & "',getdate(),'" & NulSID(Text(21)) & "','" & Log.kodeCabang & "','" & Text(24) & "','" & Text(25) & "',0)"
                  
166               Else
167                   sSQL = "insert into [" & dbReplikasiHO & "]..ProspectKary(KodeKategory,NamaKategory,Norangka,KodeKaryawan,namaKaryawan,"
168                   sSQL = sSQL & " SA,Kodepr,Nama,Alamat,Kota,Kecamatan,Kelurahan,Telp,"
169                   sSQL = sSQL & " Gsm,Cdma,TglEntry,KodeSales,kodecabang ,KodeKategoryDicari,NamaKategoryDicari,referensi) values ('" & Text(11) & "','" & Text(13) & "','" & Text(23) & "','" & gSA & "','" & NulSID(gNama_SA) & "',"
170                   sSQL = sSQL & " '1','" & kodepr & "','" & namaLengkap & "','" & NulSID(txt(5)) & "',"
171                   sSQL = sSQL & " '" & NulSID(txt(2)) & "','" & NulSID(txt(3)) & "','" & NulSID(txt(4)) & "','" & NulSID(Text(6)) & "',"
172                   sSQL = sSQL & " '" & NulSID(Text(7)) & "','" & NulSID(Text(8)) & "',getdate(),'" & NulSID(Text(21)) & "','" & Log.kodeCabang & "','" & Text(24) & "','" & Text(25) & "',0)"
173               End If
174           ElseIf Log.sentralisasiStars = 1 Then
175               If Log.flagOffBookingLocal = 1 Then
176                   sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectKary(KodeKategory,NamaKategory,Norangka,KodeKaryawan,namaKaryawan,"
177                   sSQL = sSQL & " SA,Kodepr,Nama,Alamat,Kota,Kecamatan,Kelurahan,Telp,"
178                   sSQL = sSQL & " Gsm,Cdma,TglEntry,KodeSales,kodecabang ) values ('" & Text(11) & "','" & Text(13) & "','" & Text(23) & "','" & gSA & "','" & NulSID(gNama_SA) & "',"
179                   sSQL = sSQL & " '1','" & kodepr & "','" & namaLengkap & "','" & NulSID(txt(5)) & "',"
180                   sSQL = sSQL & " '" & NulSID(txt(2)) & "','" & NulSID(txt(3)) & "','" & NulSID(txt(4)) & "','" & NulSID(Text(6)) & "',"
181                   sSQL = sSQL & " '" & NulSID(Text(7)) & "','" & NulSID(Text(8)) & "',getdate(),'" & NulSID(Text(21)) & "','" & Log.kodeCabang & "')"
                  
182               Else
183                   sSQL = "insert into [" & dbReplikasiHO & "]..ProspectKary(KodeKategory,NamaKategory,Norangka,KodeKaryawan,namaKaryawan,"
184                   sSQL = sSQL & " SA,Kodepr,Nama,Alamat,Kota,Kecamatan,Kelurahan,Telp,"
185                   sSQL = sSQL & " Gsm,Cdma,TglEntry,KodeSales,kodecabang ) values ('" & Text(11) & "','" & Text(13) & "','" & Text(23) & "','" & gSA & "','" & NulSID(gNama_SA) & "',"
186                   sSQL = sSQL & " '1','" & kodepr & "','" & namaLengkap & "','" & NulSID(txt(5)) & "',"
187                   sSQL = sSQL & " '" & NulSID(txt(2)) & "','" & NulSID(txt(3)) & "','" & NulSID(txt(4)) & "','" & NulSID(Text(6)) & "',"
188                   sSQL = sSQL & " '" & NulSID(Text(7)) & "','" & NulSID(Text(8)) & "',getdate(),'" & NulSID(Text(21)) & "','" & Log.kodeCabang & "')"
189               End If
190           Else
191               simpanHO = False
192               sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectKary(KodeKategory,NamaKategory,Norangka,KodeKaryawan,namaKaryawan,"
193               sSQL = sSQL & " SA,Kodepr,Nama,Alamat,Kota,Kecamatan,Kelurahan,Telp,"
194               sSQL = sSQL & " Gsm,Cdma,TglEntry,KodeSales) values ('" & Text(11) & "','" & Text(13) & "','" & Text(23) & "','" & gSA & "','" & NulSID(gNama_SA) & "',"
195               sSQL = sSQL & " '1','" & kodepr & "','" & namaLengkap & "','" & NulSID(txt(5)) & "',"
196               sSQL = sSQL & " '" & NulSID(txt(2)) & "','" & NulSID(txt(3)) & "','" & NulSID(txt(4)) & "','" & NulSID(Text(6)) & "',"
197               sSQL = sSQL & " '" & NulSID(Text(7)) & "','" & NulSID(Text(8)) & "',getdate(),'" & NulSID(Text(21)) & "')"
198           End If
199       End If
          
200       If Log.flagOffBookingLocal = 1 And simpanHO Then
201           Call objHT.InsertData_Ho(sSQL)
202       Else
203           Call objHT.InsertData_Bengkel(sSQL)
204       End If
          '' 20200408DEV
205       simpanHO = True
206       If Log.sentralisasiStars = 1 Then
207           If Log.flagOffBookingLocal = 1 Then
208               sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectTrn (KodePr,KodeProspect,Kodeuser,TglEntry,TglAwalEntry,SalesPr,kodecabang)"
209               sSQL = sSQL & " VALUES"
210               sSQL = sSQL & " ('" & kodepr & "','" & kodeProsp & "','" & gSA & "',getdate(), getdate(), '" & NulSID(Text(21)) & "','" & Log.kodeCabang & "')"
211           Else
212               sSQL = "insert into [" & dbReplikasiHO & "]..ProspectTrn (KodePr,KodeProspect,Kodeuser,TglEntry,TglAwalEntry,SalesPr,kodecabang)"
213               sSQL = sSQL & " VALUES"
214               sSQL = sSQL & " ('" & kodepr & "','" & kodeProsp & "','" & gSA & "',getdate(), getdate(), '" & NulSID(Text(21)) & "','" & Log.kodeCabang & "')"
215           End If
216       Else
217           simpanHO = False
218           sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectTrn (KodePr,KodeProspect,Kodeuser,TglEntry,TglAwalEntry,SalesPr)"
219           sSQL = sSQL & " VALUES"
220           sSQL = sSQL & " ('" & kodepr & "','" & kodeProsp & "','" & gSA & "',getdate(), getdate(), '" & NulSID(Text(21)) & "')"
221       End If
          
222       If isReferensi Then
223           If Log.flagOffBookingLocal = 1 Then
224               sSQL = "insert into [STARS" & Right(gDb, 2) & "]..ProspectTrn (KodePr,KodeProspect,Kodeuser,TglEntry,TglAwalEntry,SalesPr,kodecabang)"
225               sSQL = sSQL & " VALUES"
226               sSQL = sSQL & " ('" & kodepr & "','" & kodeProsp & "','" & gSA & "',getdate(), getdate(), '" & NulSID(Text(21)) & "','" & Log.kodeCabang & "')"
              
227           Else
228               sSQL = "insert into [" & dbReplikasiHO & "]..ProspectTrn (KodePr,KodeProspect,Kodeuser,TglEntry,TglAwalEntry,SalesPr,kodecabang)"
229               sSQL = sSQL & " VALUES"
230               sSQL = sSQL & " ('" & kodepr & "','" & kodeProsp & "','" & gSA & "',getdate(), getdate(), '" & NulSID(Text(21)) & "','" & Log.kodeCabang & "')"
231           End If
232       End If
          
233       If Log.flagOffBookingLocal = 1 And simpanHO Then
234           Call objHT.InsertData_Ho(sSQL)
235       Else
236           Call objHT.InsertData_Bengkel(sSQL)
237       End If
          
          '' 20200408DEV
238       If Log.sentralisasiStars = 1 Then
239           sSQL = "INSERT INTO [" & dbReplikasiHO & "]..PlanTrn (SbrPlan,KodePr,KodeProspect,TglPlan,Via,KodePlan,TypePlan,"
240           sSQL = sSQL & " idKetPlan,Penggolongan,KetPlan,[Yang Berulang Tahun],[Tanggal Lahir],"
241           sSQL = sSQL & " NoRangka,Type,[Tanggal Jatuh Tempo],KodeSales,Kodeuser,TglEntry,kodecabang)"
242           sSQL = sSQL & " VALUES"
243           sSQL = sSQL & " (2,'" & kodepr & "','" & kodeProsp & "','" & Format(isnow, "yyyy-mm-dd") & "','KUNJUNGAN','PLPR" & Right(year(isnow), 2) & "0001',1,"
244           sSQL = sSQL & " 1,'pr','Kunjungan Pertama Kali',NULL,NULL,'" & NulSID(Text(23)) & "',NULL,NULL,'" & NulSID(Text(21)) & "','AutoPlanSystem',GETDATE(),'" & Log.kodeCabang & "')"
245       Else
246           sSQL = "INSERT INTO [STARS" & Right(gDb, 2) & "]..PlanTrn (SbrPlan,KodePr,KodeProspect,TglPlan,Via,KodePlan,TypePlan,"
247           sSQL = sSQL & " idKetPlan,Penggolongan,KetPlan,[Yang Berulang Tahun],[Tanggal Lahir],"
248           sSQL = sSQL & " NoRangka,Type,[Tanggal Jatuh Tempo],KodeSales,Kodeuser,TglEntry)"
249           sSQL = sSQL & " VALUES"
250           sSQL = sSQL & " (2,'" & kodepr & "','" & kodeProsp & "','" & Format(isnow, "yyyy-mm-dd") & "','KUNJUNGAN','PLPR" & Right(year(isnow), 2) & "0001',1,"
251           sSQL = sSQL & " 1,'pr','Kunjungan Pertama Kali',NULL,NULL,'" & NulSID(Text(23)) & "',NULL,NULL,'" & NulSID(Text(21)) & "','AutoPlanSystem',GETDATE())"
252       End If
253       Call objHT.InsertData_Bengkel(sSQL)
          
254       If isReferensi Then
255           sSQL = "INSERT INTO [" & dbReplikasiHO & "]..PlanTrn (SbrPlan,KodePr,KodeProspect,TglPlan,Via,KodePlan,TypePlan,"
256           sSQL = sSQL & " idKetPlan,Penggolongan,KetPlan,[Yang Berulang Tahun],[Tanggal Lahir],"
257           sSQL = sSQL & " NoRangka,Type,[Tanggal Jatuh Tempo],KodeSales,Kodeuser,TglEntry,kodecabang)"
258           sSQL = sSQL & " VALUES"
259           sSQL = sSQL & " (2,'" & kodepr & "','" & kodeProsp & "','" & Format(isnow, "yyyy-mm-dd") & "','KUNJUNGAN','PLPR" & Right(year(isnow), 2) & "0001',1,"
260           sSQL = sSQL & " 1,'pr','Kunjungan Pertama Kali',NULL,NULL,'',NULL,NULL,'" & NulSID(Text(21)) & "','AutoPlanSystem',GETDATE(),'" & Log.kodeCabang & "')"
261       End If
          
          ''GR TO SALES
          ''--INSERT INTO NotifHansaRenkei (KodeSumber,Tipe,Nowo,NoRangka,KodeDealer,KodePesan,KodeSales,NIK,TglNotif,Proses,UserEntry,TglEntry,Keterangan)
          ''--VALUES ('GRP','LEADS','KodePr','NoRangkaSekarang','kodedatabase','NoBookCarsWorld','','NIKsales',GETDATE(),0,'UserLog',GETDATE(),'NAMA_SA#NAMA_CUST#MODEL#TELP#ALAMAT')
          
262       sSQL = "INSERT INTO [" & dbReplikasiHO & "]..NotifHansaRenkei (KodeSumber,Tipe,Nowo,NoRangka,KodeDealer,KodePesan,KodeSales,NIK,TglNotif,Proses,UserEntry,TglEntry,Keterangan)"
          ''sSQL = sSQL & " VALUES ('GRP','LEADS','" & kodepr & "','" & NulSID(Text(23)) & "','" & Right(gDb, 2) & "','SAGR:" & gNIK & " - " & Left(gNama_SA, 68) & "SALES:" & NulSID(Text(21)) & "','','" & NulSID(Text(21)) & "',GETDATE(),0,'" & gSA & "',GETDATE(),'" & gNama_SA & "#" & NulSID(Text(2)) & " " & NulSID(Text(3)) & " " & NulSID(Text(4)) & "#" & NulSID(Text(13)) & "#" & NulSID(Text(7)) & "#" & NulSID(txt(5)) & " ')"
263       sSQL = sSQL & " VALUES ('GRP','LEADS','" & kodepr & "','" & NulSID(Text(23)) & "','" & Right(gDb, 2) & "','SAGR:" & gNIK & " - " & Left(gNama_SA, 68) & "SALES:" & NulSID(Text(21)) & "','" & NulSID(Text(21)) & "','',GETDATE(),0,'" & gSA & "',GETDATE(),'" & gNama_SA & "#" & NulSID(Text(2)) & " " & NulSID(Text(3)) & " " & NulSID(Text(4)) & "#" & NulSID(Text(13)) & "#" & NulSID(Text(7)) & "#" & NulSID(txt(5)) & " ')"
264       Call objHT.InsertData_Bengkel(sSQL)
          
265       If erTranAll = False Then
266           Call objHT.CloseTransactionBengkel
267           Call hapusForm
268           MsgBox "Data inject berhasil disimpan.", vbInformation
269           Exit Sub
270       Else
271           Call objHT.CloseTransactionBengkel_err
272           Call hapusForm
273           MsgBox "Data inject gagal disimpan.", vbInformation
274           Exit Sub
275       End If

276      On Error GoTo 0
277      Exit Sub

cmdSimpan_Click_Error:
278       If isBegin Then
279           Call objHT.CloseTransactionBengkel_err
280       End If
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
281       varErlError = Erl
282       varErrNumber = err.Number
283       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmInjectDataSAtoSales")
          
End Sub

Private Sub Form_Load()
    Call hapusForm
    ''Text(0).SetFocus
End Sub

Private Sub Text_GotFocus(Index As Integer)
          Dim DiTab As Boolean
1         DiTab = False
          
2         If Text(Index).BackColor = &HC0E0FF Then
3             DiTab = True
4         Else
5             If gMainKeyCari <> "" Then
6                 If Index = 0 Then
7                     Text(23) = gMainKeyCari
8                     sSQL = "select c.NamaCust1,c.NamaCust2,c.NamaCust3,"
9                     sSQL = sSQL & " c.IDCust,b.KodeLgn,c.Propinsi,c.Kota,c.Kelurahan,c.Kecamatan,c.Alamat,c.RT,c.RW,"
10                    sSQL = sSQL & " c.KodePos,c.IdTelp,c.Telp,c.HP1,c.HP2,c.NoNPWP,"
11                    sSQL = sSQL & " a.FS_TH_MOBIL , a.MDL, a.model, a.FMDL, a.Color, a.FS_KM_AKHIR, a.NO_POLISI, a.NO_ENGINE, a.KEYNO"
12                    sSQL = sSQL & " from CUSTDAT a with (nolock)"
13                    sSQL = sSQL & " inner join [SingleIDOri]..Bridging b with (nolock) on a.FS_KD_PLG = b.KodeLgn and b.KodeDealer = '" & Right(gDb, 2) & "'"
14                    sSQL = sSQL & " inner join [SingleIDOri]..NewCustomer c with (NOLOCK) ON b.IDCust = c.IDCust"
15                    sSQL = sSQL & " where a.NO_CHASSIS = '" & NulSID(Text(23)) & "'"
                      
16                    Set rsTemp = objO.FetchData(sSQL)
17                    If Not rsTemp.EOF Then
                          Dim nDpn1 As String
                          Dim nTgh1 As String
                          Dim nBlk1 As String
18                        txt(1).Text = NulSID(rsTemp!Propinsi)
19                        txt(2).Text = NulSID(rsTemp!Kota)
20                        txt(3).Text = NulSID(rsTemp!kecamatan)
21                        txt(4).Text = NulSID(rsTemp!kelurahan)
22                        txt(5).Text = NulSID(rsTemp!Alamat)
23                        txt(6).Text = NulSID(rsTemp!rt)
24                        txt(7).Text = NulSID(rsTemp!rw)
                          
25                        Text(0) = NulSID(rsTemp!idcust)
26                        Text(1) = Left(NulSID(rsTemp!kodeLgn), 3)
27                        Text(2) = Left(NulSID(rsTemp!namacust1), 30)
28                        Text(3) = Left(NulSID(rsTemp!namacust2), 30)
29                        Text(4) = Left(NulSID(rsTemp!namacust3), 30)
30                        Text(5) = NulSID(rsTemp!KodePos)
31                        Text(6) = NulSID(rsTemp!IdTelp) & NulSID(rsTemp!telp)
32                        Text(7) = NulSID(rsTemp!hp1)
33                        Text(8) = NulSID(rsTemp!hp2)
34                        Text(9) = NulSID(rsTemp!noNPWP)
35                        Text(10) = NulSID(rsTemp!FS_TH_MOBIL)
36                        Text(11) = Left(NulSID(rsTemp!MDL), 20)
37                        Text(12) = NulSID(rsTemp!model)
38                        Text(13) = Left(NulSID(rsTemp!Fmdl), 50)
39                        Text(14) = NulSID(rsTemp!Color)
40                        Text(15) = NulSID(rsTemp!FS_KM_AKHIR)
41                        Call pecahNoPol(NulSID(rsTemp!NO_POLISI), nDpn1, nTgh1, nBlk1)
42                        Text(16) = Trim(nDpn1)
43                        Text(17) = Trim(nTgh1)
44                        Text(18) = Trim(nBlk1)
45                        Text(19) = NulSID(rsTemp!No_Engine)
46                        Text(20) = NulSID(rsTemp!KEYNO)
47                    Else
48                        MsgBox "Data tidak ditemukan dalam singleid customer.", vbInformation
49                        Call hapusForm
50                    End If
51                Else
52                    Text(Index) = gMainKeyCari
53                End If
                  
54                gMainKeyCari = ""
55                DiTab = True
56            End If
57        End If
58        If DiTab Then
59            Sendkeys ("{TAB}")
60        Else
61            Sendkeys ("{HOME}+{END}")
62        End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 21
                gMainKeyCari = Text(Index)
                If Log.FlagSalesHO = 0 Then
                    gSearchOption = "SALES"
                    frmCariSPK.show (vbModal)
                Else
                    gSearchOption = "SALES"
                    frmCariHO.show (vbModal)
                End If
            Case 24
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 26
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "CUSTOMER"
                frmCari.show (vbModal)
             
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 21
            If Text(Index) <> "" Then
                If Log.FlagSalesHO = 0 Then
                    Set oRs = objO.FetchDataSPK("select KodeSales,NamaSales from ArSales WITH (NOLOCK) where TdkAktif = 0 AND KODESALES = '" & NulSID(Text(Index)) & "'")
                Else
                    Set oRs = objO.FetchDataHO("select KodeSales,NamaSales from Stars99..ArSales WITH (NOLOCK) where KodeCabang = '" & Log.kodeCabang & "' and TdkAktif = 0 AND KODESALES = '" & NulSID(Text(Index)) & "'")
                End If
                
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SALES"
                    frmCariSPK.show (vbModal)
                Else
                    Text(Index + 1) = oRs("NamaSales")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 24
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("select Top 50 MDL, [DESC] as keterangan, MODEL, FS_GRP_MDL from [" & dbReplikasiHO & "]..model WITH (NOLOCK) WHERE FS_LOG_DEL = '' and mdl = '" & NulSID(Text(Index)) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = oRs("keterangan")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 26
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("select Top 50 MDL, [DESC] as keterangan, MODEL, FS_GRP_MDL from [" & dbReplikasiHO & "]..model WITH (NOLOCK) WHERE FS_LOG_DEL = '' and mdl = '" & NulSID(Text(Index)) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = oRs("keterangan")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
    End Select
End Sub

