VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmInjectDataBookingBP 
   Caption         =   "Inject Data SA to Booking BP"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13890
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   13890
   Begin TDBDate6Ctl.TDBDate tglBooking 
      Height          =   375
      Left            =   2040
      TabIndex        =   79
      Top             =   4920
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   661
      Calendar        =   "frmInjectDataBookingBP.frx":0000
      Caption         =   "frmInjectDataBookingBP.frx":012C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInjectDataBookingBP.frx":0198
      Keys            =   "frmInjectDataBookingBP.frx":01B6
      Spin            =   "frmInjectDataBookingBP.frx":0214
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      CursorPosition  =   0
      DataProperty    =   0
      DisplayFormat   =   "dd/mm/yyyy"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      FirstMonth      =   4
      ForeColor       =   -2147483640
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   -657434
      MousePointer    =   0
      MoveOnLRKey     =   0
      OLEDragMode     =   0
      OLEDropMode     =   0
      PromptChar      =   "_"
      ReadOnly        =   0
      ShowContextMenu =   -1
      ShowLiterals    =   0
      TabAction       =   0
      Text            =   "28/09/2020"
      ValidateMode    =   0
      ValueVT         =   7
      Value           =   44102
      CenturyMode     =   0
   End
   Begin VB.TextBox txtPekerjaan 
      Height          =   705
      Left            =   2040
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   78
      Top             =   4080
      Width           =   6015
   End
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
      Top             =   7440
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   480
      TabIndex        =   67
      Top             =   7440
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
      MICON           =   "frmInjectDataBookingBP.frx":023C
      PICN            =   "frmInjectDataBookingBP.frx":0258
      PICH            =   "frmInjectDataBookingBP.frx":6B8B
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
      MICON           =   "frmInjectDataBookingBP.frx":D842
      PICN            =   "frmInjectDataBookingBP.frx":D85E
      PICH            =   "frmInjectDataBookingBP.frx":13E57
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
      MICON           =   "frmInjectDataBookingBP.frx":1A706
      PICN            =   "frmInjectDataBookingBP.frx":1A722
      PICH            =   "frmInjectDataBookingBP.frx":20F25
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
      Top             =   7080
      Visible         =   0   'False
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
      Top             =   7080
      Visible         =   0   'False
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
      Top             =   7080
      Visible         =   0   'False
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
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   28
      Left            =   3240
      TabIndex        =   81
      Top             =   4920
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      _Version        =   393216
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
   Begin VB.Label keteranganBP 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   10200
      TabIndex        =   82
      Top             =   4320
      Width           =   3615
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Waktu Booking"
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
      Left            =   855
      TabIndex        =   80
      Top             =   5040
      Width           =   1065
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pekerjaan"
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
      Left            =   1260
      TabIndex        =   77
      Top             =   4080
      Width           =   720
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
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kode SA BP"
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
      Left            =   1155
      TabIndex        =   56
      Top             =   3765
      Width           =   825
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7575
      Left            =   0
      Picture         =   "frmInjectDataBookingBP.frx":27ACB
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmInjectDataBookingBP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmInjectDataBookingBP
' DateTime  : 05/12/2019 14:53
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

Private rsTemp      As ADODB.Recordset
Private dealerBP    As String
Private nopol       As String



Private Sub hapusForm()
    For i = 0 To 27
        Text(i) = ""
    Next
    For i = 1 To 7
        txt(i).Text = ""
    Next
    txtPekerjaan.Text = ""
    txtNama.Text = ""
    txtNoTelp.Text = ""
    ckReferensi.value = vbUnchecked
    Frame2.Visible = False
    tglBooking.value = Format(isnow, "dd/mm/yyyy")
'    If Log.hansaRenkeiBP = 1 Then
'        Label3.Visible = True
'        Text(24).Visible = True
'        MaskEdBox1.Visible = True
'        Text(25).Visible = True
'        Frame2.Visible = False
'        ckReferensi.Visible = True
'    Else
'        Label3.Visible = False
'        Text(24).Visible = False
'        MaskEdBox1.Visible = False
'        Text(25).Visible = False
'        Frame2.Visible = False
'        ckReferensi.Visible = False
'    End If
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
          
9         If NulSID(txtPekerjaan.Text) = "" Then
10            MsgBox "Data pekerjaan harus diisi.", vbInformation
11            Exit Sub
12        ElseIf Len(NulSID(txtPekerjaan.Text)) > 200 Then
13            MsgBox "Data pekerjan melebihi 200 karakter.", vbInformation
14            Exit Sub
15        End If
           
16        If NulSID(Text(21)) <> "" Then
          
17            Set rsTemp = objO.FetchData("select nik as KodeSales,[User_name] as NamaSales from [replbp99]..[user] where nik = '" & NulSID(Text(21)) & "' and dealer = '" & dealerBP & "'")
18            If Not rsTemp.EOF Then
19                Text(22) = NulSID(rsTemp!namasales)
20            Else
21                MsgBox "Data SA BP tidak ditemukan.", vbInformation
22                Text(21) = ""
23                Exit Sub
                  
24            End If
25        Else
26            Exit Sub
27        End If
          
          Dim kodepr      As String
          Dim grupCust    As Integer
          Dim namapr      As String
          Dim namaLengkap As String
          Dim kodeProsp   As String
           
28        If Log.flagOffBookingLocal = 1 Then
29            Call objHT.OpenTransactionImportHO
30        Else
31            Call objHT.OpenTransactionBengkel
32        End If
33        erTranAll = False
34        isBegin = True
          
35        kodepr = "TBP/" & year(isnow) & "/" & Format(Month(isnow), "00") & "/"
36        If Log.flagOffBookingLocal = 1 Then
37            Set rsTemp = objHT.FetchDataHO("select top 1 Convert(Numeric(19,0),Right(IsNull([NoBookingApps],0),4)) nomor from [BP99].dbo.BOOKINGONLINE where DEALER = '" & dealerBP & "' and [NoBookingApps] like 'TBP/" & year(isnow) & "/" & Format(Month(isnow), "00") & "%' order by Convert(Numeric(19,0),Right(IsNull(NoBookingApps,0),4)) desc")
38        Else
39            Set rsTemp = objHT.FetchDataBengkel("select top 1 Convert(Numeric(19,0),Right(IsNull([NoBookingApps],0),4)) nomor from [REPLBP99]..BOOKINGONLINE where DEALER = '" & dealerBP & "' and [NoBookingApps] like 'TBP/" & year(isnow) & "/" & Format(Month(isnow), "00") & "%' order by Convert(Numeric(19,0),Right(IsNull(NoBookingApps,0),4)) desc")
40        End If
41        If Not rsTemp.EOF Then
42            kodepr = kodepr & Format(rsTemp!nomor + 1, "000#")
43        Else
44            kodepr = kodepr & "0001"
45        End If
          
46        If Log.flagOffBookingLocal = 1 Then
47            sSQL = "INSERT INTO [BP99].dbo.BOOKINGONLINE ([id],[Nopol],[Nama],[Telp]"
48            sSQL = sSQL & " ,[Email],[NamaModel],[Dealer],[Kota],[TglBooking],[JenisService],[ServiceLain],[Keluhan]"
49            sSQL = sSQL & " ,[Transfer],[Proses],[TglEntry],[TransferHO],[HasilFU],[NoBookingApps],[KodePromoApp],[NoBookCarsWorld])"
50            sSQL = sSQL & " VALUES (ISNULL((select max(id)+1 noId from [BP99].dbo.BOOKINGONLINE),1),'" & nopol & "','" & NulSID(Text(2)) & " " & NulSID(Text(3)) & " " & NulSID(Text(4)) & "','" & NulSID(Text(7)) & "'"
51            sSQL = sSQL & " ,'','" & Left(NulSID(Text(13)), 20) & "','" & dealerBP & "','" & NulSID(txt(2).Text) & "','" & Format(tglBooking.Text, "yyyy-mm-dd") & "','Service BP form GR','','" & txtPekerjaan.Text & "'"
52            sSQL = sSQL & " ,0,0,getdate(),0,'','" & kodepr & "','','SAGR:" & gNIK & " - " & Left(gNama_SA, 68) & "SABP:" & NulSID(Text(21)) & "')"
53            Call objHT.InsertData_Ho(sSQL)
          
54        Else
55            sSQL = "INSERT INTO [REPLBP99]..BOOKINGONLINE ([id],[Nopol],[Nama],[Telp]"
56            sSQL = sSQL & " ,[Email],[NamaModel],[Dealer],[Kota],[TglBooking],[JenisService],[ServiceLain],[Keluhan]"
57            sSQL = sSQL & " ,[Transfer],[Proses],[TglEntry],[TransferHO],[HasilFU],[NoBookingApps],[KodePromoApp],[NoBookCarsWorld])"
58            sSQL = sSQL & " VALUES (ISNULL((select max(id)+1 noId from [REPLBP99]..BOOKINGONLINE),1),'" & nopol & "','" & NulSID(Text(2)) & " " & NulSID(Text(3)) & " " & NulSID(Text(4)) & "','" & NulSID(Text(7)) & "'"
59            sSQL = sSQL & " ,'','" & Left(NulSID(Text(13)), 20) & "','" & dealerBP & "','" & NulSID(txt(2).Text) & "','" & Format(tglBooking.Text, "yyyy-mm-dd") & "','Service BP form GR','','" & txtPekerjaan.Text & "'"
60            sSQL = sSQL & " ,0,0,getdate(),0,'','" & kodepr & "','','SAGR:" & gNIK & " - " & Left(gNama_SA, 68) & "SABP:" & NulSID(Text(21)) & "')"
61            Call objHT.InsertData_Bengkel(sSQL)
62        End If
           
63        sSQL = "INSERT INTO [" & dbReplikasiHO & "]..NotifHansaRenkei (KodeSumber,Tipe,Nowo,NoRangka,KodeDealer,KodePesan,KodeSales,NIK,TglNotif,Proses,UserEntry,TglEntry,Keterangan)"
64        sSQL = sSQL & " VALUES ('GRP','TOBP','" & NulSID(kodepr) & "','" & NulSID(Text(21)) & "','" & Right(gDb, 2) & "','SAGR:" & gNIK & " - " & Left(gNama_SA, 68) & "SABP:" & NulSID(Text(21)) & "','','','" & Format(tglNotif, "YYYY-MM-DD HH:MM") & "',0,'" & gSA & "',GETDATE(),'" & gNama_SA & "#" & NulSID(Text(2)) & " " & NulSID(Text(3)) & " " & NulSID(Text(4)) & "#" & NulSID(Text(13)) & "#" & NulSID(nopol) & "#" & NulSID(Text(7)) & "')"
65        Call objHT.InsertData_Bengkel(sSQL)
           
          ''INSERT INTO NotifHansaRenkei (KodeSumber,Tipe,Nowo,NoRangka,KodeDealer,KodePesan,KodeSales,NIK,TglNotif,Proses,UserEntry,TglEntry,Keterangan)
          ''VALUES ('GRP','TOBP','NoBookingApps','NoRangka','kodedatabase','NoBookCarsWorld','','NIKsales',GETDATE(),0,'UserLog',GETDATE(),'NAMA_SA#NAMA_CUST#MODEL#NOPOL#TELP')

66        If erTranAll = False Then
67            If Log.flagOffBookingLocal = 1 Then
68                Call objHT.CloseTransactionImportHO
69            Else
70                Call objHT.CloseTransactionBengkel
71            End If
72            Call hapusForm
73            MsgBox "Data inject berhasil disimpan.", vbInformation
74            Exit Sub
75        Else
76            If Log.flagOffBookingLocal = 1 Then
77                Call objHT.CloseTransactionImportHO_err
78            Else
79                Call objHT.CloseTransactionBengkel_err
80            End If
81            Call hapusForm
82            MsgBox "Data inject gagal disimpan.", vbInformation
83            Exit Sub
84        End If

85       On Error GoTo 0
86       Exit Sub

cmdSimpan_Click_Error:
87        If isBegin Then
88            If Log.flagOffBookingLocal = 1 Then
89                Call objHT.CloseTransactionImportHO_err
90            Else
91                Call objHT.CloseTransactionBengkel_err
92            End If
93        End If
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
94        varErlError = Erl
95        varErrNumber = err.Number
96        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmInjectDataBookingBP")
          
End Sub

Private Sub Form_Load()
    Set rsTemp = objO.FetchData("SELECT dealerbp FROM [REPLBP99]..AllowanceDealer WHERE DEALERGRP = '" & gNmCbg & "'")
    If Not rsTemp.EOF Then
        dealerBP = objO.FetchData_String("select FS_KD_DEALER from ReplBengkel99..T_DEALER where FS_LOG_DEL = '' and FS_NM_DEALER like '%" & NulSID(rsTemp!dealerBP) & "%'")
        keteranganBP.Caption = "Dealer BP : " & objO.FetchData_String("select FS_NM_DEALER from ReplBengkel99..T_DEALER where FS_LOG_DEL = '' and FS_NM_DEALER like '%" & NulSID(rsTemp!dealerBP) & "%'")
    Else
        MsgBox "Dealer belum terdaftar dalam hansa renkei BP", vbInformation
        Call CmdKeluar_Click
    End If
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
42                        nopol = NulSID(rsTemp!NO_POLISI)
43                        Text(16) = Trim(nDpn1)
44                        Text(17) = Trim(nTgh1)
45                        Text(18) = Trim(nBlk1)
46                        Text(19) = NulSID(rsTemp!No_Engine)
47                        Text(20) = NulSID(rsTemp!KEYNO)
48                    Else
49                        MsgBox "Data tidak ditemukan dalam singleid customer.", vbInformation
50                        Call hapusForm
51                    End If
52                Else
53                    Text(Index) = gMainKeyCari
54                End If
                  
55                gMainKeyCari = ""
56                DiTab = True
57            End If
58        End If
59        If DiTab Then
60            Sendkeys ("{TAB}")
61        Else
62            Sendkeys ("{HOME}+{END}")
63        End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 21
                gMainKeyCari = Text(Index)
                gSearchOption = "SABP"
                gBantuan = dealerBP
                frmCariSPK.show (vbModal)
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
                Set oRs = objO.FetchData("select nik as KodeSales,[User_name] as NamaSales from [replbp99]..[user] where nik = '" & NulSID(Text(21)) & "' and dealer = '" & dealerBP & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SABP"
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

