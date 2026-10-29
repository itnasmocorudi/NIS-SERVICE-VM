VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmHOTargetUnit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pengisian Target Semester"
   ClientHeight    =   9270
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   11565
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   5535
      Left            =   120
      TabIndex        =   30
      Top             =   3360
      Width           =   11295
      Begin VB.ComboBox cboLihatTahun 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   2175
      End
      Begin TrueOleDBGrid70.TDBGrid GridTarget 
         Height          =   4695
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8281
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   14215660
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=192,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(38)  =   "Named:id=33:Normal"
         _StyleDefs(39)  =   ":id=33,.parent=0"
         _StyleDefs(40)  =   "Named:id=34:Heading"
         _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(42)  =   ":id=34,.wraptext=-1"
         _StyleDefs(43)  =   "Named:id=35:Footing"
         _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(45)  =   "Named:id=36:Selected"
         _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(47)  =   "Named:id=37:Caption"
         _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(49)  =   "Named:id=38:HighlightRow"
         _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(51)  =   "Named:id=39:EvenRow"
         _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(53)  =   "Named:id=40:OddRow"
         _StyleDefs(54)  =   ":id=40,.parent=33"
         _StyleDefs(55)  =   "Named:id=41:RecordSelector"
         _StyleDefs(56)  =   ":id=41,.parent=34"
         _StyleDefs(57)  =   "Named:id=42:FilterBar"
         _StyleDefs(58)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   5640
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
         Index           =   16
         Left            =   7560
         TabIndex        =   38
         Top             =   240
         Width           =   3465
         _ExtentX        =   6112
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
         Left            =   7470
         TabIndex        =   39
         Top             =   240
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
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
         Left            =   5040
         TabIndex        =   35
         Top             =   285
         Width           =   465
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
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
         TabIndex        =   31
         Top             =   300
         Width           =   450
      End
   End
   Begin VB.CommandButton TblKeluar 
      Caption         =   "Tutup"
      Height          =   375
      Left            =   120
      TabIndex        =   29
      Top             =   8880
      Width           =   11295
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   3255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   11295
      Begin VB.CommandButton TblHapus 
         Caption         =   "&Hapus"
         Height          =   375
         Left            =   9600
         TabIndex        =   14
         Top             =   2760
         Width           =   1575
      End
      Begin VB.CommandButton TblSimpan 
         Caption         =   "&Tambah"
         Height          =   375
         Left            =   8040
         TabIndex        =   13
         Top             =   2760
         Width           =   1575
      End
      Begin VB.ComboBox cboTahun 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">99,9"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
         Mask            =   ">9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   5
         Top             =   1800
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
         Mask            =   ">9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
         Mask            =   ">9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   7
         Top             =   2520
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
         Mask            =   ">9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   1320
         TabIndex        =   8
         Top             =   2880
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
         Mask            =   ">9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   5880
         TabIndex        =   10
         Top             =   1440
         Width           =   1785
         _ExtentX        =   3149
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
         Mask            =   ">999999999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   5880
         TabIndex        =   11
         Top             =   1800
         Width           =   1785
         _ExtentX        =   3149
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
         Mask            =   ">999999999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   5880
         TabIndex        =   12
         Top             =   2160
         Width           =   1785
         _ExtentX        =   3149
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
         Mask            =   ">999999999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
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
         Mask            =   ">9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   5880
         TabIndex        =   9
         Top             =   1080
         Width           =   1785
         _ExtentX        =   3149
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
         Mask            =   ">999999999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   5880
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
         Index           =   12
         Left            =   5880
         TabIndex        =   36
         Top             =   600
         Width           =   3105
         _ExtentX        =   5477
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   7710
         TabIndex        =   37
         Top             =   240
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
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
         Left            =   5250
         TabIndex        =   34
         Top             =   300
         Width           =   465
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Jasa"
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
         Left            =   4860
         TabIndex        =   33
         Top             =   1125
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Sublet"
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
         Left            =   4740
         TabIndex        =   28
         Top             =   2205
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Part && Oli Non Toyota"
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
         Left            =   3630
         TabIndex        =   27
         Top             =   1845
         Width           =   2085
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Part && Oli Toyota"
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
         TabIndex        =   26
         Top             =   1485
         Width           =   1755
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target RTJ"
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
         TabIndex        =   25
         Top             =   2925
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target PDS"
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
         TabIndex        =   24
         Top             =   2565
         Width           =   810
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TargetTWC"
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
         Left            =   300
         TabIndex        =   23
         Top             =   2205
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target SBI"
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
         TabIndex        =   22
         Top             =   1845
         Width           =   765
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target GRP"
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
         Left            =   300
         TabIndex        =   21
         Top             =   1485
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target SBE"
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
         TabIndex        =   20
         Top             =   1125
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hari Kerja"
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
         Left            =   420
         TabIndex        =   19
         Top             =   645
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tahun"
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
         Left            =   675
         TabIndex        =   18
         Top             =   300
         Width           =   450
      End
   End
End
Attribute VB_Name = "frmHOTargetUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As New ADODB.Recordset
Dim rsGrid As New ADODB.Recordset

Private Sub hapusForm()

    For i = 0 To 10
        Select Case i
            Case 0
                Text(i) = "00,0"
            Case Else
                Text(i) = 0
        End Select
    Next
    
    ' ISI COMBO
    For i = 1 To 10
        For j = 1 To 12
            cboTahun.AddItem GetNamaBulan(j) & " | " & Year(gTglServer) - 5 + i
'            cboLihatTahun.AddItem GetNamaBulan(j) & " | " & Year(gTglServer) - 5 + i
        Next
        cboLihatTahun.AddItem Year(gTglServer) - 5 + i
    Next
    cboTahun = GetNamaBulan(Month(gTglServer)) & " | " & Year(gTglServer)
    cboLihatTahun = Year(gTglServer)    'GetNamaBulan(Month(gTglServer)) & " | " &
End Sub

Private Sub refreshGrid(cbo As Object)
Dim Thn As Integer
Dim Bln As Integer

    Thn = CDbl(cbo)
    'Mid(Trim(cbo), InStr(Trim(cbo), "|") + 2, Len(Trim(cbo)) - InStr(Trim(cbo), "|") - 1)
'    blN = GetAngkaBulan(Trim(Left(Trim(cbo), InStr(Trim(cbo), " |"))))
    
    GridTarget.ClearFields
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = oConnHO.Execute("SELECT CASE WHEN BULAN = 1 THEN 'JANUARI' WHEN BULAN = 2 THEN 'FEBRUARI' WHEN BULAN = 3 THEN 'MARET' WHEN BULAN = 4 THEN 'APRIL' WHEN BULAN = 5 THEN 'MEI' " _
                               & "WHEN BULAN = 6 THEN 'JUNI' WHEN BULAN = 7 THEN 'JULI' WHEN BULAN = 8 THEN 'AGUSTUS' WHEN BULAN = 9 THEN 'SEPTEMBER' WHEN BULAN = 10 THEN 'OKTOBER' " _
                               & "WHEN BULAN = 11 THEN 'NOVEMBER' WHEN BULAN = 12 THEN 'DESEMBER' END AS BULAN, HARIKERJA, SBE, GRP, SBI, TWC, PDS, RTJ, JASA, PARTGENUINE, PARTNONGENUINE, SUBLET " _
                               & "FROM TARGETGRP WHERE TAHUN = '" & CDbl(Thn) & "' AND FS_KD_DEALER = '" & Text(15) & "' ORDER BY BULAN")
    If Not rsGrid.EOF Then
        GridTarget.DataSource = rsGrid
        Call gridHeader
    End If
End Sub

Private Sub gridHeader()
    
    GridTarget.Columns(0).Caption = "BULAN"
    GridTarget.Columns(0).Width = 1000
    GridTarget.Columns(0).DataField = "BULAN"
    GridTarget.Columns(1).Caption = "HARI KERJA"
    GridTarget.Columns(1).Width = 800
    GridTarget.Columns(1).DataField = "HARIKERJA"
    GridTarget.Columns(2).Caption = "TARGET SBE"
    GridTarget.Columns(2).Width = 2000
    GridTarget.Columns(2).DataField = "SBE"
    GridTarget.Columns(3).Caption = "TARGET GRP"
    GridTarget.Columns(3).Width = 2000
    GridTarget.Columns(3).DataField = "GRP"
    GridTarget.Columns(4).Caption = "TARGET SBI"
    GridTarget.Columns(4).Width = 2000
    GridTarget.Columns(4).DataField = "SBI"
    GridTarget.Columns(5).Caption = "TARGET TWC"
    GridTarget.Columns(5).Width = 2000
    GridTarget.Columns(5).DataField = "TWC"
    GridTarget.Columns(6).Caption = "TARGET PDS"
    GridTarget.Columns(6).Width = 2000
    GridTarget.Columns(6).DataField = "PDS"
    GridTarget.Columns(7).Caption = "REVENUE JASA"
    GridTarget.Columns(7).Width = 2500
    GridTarget.Columns(7).DataField = "JASA"
    GridTarget.Columns(8).Caption = "REVENUE PART&OLI TOYOTA"
    GridTarget.Columns(8).Width = 2500
    GridTarget.Columns(8).DataField = "PARTGENUINE"
    GridTarget.Columns(9).Caption = "REVENUE OLI&BAHAN NON TOYOTA"
    GridTarget.Columns(9).Width = 2500
    GridTarget.Columns(9).DataField = "PARTNONGENUINE"
    GridTarget.Columns(10).Caption = "REVENUE SUBLET"
    GridTarget.Columns(10).Width = 2500
    GridTarget.Columns(10).DataField = "SUBLET"
End Sub

Private Sub cboLihatTahun_Click()
    If TblSimpan.Caption = "&Simpan" Then
        Call refreshGrid(cboLihatTahun)
    End If
End Sub

Private Sub cboLihatTahun_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cboTahun_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub Form_Load()
    Call hapusForm
    Call refreshGrid(cboLihatTahun)
End Sub

Private Sub GridTarget_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Text(15) <> "" Then
        With GridTarget
            If rsGrid.RecordCount > 0 Then
                cboTahun = Trim(.Columns(0)) & " | " & CDbl(cboLihatTahun)
                Text(11) = Trim(Text(15))
                Text(12) = Trim(Text(16))
                Text(0) = Format(Replace(.Columns(1), ".", ","), "00.0")
                Text(1) = .Columns(2)
                Text(2) = .Columns(3)
                Text(3) = .Columns(4)
                Text(4) = .Columns(5)
                Text(5) = .Columns(6)
                Text(7) = .Columns(7)
                Text(8) = .Columns(8)
                Text(9) = .Columns(9)
                Text(10) = .Columns(10)
            End If
        End With
    End If
End Sub

Private Sub tblKeluar_Click()
    Unload Me
End Sub

Private Sub TblSimpan_Click()
Dim Thn As Integer
Dim Bln As Integer

    If TblSimpan.Caption = "&Tambah" Then
        Call hapusForm
        Text(0).SetFocus
        TblSimpan.Caption = "&Simpan"
    ElseIf TblSimpan.Caption = "&Simpan" Then
        Thn = Mid(Trim(cboTahun), InStr(Trim(cboTahun), "|") + 2, Len(Trim(cboTahun)) - InStr(Trim(cboTahun), "|") - 1)
        Bln = GetAngkaBulan(Trim(Left(Trim(cboTahun), InStr(Trim(cboTahun), " |"))))
        
        If rsTemp.State Then rsTemp.Close
        Set rsTemp = oConnHO.Execute("SELECT * FROM TARGETGRP WHERE TAHUN = '" & CDbl(Thn) & "' AND BULAN = '" & CDbl(Bln) & "' AND FS_KD_DEALER = '" & Text(11) & "'")
        If Not rsTemp.EOF Then
            If MsgBox("Akan Merubah Data ?", vbYesNo, "UPDATE TARGET") = vbYes Then
                oConnHO.Execute ("UPDATE TARGETGRP SET HARIKERJA = '" & Replace(CDbl(Text(0)), ",", ".") & "', " _
                           & "SBE = '" & CDbl(Text(1)) & "', GRP = '" & CDbl(Text(2)) & "', SBI = '" & CDbl(Text(3)) & "', TWC = '" & CDbl(Text(4)) & "', PDS = '" & CDbl(Text(5)) & "', " _
                           & "RTJ = '" & CDbl(Text(6)) & "', JASA = " & CDbl(Text(7)) & ", PARTGENUINE = " & CDbl(Text(8)) & ", PARTNONGENUINE = " & CDbl(Text(9)) & ", SUBLET = " & CDbl(Text(10)) & " " _
                           & "WHERE TAHUN = '" & CDbl(Thn) & "' AND BULAN = '" & CDbl(Bln) & "' AND FS_KD_DEALER = '" & Text(11) & "'")
            End If
        Else
            oConnHO.Execute ("INSERT INTO TARGETGRP VALUES ('" & CDbl(Thn) & "','" & CDbl(Bln) & "','" & Replace(CDbl(Replace(Text(0), ">", "")), ",", ".") & "', '" & Trim(Text(11)) & "', " _
                           & "'" & Trim(Text(12)) & "', '" & CDbl(Text(1)) & "', '" & CDbl(Text(2)) & "', '" & CDbl(Text(3)) & "', '" & CDbl(Text(4)) & "', '" & CDbl(Text(5)) & "', " _
                           & "'" & CDbl(Text(6)) & "', " & CDbl(Text(7)) & ", " & CDbl(Text(8)) & ", " & CDbl(Text(9)) & ", " & CDbl(Text(10)) & ")")
        End If
        rsTemp.Close
        Call refreshGrid(cboLihatTahun)
        
        TblSimpan.Caption = "&Tambah"
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

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 11, 15
                gMainKeyCari = Text(Index)
                gBantuan = "INTERNAL"
                gSearchOption = "DEALER"
                frmCariHO.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 0
            If Text(Index) = "" Then
                Text(Index) = 0
            End If
    End Select
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                If Replace(Replace(Text(Index), ">", ""), ",", ".") > CDbl(Left(GetTglAkhir(Month(gTglServer), Year(gTglServer), 0), InStr(GetTglAkhir(Month(gTglServer), Year(gTglServer), 0), "-") - 1)) Then
                    MsgBox "Jumlah Hari tidak bisa lebih besar dari " & CDbl(Left(GetTglAkhir(Month(gTglServer), Year(gTglServer), 0), InStr(GetTglAkhir(Month(gTglServer), Year(gTglServer), 0), "-") - 1))
                    Text(Index) = "00,0"
                    SendKeys ("+{TAB}")
                    Exit Sub
                End If
            Else
                Text(Index) = 0
                Text(Index).SetFocus
                Exit Sub
            End If
        Case 11, 15
            If Text(Index) <> "" Then
                sSQL = "SELECT * FROM T_DEALER WHERE INTERNAL = 1 AND FS_KD_DEALER = '" & Text(Index) & "'"
                oRs.Open sSQL, oConnHO, adOpenStatic
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gBantuan = "INTERNAL"
                    gSearchOption = "DEALER"
                    frmCariHO.Show (vbModal)
                Else
                    Text(Index + 1) = oRs("FS_NM_DEALER")
                    If Index = 15 Then
                        Call refreshGrid(cboLihatTahun)
                    End If
                End If
                oRs.Close
            End If
    End Select
End Sub
