VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmHOInsentifBobotSparepart 
   BackColor       =   &H00FFFFFF&
   Caption         =   "BOBOT SPAREPART"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12930
   ScaleWidth      =   21360
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12840
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   6495
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   12015
      Begin VB.Frame frmeGanti 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   96000
         TabIndex        =   11
         Top             =   1920
         Visible         =   0   'False
         Width           =   10335
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   12
            Top             =   765
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
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
            Index           =   8
            Left            =   915
            TabIndex        =   13
            Top             =   405
            Width           =   2385
            _ExtentX        =   4207
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   9
            Left            =   6675
            TabIndex        =   15
            Top             =   720
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Index           =   10
            Left            =   6675
            TabIndex        =   14
            Top             =   360
            Width           =   2385
            _ExtentX        =   4207
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
         Begin AIFCmp1.asxPowerButton cmdProses 
            Height          =   360
            Left            =   7965
            TabIndex        =   16
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            BackColor       =   16761024
            BorderStyle     =   3
            Caption         =   "&Proses"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   4680
            Picture         =   "frmHOInsentifBobotSparepart.frx":0000
            Stretch         =   -1  'True
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Baru"
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
            Left            =   5880
            TabIndex        =   18
            Top             =   405
            Width           =   675
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   3315
            Top             =   405
            Width           =   60
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Lama"
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
            Left            =   120
            TabIndex        =   17
            Top             =   450
            Width           =   720
         End
      End
      Begin TrueOleDBGrid70.TDBGrid GridPart 
         Height          =   4335
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   14933984
         RowDividerColor =   14933984
         RowSubDividerColor=   14933984
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H83F3EE&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Named:id=33:Normal"
         _StyleDefs(41)  =   ":id=33,.parent=0"
         _StyleDefs(42)  =   "Named:id=34:Heading"
         _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(44)  =   ":id=34,.wraptext=-1"
         _StyleDefs(45)  =   "Named:id=35:Footing"
         _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(47)  =   "Named:id=36:Selected"
         _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(49)  =   "Named:id=37:Caption"
         _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(51)  =   "Named:id=38:HighlightRow"
         _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=39:EvenRow"
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
         _StyleDefs(61)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(62)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(63)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(64)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   0
         Top             =   240
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   6
         Left            =   4920
         TabIndex        =   7
         Top             =   240
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   3120
         TabIndex        =   22
         Top             =   960
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "99,99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   3120
         TabIndex        =   24
         Top             =   600
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   2
         Left            =   4920
         TabIndex        =   25
         Top             =   600
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Index           =   5
         Left            =   3120
         TabIndex        =   27
         Top             =   1320
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "99,99"
         PromptChar      =   ">"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bobot TMS"
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
         Left            =   1935
         TabIndex        =   28
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sparepart - Kode"
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
         Left            =   1710
         TabIndex        =   26
         Top             =   645
         Width           =   1230
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4800
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bobot"
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
         Left            =   2280
         TabIndex        =   23
         Top             =   960
         Width           =   420
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4800
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posisi - Kode"
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
         Left            =   2070
         TabIndex        =   8
         Top             =   285
         Width           =   900
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7275
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   12832
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "  Tabel Insentif Bobot Sparepart"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   4194304
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   503
      BackColor       =   16777215
      HighlightColor  =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   7920
      TabIndex        =   2
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Tambah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdUbah 
      Height          =   360
      Left            =   8880
      TabIndex        =   4
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Ubah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   9840
      TabIndex        =   5
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdKeluar 
      Height          =   360
      Left            =   10800
      TabIndex        =   6
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdDownloadExcel 
      Height          =   360
      Left            =   6480
      TabIndex        =   9
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Download Excel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdGanti 
      Height          =   360
      Left            =   480
      TabIndex        =   10
      Top             =   7080
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Ganti Part"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   5000
      TabIndex        =   19
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   12648447
      Caption         =   "Import Excel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ARProgBarCtrl.ARProgressBar ProgBar 
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   7080
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   450
      Value           =   0
      ForeColor       =   52479
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionForeColor=   128
      IniColor        =   52479
   End
   Begin TrueOleDBGrid70.TDBGrid gridDownload 
      Height          =   1215
      Left            =   4680
      TabIndex        =   21
      Top             =   7440
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2143
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NAMA PEKERJAAN"
      Columns(0).DataField=   "NAMAPEKERJAAN"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "KODE MODEL"
      Columns(1).DataField=   "KODE MODEL"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "KODE SUBLET"
      Columns(2).DataField=   "KODESUBLET"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "QTY"
      Columns(3).DataField=   "QTY"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "KETERANGAN"
      Columns(4).DataField=   "KETERANGAN"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   14933984
      RowDividerColor =   14933984
      RowSubDividerColor=   14933984
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=32,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H83F3EE&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
      _StyleDefs(77)  =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(78)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(79)  =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(80)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
   End
End
Attribute VB_Name = "frmHOInsentifBobotSparepart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOInsentifBobotSparepart
' DateTime  : 13/05/2020 10:01
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

 
Dim rsGrid      As New ADODB.Recordset
Dim rsDownload As New ADODB.Recordset
Dim rsModel     As ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim newRecord   As Boolean
Dim KDPART_LAMA As String
Dim boleh       As Boolean
Dim pesan       As String
Dim cols        As TrueOleDBGrid70.Columns

Private Sub ber_sih()
    For i = 0 To 6
        Select Case i
            Case 5
                Text(i) = ">>,>>"
                Text(i).Mask = "99,99"
            Case 3
                Text(i) = ">>,>>"
                Text(i).Mask = "99,99"
            Case Else
                Text(i) = ""
        End Select
    Next
    'cboPerhitungan.ListIndex = 0
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    Text(0).Enabled = Flg
    Text(1).Enabled = Flg
    Text(2).Enabled = False
    Text(3).Enabled = flg1
    Text(5).Enabled = flg1
    Text(6).Enabled = False
End Sub

Private Sub refreshGrid()
    GridPart.ClearFields
    Set rsGrid = objO.FetchDataHO("select a.KdTipe,c.Tipe,a.BobotPart,a.BobotPartTMS,a.KdItem,isnull(b.NamaBarang,'') NamaBarang from INSENTIF..GRMstBobotPart a inner join INSENTIF..GRTipeMP c on c.KdTipe = a.KdTipe LEFT JOIN " & gAlamatDepo & "[depo00-" & Format(Now(), "yyyy") & Format(Now(), "mm") & "].dbo.inmst b ON b.KodeItem = a.KdItem where ISNULL(a.UserDel,'') = ''")
    If Not rsGrid.EOF Then
        GridPart.DataSource = rsGrid
        'Call HeaderColom
    End If
End Sub

Private Sub HeaderColom()
    GridPart.Columns(0).Caption = "Kode Tipe"
    GridPart.Columns(0).Width = 2000
    GridPart.Columns(1).Caption = "Tipe"
    GridPart.Columns(1).Width = 3500
    GridPart.Columns(2).Caption = "Bobot Posisi"
    GridPart.Columns(2).Width = 2000
    GridPart.Columns(3).Caption = "Bobot Result"
    GridPart.Columns(3).Width = 2000
    GridPart.Columns(4).Caption = "Bobot Process"
    GridPart.Columns(4).Width = 2000
End Sub


Private Sub ProsesExport(onV As String)
          Dim oApp    As Excel.Application
          Dim oWB     As Excel.Workbook
          Dim nAsc    As Integer
1         Set oApp = New Excel.Application
2         oApp.Visible = False
3         Set oWB = oApp.Workbooks.Add
4         With oWB
      '        For C = 0 To rsDownload.Fields.Count - 1
      '            .Sheets("Sheet1").Cells(1, C + 1).value = rsDownload.Fields(C).Name
      '            nAsc = nAsc + 1
      '        Next
        
5       .Sheets("Sheet1").Cells(1, 1).value = "POSISI (Max 50 kar - sesuai master Posisi)"
6       .Sheets("Sheet1").Cells(1, 2).value = "KODE PART"
7       .Sheets("Sheet1").Cells(1, 3).value = "NAMA PART"
8       .Sheets("Sheet1").Cells(1, 4).value = "BOBOT PART (gunakan koma untuk desimal)"
9       .Sheets("Sheet1").Cells(1, 5).value = "BOBOT PART TMS (gunakan koma untuk desimal)"
        '.Sheets("Sheet1").Cells(1, 5).value = "KETERANGAN (max 30 kar)"
          

10      .Sheets("Sheet1").Cells(2, 1).CopyFromRecordset rsDownload
11      .Sheets("Sheet1").Range("1:1").Font.Bold = True
12      .Sheets("Sheet1").Select
13      .Close SaveChanges:=True, FileName:=onV
14        End With
15        Set oWB = Nothing
16        oApp.Quit
17        Set oApp = Nothing
End Sub
 

Private Sub cmdDownloadExcel_Click()
          
1        On Error GoTo cmdDownloadExcel_Click_Error
          Dim Er As New ErrInfo
2         If rsDownload.State Then rsDownload.Close
3         Set rsDownload = Nothing
4         With rsDownload
5             .Fields.Append "POSISI", adVarChar, 50
6             .Fields.Append "BOBOT INSENTIF", adVarChar, 50
7             .Fields.Append "BOBOT RESULT", adVarChar, 50
8             .Fields.Append "BOBOT PROSES", adVarChar, 50
9             .Open
      '
      '        .AddNew
      '        !namaPekerjaan = "NAMA PEKERJAAN"
      '        !KODEMODEL = "KODE MODEL"
      '        !KODESUBLET = "KODE SUBLET"
      '        !qty = "QTY SUBLET"
      '        !keterangan = "KETERANGAN"
      '        .Update
10        End With
            
          Dim MyPath As String
          Dim onVal As String
11        MyPath = ""
          Dim f As New FormExplorer
12        f.show (vbModal)
13        MyPath = f.PathSaving
14        If MyPath = "" Then Exit Sub
15        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
16        If f.FileName = "" Then Exit Sub
17        onVal = MyPath & f.FileName & ".xls"
18        Call ProsesExport(onVal)
19                Er.HeadCap = "Export to excel !"
20                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
21                Beep
22                Er.show (vbModal)
          
23       On Error GoTo 0
24       Exit Sub

cmdDownloadExcel_Click_Error:

25        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdDownloadExcel_Click of Form frmHOInsentifBobotSparepart"


End Sub

Private Sub cmdGanti_Click()
    Text(8) = ""
    Text(4) = ""
    Text(10) = ""
    Text(9) = ""
    frmeGanti.Visible = True
End Sub

Private Sub cmdHapus_Click()
          Dim isBegin As Boolean
1        On Error GoTo cmdHapus_Click_Error
2         isBegin = False
3         Set rsTemp = objO.FetchDataHO("SELECT * FROM INSENTIF..GRMstBobotPart WITH (NOLOCK) WHERE KdTipe = '" & Text(0) & "' and KdItem = '" & Text(1) & "'")
4         If Not rsTemp.EOF Then
5             temp = MsgBox("Yakin akan menghapus data ini ?", vbYesNo, "HAPUS DATA")
6             If temp = vbYes Then
7                 Call objHT.OpenTransactionHO
8                 isBegin = True
9                 Call objHT.InsertData_Ho("UPDATE INSENTIF..GRMstBobotPart SET TglDel = getdate(), UserDel = '" & gSA & "' WHERE KdTipe = '" & Text(0) & "' and KdItem = '" & Text(1) & "'")
                  
10                Call objHT.CloseTransactionHO
11            End If
12            Call ber_sih
13            Call refreshGrid
14        Else
15            MsgBox "Pilih data yang ingin dihapus"
16            Exit Sub
17        End If

18       On Error GoTo 0
19       Exit Sub

cmdHapus_Click_Error:
20        If isBegin Then
21            Call objHT.CloseTransactionHO_err
22        End If
23        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdHapus_Click of Form frmHOInsentifBobotSparepart"


End Sub

Private Sub cmdImport_Click()
      Dim ErrBegin As Boolean, konter As Long
1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim rsTempExcel As ADODB.Recordset
         Dim Jml As Integer
         Dim itemPart As String
         Dim isBegin As Boolean
2        isBegin = False
         
        ' Dim konter As Integer
         
3       konter = 0
4       ProgBar.Min = 0
5       CommonDialog1.FileName = ""
6       CommonDialog1.InitDir = App.Path
7       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
8       CommonDialog1.ShowOpen
9       If CommonDialog1.FileName <> "" Then
10         If CnExel.State Then CnExel.Close
            'CnExel.Open "Provider=Microsoft.ACE.OLEDB.12.0;" &
                 ' "Data Source=" & CommonDialog1.FileName & ";" & _
                  '"Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
11         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
12         GoSub Proses
13     Else
14        Exit Sub
15     End If
16     Exit Sub
       
Proses:
17    ErrBegin = True
         Dim VarNmFile As String
         Dim kodePosisi As String
         Dim namaPart   As String
         Dim rsModel  As ADODB.Recordset
18       Call objHT.OpenTransactionHO
19       isBegin = True
20        gBantuan = "SBE"
21       If RcExcel.State Then RcExcel.Close
22       RcExcel.Open "Select * from [Part$]", CnExel, adOpenKeyset
23       If RcExcel.RecordCount > 0 Then
24          ProgBar.Visible = True
25          ProgBar.Max = RcExcel.RecordCount
26          RcExcel.MoveFirst
27          Do While Not RcExcel.EOF
28             DoEvents
29             konter = konter + 1
30             ProgBar.value = konter
31             kodePosisi = ""
32             namaPart = ""
              
33             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
               ''joborder
34             If objO.FetchDataHO_Count("SELECT * FROM fkaryawan..DataKaryawan where (Divisi='service' or posisi like '%ADMINISTRATION SECTION HEAD%' or posisi like '%ADH%' or posisi like '%Branch%' ) and posisi = '" & Trim(RcExcel.Fields(0).value) & "'") = 0 Then
35                Call objHT.CloseTransactionHO_err
36                ProgBar.Visible = False
37                MsgBox "Posisi = " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " salah!", vbCritical, "Error"
38                Exit Sub
39             Else
40                kodePosisi = objO.FetchData_String("SELECT KdTipe FROM INSENTIF..GRTipeMP where ISNULL(UserDel,'') = '' and Tipe = '" & Trim(RcExcel.Fields(0).value) & "'")
41             End If
               
42            If objO.FetchDataHO_Count("select distinct KodeItem,NamaBarang from " & gAlamatDepo & "[depo00-" & Format(Now(), "yyyy") & Format(Now(), "mm") & "].dbo.inmst where KodeItem = '" & Trim(RcExcel.Fields(1).value) & "'") Then
43                Call objHT.CloseTransactionHO_err
44                ProgBar.Visible = False
45                MsgBox "Kode Item part baris ke-" & konter & " tidak terdaftar dalam depo!", vbCritical, "Error"
46                Exit Sub
47            Else
48                namaPart = objO.FetchData_String("select distinct NamaBarang from " & gAlamatDepo & "[depo00-" & Format(Now(), "yyyy") & Format(Now(), "mm") & "].dbo.inmst where KodeItem = '" & Trim(RcExcel.Fields(1).value) & "'")
49            End If
50            If setHurufAngka(Nul(RcExcel.Fields(3).value)) < 0 Then
51                Call objHT.CloseTransactionHO_err
52                ProgBar.Visible = False
53                MsgBox "Bobot Result baris ke-" & konter & " salah!", vbCritical, "Error"
54                Exit Sub
55            End If
56            If setHurufAngka(Nul(RcExcel.Fields(5).value)) < 0 Then
57                Call objHT.CloseTransactionHO_err
58                ProgBar.Visible = False
59                MsgBox "Bobot Result TMS baris ke-" & konter & " salah!", vbCritical, "Error"
60                Exit Sub
61            End If
               
               
62                Set rsTemp = objO.FetchDataHO("SELECT * FROM INSENTIF..GRMstBobotPart WITH (NOLOCK) WHERE KdTipe = '" & Trim(kodePosisi) & "' and KdItem = '" & Trim(RcExcel.Fields(1).value) & "' and isnull(userdel,'') = ''")
                  
63                If rsTemp.EOF Then
64                    Call objHT.InsertData_Ho("INSERT INTO INSENTIF..GRMstBobotPart (KdTipe,KdItem,BobotPart,UserEntry,TglEntry,BobotPartTMS) VALUES ('" & Trim(kodePosisi) & "','" & Trim(RcExcel.Fields(1).value) & "', '" & setHurufAngka(RcExcel.Fields(3).value * 100) & "','" & gSA & "',getdate(), '" & setHurufAngka(RcExcel.Fields(4).value * 100) & "')")
                          
65                Else
66                    Call objHT.UpdateData_Ho("UPDATE INSENTIF..GRMstBobotPart SET BobotPart = '" & setHurufAngka(setHurufAngka(RcExcel.Fields(3).value * 100)) & "',BobotPartTMS = '" & setHurufAngka(RcExcel.Fields(3).value * 100) & "' WHERE KdTipe = '" & Trim(kodePosisi) & "' and KdItem = '" & Trim(RcExcel.Fields(1).value) & "' and isnull(userdel,'') = ''")
                      
67                End If
                 ' Call objHT.InsertData_Ho(sSQL)
                       
68             RcExcel.MoveNext
69          Loop
70       End If
71       RcExcel.Close
72       CnExel.Close
      '   ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
73        Call objHT.CloseTransactionHO
74       MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
75       ProgBar.Visible = False
76       Call ber_sih
77       Call buka_Tutup(False, False)
78       Call refreshGrid
79       cmdUbah.Caption = "&Ubah"
80       cmdSimpan.Caption = "&Tambah"
      '   GridAFI.Refresh
81       Return
82       Exit Sub
vb_error:
83       If ErrBegin = True Then
           ' DbCon.RollbackTrans
84       End If
85       If isBegin Then
86        Call objHT.CloseTransactionHO_err
87       End If
88       ProgBar.Visible = False
89       MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdProses_Click()
    Call objO.InsertDataHO("INSERT INTO HO_PARTPACK  (JOB,FS_ITEM,MDL,MODEL,KDSUBLET,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL) SELECT DISTINCT JOB, FS_ITEM, MDL, MODEL, '" & Trim(Text(10)) & "', '" & Trim(Text(9)) & "', QTY, KETERANGAN, FS_LOG_DEL, FD_TGL_DEL FROM HO_PARTPACK WITH (NOLOCK) WHERE KDSUBLET = '" & Trim(Text(8)) & "'")
    Call objO.UpdateDataHO("UPDATE HO_PARTPACK SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE KDSUBLET = '" & Trim(Text(8)) & "'")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(Text(10)) & "', 'S')")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(Text(8)) & "', 'D')")
    frmeGanti.Visible = False
End Sub

Private Sub cmdSimpan_Click()
          Dim isBegin As Boolean
          Dim kode As String
          Dim perhitungan As Integer
          
1        On Error GoTo cmdSimpan_Click_Error
2         isBegin = False
3         If cmdSimpan.Caption = "&Tambah" Then
4             newRecord = True
5             Call ber_sih
6             Call buka_Tutup
7             Text(0).SetFocus
8             isiId
9             cmdUbah.Caption = "&Batal"
10            cmdSimpan.Caption = "&Simpan"
11        ElseIf cmdSimpan.Caption = "&Simpan" Then
12            boleh = True
13            pesan = "Lengkapi data berikut "
              
      '14            If setHurufAngka(Text(1)) <= 0 Then
      '15                pesan = pesan + "Bobot insentif "
      '16                Text(1) = 0
      '17                Text(1).SetFocus
      '18                boleh = False
      '19            End If
14            If Text(0) = "" Then
15                boleh = False
16                pesan = pesan + "Posisi. "
17            Else
18                Set oRs = objO.FetchDataHO("SELECT * FROM INSENTIF..GRTipeMP where ISNULL(UserDel,'') = '' and KdTipe = '" & Text(0) & "'")
19                If oRs.BOF Then
20                    boleh = False
21                    pesan = pesan + "Posisi. "
22                    Text(0) = ""
23                    Text(6) = ""
24                Else
25                    Text(6) = NulSID(oRs!tipe)
26                End If
27                oRs.Close
28            End If
               
29            If setHurufAngka(Text(3)) < 0 Then
30                boleh = False
31                pesan = pesan + "Proporsi. "
32            End If
               
33            If setHurufAngka(Text(5)) < 0 Then
34                boleh = False
35                pesan = pesan + "Proporsi TMS. "
36            End If
37            If boleh Then
lanjutLooping:
38                Call objHT.OpenTransactionHO
39                isBegin = True
40                If newRecord Then
41                    Call objHT.InsertData_Ho("INSERT INTO INSENTIF..GRMstBobotPart (KdTipe,KdItem,BobotPart,UserEntry,TglEntry,BobotPartTMS) VALUES ('" & Text(0) & "','" & Text(1) & "', '" & setHurufAngka(Text(3)) & "','" & gSA & "',getdate(), '" & setHurufAngka(Text(5)) & "')")
                          
42                Else
43                    Call objHT.UpdateData_Ho("UPDATE INSENTIF..GRMstBobotPart SET BobotPart = '" & setHurufAngka(Text(3)) & "',BobotPartTMS = '" & setHurufAngka(Text(5)) & "' WHERE KdTipe = '" & Text(0) & "' and KdItem = '" & Text(1) & "'")
                      
44                End If
45                Call objHT.CloseTransactionHO
              
46                Call ber_sih
47                Call buka_Tutup(False, False)
48                Call refreshGrid
49                cmdUbah.Caption = "&Ubah"
50                cmdSimpan.Caption = "&Tambah"
51            Else
52                MsgBox pesan, vbCritical, "VALID DATA"
53                Exit Sub
54            End If
55        End If

56       On Error GoTo 0
57       Exit Sub

cmdSimpan_Click_Error:
58        If isBegin Then
59            Call objHT.CloseTransactionHO_err
60        End If
61        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpan_Click of Form frmHOInsentifBobotSparepart"


End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        newRecord = False
        Call buka_Tutup(False, True)
        KDPART_LAMA = Text(1)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
    End If
End Sub

Private Sub Form_Load()
'    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridPart.Width = Frame1.Width - 200
    GridPart.Height = Frame1.Height - Text(3).Top - Text(3).Height - 200
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdDownloadExcel.Top = cmdKeluar.Top
    cmdGanti.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    cmdDownloadExcel.Top = cmdKeluar.Top
    ProgBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdDownloadExcel.Left = cmdSimpan.Left - cmdDownloadExcel.Width - 100
    cmdImport.Left = cmdDownloadExcel.Left - cmdImport.Width - 100
    ProgBar.Left = cmdImport.Left - ProgBar.Width - 100
End Sub

Private Sub GridPart_FilterChange()
    On Error GoTo errHandler
    Set cols = GridPart.Columns
    Dim C As Integer
    
    C = GridPart.col
    GridPart.HoldFields
    rsGrid.filter = getFilter()
    GridPart.col = C
    GridPart.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        
        End If
    Next col
    getFilter = tmp
End Function

Private Sub GridPart_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF Then
        Text(6) = rsGrid.Fields(1)       'GridPart.Columns(0)
        Text(0) = rsGrid.Fields(0)      'GridPart.Columns(1)    'GridPart.Columns(2)
        Text(1) = rsGrid.Fields(4)       'GridPart.Columns(3)
        Text(2) = rsGrid.Fields(5)       'GridPart.Columns(5)
        Text(3) = Format(rsGrid.Fields(2), "00.00")
        Text(5) = Format(rsGrid.Fields(3), "00.00")
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "DATATIPEMP"
                frmCariHO.show (vbModal)
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "PARTDEPO"
                frmCariHO.show (vbModal)
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
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                'Set oRs = objO.FetchDataHO("SELECT * FROM jobOrder WITH (NOLOCK) WHERE FS_Item = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                Set oRs = objO.FetchDataHO("SELECT * FROM INSENTIF..GRTipeMP where ISNULL(UserDel,'') = '' and KdTipe = '" & Text(Index) & "'")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "DATATIPEMP"
                    gBantuan = "SBE"
                    frmCariHO.show (vbModal)
                Else
                    Text(6) = NulSID(oRs!tipe)
                End If
                oRs.Close
            End If
        Case 1
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("select distinct KodeItem,NamaBarang from " & gAlamatDepo & "[depo00-" & Format(Now(), "yyyy") & Format(Now(), "mm") & "].dbo.inmst where KodeItem = '" & Text(Index) & "'")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "DATATIPEMP"
                    gBantuan = "SBE"
                    frmCariHO.show (vbModal)
                Else
                    Text(2) = NulSID(oRs!NamaBarang)
                End If
                oRs.Close
            End If
    End Select
End Sub

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchDataHO("SELECT TOP 1 * FROM INSENTIF..GRMstKPIPosisi WITH (NOLOCK) ORDER BY RIGHT(KdTipe,3) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KdTipe"), 3))
        noUrut = noUrut + 1
        Text(6) = "GR-" & Format(noUrut, "00#")
    Else
        noUrut = 1
        Text(6) = "GR-" & Format(noUrut, "00#")
    End If
End Sub

Private Function fungsiIsiId() As String
    Dim noUrut As Integer
    Dim hasil As String
    
    Set rsTemp = objO.FetchDataHO("SELECT TOP 1 * FROM INSENTIF..GRMstKPIPosisi WITH (NOLOCK) ORDER BY RIGHT(KdTipe,3) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KdTipe"), 3))
        noUrut = noUrut + 1
        hasil = "GR-" & Format(noUrut, "00#")
    Else
        noUrut = 1
        hasil = "GR-" & Format(noUrut, "00#")
    End If
    fungsiIsiId = hasil
End Function
