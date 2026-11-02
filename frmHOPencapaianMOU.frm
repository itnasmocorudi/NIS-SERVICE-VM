VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmHOPencapaianMOU 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MOU"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10545
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   7095
      Left            =   480
      TabIndex        =   4
      Top             =   480
      Width           =   9975
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   9735
         Begin TDBDate6Ctl.TDBDate tglAwal 
            Height          =   375
            Left            =   2280
            TabIndex        =   12
            Top             =   120
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            Calendar        =   "frmHOPencapaianMOU.frx":0000
            Caption         =   "frmHOPencapaianMOU.frx":012C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmHOPencapaianMOU.frx":0198
            Keys            =   "frmHOPencapaianMOU.frx":01B6
            Spin            =   "frmHOPencapaianMOU.frx":0214
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
            Text            =   "09/23/2026"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   46288
            CenturyMode     =   0
         End
         Begin TDBDate6Ctl.TDBDate tglAkhir 
            Height          =   375
            Left            =   4200
            TabIndex        =   13
            Top             =   120
            Width           =   1455
            _Version        =   65536
            _ExtentX        =   2566
            _ExtentY        =   661
            Calendar        =   "frmHOPencapaianMOU.frx":023C
            Caption         =   "frmHOPencapaianMOU.frx":0368
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmHOPencapaianMOU.frx":03D4
            Keys            =   "frmHOPencapaianMOU.frx":03F2
            Spin            =   "frmHOPencapaianMOU.frx":0450
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
            Text            =   "23/09/2026"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   46288
            CenturyMode     =   0
         End
         Begin AIFCmp1.asxPowerButton asxPowerButton1 
            Height          =   360
            Left            =   5760
            TabIndex        =   15
            Top             =   120
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   12632256
            BorderStyle     =   3
            Caption         =   "Refresh"
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   2
            Left            =   1830
            TabIndex        =   16
            Top             =   600
            Width           =   2865
            _ExtentX        =   5054
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
            Index           =   3
            Left            =   1830
            TabIndex        =   18
            Top             =   960
            Width           =   2865
            _ExtentX        =   5054
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
            Index           =   4
            Left            =   6510
            TabIndex        =   20
            Top             =   600
            Width           =   2865
            _ExtentX        =   5054
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
            Index           =   5
            Left            =   6510
            TabIndex        =   22
            Top             =   960
            Width           =   2865
            _ExtentX        =   5054
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
            Index           =   6
            Left            =   1830
            TabIndex        =   24
            Top             =   1320
            Width           =   2865
            _ExtentX        =   5054
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
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sisa Mou (Hari)"
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
            TabIndex        =   25
            Top             =   1365
            Width           =   1080
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPUS Avg"
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
            Left            =   5535
            TabIndex        =   23
            Top             =   1005
            Width           =   720
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CPUS Total"
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
            Left            =   5460
            TabIndex        =   21
            Top             =   645
            Width           =   795
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revenue Avg"
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
            TabIndex        =   19
            Top             =   1005
            Width           =   975
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Revenue Total"
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
            TabIndex        =   17
            Top             =   645
            Width           =   1050
         End
         Begin VB.Label Label3 
            Caption         =   " s/d "
            Height          =   375
            Left            =   3840
            TabIndex        =   14
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Range Periode"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   120
            Width           =   1935
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   4080
         TabIndex        =   5
         Top             =   600
         Width           =   2865
         _ExtentX        =   5054
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
      Begin TrueOleDBGrid70.TDBGrid GridUnitMOU 
         Height          =   4095
         Left            =   120
         TabIndex        =   6
         Top             =   2880
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7223
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFF80&"
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
         Left            =   4080
         TabIndex        =   7
         Top             =   240
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   57
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perusahaan"
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
         Left            =   2970
         TabIndex        =   9
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. MOU"
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
         Left            =   3180
         TabIndex        =   8
         Top             =   285
         Width           =   645
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7995
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   14102
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "  Detail Pencapaian MOU"
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
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   5880
      TabIndex        =   0
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12632256
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
      Left            =   6840
      TabIndex        =   1
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12632256
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
      Left            =   7800
      TabIndex        =   2
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12632256
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
      Left            =   8760
      TabIndex        =   3
      Top             =   7680
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12632256
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
End
Attribute VB_Name = "frmHOPencapaianMOU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 06 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : -
'=======================================

Dim rsGrid As New ADODB.Recordset
Dim rsTemp As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns

Private Sub ber_sih()
'    For i = 2 To 6
'        n = Text(i).Mask
'        Text(i).Mask = ""
'        Text(i) = ""
'        Text(i).Mask = n
'    Next
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
'    For i = 2 To 6
'        Text(i).Enabled = Flg
'    Next
'    cboCountry.Enabled = Flg
End Sub

Private Sub refreshGrid()
    Dim cpus, iDealer, avgCpus, totalBulan, revenueTot As Double
    totalBulan = 0
    revenueTot = 0
    sSQL = "select isnull(sum(TOTAL),0)  from T_INVOICEnew with (nolock) where MOU = '" & NulSID(Text(0)) & "' and INVDATE between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & "' and fs_rc <> 'R001' and TPSVC <> 'W'"
    revenueTot = objO.FetchDataHO_Double(sSQL)
    
    Text(2) = Format(revenueTot, "#,#0")
    totalBulan = DateDiff("m", Format(tglAwal, "yyyy-mm-dd"), Format(tglAkhir, "yyyy-mm-dd")) + 1
    
    sSQL = "select isnull(avg(TOTAL),0)  from T_INVOICEnew with (nolock) where MOU = '" & NulSID(Text(0)) & "' and INVDATE between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & "' and fs_rc <> 'R001' and TPSVC <> 'W'"
    Text(3) = Format(Round(revenueTot / totalBulan, 0), "#,#0") 'objO.FetchDataHO_Double(sSQL)
    sSQL = "select isnull(count(NO_WO),0) from T_INVOICEnew with (nolock) where MOU = '" & NulSID(Text(0)) & "' and INVDATE between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & "' and fs_rc <> 'R001' and TPSVC <> 'W'"
    cpus = objO.FetchDataHO_Double(sSQL)
    
    sSQL = "select distinct FS_KD_DEALER from T_INVOICEnew with (nolock) where MOU = '" & NulSID(Text(0)) & "' and INVDATE between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & "' and fs_rc <> 'R001' and TPSVC <> 'W'"
    iDealer = objO.FetchDataHO_Count(sSQL)
    avgCpus = 0
    If totalBulan > 0 Then
        avgCpus = Round(cpus / totalBulan, 2)
    End If
    Text(4) = cpus
    Text(5) = Format(avgCpus, "#,#0")
    
    GridUnitMOU.ClearFields
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = objO.FetchDataHO("select CUSTOMER,FMDL,NO_CHASSIS,NO_POLISI,TOTAL from T_INVOICEnew with (nolock) where MOU = '" & NulSID(Text(0)) & "' and INVDATE between '" & Format(tglAwal, "yyyy-mm-dd") & "' and '" & Format(tglAkhir, "yyyy-mm-dd") & "' and fs_rc <> 'R001' and TPSVC <> 'W'")
    'If Not rsGrid.EOF Then
        GridUnitMOU.DataSource = rsGrid
    '    Call HeaderColom
    'End If
End Sub

Private Sub HeaderColom()
    GridUnitMOU.Columns(0).Caption = "MODEL"
    GridUnitMOU.Columns(0).Width = 2000
    GridUnitMOU.Columns(0).Caption = "MODEL"
    GridUnitMOU.Columns(0).Width = 2000
    'MODEL, model as KETERANGAN, no_polisi as [NO. POLISI], no_Rangka as [NO. RANGKA]
End Sub
 

Private Sub asxPowerButton1_Click()
 Call refreshGrid
 
End Sub

Private Sub cmdHapus_Click()
    If rsTemp.State Then rsTemp.Close
    Set rsTemp = oConn.Execute("SELECT * FROM tempMOUdtl WHERE no_mou = '" & Text(0) & "' AND mdl = '" & Text(2) & "' " _
               & "AND no_polisi = '" & Text(4) & "' AND no_rangka = '" & cboCountry & Text(5) & Text(6) & "'")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus unit kendaraan dengan No. Rangka : " & cboCountry & Text(5) & Text(6) & " ?")
        If temp = vbYes Then
            oConn.Execute ("DELETE tempMOUdtl WHERE no_mou = '" & Text(0) & "' AND mdl = '" & Text(2) & "' " _
                          & "AND no_polisi = '" & Text(4) & "' AND no_rangka = '" & cboCountry & Text(5) & Text(6) & "'")
        End If
    Else
        MsgBox "Data Belum ada / belum tersimpan"
        Exit Sub
    End If
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
    If cmdSimpan.Caption = "&Tambah" Then
        Call ber_sih
        Call buka_Tutup
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Ubah"
        newRecord = True
    ElseIf cmdSimpan.Caption = "&Simpan" Then
        
        Call ClearsSql
        gsSql1(0) = "NO_MOU"
        gsSql2(0) = Text(0)
        gsSql1(1) = "MDL"
        gsSql2(1) = Text(2)
        gsSql1(2) = "MODEL"
        gsSql2(2) = Text(3)
        gsSql1(3) = "NO_POLISI"
        gsSql2(3) = Trim(Replace(Replace(Text(4), ">", ""), " ", ""))
        gsSql1(4) = "NO_RANGKA"
        gsSql2(4) = cboCountry & Text(5) & Replace(Text(6), ">", "")
    
        ' SIMPAN DATA
        If rsTemp.State Then rsTemp.Close
        If Text(4) = "" Then
            Set rsTemp = oConn.Execute("SELECT * FROM tempMOUdtl WHERE no_MOU = '" & Text(0) & "' AND mdl = '" & Text(2) & "' " _
                                     & "AND no_polisi = '" & Replace(Text(4), ">", "") & "' AND no_rangka = '" & cboCountry & Text(5) & Replace(Text(6), ">", "") & "'")
        Else
            Set rsTemp = oConn.Execute("SELECT * FROM tempMOUdtl WHERE no_MOU = '" & Text(0) & "' AND mdl = '" & Text(2) & "' " _
                                     & "AND no_rangka = '" & cboCountry & Text(5) & Replace(Text(6), ">", "") & "'")
        End If
        If Not rsTemp.EOF Then
            ' DATA SUDAH ADA
            If newRecord Then
                MsgBox "DATA SUDAH ADA. Jika ingin merubah pilih UBAH", vbCritical, "DUPLIKASI DATA"
            Else
                temp = MsgBox("Yakin akan merubah Data ?", vbYesNo, "DUPLIKASI DATA")
                If temp = vbYes Then
                    sSQL = GensSql2("tempMOUdtl", True) + " where NO_MOU = '" & Text(0) & "' AND mdl = '" & Text(2) & "' " _
                         & "AND no_polisi = '" & Replace(Text(4), ">", "") & "'"
                End If
            End If
        Else
            ' DATA BARU
            sSQL = GensSql2("tempMOUdtl", False)
        End If
        oConn.Execute sSQL
        
        Call ber_sih
        Call buka_Tutup(False, False)
        Call refreshGrid
        
        cmdSimpan.Caption = "&Tambah"
        cmdUbah.Caption = "&Ubah"
    End If
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call buka_Tutup
        Text(2).SetFocus
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False)
        cmdSimpan.Caption = "&Tambah"
        cmdUbah.Caption = "&Ubah"
    End If
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False) '
    Text(0) = frmHOListMOU.Text(0)
    Text(1) = frmHOListMOU.Text(1)
   ' Call refreshGrid
End Sub

Private Sub GridUnitMOU_FilterChange()
    On Error GoTo errHandler
    Set cols = GridUnitMOU.Columns
    Dim C As Integer
    
    C = GridUnitMOU.col
    GridUnitMOU.HoldFields
    rsGrid.filter = getFilter()
    GridUnitMOU.col = C
    GridUnitMOU.EditActive = True
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

Private Sub GridUnitMOU_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Then
        'GridUnitMOU_RowColChange
    End If
End Sub

Private Sub GridUnitMOU_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    Text(2) = GridUnitMOU.Columns(0)
'    Text(3) = GridUnitMOU.Columns(1)
'    Text(4).Mask = ""
'    Text(4) = GridUnitMOU.Columns(2)
'    cboCountry = Trim(Left(GridUnitMOU.Columns(3), 3))
'    Text(5) = Mid(GridUnitMOU.Columns(3), 4, 6)
'    Text(6) = Mid(GridUnitMOU.Columns(3), 10, Len(GridUnitMOU.Columns(3)) - 9)
End Sub
