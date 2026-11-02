VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOProgramFOPUbah 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MASTER FOP PROGRAM"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   20220
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12840
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   6495
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   14055
      Begin TrueOleDBGrid70.TDBGrid gridData 
         Height          =   2295
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   4048
         _LayoutType     =   4
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
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=20"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=20"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=20"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=20"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=3281"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=20"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=20"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=20"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=3281"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=20"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0,.valignment=2"
         _StyleDefs(74)  =   "Named:id=34:Heading"
         _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=34,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=35:Footing"
         _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   "Named:id=36:Selected"
         _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=37:Caption"
         _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(83)  =   "Named:id=38:HighlightRow"
         _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=39:EvenRow"
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
         _StyleDefs(93)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(94)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(95)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(96)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         _StyleDefs(97)  =   "Named:id=87:SudahDitagih"
         _StyleDefs(98)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(99)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(100) =   ":id=87,.fontname=Tahoma"
      End
      Begin VB.Frame Frame2 
         Caption         =   "Detail Data"
         Height          =   2295
         Left            =   5400
         TabIndex        =   26
         Top             =   240
         Width           =   7455
         Begin VB.ComboBox cbTagihan 
            Height          =   315
            ItemData        =   "frmHOProgramFOPUbah.frx":0000
            Left            =   1320
            List            =   "frmHOProgramFOPUbah.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   840
            Width           =   2415
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbKodeNik 
            Height          =   285
            Left            =   1320
            TabIndex        =   27
            Top             =   1185
            Width           =   1455
            AllowInput      =   0   'False
            _Version        =   196616
            BackColorOdd    =   8438015
            Columns(0).Width=   3200
            _ExtentX        =   2566
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin TDBDate6Ctl.TDBDate tglRetailAwal 
            Height          =   285
            Left            =   1320
            TabIndex        =   28
            Top             =   1485
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            Calendar        =   "frmHOProgramFOPUbah.frx":0018
            Caption         =   "frmHOProgramFOPUbah.frx":0144
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmHOProgramFOPUbah.frx":01B0
            Keys            =   "frmHOProgramFOPUbah.frx":01CE
            Spin            =   "frmHOProgramFOPUbah.frx":022C
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
            Text            =   "22/05/2020"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   43973
            CenturyMode     =   0
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   30
            Top             =   240
            Width           =   5865
            _ExtentX        =   10345
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
            Index           =   1
            Left            =   1320
            TabIndex        =   31
            Top             =   540
            Width           =   2385
            _ExtentX        =   4207
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
            Index           =   2
            Left            =   3840
            TabIndex        =   32
            Top             =   540
            Width           =   3465
            _ExtentX        =   6112
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
         Begin TDBDate6Ctl.TDBDate tglRetailAkhir 
            Height          =   285
            Left            =   2880
            TabIndex        =   33
            Top             =   1485
            Width           =   1095
            _Version        =   65536
            _ExtentX        =   1931
            _ExtentY        =   503
            Calendar        =   "frmHOProgramFOPUbah.frx":0254
            Caption         =   "frmHOProgramFOPUbah.frx":0380
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmHOProgramFOPUbah.frx":03EC
            Keys            =   "frmHOProgramFOPUbah.frx":040A
            Spin            =   "frmHOProgramFOPUbah.frx":0468
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
            Text            =   "22/05/2020"
            ValidateMode    =   0
            ValueVT         =   7
            Value           =   43973
            CenturyMode     =   0
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   5
            Left            =   4845
            TabIndex        =   34
            Top             =   1440
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   ">9999999999999"
            PromptChar      =   ">"
         End
         Begin AIFCmp1.asxPowerButton cmdUpdate 
            Height          =   360
            Left            =   1320
            TabIndex        =   42
            Top             =   1800
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Ubah Data"
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
         Begin VB.Shape Shape2 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   7200
            Top             =   240
            Width           =   60
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Order"
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
            Left            =   585
            TabIndex        =   41
            Top             =   285
            Width           =   720
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model Kendaran"
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
            TabIndex        =   40
            Top             =   585
            Width           =   1155
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   3720
            Top             =   540
            Width           =   60
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tagihan"
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
            Left            =   735
            TabIndex        =   39
            Top             =   900
            Width           =   570
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Tahun NIK"
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
            TabIndex        =   38
            Top             =   1230
            Width           =   1155
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periode Retail"
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
            TabIndex        =   37
            Top             =   1530
            Width           =   990
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "s / d"
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
            Left            =   2520
            TabIndex        =   36
            Top             =   1530
            Width           =   315
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Deposit"
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
            Left            =   4200
            TabIndex        =   35
            Top             =   1485
            Width           =   540
         End
      End
      Begin VB.Frame frameNikKendaraan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode NIK Kendaraan"
         Height          =   3615
         Left            =   14400
         TabIndex        =   23
         Top             =   1080
         Visible         =   0   'False
         Width           =   5415
         Begin TrueOleDBGrid70.TDBGrid GridNik 
            Height          =   3015
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   5318
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE"
            Columns(0).DataField=   "kode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "TAHUN"
            Columns(1).DataField=   "tahun"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   4
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "OK"
            Columns(2).DataField=   "IKUT"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   14737632
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=112,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Named:id=33:Normal"
            _StyleDefs(49)  =   ":id=33,.parent=0"
            _StyleDefs(50)  =   "Named:id=34:Heading"
            _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(52)  =   ":id=34,.wraptext=-1"
            _StyleDefs(53)  =   "Named:id=35:Footing"
            _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(55)  =   "Named:id=36:Selected"
            _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(57)  =   "Named:id=37:Caption"
            _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(59)  =   "Named:id=38:HighlightRow"
            _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=39:EvenRow"
            _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(63)  =   "Named:id=40:OddRow"
            _StyleDefs(64)  =   ":id=40,.parent=33"
            _StyleDefs(65)  =   "Named:id=41:RecordSelector"
            _StyleDefs(66)  =   ":id=41,.parent=34"
            _StyleDefs(67)  =   "Named:id=42:FilterBar"
            _StyleDefs(68)  =   ":id=42,.parent=33"
         End
         Begin AIFCmp1.asxPowerButton cmdTutup 
            Height          =   360
            Left            =   120
            TabIndex        =   25
            Top             =   3240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "&Tutup"
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
      Begin AIFCmp1.asxPowerButton cmdTampilkanKodeNik 
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Top             =   1125
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Pilih NIK"
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
      Begin VB.CheckBox chPart 
         BackColor       =   &H000080FF&
         Caption         =   "Free Part"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   11400
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frmeGanti 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   960
         TabIndex        =   10
         Top             =   3000
         Visible         =   0   'False
         Width           =   10335
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   11
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
            TabIndex        =   12
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
            TabIndex        =   14
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
            TabIndex        =   13
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
            TabIndex        =   15
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
            Picture         =   "frmHOProgramFOPUbah.frx":0490
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   450
            Width           =   720
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   1440
         TabIndex        =   0
         Top             =   1125
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
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
      Begin AIFCmp1.asxPowerButton cmdRefresh 
         Height          =   360
         Left            =   4200
         TabIndex        =   1
         Top             =   1345
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Tampilkan Paket Item"
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
      Begin TrueOleDBGrid70.TDBGrid gridPart 
         Height          =   3735
         Left            =   120
         TabIndex        =   2
         Top             =   2640
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   6588
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Terlibat"
         Columns(0).DataField=   "terlibat"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=20"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=20"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=20"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=20"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=91,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=88,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=89,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=90,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0,.valignment=2"
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
         _StyleDefs(81)  =   "Named:id=87:SudahDitagih"
         _StyleDefs(82)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(83)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(84)  =   ":id=87,.fontname=Tahoma"
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   11400
         Top             =   1680
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Part"
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
         Left            =   10305
         TabIndex        =   21
         Top             =   270
         Visible         =   0   'False
         Width           =   960
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
      FormatString    =   "  Tabel Master FOP Program"
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
      Left            =   8880
      TabIndex        =   4
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Simpan"
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
      Left            =   9840
      TabIndex        =   5
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Batal"
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
      Left            =   6720
      TabIndex        =   6
      Top             =   7440
      Visible         =   0   'False
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
      TabIndex        =   7
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
   Begin AIFCmp1.asxPowerButton cmdCetak 
      Height          =   360
      Left            =   6720
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Template"
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
      TabIndex        =   9
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
      Left            =   5040
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
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
      TabIndex        =   19
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
End
Attribute VB_Name = "frmHOProgramFOPUbah"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOProgramFOPUbah
' DateTime  : 26/05/2020 10:48
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

Dim rsDownload  As New ADODB.Recordset
Dim rsData      As ADODB.Recordset
Dim rsDataUbah  As ADODB.Recordset
Dim rsGridUbah  As ADODB.Recordset
Dim rsGrid      As New ADODB.Recordset
Dim rsGridSub   As New ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim rsModel     As ADODB.Recordset
Dim newRecord   As Boolean
Dim KDPART_LAMA As String
Dim boleh       As Boolean
Dim pesan       As String
Dim cols        As TrueOleDBGrid70.Columns
Dim rsNik       As New ADODB.Recordset
Dim daftarNIK   As String
Dim daftarFOP   As String
Dim daftarModel As String
Dim pekerjaan   As String


Public Property Let GetDaftarModel(value As String)
    daftarModel = value
End Property

Public Property Let getRSData(value As ADODB.Recordset)
    Set rsDataUbah = value
End Property

Public Property Let getDataPekerjaan(value As String)
    pekerjaan = NulSID(value)
End Property



Private Sub ber_sih()
    For i = 1 To 5
        Select Case i
            Case 5
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    tglRetailAwal = isnow
    tglRetailAkhir = isnow
    ''cbTagihan.Text = ""
    cbTagihan.Clear
    cbTagihan.AddItem "NRM"
    cbTagihan.AddItem "TAM"
    
    Call refreshComboNIK
    chPart.value = vbUnchecked
    
    Label15.Visible = False
    Text(5).Visible = False
     
    cmdHapus.Enabled = False
    
    gridPart.DataSource = Nothing
End Sub

Private Sub buka_tutupSubtitusi(Flg As Boolean)
     
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    Text(0).Enabled = Not Flg
    Text(1).Enabled = Not flg1
    Text(2).Enabled = Not flg1
    Text(3).Enabled = Not flg1
    Text(5).Enabled = Not flg1
    chPart.Enabled = Not flg1
    cbKodeNik.Enabled = Not flg1
    
    tglRetailAwal.Enabled = Not flg1
    tglRetailAkhir.Enabled = Not flg1
    cbTagihan.Enabled = Not flg1
     
End Sub

Private Sub refreshGridModel()
    Call createRSGridUbah
    If rsDataUbah.State = adStateClosed Then Exit Sub
    If rsDataUbah.RecordCount > 0 Then
        rsDataUbah.MoveFirst
        
        daftarNIK = ""
        daftarFOP = ""
        Do While Not rsDataUbah.EOF
            With rsGridUbah
                .AddNew
                !model = NulSID(rsDataUbah!model)
                !fullmodel = NulSID(Left(objO.FetchDataHO_String("Select [desc] from MODEL WITH (NOLOCK) where MDL = '" & !model & "' and FS_LOG_DEL = ''"), 100))
                !jobOrder = NulSID(rsDataUbah!jobOrder)
                !program = NulSID(rsDataUbah!program)
                daftarFOP = daftarFOP & "'" & !program & "',"
                !nik = objO.FetchDataHO_String("SELECT KODE FROM listTahunNikFOP WITH (NOLOCK) WHERE TAHUN = '" & NulSID(rsDataUbah!nik) & "'")
                daftarNIK = daftarNIK & "'" & !nik & "',"
                sSQL = "select RETAILAWAL from HO_FOP_Program with (nolock) where FS_LOG_DEL = '' and fs_item = '" & !jobOrder & "' and MDL = '" & !model & "' and FOP = '" & !program & "' and NIK = '" & !nik & "'"
                !retailawal = IIf(objO.FetchDataHO_String(sSQL) = "", "01/01/1900", Format(objO.FetchDataHO_String(sSQL), "dd/mm/yyyy"))
                sSQL = "select RETAILAKHIR from HO_FOP_Program with (nolock) where FS_LOG_DEL = '' and fs_item = '" & !jobOrder & "' and MDL = '" & !model & "' and FOP = '" & !program & "' and NIK = '" & !nik & "'"
                !retailakhir = IIf(objO.FetchDataHO_String(sSQL) = "", "01/01/1900", Format(objO.FetchDataHO_String(sSQL), "dd/mm/yyyy"))
                
                If !program = "NRM" Then
                    SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(!jobOrder) & "'")
                    If SBE = "R010" Then
                        SBE = "sb10k"
                    ElseIf SBE = "R020" Then
                        SBE = "sb20k"
                    ElseIf SBE = "R030" Then
                        SBE = "sb30k"
                    ElseIf SBE = "R040" Then
                        SBE = "sb40k"
                    ElseIf SBE = "R050" Then
                        SBE = "sb50k"
                    Else
                        SBE = "DEPOSIT"
                    End If
                    !NilaiDeposit = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(!model) & "'")
                Else
                    !NilaiDeposit = 0
                End If
               '' !NilaiDeposit = NulSID(rsDataUbah!NilaiDeposit)
                !programasli = !program
                !nikasli = !nik
                .Update
            End With
            rsDataUbah.MoveNext
        Loop
    End If
    Call HeaderColomGridData
    Call refreshGrid
    
End Sub


Private Sub HeaderColomGridData()
    With gridData
        .Columns(0).Caption = "Model"
        .Columns(0).Width = 2000
        .Columns(0).Alignment = dbgLeft
        .Columns(0).DataField = "model"
        
        .Columns(1).Caption = "Nama Pekerjaan"
        .Columns(1).DataField = "jobOrder"
        .Columns(1).Width = 3500
        .Columns(1).Alignment = dbgLeft
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Caption = "FOP"
        .Columns(2).Width = 500
        .Columns(2).DataField = "program"
        .Columns(2).Alignment = dbgLeft
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Caption = "NIK"
        .Columns(3).Width = 500
        .Columns(3).DataField = "Nik"
'        .Columns(3).Alignment = dbgRight
'        .Columns(3).NumberFormat = "#,#0"
        .Columns(3).Locked = True
        
        .Columns(4).Caption = "Retail Awal"
        .Columns(4).Width = 1500
        .Columns(4).DataField = "RETAILAWAL"
        .Columns(4).Locked = True
        
        .Columns(5).Caption = "Retail Akhir"
        .Columns(5).Width = 1500
        .Columns(5).DataField = "RETAILAkhir"
        .Columns(5).Locked = True
        
        .Columns(6).Caption = "Deposit"
        .Columns(6).Width = 1500
        .Columns(6).DataField = "nilaideposit"
        .Columns(6).Alignment = dbgRight
        .Columns(6).NumberFormat = "#,#0"
        .Columns(6).Locked = True
        
        .Columns(7).Caption = "Retail Akhir"
        .Columns(7).Width = 1000
        .Columns(7).DataField = "programasli"
        .Columns(7).Visible = False
        
        .Columns(8).Caption = "Retail Akhir"
        .Columns(8).Width = 1000
        .Columns(8).DataField = "nikasli"
        .Columns(8).Visible = False
    End With
End Sub

Private Sub createRSGridUbah()
    Set rsGridUbah = New ADODB.Recordset
    If rsGridUbah.State = adStateOpen Then
        rsGridUbah.Close
    End If
    With rsGridUbah
        .Fields.Append "Model", adVarChar, 50
        .Fields.Append "FullModel", adVarChar, 100
        .Fields.Append "JobOrder", adVarChar, 200
        .Fields.Append "Program", adVarChar, 50
        .Fields.Append "NIK", adVarChar, 50
        .Fields.Append "retailawal", adDate
        .Fields.Append "retailAkhir", adDate
        .Fields.Append "NilaiDeposit", adDouble
        .Fields.Append "ProgramAsli", adVarChar, 50
        .Fields.Append "NIKAsli", adVarChar, 50
        .Fields.Append "Fop", adVarChar, 50
        .Open
    End With
    
    gridData.DataSource = rsGridUbah
    ''Call HeaderColom
End Sub

Private Sub refreshGrid()
    gridPart.ClearFields
    If Right(daftarNIK, 1) = "," Then
        daftarNIK = Left(daftarNIK, Len(daftarNIK) - 1)
    End If
    If Right(daftarFOP, 1) = "," Then
        daftarFOP = Left(daftarFOP, Len(daftarFOP) - 1)
    End If
    If Right(daftarModel, 1) = "," Then
        daftarModel = Left(daftarModel, Len(daftarModel) - 1)
    End If
    
    ''sSQL = "select distinct a.* from (select ITEM_PART kode,NAMA_PART nama ,QTY,ISNULL((select '1' from HO_FOP_Program with (nolock) where FS_ITEM = a.FS_ITEM and MDL = a.MDL and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & cbTagihan.Text & "' and ITEM_PART = a.ITEM_PART ),0) as terlibat from HO_PARTPACK a with (nolock) where rekomendasi = 0 and FS_ITEM = '" & NulSID(Text(0)) & "' and MDL = '" & NulSID(Text(1)) & "'"
    ''sSQL = sSQL & " Union"
    ''sSQL = sSQL & " select KDSUBLET,NAMASUBLET,QTY,ISNULL((select '1' from HO_FOP_Program with (nolock) where FS_ITEM = b.FS_ITEM and MDL = b.MDL and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & cbTagihan.Text & "' and ITEM_PART = b.KDSUBLET),0) as terlibat from HO_SUBLETPACK b with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL = '" & NulSID(Text(1)) & "' ) a "
    sSQL = "select distinct a.* from (select ITEM_PART kode,NAMA_PART nama ,QTY,ISNULL((select '1' from HO_FOP_Program with (nolock) where FS_ITEM = a.FS_ITEM and MDL = a.MDL and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & cbTagihan.Text & "' and ITEM_PART = a.ITEM_PART ),0) as terlibat from HO_PARTPACK a with (nolock) where rekomendasi = 0 and FS_ITEM = '" & NulSID(Text(0)) & "' and MDL like '" & NulSID(Text(1)) & "%'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select KDSUBLET,NAMASUBLET,QTY,ISNULL((select '1' from HO_FOP_Program with (nolock) where FS_ITEM = b.FS_ITEM and MDL = b.MDL and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & cbTagihan.Text & "' and ITEM_PART = b.KDSUBLET),0) as terlibat from HO_SUBLETPACK b with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL like '" & NulSID(Text(1)) & "%' ) a "
    
    sSQL = "select distinct a.* from (select ITEM_PART kode,NAMA_PART nama ,QTY,"
    ''sSQL = sSQL & " ISNULL((select top 1 '1' from HO_FOP_Program with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL like '" & NulSID(Text(1)) & "%' and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & NulSID(cbTagihan.Text) & "' and ITEM_PART = a.ITEM_PART  and fs_log_del = ''),0) as terlibat"
    sSQL = sSQL & " ISNULL((select top 1 '1' from HO_FOP_Program with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL in (" & daftarModel & ") and NIK IN (" & daftarNIK & ") and FOP in (" & daftarFOP & ") and ITEM_PART = a.ITEM_PART  and fs_log_del = ''),0) as terlibat"
    sSQL = sSQL & " from HO_PARTPACK a with (nolock)"
    sSQL = sSQL & " where rekomendasi = 0 and FS_ITEM = '" & NulSID(Text(0)) & "'  and FS_LOG_DEL = ''"
    sSQL = sSQL & " and MDL in (" & daftarModel & ")"
    sSQL = sSQL & " Union"
    ''sSQL = sSQL & " select sub.KDSUBLET,sub.NAMASUBLET,sub.QTY,ISNULL((select top 1 '1' from HO_FOP_Program with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL like '" & NulSID(Text(1)) & "%'  and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & NulSID(cbTagihan.Text) & "' and ITEM_PART = sub.KDSUBLET and fs_log_del = ''),0) as terlibat  from ("
    sSQL = sSQL & " select sub.KDSUBLET,sub.NAMASUBLET,sub.QTY,ISNULL((select top 1 '1' from HO_FOP_Program with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL like '" & NulSID(Text(1)) & "%'  and NIK in (" & daftarNIK & ") and FOP in (" & daftarFOP & ") and ITEM_PART = sub.KDSUBLET and fs_log_del = ''),0) as terlibat  from ("
    sSQL = sSQL & " select b.KDSUBLET,c.NAMASUBLET,max(b.QTY) QTY"
    sSQL = sSQL & " from HO_SUBLETPACK b with (nolock) inner join MSTSUBLET c on b.KDSUBLET = c.KDSUBLET  and c.FS_LOG_DEL = '' where b.FS_ITEM = '" & NulSID(Text(0)) & "' and b.MDL in (" & daftarModel & ") and b.FS_LOG_DEL = ''"
    sSQL = sSQL & " group by b.KDSUBLET,c.NAMASUBLET ) sub"
    sSQL = sSQL & " ) a"
    
    Set rsGrid = objO.FetchDataHO(sSQL)
    Call createRSData
    ''Call HeaderColom
    If Not rsGrid.EOF Then
        rsGrid.MoveFirst
        Do While Not rsGrid.EOF
            
            With rsData
                .AddNew
                !kode = NulSID(rsGrid!kode)
                !Nama = NulSID(rsGrid!Nama)
                !qty = NulSID(rsGrid!qty)
                !qtyasli = NulSID(rsGrid!qty)
                !terlibat = IIf(NulSID(rsGrid!terlibat) = 1, True, False)
                .Update
            End With
            'Debug.Print gridPart.Splits(0).Columns(0).value
            'Call createRSData
            rsGrid.MoveNext
        Loop
    End If
   ' gridPart.Splits(0).Columns(0).ValueItems.Presentation = dbgCheckBox
   '' Set gridPart.DataSource = rsData
    Call HeaderColom
    ''gridPart.ReBind
End Sub

Private Sub createRSData()
    Set rsData = New ADODB.Recordset
    If rsData.State = adStateOpen Then rsData.Close
    With rsData
        .Fields.Append "terlibat", adBoolean
        .Fields.Append "kode", adVarChar, 50
        .Fields.Append "nama", adVarChar, 200
        .Fields.Append "qty", adInteger
        .Fields.Append "qtyAsli", adInteger
        .Open
    End With
    
    gridPart.DataSource = rsData
    ''Call HeaderColom
End Sub


Private Sub HeaderColom()
    With gridPart
        .Columns(0).Caption = "Terlibat"
        .Columns(0).Width = 1000
        .Columns(0).Alignment = dbgCenter
        .Columns(0).DataField = "terlibat"
        .Splits(0).Columns(0).Locked = False
        .Splits(0).Columns(0).ValueItems.Presentation = dbgCheckBox
        
        .Columns(1).Caption = "Kode Item"
        .Columns(1).DataField = "kode"
        .Columns(1).Width = 2000
        .Columns(1).Alignment = dbgLeft
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Caption = "Nama Item"
        .Columns(2).Width = 4000
        .Columns(2).DataField = "nama"
        .Columns(2).Alignment = dbgLeft
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Caption = "Qty"
        .Columns(3).Width = 500
        .Columns(3).DataField = "QTY"
        .Columns(3).Alignment = dbgRight
        .Columns(3).NumberFormat = "#,#0"
        .Columns(3).Locked = False
        
        .Columns(4).Caption = "Qty Asli"
        .Columns(4).Width = 500
        .Columns(4).DataField = "QTYasli"
        .Columns(4).Visible = False
    End With
End Sub


Private Sub HeaderColomSubstitusi()
    
    gridSub.Columns(0).Caption = "KODE PART"
    gridSub.Columns(0).Width = 2000
    gridSub.Columns(1).Caption = "NAMA PART"
    gridSub.Columns(1).Width = 3000
    gridSub.Columns(2).Caption = "QTY"
    gridSub.Columns(2).Width = 600
    gridSub.Columns(2).NumberFormat = "#,#0"
End Sub

Private Sub cbTagihan_Change()
    If cbTagihan.Text = "NRM" Then
        Label15.Visible = True
        Text(5).Visible = True
        SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "'")
        If SBE = "R010" Then
            SBE = "sb10k"
        ElseIf SBE = "R020" Then
            SBE = "sb20k"
        ElseIf SBE = "R030" Then
            SBE = "sb30k"
        ElseIf SBE = "R040" Then
            SBE = "sb40k"
        ElseIf SBE = "R050" Then
            SBE = "sb50k"
        Else
            SBE = "DEPOSIT"
        End If
        Text(15) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(rsGrid.Fields(4)) & "'")
    Else
        Label15.Visible = False
        Text(5).Visible = False
    End If
End Sub

Private Sub cbTagihan_Click()
    If cbTagihan.Text = "NRM" Then
        Label15.Visible = True
        Text(5).Visible = True
        SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(Text(0)) & "'")
        If SBE = "R010" Then
            SBE = "sb10k"
        ElseIf SBE = "R020" Then
            SBE = "sb20k"
        ElseIf SBE = "R030" Then
            SBE = "sb30k"
        ElseIf SBE = "R040" Then
            SBE = "sb40k"
        ElseIf SBE = "R050" Then
            SBE = "sb50k"
        Else
            SBE = "DEPOSIT"
        End If
        'Text(5) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(Text(1)) & "'")
        Text(5) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory like '" & NulSID(Text(1)) & "%'")

    Else
        Label15.Visible = False
        Text(5).Visible = False
    End If
End Sub

Private Sub cmdCetak_Click()

   On Error GoTo cmdCetak_Click_Error
    Dim Er As New ErrInfo
    If rsDownload.State Then rsDownload.Close
    Set rsDownload = Nothing
    With rsDownload
        .Fields.Append "NAMAPEKERJAAN", adVarChar, 50
        .Fields.Append "KODEMODEL", adVarChar, 50
        .Fields.Append "TAGIHAN", adVarChar, 50
        .Fields.Append "KODENIK", adVarChar, 50
        .Fields.Append "TGLRETAILAWAL", adVarChar, 50
        .Fields.Append "TGLRETAILAKHIR", adVarChar, 50
        .Fields.Append "KODEITEM", adVarChar, 50
        .Fields.Append "NOMINALDEPOSITO", adVarChar, 50
        .Open
 
    End With
      
    Dim MyPath As String
    Dim onVal As String
    MyPath = ""
    Dim f As New FormExplorer
    f.show (vbModal)
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    onVal = MyPath & f.FileName & ".xls"
    Call ProsesExport(onVal)
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            
            Beep
            Er.show (vbModal)
    
    
    

   On Error GoTo 0
   Exit Sub

cmdCetak_Click_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdCetak_Click of Form frmHOProgramFOPUbah"


End Sub

Private Sub cmdGanti_Click()
    Text(8) = ""
    Text(4) = ""
    Text(10) = ""
    Text(9) = ""
    frmeGanti.Visible = True
End Sub

Private Sub cmdHapus_Click()
    Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_fop_program WITH (NOLOCK) WHERE FOP = '" & cbTagihan.Text & "' and NIK = '" & NulSID(cbKodeNik.Text) & "' and fs_item = '" & Text(0) & "' AND mdl like '" & Text(1) & "%' and FS_LOG_DEL = ''")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus data " & rsTemp.RecordCount & " model kendaraan ?" & vbCr & "Untuk Job Order : " & Text(0), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE ho_fop_program SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE FOP = '" & cbTagihan.Text & "' and NIK = '" & NulSID(cbKodeNik.Text) & "' and fs_item = '" & Text(0) & "' AND mdl like '" & Text(1) & "%' and FS_LOG_DEL = ''")
            ''Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('PARTPACK', '" & Text(0) & "', 'D')")
        End If
        Call ber_sih
        ''Call refreshGrid
    Else
        MsgBox "Pilih data yang ingin dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdHapusSub_Click()
    Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and mdl = '" & NulSID(rsGrid.Fields(4)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "'  and ITEM_PARTSUB = '" & NulSID(rsGridSub.Fields(0)) & "' and NAMA_PARTSUB = '" & NulSID(rsGridSub.Fields(1)) & "' and QTYSUB = '" & NulSID(rsGridSub.Fields(2)) & "' and FS_LOG_DEL = ''")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus data sub part : " & NulSID(rsGridSub.Fields(0)) & " ?" & vbCr & "Untuk Job Order : " & Text(0), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE HO_FOP_PROGRAM_SUB SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and mdl = '" & NulSID(rsGrid.Fields(4)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "'  and ITEM_PARTSUB = '" & NulSID(rsGridSub.Fields(0)) & "' and NAMA_PARTSUB = '" & NulSID(rsGridSub.Fields(1)) & "' and QTYSUB = '" & NulSID(rsGridSub.Fields(2)) & "' and FS_LOG_DEL = ''")
        End If
        Text(11) = ""
        Text(12) = ""
        Text(13) = 0
        gridSub.ClearFields
        Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(8)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and MDL = '" & NulSID(rsGrid.Fields(4)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
        gridSub.DataSource = rsGridSub
        Call HeaderColomSubstitusi
        gridSub.ReBind
    Else
        MsgBox "Pilih data yang ingin dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdImport_Click()
      Dim ErrBegin As Boolean, konter As Long

1       On Error GoTo vb_error
          Dim CnExel As New ADODB.Connection
          Dim RcExcel As New ADODB.Recordset
          Dim rsTempExcel As ADODB.Recordset
          Dim Jml As Integer
          Dim itemPart As String
          Dim tglAwal  As Date
          Dim tglAkhir As Date
          Dim kodePart As String
          Dim namaPart As String
          Dim qty      As Double
          Dim isBegin  As Boolean
          Dim SBE As String
          Dim namaPekerjaan As String
          Dim itemPekerjaan As String
2         isBegin = False
         
        ' Dim konter As Integer
         
3       konter = 0
4       ProgBar.Min = 0
5       CommonDialog1.FileName = ""
6       CommonDialog1.InitDir = App.Path
7       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
8       CommonDialog1.ShowOpen
9       If CommonDialog1.FileName <> "" Then
10         If CnExel.State Then CnExel.Close
11         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
12         GoSub Proses
13     Else
14        Exit Sub
15     End If
16     Exit Sub
       
Proses:
17    ErrBegin = True
         Dim VarNmFile As String
         Dim model As String
         
18        gBantuan = "SBE"
19       If RcExcel.State Then RcExcel.Close
20       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
21       If RcExcel.RecordCount > 0 Then
22            Call objHT.OpenTransactionHO
23            isBegin = True

24          ProgBar.Visible = True
25          ProgBar.Max = RcExcel.RecordCount
26          RcExcel.MoveFirst
27          Do While Not RcExcel.EOF
28             DoEvents
29             konter = konter + 1
30             ProgBar.value = konter
31             model = ""
              
32             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
33             If objO.FetchDataHO_Count("select * from HO_FOP_JOB WITH(NOLOCK) WHERE FS_LOG_DEL = '' and FS_ITEM = '" & Trim(RcExcel.Fields(0).value) & "'") = 0 Then
34                Call objHT.CloseTransactionHO_err
35                ProgBar.Visible = False
36                MsgBox "JOB ORDER = " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " salah!", vbCritical, "Error"
37                Exit Sub
38             End If
               
39             If Nul(RcExcel.Fields(1).value) = "" Then
40                Call objHT.CloseTransactionHO_err
41                ProgBar.Visible = False
42                MsgBox "KODE MODEL baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
43                Exit Sub
44             Else
45                Set rsTempExcel = objO.FetchDataHO("SELECT * FROM MODEL WHERE FS_LOG_DEL = '' AND MDL = '" & Trim(RcExcel.Fields(1).value) & "'")
46                If rsTempExcel.RecordCount > 0 Then
47                    model = Trim(rsTempExcel("DESC"))
48                Else
                      'jika tidak ada data
49                    Call objHT.CloseTransactionHO_err
50                    ProgBar.Visible = False
51                    MsgBox "Model baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
52                    Exit Sub
53                End If
54             End If
               
55             If Nul(RcExcel.Fields(2).value) = "" Then
56                Call objHT.CloseTransactionHO_err
57                ProgBar.Visible = False
58                MsgBox "TAGIHAN baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
59                Exit Sub
60             ElseIf Not (NulSID(RcExcel.Fields(2).value) = "TAM" Or NulSID(RcExcel.Fields(2).value) = "NRM") Then
61                Call objHT.CloseTransactionHO_err
62                ProgBar.Visible = False
63                MsgBox "TAGIHAN baris ke-" & konter & " harus NRM / TAM !", vbCritical, "Error"
64                Exit Sub
65             End If
               
               
66             If Nul(RcExcel.Fields(3).value) = "" Then
67                Call objHT.CloseTransactionHO_err
68                ProgBar.Visible = False
69                MsgBox "KODE TAHUN NIK Kendaraan baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
70                Exit Sub
71             Else
72                Set rsTempExcel = objO.FetchDataHO("SELECT  kode,tahun from listTahunNikFOP WHERE kode = '" & Trim(RcExcel.Fields(3).value) & "'")
73                If rsTempExcel.RecordCount = 0 Then
                      'jika tidak ada data
74                    Call objHT.CloseTransactionHO_err
75                    ProgBar.Visible = False
76                    MsgBox "KODE TAHUN NIK baris ke-" & konter & " tidak terdaftar !", vbCritical, "Error"
77                    Exit Sub
78                End If
79             End If
               
               
80            If NulSID(RcExcel.Fields(2).value) = "NRM" Then
81                If Nul(RcExcel.Fields(7).value) = "" Then
82                   Call objHT.CloseTransactionHO_err
83                   ProgBar.Visible = False
84                   MsgBox "Deposit baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
85                   Exit Sub
86                ElseIf IsNumeric(Nul(RcExcel.Fields(7).value)) = False Then
87                   Call objHT.CloseTransactionHO_err
88                   ProgBar.Visible = False
89                   MsgBox "Deposit baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
90                   Exit Sub
91                End If
92            End If
                  
93             tglAwal = "01/01/1900"
94             tglAkhir = "01/01/1900"
95             If NulSID(RcExcel.Fields(4).value) <> "" Then
96                If IsDate(NulSID(RcExcel.Fields(4).value)) Then
97                    tglAwal = Format(NulSID(RcExcel.Fields(4).value), "dd/mm/yyyy")
98                End If
99             End If
               
100            If NulSID(RcExcel.Fields(5).value) <> "" Then
101               If IsDate(NulSID(RcExcel.Fields(5).value)) Then
102                   tglAwal = Format(NulSID(RcExcel.Fields(5).value), "dd/mm/yyyy")
103               End If
104            End If
               
105            kodePart = ""
106            namaPart = ""
107            qty = 0
108            If NulSID(RcExcel.Fields(6).value) = "" Then
109               Call objHT.CloseTransactionHO_err
110               ProgBar.Visible = False
111               MsgBox "ITEM PART / SUBLET baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
112               Exit Sub
113            Else
                  ''cek part terdaftar / tidak
114               sSQL = "select ITEM_PART kode ,NAMA_PART nama,QTY from HO_PARTPACK where FS_ITEM = '" & NulSID(RcExcel.Fields(0).value) & "' and MDL = '" & NulSID(RcExcel.Fields(1).value) & "' and ITEM_PART = '" & NulSID(RcExcel.Fields(6).value) & "' and FS_LOG_DEL = ''"
115               Set rsTemp = objHT.FetchDataHO(sSQL)
116               If rsTemp.RecordCount = 0 Then
                      ''cek sublet terdaftar ./ tidak
117                   sSQL = "select KDSUBLET kode ,NAMASUBLET nama,QTY from HO_SUBLETPACK where FS_ITEM = '" & NulSID(RcExcel.Fields(0).value) & "' and MDL = '" & NulSID(RcExcel.Fields(1).value) & "' and KDSUBLET = '" & NulSID(RcExcel.Fields(6).value) & "' and FS_LOG_DEL = ''"
118                   Set rsTemp = objHT.FetchDataHO(sSQL)
119                   If rsTemp.RecordCount = 0 Then
120                       Call objHT.CloseTransactionHO_err
121                       ProgBar.Visible = False
122                       MsgBox "ITEM PART / SUBLET baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
123                       Exit Sub
124                   Else
125                       GoTo isiDataPartSublet
126                   End If
127               Else
isiDataPartSublet:
128                   kodePart = NulSID(rsTemp!kode)
129                   namaPart = NulSID(rsTemp!Nama)
130                   qty = NulSID(rsTemp!qty)
131               End If
132            End If
              ''Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_FOP_Program WITH (NOLOCK) WHERE FOP = '" & Trim(RcExcel.Fields(2).value) & "' and NIK = '" & Trim(RcExcel.Fields(3).value) & "' and fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(RcExcel.Fields(1).value) & "' AND item_part = '" & Trim(kodePart) & "' and FS_LOG_DEL = ''")
133           Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_FOP_Program WITH (NOLOCK) WHERE NIK = '" & Trim(RcExcel.Fields(3).value) & "' and fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(RcExcel.Fields(1).value) & "' AND item_part = '" & Trim(kodePart) & "' and FS_LOG_DEL = ''")
134           If Not rsTemp.EOF Then
135               sSQL = "UPDATE HO_FOP_Program SET MODEL = '" & NulSID(model) & "',RETAILAWAL = '" & Format(tglAwal, "mm/dd/yyyy") & "',RETAILAKHIR = '" & Format(tglAkhir, "mm/dd/yyyy") & "',NAMA_PART = '" & namaPart & "',QTY = " & k2t(qty) & " WHERE FOP = '" & Trim(RcExcel.Fields(2).value) & "' and NIK = '" & Trim(RcExcel.Fields(3).value) & "' and fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(RcExcel.Fields(1).value) & "' AND item_part = '" & Trim(kodePart) & "' and FS_LOG_DEL = ''"
136           Else
137               sSQL = "INSERT INTO HO_FOP_Program (JOB,FS_ITEM,MDL,MODEL,FOP,NIK,RETAILAWAL,RETAILAKHIR,ITEM_PART,NAMA_PART,QTY,KETERANGAN,USERENTRY,TGLENTRY,FS_LOG_DEL,FD_TGL_DEL) VALUES "
138               sSQL = sSQL & " ('" & gBantuan & "','" & Trim(RcExcel.Fields(0).value) & "','" & Trim(RcExcel.Fields(1).value) & "','" & model & "','" & Trim(RcExcel.Fields(2).value) & "','" & Trim(RcExcel.Fields(3).value) & "','" & Format(tglAwal, "mm/dd/yyyy") & "','" & Format(tglAkhir, "mm/dd/yyyy") & "','" & Trim(kodePart) & "','" & Trim(namaPart) & "'," & k2t(qty) & ",'Import Excel','" & gSA & "',getdate(),'','')"
139           End If
140           Call objHT.InsertData_Ho(sSQL)
              
141           If NulSID(RcExcel.Fields(2).value) = "NRM" Then
                  ''Dim sbe As String
                  
142               Set rsTemp = objHT.FetchDataHO("SELECT * FROM mvp..m_deposit WITH (NOLOCK) WHERE KodeKategory = '" & Left(NulSID(RcExcel.Fields(1).value), 20) & "' ")
143               SBE = objHT.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(RcExcel.Fields(0).value) & "'")
144               If SBE = "R010" Then
145                   SBE = "sb10k"
146               ElseIf SBE = "R020" Then
147                   SBE = "sb20k"
148               ElseIf SBE = "R030" Then
149                   SBE = "sb30k"
150               ElseIf SBE = "R040" Then
151                   SBE = "sb40k"
152               ElseIf SBE = "R050" Then
153                   SBE = "sb50k"
154               Else
155                   SBE = "DEPOSIT"
156               End If
157               If Not rsTemp.EOF Then
158                   If NulSID(SBE) = NulSID("DEPOSIT") Then
159                       sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(RcExcel.Fields(7).value) & " ,DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(RcExcel.Fields(7).value) & " WHERE KodeKategory = '" & Left(NulSID(RcExcel.Fields(1).value), 20) & "' "
160                   Else
161                       itemPekerjaan = ""
162                       Set rsTemp = objHT.FetchDataHO("SELECT * FROM mvp..m_deposit WITH (NOLOCK) WHERE KodeKategory = '" & Left(NulSID(model), 20) & "' and  ket_" & SBE & " like '%" & itemPekerjaan & "%'")
163                       If rsTemp.EOF Then
164                           itemPekerjaan = namaPart
165                           sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(RcExcel.Fields(7).value) & " ,ket_" & SBE & " = CONVERT(VARCHAR(MAX), ket_" & SBE & ") + ' " & itemPekerjaan & "',DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(RcExcel.Fields(7).value) & " WHERE KodeKategory = '" & Left(NulSID(RcExcel.Fields(1).value), 20) & "' "
166                       Else
167                           sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(RcExcel.Fields(7).value) & ",DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(RcExcel.Fields(7).value) & " WHERE KodeKategory = '" & Left(NulSID(RcExcel.Fields(1).value), 20) & "' "
168                       End If
                          
169                   End If
170               Else
171                   If NulSID(SBE) = NulSID("DEPOSIT") Then
172                       sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & ", ket_" & SBE & ",deposit,tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(RcExcel.Fields(1).value), 20) & "'," & k2t(RcExcel.Fields(7).value) & ",'" & itemPekerjaan & "'," & k2t(RcExcel.Fields(7).value) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
173                   Else
174                       sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & " ,deposit,tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(RcExcel.Fields(1).value), 20) & "'," & k2t(RcExcel.Fields(7).value) & " ," & k2t(RcExcel.Fields(7).value) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
175                   End If
176               End If
177               Call objHT.InsertData_Ho(sSQL)
178           End If
              
179            RcExcel.MoveNext
180         Loop
181      End If
182      RcExcel.Close
183      CnExel.Close
      '   ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
184       Call objHT.CloseTransactionHO
185      MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
186      ProgBar.Visible = False
187      Call ber_sih
188      Call buka_Tutup(False, False)
         'Call refreshGrid
         'cmdUbah.Caption = "&Ubah"
         'cmdSimpan.Caption = "&Tambah"
      '   GridAFI.Refresh
189      Return
190      Exit Sub
vb_error:
191      If isBegin = True Then
           ' DbCon.RollbackTrans
192       Call objHT.CloseTransactionHO_err
193      End If
194      ProgBar.Visible = False
195       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdImport_Click of Form frmHOProgramFOPUbah"

196      On Error GoTo 0
197      Exit Sub
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdProses_Click()
    Call objO.InsertDataHO("INSERT INTO HO_PARTPACK  (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL) SELECT DISTINCT JOB, FS_ITEM, MDL, MODEL, '" & Trim(Text(10)) & "', '" & Trim(Text(9)) & "', QTY, KETERANGAN, FS_LOG_DEL, FD_TGL_DEL FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & Trim(Text(8)) & "'")
    Call objO.UpdateDataHO("UPDATE HO_PARTPACK SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE ITEM_PART = '" & Trim(Text(8)) & "'")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(Text(10)) & "', 'S')")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(Text(8)) & "', 'D')")
    frmeGanti.Visible = False
End Sub

Private Sub cmdRefresh_Click()
    Dim boleh As Boolean
    Dim pesan As String
    pesan = ""
    boleh = True
    daftarNIK = ""
    
    If cbTagihan.Text = "" Then
         pesan = pesan + "Jenis Tagihan. "
         cbTagihan.SetFocus
         boleh = False
    End If
     
'    If NulSID(cbKodeNik.Text) = "" Then ''If Text(3) = "" Then
'        boleh = False
'        pesan = pesan + "Kode NIK Tahun. "
'    End If

    If rsNik.State = adStateClosed Then Call createTbl
    
    If Not rsNik.EOF Then
        rsNik.filter = "ikut=true"
        If rsNik.RecordCount = 0 Then
            boleh = False
            pesan = pesan + "Kode NIK Tahun. "
        Else
            rsNik.MoveFirst
            Do While Not rsNik.EOF
                daftarNIK = daftarNIK & "'" & rsNik!kode & "',"
                
                rsNik.MoveNext
            Loop
        End If
        rsNik.filter = ""
    Else
        boleh = False
        pesan = pesan + "Kode NIK Tahun. "
    End If
    
    If boleh Then
        Call refreshGrid
        
        cmdSimpan.Enabled = True
        cmdUbah.Enabled = True
        cmdHapus.Enabled = True
    Else
        MsgBox pesan, vbCritical, "VALID DATA"
    End If
End Sub

Private Sub cmdSimpan_Click()
    Dim isBegin As Boolean
    Dim SBE As String
    Dim namaPekerjaan As String
    Dim itemPekerjaan As String
    
   On Error GoTo cmdSimpan_Click_Error

    isBegin = False
     
        boleh = True
        pesan = "Lengkapi data berikut "
        rsGridUbah.MoveFirst
        Do While Not rsGridUbah.EOF
            If rsGridUbah!program = "" Then
                pesan = pesan + "Jenis Tagihan. "
                boleh = False
            End If
            If rsGridUbah!nik = "" Then
                pesan = pesan + "Kode NIK Tahun. "
                boleh = False
            End If
            
            rsGridUbah.MoveNext
        Loop
        rsGridUbah.MoveFirst
          
        If rsData.State = adStateClosed Then
            boleh = False
            pesan = pesan + "Item part / sublet harus diisi. "
        Else
            rsData.MoveFirst
            rsData.filter = "terlibat=true"
            If rsData.RecordCount = 0 Then
                boleh = False
                pesan = pesan + "Item part / sublet harus diisi. "
            End If
            rsData.filter = ""
        End If
        
        If boleh Then
             
        rsGridUbah.MoveFirst
        Call objHT.OpenTransactionHO
        isBegin = True
        pesan = ""
        Do While Not rsGridUbah.EOF
            itemPekerjaan = ""
            ''hapus data ho_fop_program atas model,nik,tagihan dan pekerjaan kemudian insert data baru
            sSQL = "DELETE FROM ho_fop_program  WHERE NIK = '" & NulSID(rsGridUbah!nikasli) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & Left(NulSID(rsGridUbah!model), 20) & "'   and FS_LOG_DEL = ''"
            Call objHT.InsertData_Ho(sSQL)
            Do While Not rsData.EOF
                If rsData!terlibat Then
                    namaPekerjaan = ""
                    namaPekerjaan = objHT.FetchDataHO_String("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE mdl = '" & Left(NulSID(rsGridUbah!model), 20) & "' and fs_item = '" & Text(0) & "' AND item_part = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = '' ")
                    If NulSID(namaPekerjaan, "") = "" Then
                        namaPekerjaan = objHT.FetchDataHO_String("SELECT NAMASUBLET FROM HO_SUBLETPACK WITH (NOLOCK) WHERE mdl = '" & Left(NulSID(rsGridUbah!model), 20) & "' and fs_item = '" & Text(0) & "' AND KDSUBLET = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = '' ")
                    End If
                    If namaPekerjaan <> "" Then
                        Set rsTemp = objHT.FetchDataHO("SELECT * FROM ho_fop_program WITH (NOLOCK) WHERE NIK = '" & NulSID(rsGridUbah!nik) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & Left(NulSID(rsGridUbah!model), 20) & "' AND item_part = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = ''")
                        If Not rsTemp.EOF Then
                            Call objHT.InsertData_Ho("UPDATE ho_fop_program SET item_part = '" & NulSID(rsData!kode) & "', nama_part = '" & NulSID(namaPekerjaan) & "' , QTY = '" & CDbl(rsData!qty) & "',FOP = '" & NulSID(rsGridUbah!program) & "',NIK = '" & NulSID(rsGridUbah!nik) & "',RETAILAWAL = '" & Format(rsGridUbah!retailawal, "yyyy-mm-dd") & "',RETAILAKHIR = '" & Format(rsGridUbah!retailakhir, "yyyy-mm-dd") & "' WHERE FOP = '" & NulSID(rsGridUbah!program) & "' and NIK = '" & NulSID(rsGridUbah!nik) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & Left(NulSID(rsGridUbah!model), 20) & "' AND item_part = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = ''")
                        Else
                            sSQL = "INSERT INTO ho_fop_program (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL,FOP,NIK,RETAILAWAL,RETAILAKHIR,USERENTRY,TGLENTRY) "
                            sSQL = sSQL & " VALUES ('" & gBantuan & "', '" & Text(0) & "','" & Left(NulSID(rsGridUbah!model), 20) & "','" & Left(NulSID(rsGridUbah!fullmodel), 50) & "', '" & NulSID(rsData!kode) & "','" & NulSID(namaPekerjaan) & "','" & CDbl(rsData!qty) & "','','','','" & NulSID(rsGridUbah!program) & "','" & NulSID(rsGridUbah!nik) & "','" & Format(rsGridUbah!retailawal, "yyyy-mm-dd") & "','" & Format(rsGridUbah!retailakhir, "yyyy-mm-dd") & "','" & gSA & "',getdate())"
                            Call objHT.InsertData_Ho(sSQL)
                        End If
                        itemPekerjaan = itemPekerjaan & " " & NulSID(namaPekerjaan)
                    Else
                        pesan = pesan & "Item " & NulSID(rsData!kode) & " tidak terdaftar untuk model " & Left(NulSID(rsGridUbah!model), 20) & vbCrLf
                    End If
                End If
                rsData.MoveNext
            Loop
            ''cek mvp m_deposit
            ''--------------------------------------------------------------
            If NulSID(rsGridUbah!program) = "NRM" Then
                ''Dim sbe As String
                
                Set rsTemp = objHT.FetchDataHO("SELECT * FROM mvp..m_deposit WITH (NOLOCK) WHERE KodeKategory = '" & Left(NulSID(rsGridUbah!model), 20) & "' ")
                SBE = objHT.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(Text(0)) & "'")
                If SBE = "R010" Then
                    SBE = "sb10k"
                ElseIf SBE = "R020" Then
                    SBE = "sb20k"
                ElseIf SBE = "R030" Then
                    SBE = "sb30k"
                ElseIf SBE = "R040" Then
                    SBE = "sb40k"
                ElseIf SBE = "R050" Then
                    SBE = "sb50k"
                Else
                    SBE = "DEPOSIT"
                End If
                If Not rsTemp.EOF Then
                    If NulSID(SBE) = NulSID("DEPOSIT") Then
                        sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(rsGridUbah!NilaiDeposit) & " ,DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(rsGridUbah!NilaiDeposit) & " WHERE KodeKategory = '" & Left(NulSID(rsGridUbah!model), 20) & "' "
                    Else
                        sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(rsGridUbah!NilaiDeposit) & " ,ket_" & SBE & " = '" & itemPekerjaan & "',DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(rsGridUbah!NilaiDeposit) & " WHERE KodeKategory = '" & Left(NulSID(rsGridUbah!model), 20) & "' "
                    End If
                Else
                    If NulSID(SBE) = NulSID("DEPOSIT") Then
                        sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & ", ket_" & SBE & ",deposit,tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(rsGridUbah!model), 20) & "'," & k2t(rsGridUbah!NilaiDeposit) & ",'" & itemPekerjaan & "'," & k2t(rsGridUbah!NilaiDeposit) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
                    Else
                        sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & " ,deposit,tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(rsGridUbah!model), 20) & "'," & k2t(rsGridUbah!NilaiDeposit) & " ," & k2t(rsGridUbah!NilaiDeposit) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
                    End If
                End If
                Call objHT.InsertData_Ho(sSQL)
            End If
            rsGridUbah.MoveNext
        Loop
        Call objHT.CloseTransactionHO
        If pesan <> "" Then
            MsgBox pesan, vbInformation
        End If
        Unload Me
    Else
        MsgBox pesan, vbCritical, "VALID DATA"
        Exit Sub
    End If

   On Error GoTo 0
   Exit Sub

cmdSimpan_Click_Error:
    If isBegin Then
        Call objHT.CloseTransactionHO_err
    End If
    MsgBox "Error " & Erl & " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmHOProgramFOPUbah"
End Sub

Private Sub cmdSimpanSub_Click()
    Dim isBegin     As Boolean
   On Error GoTo cmdSimpanSub_Click_Error

    isBegin = False
    boleh = True
    
    If CDbl(Text(13)) <= 0 Then
        pesan = pesan + "Jumlah item (QTY) substitusi "
        Text(13) = 0
        Text(13).SetFocus
        boleh = False
    End If
    
    If Text(12) = "" Then
        boleh = False
        pesan = pesan + "Part substitusi. "
    End If
    
    If boleh Then
        Call objHT.OpenTransactionHO
        isBegin = True
        sSQL = "SELECT * FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' AND  MDL = '" & NulSID(rsGrid.Fields(4)) & "' AND ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' AND QTY = '" & NulSID(rsGrid.Fields(9)) & "' AND ITEM_PARTSUB = '" & NulSID(Text(12)) & "'"
        Set rsTemp = objHT.FetchDataHO(sSQL)
        If Not rsTemp.EOF Then
            sSQL = "UPDATE HO_FOP_PROGRAM_SUB SET ITEM_PARTSUB = '" & NulSID(Text(12)) & "',NAMA_PARTSUB = '" & NulSID(Text(11)) & "',QTYSUB = '" & NulSID(Text(13)) & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' AND  MDL = '" & NulSID(rsGrid.Fields(4)) & "' AND ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' AND QTY = '" & NulSID(rsGrid.Fields(9)) & "' AND ITEM_PARTSUB = '" & NulSID(Text(12)) & "'"
        Else
            sSQL = "INSERT INTO HO_FOP_PROGRAM_SUB (FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,ITEM_PARTSUB,NAMA_PARTSUB,QTYSUB) VALUES ('" & NulSID(rsGrid.Fields(6)) & "','" & NulSID(rsGrid.Fields(4)) & "','" & NulSID(rsGrid.Fields(5)) & "','" & NulSID(rsGrid.Fields(7)) & "','" & NulSID(rsGrid.Fields(8)) & "','" & NulSID(rsGrid.Fields(9)) & "','" & NulSID(Text(12)) & "','" & NulSID(Text(11)) & "','" & NulSID(Text(13)) & "')"
        End If
        Call objHT.InsertData_Ho(sSQL)
        Call objHT.CloseTransactionHO
        
        Text(11) = ""
        Text(12) = ""
        Text(13) = 0
        gridSub.ClearFields
        Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(8)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and MDL = '" & NulSID(rsGrid.Fields(4)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
        gridSub.DataSource = rsGridSub
        Call HeaderColomSubstitusi
        gridSub.ReBind
    Else
        MsgBox pesan, vbCritical, "VALID DATA"
        Exit Sub
    End If

   On Error GoTo 0
   Exit Sub

cmdSimpanSub_Click_Error:
    If isBegin Then
        Call objHT.CloseTransactionHO_err
    End If
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSimpanSub_Click of Form frmHOProgramFOPUbah"
End Sub

Private Sub cmdTampilkanKodeNik_Click()
    Call refreshNikKendaraan
    frameNikKendaraan.Visible = True
    frameNikKendaraan.Left = 1440
End Sub


Private Sub createTbl()
    
    If rsNik.State Then rsNik.Close
    Set rsNik = Nothing
    With rsNik
        .Fields.Append "kode", adVarChar, 15
        .Fields.Append "tahun", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridNik.DataSource = rsNik
    GridNik.AlternatingRowStyle = True
End Sub

Private Sub refreshNikKendaraan()
    Set rsTemp = objO.FetchDataHO("select kode,tahun from listTahunNikFOP with (nolock)")
    Call createTbl
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            rsNik.AddNew
            rsNik!kode = Trim(rsTemp("kode"))
            rsNik!Tahun = Trim(rsTemp("tahun"))
            rsNik!ikut = False
            rsNik.Update
            rsTemp.MoveNext
        Loop

        rsNik.MoveFirst
        GridNik.Columns(0).Width = 1200
        GridNik.Columns(1).Width = 2000
        GridNik.Columns(2).Alignment = dbgCenter
        GridNik.Columns(2).Width = 800
    End If
End Sub

Private Sub cmdTutup_Click()
    frameNikKendaraan.Visible = False
    frameNikKendaraan.Left = 14400
End Sub

Private Sub cmdUbah_Click()
    Call refreshGridModel
    Call refreshGrid
End Sub

Private Sub cmdUpdate_Click()
        
    rsGridUbah!program = NulSID(cbTagihan.Text)
    rsGridUbah!nik = NulSID(cbKodeNik.value)
    rsGridUbah!retailawal = Format(tglRetailAwal.value, "dd/mm/yyyy")
    rsGridUbah!retailakhir = Format(tglRetailAkhir.value, "dd/mm/yyyy")
    rsGridUbah!NilaiDeposit = NulSID(Text(5))
    gridData.ReBind
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False, False)
    Call buka_tutupSubtitusi(False)
    Text(0) = pekerjaan
    Call refreshGridModel
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    
    gridPart.Width = Frame1.Width - 200
    gridPart.Height = Frame1.Height - Text(5).Top - Text(5).Height - 200
    
    gridData.Width = (Frame1.Width - 200) - Frame2.Width
    Frame2.Left = gridData.Width + 150
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdGanti.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    ProgBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    ''cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdKeluar.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    cmdImport.Left = cmdCetak.Left - cmdImport.Width - 100
    ProgBar.Left = cmdImport.Left - ProgBar.Width - 100
End Sub

Private Sub gridData_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     
    Text(0) = NulSID(rsGridUbah!jobOrder)
    Text(1) = NulSID(rsGridUbah!model)
    Text(2) = NulSID(rsGridUbah!fullmodel) 'objO.FetchDataHO_String("Select [desc] from MODEL WITH (NOLOCK) where MDL = '" & Text(1) & "' and FS_LOG_DEL = ''")
    cbTagihan.Text = Trim(rsGridUbah!program)
    cbKodeNik.value = Trim(rsGridUbah!nik)
    tglRetailAwal = Format(NulSID(rsGridUbah!retailawal), "dd/mm/yyyy")
    tglRetailAkhir = Format(NulSID(rsGridUbah!retailakhir), "dd/mm/yyyy")
    If cbTagihan.Text = "NRM" Then
        Label15.Visible = True
        Text(5).Visible = True
        SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGridUbah!jobOrder) & "'")
        If SBE = "R010" Then
            SBE = "sb10k"
        ElseIf SBE = "R020" Then
            SBE = "sb20k"
        ElseIf SBE = "R030" Then
            SBE = "sb30k"
        ElseIf SBE = "R040" Then
            SBE = "sb40k"
        ElseIf SBE = "R050" Then
            SBE = "sb50k"
        Else
            SBE = "DEPOSIT"
        End If
        Text(5) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(rsGridUbah!model) & "'")
    Else
        Label15.Visible = False
        Text(5).Visible = False
    End If
End Sub

Private Sub gridPart_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 3 Then
        If gridPart.Columns(ColIndex).value > rsData!qtyasli Then
            gridPart.Columns(ColIndex).value = rsData!qtyasli
            rsData!qty = rsData!qtyasli
        ElseIf gridPart.Columns(ColIndex).value < 1 Then
            gridPart.Columns(ColIndex).value = rsData!qtyasli
            rsData!qty = rsData!qtyasli
        End If
    End If
   ' gridPart.ReBind
End Sub

Private Sub GridPart_FilterChange()
    On Error GoTo errHandler
    Set cols = gridPart.Columns
    Dim C As Integer
    
    C = gridPart.col
    gridPart.HoldFields
    rsGrid.filter = getFilter()
    gridPart.col = C
    gridPart.EditActive = True
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
 
Private Sub gridPart_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        If rsData.State Then
            rsData.MoveFirst
            Do While Not rsData.EOF
                rsData!terlibat = IIf(rsData!terlibat = True, False, True)
                rsData.MoveNext
            Loop
            rsData.MoveFirst
        End If
    End If
End Sub

Private Sub GridPart_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If Not rsGrid.EOF Then
'        Text(5) = rsGrid.Fields(4)      ''model kendaraan   'GridPart.Columns(0)
'        Text(6) = rsGrid.Fields(5)      ''nama model        'GridPart.Columns(1)
'        Text(0) = rsGrid.Fields(6)      ''job order         'GridPart.Columns(2)
'        Text(1) = rsGrid.Fields(7)      ''kode part         'GridPart.Columns(3)
'        Text(2) = rsGrid.Fields(8)      ''nama part         'GridPart.Columns(4)
'        Text(3) = rsGrid.Fields(10)     ''keterangan        'GridPart.Columns(6)
'        Text(7) = rsGrid.Fields(9)      ''qty               'GridPart.Columns(5)
'        ''chPart.Enabled = IIf(rsGrid.Fields(7).value, vbChecked, vbUnchecked)
'
'        Text(14) = rsGrid.Fields(1)     '' NIK TAHUN
'        tglRetailAwal.value = Format(rsGrid.Fields(2), "")
'        tglRetailAkhir.value = Format(rsGrid.Fields(3), "")
'        cbTagihan.Text = rsGrid.Fields(0)
'        Dim SBE As String
'
'        If cbTagihan.Text = "NRM" Then
'            Label15.Visible = True
'            Text(15).Visible = True
'            SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "'")
'            If SBE = "R010" Then
'                SBE = "sb10k"
'            ElseIf SBE = "R020" Then
'                SBE = "sb20k"
'            ElseIf SBE = "R030" Then
'                SBE = "sb30k"
'            ElseIf SBE = "R040" Then
'                SBE = "sb40k"
'            ElseIf SBE = "R050" Then
'                SBE = "sb50k"
'            Else
'                SBE = "DEPOSIT"
'            End If
         '  Text(15) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(rsGrid.Fields(4)) & "'")
'        Else
'            Label15.Visible = False
'            Text(15).Visible = False
'
'        End If
'
'        gridSub.ClearFields
'        Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(8)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and MDL = '" & NulSID(rsGrid.Fields(4)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
'        ''If Not rsGridSub.EOF Then
'            Text(12) = ""
'            Text(11) = ""
'            Text(13) = 0
'
'            gridSub.DataSource = rsGridSub
'            Call HeaderColomSubstitusi
'            gridSub.ReBind
'        ''End If
'    End If
End Sub

 
Private Sub gridSub_FilterChange()
On Error GoTo errHandler
    Set cols = gridSub.Columns
    Dim C As Integer
    
    C = gridSub.col
    gridSub.HoldFields
    rsGridSub.filter = getFilter()
    gridSub.col = C
    gridSub.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub gridSub_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGridSub.EOF Then
        Text(12) = rsGridSub.Fields(0)  ' kodeitem
        Text(11) = rsGridSub.Fields(1)  ' nama part
        Text(13) = rsGridSub.Fields(2)  ' qty
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
                gSearchOption = "JOB_ORDER_FOP"
                gBantuan = "SBE"    '/ "ERC" '/
                frmCariHO.show (vbModal)
            Case 8, 12
                ' PART DARI HO
                gMainKeyCari = Text(Index)
                gSearchOption = "PARTHO"
                frmCariHO.show (vbModal)
            Case 1
                Cancel = True
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
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
                Set oRs = objO.FetchDataHO("SELECT * FROM ho_fop_job WITH (NOLOCK) WHERE FS_Item = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER_FOP"
                    gBantuan = "SBE"
                    frmCariHO.show (vbModal)
                End If
                oRs.Close
            End If
        Case 8
            If Text(Index) <> "" Then
            ' PART MANUAL INPUTAN HO
                Set oRs = objO.FetchDataHO("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & Text(Index) & "'")
                If Not oRs.BOF Then
                    Text(4) = Trim(oRs("NAMA_PART"))
                End If
                oRs.Close
            End If
        Case 12
            If Text(Index) <> "" Then
            ' PART MANUAL INPUTAN HO
                Set oRs = objO.FetchDataHO("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & Text(Index) & "'")
                If Not oRs.BOF Then
                    Text(11) = Trim(oRs("NAMA_PART"))
                End If
                oRs.Close
            End If
        Case 1
            If Text(Index) <> "" Then
                'Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL like '" & Text(1) & "%' and FS_LOG_DEL = ''")
                '' ttidak bisa menggunakan like karena bisa jadi data ho_part / ho_sublet tidak ada d model tertentu
                
                ''Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL = '" & Text(1) & "' and FS_LOG_DEL = ''")
                ''20200625 Dibuka kembali
                Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL like '" & Text(1) & "%' and FS_LOG_DEL = ''")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCariHO.show (vbModal)
                Else
                    Text(2) = oRs("Desc")
                End If
                oRs.Close
            End If
    End Select
End Sub


Private Sub ProsesExport(onV As String)
    Dim oApp    As Excel.Application
    Dim oWB     As Excel.Workbook
    Dim nAsc    As Integer
    Set oApp = New Excel.Application
    oApp.Visible = False
    Set oWB = oApp.Workbooks.Add
    With oWB
'        For C = 0 To rsDownload.Fields.Count - 1
'            .Sheets("Sheet1").Cells(1, C + 1).value = rsDownload.Fields(C).Name
'            nAsc = nAsc + 1
'        Next
        .Sheets("Sheet1").Cells(1, 1).value = "NAMAPEKERJAAN (max 50 Kar - sesuai master job)"
        .Sheets("Sheet1").Cells(1, 2).value = "KODEMODEL (max 20 kar - sesuai master model)"
        .Sheets("Sheet1").Cells(1, 3).value = "TAGIHAN ( pilihan NRM / TAM)"
        .Sheets("Sheet1").Cells(1, 4).value = "KODENIK (inisial Tahun NIK - sesuai master)"
        .Sheets("Sheet1").Cells(1, 5).value = "TGLRETAILAWAL (DD/MM/YYYY)"
        .Sheets("Sheet1").Cells(1, 6).value = "TGLRETAILAKHIR (DD/MM/YYYY)"
        .Sheets("Sheet1").Cells(1, 7).value = "KODEITEM (max 15 - sesuai master part)"
        .Sheets("Sheet1").Cells(1, 8).value = "NOMINAL DEPOSITO"

        .Sheets("Sheet1").Cells(2, 1).CopyFromRecordset rsDownload
        .Sheets("Sheet1").Range("1:1").Font.Bold = True
        .Sheets("Sheet1").Select
        .Close SaveChanges:=True, FileName:=onV
    End With
    Set oWB = Nothing
    oApp.Quit
    Set oApp = Nothing
End Sub

Private Sub refreshComboNIK()
    sSQL = "select kode,tahun from listTahunNikFOP with (nolock) "

    Adodc1.ConnectionString = SQLCONBengkel
    Adodc1.RecordSource = sSQL
    cbKodeNik.DataSourceList = Adodc1
    cbKodeNik.DataFieldList = "kode"
    
    cbKodeNik.Columns(0).Width = 2000
    cbKodeNik.Columns(1).Width = 3500
    cbKodeNik.Columns(0).Caption = "Kode"
    cbKodeNik.Columns(1).Caption = "Tahun"
    cbKodeNik.Text = ""
End Sub
