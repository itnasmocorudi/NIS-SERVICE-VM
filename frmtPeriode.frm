VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtPeriode 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PERIODE"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   11280
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8805
      Left            =   600
      TabIndex        =   0
      Top             =   -360
      Width           =   11295
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3480
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   5280
         Visible         =   0   'False
         Width           =   255
      End
      Begin TDBDate6Ctl.TDBDate tglAwal 
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   5160
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         Calendar        =   "frmtPeriode.frx":0000
         Caption         =   "frmtPeriode.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmtPeriode.frx":0198
         Keys            =   "frmtPeriode.frx":01B6
         Spin            =   "frmtPeriode.frx":0214
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
         Format          =   "mm/dd/yyyy"
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
         Text            =   "08/08/2016"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   42590
         CenturyMode     =   0
      End
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7435
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tahun"
         Columns(0).DataField=   "tahun"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Bulan"
         Columns(1).DataField=   "bulan"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Tanggal Awal"
         Columns(2).DataField=   "tglAwal"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Tanggal Akhir"
         Columns(3).DataField=   "tglAkhir"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "id"
         Columns(4).DataField=   "id"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
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
         Splits(0)._ColumnProps(20)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         WrapCellPointer =   -1  'True
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H80000016&,.bold=-1"
         _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
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
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
      End
      Begin ucXPButton.XPButton cmdDelete 
         Height          =   495
         Left            =   5520
         TabIndex        =   5
         Top             =   6120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&d"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtPeriode.frx":023C
         PICN            =   "frmtPeriode.frx":0258
         PICH            =   "frmtPeriode.frx":07E0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSave 
         Height          =   495
         Left            =   4320
         TabIndex        =   6
         Top             =   6120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&s"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtPeriode.frx":0E1C
         PICN            =   "frmtPeriode.frx":0E38
         PICH            =   "frmtPeriode.frx":134B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdAdd 
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   6120
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtPeriode.frx":18ED
         PICN            =   "frmtPeriode.frx":1909
         PICH            =   "frmtPeriode.frx":1D55
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdEdit 
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   6120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&e"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtPeriode.frx":221D
         PICN            =   "frmtPeriode.frx":2239
         PICH            =   "frmtPeriode.frx":26E4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdExit 
         Height          =   495
         Left            =   9840
         TabIndex        =   9
         Top             =   6120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&x"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtPeriode.frx":2C09
         PICN            =   "frmtPeriode.frx":2C25
         PICH            =   "frmtPeriode.frx":30F9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBDate6Ctl.TDBDate tglAkhir 
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   5520
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   661
         Calendar        =   "frmtPeriode.frx":3644
         Caption         =   "frmtPeriode.frx":3770
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmtPeriode.frx":37DC
         Keys            =   "frmtPeriode.frx":37FA
         Spin            =   "frmtPeriode.frx":3858
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
         Format          =   "mm/dd/yyyy"
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
         Text            =   "08/08/2016"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   42590
         CenturyMode     =   0
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Akhir :"
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
         Left            =   500
         TabIndex        =   13
         Top             =   5640
         Width           =   1485
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Awal :"
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
         Left            =   555
         TabIndex        =   12
         Top             =   5200
         Width           =   1425
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: LIST PERIODE ::."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   4
         Top             =   60
         Width           =   10800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   -120
         X2              =   11160
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: DATA PERIODE ::."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   4560
         TabIndex        =   3
         Top             =   4740
         Width           =   2175
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         X1              =   -120
         X2              =   11160
         Y1              =   6000
         Y2              =   6000
      End
      Begin VB.Label lblJml 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   7680
         TabIndex        =   2
         Top             =   60
         Width           =   1380
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11055
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4680
         Width           =   11055
      End
   End
End
Attribute VB_Name = "frmtPeriode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsM As ADODB.Recordset
Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String
Private C As Integer
Private isOnWar As Boolean

Private Sub cmdEdit_Click()
   Call BindToControl(False, True, True)
End Sub

Private Sub Form_Load()
    Call SetControlForm
    Call FillGridClass
    Call BindToControl(True, False, False)
End Sub
'
Private Sub SetControlForm()

End Sub
'    With Me
'        .cmdAdd.Enabled = True
'       ' .cmdCreate.Visible = False
'        .cmdEdit.Enabled = True
'        .cmdSave.Visible = True
'        .cmdSave.Enabled = False
'        .cmdCancel.Visible = False
'        .cmdDelete.Visible = True
'        .cmdDelete.Enabled = True
'        .Grid.Enabled = True
'        .txt(0).BackColor = fColor
'        .txt(1).BackColor = fColor
'        .txt(2).BackColor = fColor
'        .txtRp(10).BackColor = fColor
'        .txtRp(11).BackColor = fColor
'        .txtRp(12).BackColor = fColor
'        .txtRp(13).BackColor = fColor
'        .cboWar.BackColor = fColor
'        .cboWar.Clear
'        .cboWar.AddItem "NO"
'        .cboWar.AddItem "YES"
'        .cboWar.Text = "NO"
'        .cmdExit.Enabled = True
'    End With
''End Sub

Private Sub FillGridClass()
    Set rsM = objO.FetchData("SELECT * FROM periode")
    If rsM.RecordCount = 0 Then
        With Me
'            .cmdAdd.Enabled = True
'          '  .cmdCreate.Visible = False
'            .cmdEdit.Enabled = False
'            .cmdSave.Visible = True
'            .cmdSave.Enabled = False
'            .cmdCancel.Visible = False
'            .cmdDelete.Visible = False
'            .cmdDelete.Enabled = True
'            .Grid.Enabled = False
'            .Grid.DataSource = Nothing
'            .lblJml.Caption = "0"
'            .txt(0).BackColor = fColor
'            .txt(1).BackColor = fColor
'            .txt(2).BackColor = fColor
'            .txtRp(10).BackColor = fColor
'            .txtRp(11).BackColor = fColor
'            .txtRp(12).BackColor = fColor
'            .txtRp(13).BackColor = fColor
'            .cboWar.BackColor = fColor
'            .txt(0).Text = ""
'            .txt(1).Text = ""
'            .txt(2).Text = ""
'            .txtRp(10) = 0
'            .txtRp(11) = 0
'            .txtRp(12) = 0
'            .txtRp(13) = 0
'            .cboWar.Text = "NO"
'            .cmdExit.Enabled = True
        End With
        Exit Sub
    End If
    Set Grid.DataSource = rsM
    lblJml.Caption = rsM.RecordCount
 
         

End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    tglAwal.value = Format(Grid.Columns("tglAwal").value, "dd/mm/yyyy")
    tglAkhir.value = Format(Grid.Columns("tglAkhir").value, "dd/mm/yyyy")
    Text1.Text = Grid.Columns("id").value
End Sub
Private Sub BindToControl(Edit As Boolean, save As Boolean, batal As Boolean)
On Error GoTo ErrNull
    cmdEdit.Enabled = Edit
    cmdSave.Enabled = save
    cmdDelete.Enabled = batal
    
    If Edit = True Then
      tglAwal.Enabled = False
      tglAkhir.Enabled = False
    Else
      tglAwal.Enabled = True
      tglAkhir.Enabled = True
    End If
ErrNull:
End Sub

'===============================================================================================
'========= A D D ======== A D D ======== A D D =================================================
'===============================================================================================
'Private Sub CmdAdd_Click()
'    With Me
'        .cmdAdd.Enabled = False
'       ' .cmdCreate.Visible = True
'       ' .cmdCreate.Enabled = True
'        .cmdEdit.Enabled = False
'        .cmdSave.Visible = False
'        .cmdCancel.Visible = True
'        .cmdCancel.Enabled = True
'        .cmdDelete.Visible = False
'        .Grid.Enabled = False
'        .txt(0).BackColor = tColor
'        .txt(1).BackColor = tColor
'        .txt(2).BackColor = tColor
'        .txtRp(10).BackColor = tColor
'        .txtRp(11).BackColor = tColor
'        .txtRp(12).BackColor = tColor
'        .txtRp(13).BackColor = tColor
'        .cboWar.BackColor = tColor
'        .txt(0).Text = ""
'        .txt(1).Text = ""
'        .txt(2).Text = ""
'        .txtRp(10) = 0
'        .txtRp(11) = 0
'        .txtRp(12) = 0
'        .txtRp(13) = 0
'        .cboWar.Text = "NO"
'        .cmdExit.Enabled = False
'        .txt(0).SetFocus
'
'        If isOnWar = True Then
'            .cboWar.BackColor = fColor
'        Else
'            .cboWar.BackColor = tColor
'        End If
'    End With
'End Sub
'
'Private Sub cmdCreate_Click()
'    If FillNotNull = True Then
'        rsM.Find "Kode = '" & Trim(txt(0).Text) & "'", , adSearchForward, 1
'        If Not rsM.EOF Then
'            MsgBox "Data Kelas " & txt(0).Text & " sudah terdaftar, Ganti Kode !", vbInformation + vbOKOnly, "Car Class"
'            Exit Sub
'        End If
'        If MsgBox("Simpan Data Kelas Baru ?", vbQuestion + vbYesNo, "Car Class") = vbYes Then
'            Call objO.InsertDataHO("INSERT INTO CarClass(ClassCode, ClassName, Price,priceTAM, Note, DiscNRM, DiscNON, war, FS_LOG_DEL, FD_TGL_DEL) " & _
'                                   "VALUES('" & Trim(txt(0).Text) & "','" & Trim(txt(1).Text) & "'," & RepRp(txtRp(10)) & "," & RepRp(txtRp(11)) & ",'" & Trim(txt(2).Text) & "'," & RepRp(txtRp(12)) & "," & RepRp(txtRp(13)) & "," & IIf(cboWar.Text = "YES", 1, 0) & ",'','01/01/3000')")
'            Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('KELAS', '" & Trim(txt(0).Text) & "','S')")
'            Call SetControlForm
'            Call FillGridClass
'        End If
'    Else
'        Call NotValid_Input
'    End If
'End Sub

'===============================================================================================
'======= E D I T =========== E D I T =========== E D I T =======================================
'===============================================================================================
'Private Sub cmdEdit_Click()
'    With Me
'        .cmdAdd.Enabled = False
'    '    .cmdCreate.Visible = False
'        .cmdEdit.Enabled = False
'        .cmdSave.Visible = True
'        .cmdSave.Enabled = True
'        .cmdCancel.Visible = True
'        .cmdCancel.Enabled = True
'        .cmdDelete.Visible = False
'        .Grid.Enabled = False
'        .txt(0).BackColor = fColor
'        .txt(1).BackColor = tColor
'        .txt(2).BackColor = tColor
'        .txtRp(10).BackColor = tColor
'        .txtRp(11).BackColor = tColor
'        .txtRp(12).BackColor = tColor
'        .txtRp(13).BackColor = tColor
'        .cboWar.BackColor = tColor
'        .cmdExit.Enabled = False
'        .txt(1).SetFocus
'
'        If isOnWar = True Then
'            If .cboWar.Text = "YES" Then
'                .cboWar.BackColor = tColor
'            Else
'                .cboWar.BackColor = fColor
'            End If
'        Else
'            .cboWar.BackColor = tColor
'        End If
'    End With
'End Sub

Private Sub cmdSave_Click()
   Call objO.InsertData("UPDATE periode SET tglAwal = '" & Format(tglAwal.value, "mm/dd/yyyy") & "',tglAkhir = '" & Format(tglAkhir.value, "mm/dd/yyyy") & "' WHERE id = '" & Trim(Text1.Text) & "' ")
   MsgBox "Data berhasil disimpan", vbInformation
   Call SetControlForm
    Call FillGridClass
    Call BindToControl(True, False, False)
End Sub

'===============================================================================================
'===============================================================================================
'Private Function FillNotNull() As Boolean
'    FillNotNull = True
'    If txt(0).Text = "" Or txt(1).Text = "" Then
'        FillNotNull = False
'    End If
'End Function
Private Sub NotValid_Input()
    MsgBox "Data Yang Anda Inputkan Belum Lengkap !", vbInformation + vbOKOnly, "Car Class"
End Sub
'===============================================================================================
'===============================================================================================

'===============================================================================================
'========= C A N C E L ========= C A N C E L ========= C A N C E L =============================
'===============================================================================================
Private Sub cmdCancel_Click()
    Call SetControlForm
    Call FillGridClass
End Sub

'===============================================================================================
'========= D E L E T E ========= D E L E T E ========= D E L E T E =============================
'===============================================================================================
Private Sub CmdDelete_Click()
Call SetControlForm
    Call FillGridClass
    Call BindToControl(True, False, False)
End Sub

Private Sub grid_FilterChange()
On Error GoTo err
   Set cols = Grid.Columns
   
   C = Grid.col
   Grid.HoldFields
   rsM.filter = getFilter()
   If rsM.EOF = True Then
      For Each col In Grid.Columns
         col.FilterText = ""
      Next col
      rsM.filter = ""
   End If
   Grid.col = C
   Grid.EditActive = True
   lblJml.Caption = rsM.RecordCount
   Exit Sub
err:
   Resume Next
End Sub

Private Function getFilter() As String
   On Error GoTo err

On Error GoTo err
Dim tmp As String
Dim n As Integer

   For Each col In cols
      If Trim(col.FilterText) <> "" Then
         n = n + 1
         If n > 1 Then
            tmp = tmp & " AND "
         End If
         tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub grid_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 6 Then rsM.Sort = "[" & Grid.Columns(ColIndex).DataField & "]" & " " & Urut
4        Grid.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        lblJml.Caption = rsM.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter ")
End Sub
 

Private Sub cmdExit_Click()
    Unload Me
End Sub
 
