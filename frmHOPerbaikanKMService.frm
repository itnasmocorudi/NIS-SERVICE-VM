VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOPerbaikanKMService 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "PERBAIKAN KM SERVICE"
   ClientHeight    =   7875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9120
      TabIndex        =   36
      Top             =   4560
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdGanti 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   6000
      TabIndex        =   40
      Top             =   4560
      Visible         =   0   'False
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Ganti KM Service"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      Height          =   7485
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   13455
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2535
         Left            =   120
         TabIndex        =   46
         Top             =   4680
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4471
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
         Splits(0).RecordSelectorWidth=   953
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   15790320
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
         PrintInfos(0)._StateFlags=   3
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
         DeadAreaBackColor=   15790320
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
      End
      Begin MSAdodcLib.Adodc AdcKode 
         Height          =   330
         Left            =   9120
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DCKodeGroup 
         Height          =   375
         Left            =   2520
         TabIndex        =   0
         Top             =   600
         Width           =   3195
         AllowInput      =   0   'False
         AllowNull       =   0   'False
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RowHeight       =   423
         Columns(0).Width=   3200
         _ExtentX        =   5644
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "- Pilih Nama Dealer -"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   13320
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   4
            Left            =   930
            TabIndex        =   4
            Top             =   510
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   5
            Left            =   930
            TabIndex        =   5
            Top             =   800
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   503
            _Version        =   393216
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   6
            Left            =   930
            TabIndex        =   6
            Top             =   1100
            Width           =   3405
            _ExtentX        =   6006
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   7
            Left            =   930
            TabIndex        =   7
            Top             =   1400
            Width           =   1845
            _ExtentX        =   3254
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
            Mask            =   ">9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   10
            Left            =   6390
            TabIndex        =   8
            Top             =   450
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   12
            Left            =   6390
            TabIndex        =   9
            Top             =   1110
            Width           =   2295
            _ExtentX        =   4048
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   15
            Left            =   6390
            TabIndex        =   10
            Top             =   1710
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   14
            Left            =   6390
            TabIndex        =   11
            Top             =   1410
            Width           =   3135
            _ExtentX        =   5530
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   13
            Left            =   8730
            TabIndex        =   12
            Top             =   1110
            Width           =   795
            _ExtentX        =   1402
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
            Mask            =   ">CCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   16
            Left            =   930
            TabIndex        =   13
            Top             =   1980
            Width           =   1170
            _ExtentX        =   2064
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   17
            Left            =   930
            TabIndex        =   14
            Top             =   2280
            Width           =   3210
            _ExtentX        =   5662
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   11
            Left            =   6390
            TabIndex        =   15
            Top             =   795
            Width           =   1215
            _ExtentX        =   2143
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   8
            Left            =   930
            TabIndex        =   16
            Top             =   1680
            Width           =   1845
            _ExtentX        =   3254
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
            Mask            =   ">9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   9
            Left            =   6390
            TabIndex        =   17
            Top             =   2010
            Width           =   1845
            _ExtentX        =   3254
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
            Mask            =   ">9999-9999999999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   18
            Left            =   6390
            TabIndex        =   18
            Top             =   2310
            Width           =   3210
            _ExtentX        =   5662
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
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   2
            Left            =   10800
            TabIndex        =   41
            Top             =   1035
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox text 
            Height          =   345
            Index           =   20
            Left            =   10800
            TabIndex        =   42
            Top             =   1440
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   609
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "KM Akhir"
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
            Left            =   9720
            TabIndex        =   44
            Top             =   1365
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Service"
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
            Left            =   9570
            TabIndex        =   43
            Top             =   1080
            Width           =   1140
         End
         Begin VB.Line Line8 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   3840
            Y1              =   195
            Y2              =   195
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   3840
            Y1              =   75
            Y2              =   75
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Customer && Kendaraan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   4200
            TabIndex        =   33
            Top             =   0
            Width           =   3675
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   8520
            X2              =   13320
            Y1              =   75
            Y2              =   75
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000000FF&
            X1              =   8520
            X2              =   13320
            Y1              =   195
            Y2              =   195
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   13335
         End
         Begin VB.Label Label3 
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
            Left            =   480
            TabIndex        =   32
            Top             =   1725
            Width           =   315
         End
         Begin VB.Label Label27 
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
            Left            =   375
            TabIndex        =   31
            Top             =   555
            Width           =   405
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Left            =   285
            TabIndex        =   30
            Top             =   855
            Width           =   495
         End
         Begin VB.Label Label23 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Left            =   450
            TabIndex        =   29
            Top             =   1155
            Width           =   330
         End
         Begin VB.Label Label22 
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
            Left            =   285
            TabIndex        =   28
            Top             =   1440
            Width           =   375
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
            Left            =   5595
            TabIndex        =   27
            Top             =   495
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
            Left            =   5820
            TabIndex        =   26
            Top             =   1125
            Width           =   420
         End
         Begin VB.Label Label5 
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
            Left            =   5760
            TabIndex        =   25
            Top             =   1725
            Width           =   480
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Left            =   240
            TabIndex        =   24
            Top             =   2040
            Width           =   645
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Frame."
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
            TabIndex        =   23
            Top             =   2280
            Width           =   510
         End
         Begin VB.Label Label55 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NO INVOICE"
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
            Left            =   5445
            TabIndex        =   22
            Top             =   840
            Width           =   900
         End
         Begin VB.Label Label2 
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
            Left            =   5880
            TabIndex        =   21
            Top             =   2085
            Width           =   435
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Engine."
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
            Left            =   5820
            TabIndex        =   20
            Top             =   2400
            Width           =   540
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   3960
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   3960
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label74 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data Customer && Kendaraan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   4320
            TabIndex        =   19
            Top             =   45
            Width           =   3675
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   8640
            X2              =   12720
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   8640
            X2              =   12720
            Y1              =   240
            Y2              =   240
         End
      End
      Begin MSMask.MaskEdBox text 
         Height          =   345
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Top             =   1050
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   345
         Index           =   1
         Left            =   2640
         TabIndex        =   38
         Top             =   4200
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   25
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERBAIKAN KM SERVICE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   5280
         TabIndex        =   45
         Top             =   120
         Width           =   3540
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BorderColor     =   &H80000002&
         FillColor       =   &H0080FF80&
         FillStyle       =   0  'Solid
         Height          =   525
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   15
         Width           =   13455
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "KM Service Benar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   645
         TabIndex        =   39
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NO WO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1560
         TabIndex        =   35
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Dealer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   34
         Top             =   650
         Width           =   1215
      End
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
End
Attribute VB_Name = "frmHOPerbaikanKMService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private dealer As String
Private dBase As String
Private koneksi As String
Private usernameDealer As String
Private passwordDealer As String
Private usernameHO As String
Private passwordHO As String
Private usernameGlobal As String
Private passwordGlobal As String
Private isDoubleCustDat As Boolean
Private isDoubleCustDatHO As Boolean
Private isGR As Boolean
Private FS_KD_DEALER As String
Private serviceSekarang As Integer
Private kelipatan As Integer
Private SBE As Double
Dim rsHistory As New ADODB.Recordset

Private isDealerNewEnkrip As Boolean


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGanti_Click()
          Dim boleh As Boolean
1         boleh = True
          
2         If Len(Trim(Text(0))) = 0 Then
3             MsgBox "No WO Harus diisi", vbCritical, "Informasi"
4             Exit Sub
5         ElseIf Len(Trim(Text(1))) = 0 Then
6             MsgBox "KM Service yang benar harus diisi!", vbCritical, "Informasi"
7             Exit Sub
8         Else
9             If Not rsHistory.EOF Then
10                rsHistory.MoveFirst
11                Do While Not rsHistory.EOF
12                    If CDbl(rsHistory!km) > CDbl(Trim(Text(1))) Then
13                        MsgBox "KM Service tidak boleh lebih kecil dari history sebelumnya!", vbCritical, "Informasi"
14                        Exit Sub
15                    End If
16                    rsHistory.MoveNext
17                Loop
18            End If
19        End If
          
20        If MsgBox("Apakah yakin mengganti KM Service " & Trim(Text(20)) & " dengan " & Format(CDbl(Trim(Text(1))), "#,#") & " ?", vbYesNo, "Konfirmasi Penggantian KM Service") = vbYes Then
              'lakukan pergantian data
21            On Error GoTo vbError
22            koneksi = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer
23            erTranAll = False
              'perbaikan kecabang
24            Call objHT.OpenTransactionPerbaikanDataHO(koneksi)
              
              'ganti no rangka
              
25            Call objHT.UpdatePerbaikanDataHO("UPDATE JOBDTA SET KM = " & Format(CDbl(Trim(Text(1))), "") & " WHERE NO_WO = '" & Trim(Text(0)) & "'")
26            Call objHT.UpdatePerbaikanDataHO("UPDATE T_INVOICE SET KM = " & Format(CDbl(Trim(Text(1))), "") & " WHERE NO_WO = '" & Trim(Text(0)) & "'")
27            Call objHT.UpdatePerbaikanDataHO("UPDATE HISTORY SET KM = " & Format(CDbl(Trim(Text(1))), "") & " WHERE NO_WO = '" & Trim(Text(0)) & "'")
28            Call objHT.UpdatePerbaikanDataHO("UPDATE CUSTDAT SET FS_KM_AKHIR = " & Format(CDbl(Trim(Text(1))), "") & " WHERE NO_POLISI = '" & Trim(Text(16)) & "' AND NO_CHASSIS = '" & Trim(Text(17)) & "'")

              
                  
29            If isGR = False Then
                  'lakukan perhitungan estimasi next service
                  'parameter nextestimasi NORANGKA, TANGGAL_SERVICE, KM,SBE,KELIPATAN,NOINV,NOWO,CUSTOMER,CUSTOMER_ADDRESS, NO_POLISI,TELEPON,GSM, CDMA
30                SBE = CDbl(serviceSekarang) * 1000
                  Dim rsNext As ADODB.Recordset
                  Dim nextS As String
31                Set rsNext = objHT.FetchDataPerbaikanDataHO("SELECT * FROM T_PROFIL")
32                nextS = IIf(Trim(rsNext("ESTIMASIDEALER")) <> "", Trim(rsNext("ESTIMASIDEALER")), "NextEstimasi")
33                Call objHT.UpdatePerbaikanDataHO("exec " & nextS & " '" & Trim(Text(17)) & "','" & TglYYYYMMDD(Text(2)) & "'," & Format(CDbl(Text(1)), "") & ",'" & SBE & "'," & kelipatan & ",'" & Text(11) & "','" & Trim(Text(0)) & "','" & Trim(Text(4)) & "','" & Trim(Text(5)) & "','" & Trim(Text(16)) & "','" & Trim(Text(7)) & "','" & Trim(Text(8)) & "','" & Trim(Text(9)) & "'")
                  
34            End If
              
              ''2018-08-31
              ''========================================================
35            Call objHT.UpdatePerbaikanDataHO("UPDATE [" & dbReplikasiHO & "].dbo.JOBDTA SET KM = '" & Format(CDbl(Trim(Text(1))), "") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND FS_KD_DEALER = '" & Trim(FS_KD_DEALER) & "'")
36            Call objHT.UpdatePerbaikanDataHO("UPDATE [" & dbReplikasiHO & "].dbo.T_INVOICENEW SET KM = '" & Format(CDbl(Trim(Text(1))), "") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND CABANG = '" & Trim(dBase) & "'")
37            Call objHT.UpdatePerbaikanDataHO("UPDATE [" & dbReplikasiHO & "].dbo.HISTORYNEW SET KM = '" & Format(CDbl(Trim(Text(1))), "") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND CABANG = '" & Trim(dBase) & "'")

              
              ''========================================================
              
38            Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE [BENGKEL99].dbo.JOBDTA SET KM = '" & Format(CDbl(Trim(Text(1))), "") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND FS_KD_DEALER = '" & Trim(FS_KD_DEALER) & "'")
39            Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE [BENGKEL99].dbo.T_INVOICENEW SET KM = '" & Format(CDbl(Trim(Text(1))), "") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND CABANG = '" & Trim(dBase) & "'")
40            Call objHT.UpdatePerbaikanDataHoConnectionHO("UPDATE [BENGKEL99].dbo.HISTORYNEW SET KM = '" & Format(CDbl(Trim(Text(1))), "") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND CABANG = '" & Trim(dBase) & "'")
             'lakukan penyimpanan log d profilbengkel cabang
              
41            Call objHT.UpdatePerbaikanDataHO("INSERT INTO PROFILBENGKEL.dbo.LOGUPDATE VALUES ('UpdateKM','" & Trim(Text(0)) & "','" & Trim(Text(17)) & "','" & gSA & "',GETDATE(),'" & Trim(Text(1)) & "','" & Trim(Text(20)) & "')")
              
42            If erTranAll = False Then
43                Call objHT.CloseTransactionPerbaikanDataHO
44                Call saveLogMenu(gSA, ComptName, gVersi, "frmHOPerbaikanKMService", "gantiKMService", 1)
45                MsgBox "KM Service berhasil diubah!", vbInformation, "Informasi Berhasil Ganti KM Service"
46                Call hapusForm
47                Text(0) = ""
48                Call FetchCombo
49                Exit Sub
50            Else
51                Call objHT.CloseTransactionPerbaikanDataHO_err
52                Call saveLogMenu(gSA, ComptName, gVersi, "frmHOPerbaikanKMService", "gantiKMService", 2)
53                MsgBox "Terjadi kesalahan saat mengganti No Rangka!", vbInformation, "Gagal Simpan"
54                Exit Sub
55            End If
vbError:
56            Call objHT.CloseTransactionPerbaikanDataHO_err
57            Call saveLogMenu(gSA, ComptName, gVersi, "CloseTransactionPerbaikanDataHO_err", "gantiNoRangka", 2)
58            MsgBox "Terjadi kesalahan saat mengganti No Rangka!", vbInformation, "Gagal Simpan"
59            Exit Sub
60        End If
End Sub

Private Sub DCKodeGroup_Change()
    'dealer = ip dealer
    'dbase  = database dealer
    
    isDealerNewEnkrip = False
    
    FS_KD_DEALER = DCKodeGroup.Columns(0).value
    dealer = DCKodeGroup.Columns(2).value
    dBase = DCKodeGroup.Columns(3).value
    
    If DCKodeGroup.Columns(6).value = "1" Then isDealerNewEnkrip = True
    
    If isDealerNewEnkrip Then
        usernameDealer = Dekripsi(DCKodeGroup.Columns(4).value, KataKunci(3))
        passwordDealer = Dekripsi(DCKodeGroup.Columns(5).value, KataKunci(5))
    Else
        usernameDealer = "sa" 'DekripsiGamaLama(DCKodeGroup.Columns(4).value, KataKunciGamaLama(3))
        passwordDealer = "nasmoco~123" 'DekripsiGamaLama(DCKodeGroup.Columns(5).value, KataKunciGamaLama(5))
    End If
End Sub

Private Sub DCKodeGroup_Click()
    Call DCKodeGroup_Change
End Sub

Private Sub Form_Load()
    Call FetchCombo
End Sub

Private Sub FetchCombo()
    AdcKode.ConnectionString = SQLConHO
    AdcKode.RecordSource = "select FS_KD_DEALER, FS_NM_DEALER,VPNCABANG,DBASECBG,USERCBG,PASSCBG,NEWENKRIP from t_dealer WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1"
    DCKodeGroup.DataSourceList = AdcKode
    DCKodeGroup.DataFieldList = "FS_NM_DEALER"
    DCKodeGroup.Columns(0).Width = 1200
    DCKodeGroup.Columns(1).Width = 4000
    DCKodeGroup.Columns(0).Caption = "Kode Dealer"
    DCKodeGroup.Columns(1).Caption = "Nama Dealer"
    DCKodeGroup.Columns(2).Visible = False
    DCKodeGroup.Columns(3).Visible = False
    DCKodeGroup.Columns(4).Visible = False
    DCKodeGroup.Columns(5).Visible = False
    DCKodeGroup.Columns(6).Visible = False
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
'
'
'        Select Case Index
'            Case 23, 33, 95
'                SSTab1.Tab = 0
'            Case 34, 43
'                SSTab1.Tab = 1
'            Case 44
'                SSTab1.Tab = 2
'            Case 54, 53
'                SSTab1.Tab = 3
'        End Select
    End If
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call text_Validate(Index, True)
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If dealer <> "" Then
                If Text(Index) <> "" Then
                    Call hapusForm
                    Call PencarianKMServiceTerakhir
                Else
                    MsgBox "NO WO harus diisi!", vbCritical, "Informasi"
                End If
            Else
                MsgBox "Dealer harus dipilih!", vbCritical, "Informasi"
            End If
            
    End Select
End Sub

Private Sub PencarianKMServiceTerakhir()
          Dim noWO As String
          Dim sqlSql As String
1         On Error GoTo errHandler
          
2         noWO = Trim(Text(0))
3         isGR = False
          
          Dim rsJobDta As New ADODB.Recordset
          
          'dealer = "10.10.26.171"
          'dbase = "bengkel05"
                    
4         usernameGlobal = gUserKoneksi
5         passwordGlobal = gPasswordKonesi

          Dim rs As New ADODB.Recordset
       
          'mengambil data dealer
6         rsJobDta.Open "SELECT J.*,C.GSM,C.CDMA FROM JOBDTA J WITH (NOLOCK) INNER JOIN CUSTDAT C WITH (NOLOCK) ON C.NO_CHASSIS = J.NO_CHASSIS WHERE NO_WO = '" & Trim(noWO) & "'", ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
          
7         If Not rsJobDta.EOF Then
              'jika ada data cusdat no rangka benar maka tampilkan
8             Text(4) = rsJobDta("CUSTOMER")
9             Text(5) = Left(rsJobDta("CUST_ADRES"), 60)
10            Text(6) = rsJobDta("KOTA")
11            Text(7) = rsJobDta("TELPON")
12            Text(8) = rsJobDta("GSM")
13            Text(9) = rsJobDta("CDMA")
14            Text(10) = rsJobDta("FS_TH_MOBIL")
15            Text(11) = rsJobDta("NOINV")
16            Text(12) = rsJobDta("MDL")
17            Text(13) = rsJobDta("MODEL")
18            Text(14) = rsJobDta("FMDL")
19            Text(15) = rsJobDta("COLOR")
20            Text(16) = rsJobDta("NO_POLISI")
21            Text(17) = rsJobDta("NO_CHASSIS")
22            Text(18) = rsJobDta("NO_ENGINE")
              
23            If rsJobDta("FS_RC") = "" Then
24                isGR = True
                  
25            Else
26                serviceSekarang = CInt(Right(rsJobDta("FS_RC"), 3))
                  
27                If serviceSekarang = 1 Then
28                    kelipatan = 10000
29                ElseIf serviceSekarang Mod 10 = 0 Then
                      'kelipatan 10000
30                    kelipatan = 10000
31                ElseIf serviceSekarang Mod 10 = 5 Then
                      'kelipatan = 5000
32                    kelipatan = 5000
33                End If
34            End If
              
35            strSql = "select [date] as TGLWO, NO_WO AS NOWO, NOINV ,fs_nm_dealer as DEALER,km as KM,fs_rc as BERKALA from [" & dbReplikasiHO & "]..T_INVOICEnew where NO_CHASSIS = '" & Trim(rsJobDta("NO_CHASSIS")) & "' and  "
36            strSql = strSql & " [date] < ("
37            strSql = strSql & " select top 1 [date]"
38            strSql = strSql & " from [" & dbReplikasiHO & "]..T_INVOICEnew where NO_WO = '" & noWO & "' and CABANG = '" & dBase & "'"
39            strSql = strSql & " ) ORDER BY [date] DESC "
              
40            If rsHistory.State Then rsHistory.Close
              
41            rsHistory.Open strSql, ("Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & usernameDealer & ";password=" & passwordDealer & ";Initial Catalog=" & dBase & ";server=" & dealer), adOpenStatic, adLockBatchOptimistic, adCmdText
42            TDBGrid1.DataSource = rsHistory
      '        If Not rsHistory.EOF Then
43            Text(2).Mask = ""
44            Text(2) = TglDDMMYYYY(rsJobDta("date"))
45            Text(20) = Format(CDbl(rsJobDta("km")), "#,#")
      '        End If
46            Label8.Visible = True
47            Text(1).Visible = True
48            Text(1).SetFocus
49            cmdGanti.Visible = True
              
50        Else
51            MsgBox "Tidak ada NO WO yang dicari", vbOKOnly, "Informasi"
52        End If
53        Exit Sub
errHandler:
54        MsgBox "Terjadi kesalahan dalam mengambil data!", vbCritical, "Informasi" 'err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub hapusForm()
    For i = 1 To 20
        If i = 11 Or i = 2 Then
        Text(i).Mask = ""
        ElseIf i = 3 Or i = 19 Then
        'text tersebut tidak ada index
        Else
        Text(i) = ""
        End If
    Next
    Label8.Visible = False
    Text(1).Visible = False
    cmdGanti.Visible = False
    
End Sub
