VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmOSublet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "ORDER SUBLET"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7380
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   360
      TabIndex        =   10
      Top             =   0
      Width           =   9855
      Begin VB.CheckBox chAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Add Job"
         Enabled         =   0   'False
         Height          =   255
         Left            =   3600
         TabIndex        =   36
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox chkPart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Part"
         Height          =   255
         Left            =   3240
         TabIndex        =   5
         Top             =   3000
         Width           =   855
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   11
         Top             =   525
         Width           =   3750
         _ExtentX        =   6615
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
         Index           =   7
         Left            =   1080
         TabIndex        =   0
         Top             =   1920
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   1
         Left            =   1065
         TabIndex        =   13
         Top             =   885
         Width           =   2310
         _ExtentX        =   4075
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
         Index           =   3
         Left            =   1065
         TabIndex        =   15
         Top             =   1245
         Width           =   3750
         _ExtentX        =   6615
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
         Index           =   2
         Left            =   3480
         TabIndex        =   17
         Top             =   885
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   5880
         TabIndex        =   18
         Top             =   525
         Width           =   3855
         _ExtentX        =   6800
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
         Index           =   5
         Left            =   5880
         TabIndex        =   19
         Top             =   885
         Width           =   3855
         _ExtentX        =   6800
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   1
         Top             =   2280
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   9
         Left            =   1080
         TabIndex        =   2
         Top             =   2640
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   11
         Left            =   1080
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
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
         Format          =   "#,##0.00##"
         Mask            =   ">9999999999"
         PromptChar      =   ">"
      End
      Begin TrueOleDBGrid70.TDBGrid GridSublet 
         Height          =   3375
         Left            =   120
         TabIndex        =   24
         Top             =   3360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5953
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
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).DataField=   ""
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).DataField=   ""
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).DataField=   ""
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   13
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=13"
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
         Splits(0)._ColumnProps(21)=   "Column(5).Width=3281"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=3281"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
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
         DeadAreaBackColor=   12577002
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=46,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=82,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=79,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=80,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=81,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=28,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=32,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=29,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=30,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=31,.parent=17"
         _StyleDefs(88)  =   "Named:id=33:Normal"
         _StyleDefs(89)  =   ":id=33,.parent=0"
         _StyleDefs(90)  =   "Named:id=34:Heading"
         _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(92)  =   ":id=34,.wraptext=-1"
         _StyleDefs(93)  =   "Named:id=35:Footing"
         _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(95)  =   "Named:id=36:Selected"
         _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(97)  =   "Named:id=37:Caption"
         _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(99)  =   "Named:id=38:HighlightRow"
         _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(101) =   "Named:id=39:EvenRow"
         _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(103) =   "Named:id=40:OddRow"
         _StyleDefs(104) =   ":id=40,.parent=33"
         _StyleDefs(105) =   "Named:id=41:RecordSelector"
         _StyleDefs(106) =   ":id=41,.parent=34"
         _StyleDefs(107) =   "Named:id=42:FilterBar"
         _StyleDefs(108) =   ":id=42,.parent=33"
         _StyleDefs(109) =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(110) =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(111) =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(112) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   4320
         TabIndex        =   28
         Top             =   1920
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
         Index           =   10
         Left            =   3240
         TabIndex        =   3
         Top             =   2640
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
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
         Left            =   5880
         TabIndex        =   6
         Top             =   1920
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   6
         Left            =   5880
         TabIndex        =   31
         Top             =   1245
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   13
         Left            =   5880
         TabIndex        =   7
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   8400
         TabIndex        =   37
         Top             =   1245
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   5880
         TabIndex        =   41
         Top             =   2640
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   17
         Left            =   5880
         TabIndex        =   8
         Top             =   3000
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12648384
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
         Mask            =   "99:99"
         PromptChar      =   ">"
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estimasi Pengerjaan"
         Height          =   195
         Left            =   4320
         TabIndex        =   44
         Top             =   3045
         Width           =   1425
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   7080
         X2              =   9720
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   7080
         X2              =   9720
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   120
         X2              =   2760
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   120
         X2              =   2760
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Order"
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
         Left            =   4440
         TabIndex        =   40
         Top             =   0
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   7080
         X2              =   9720
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   120
         X2              =   2760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penambahan Order Sublet"
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
         Left            =   3360
         TabIndex        =   39
         Top             =   1530
         Width           =   3330
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
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
         TabIndex        =   38
         Top             =   1290
         Width           =   585
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Nota"
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
         Left            =   5235
         TabIndex        =   34
         Top             =   2325
         Width           =   585
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe "
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
         TabIndex        =   33
         Top             =   3045
         Width           =   345
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   5310
         TabIndex        =   32
         Top             =   1290
         Width           =   510
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   4980
         TabIndex        =   30
         Top             =   1965
         Width           =   840
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
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
         Left            =   2640
         TabIndex        =   29
         Top             =   2685
         Width           =   510
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Harga"
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
         TabIndex        =   23
         Top             =   3045
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Left            =   570
         TabIndex        =   22
         Top             =   2685
         Width           =   270
      End
      Begin VB.Label Label4 
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
         Left            =   150
         TabIndex        =   21
         Top             =   2325
         Width           =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No PO"
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
         TabIndex        =   20
         Top             =   1965
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Left            =   5130
         TabIndex        =   16
         Top             =   525
         Width           =   690
      End
      Begin VB.Label Label1 
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
         Left            =   510
         TabIndex        =   14
         Top             =   930
         Width           =   420
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rangka"
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
         Left            =   90
         TabIndex        =   12
         Top             =   570
         Width           =   840
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   255
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   1560
         Width           =   9855
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   9855
      End
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   5640
      TabIndex        =   9
      Top             =   6960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Tambah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
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
      Left            =   6600
      TabIndex        =   25
      Top             =   6960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Ubah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
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
      Left            =   7560
      TabIndex        =   26
      Top             =   6960
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
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
      Left            =   8520
      TabIndex        =   27
      Top             =   6960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   7215
      Left            =   11880
      TabIndex        =   35
      Top             =   120
      Width           =   5415
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   270
      Index           =   16
      Left            =   6000
      TabIndex        =   42
      Top             =   0
      WhatsThisHelpID =   3000
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      BackColor       =   12648384
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
      Mask            =   "99:99"
      PromptChar      =   ">"
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Adjust"
      Height          =   195
      Left            =   6000
      TabIndex        =   43
      Top             =   45
      WhatsThisHelpID =   3000
      Width           =   435
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   7335
      Left            =   0
      Picture         =   "frmOSublet.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmOSublet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 04 Maret 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsTemp As ADODB.Recordset
Dim rsGrid As New ADODB.Recordset
Dim newRecord As Boolean
Dim boleh As Boolean
Dim pesan As String
Dim nNotaLama As String
Dim idLihat As Integer
Dim nHPP As Double
Dim cols As TrueOleDBGrid70.Columns

Private Sub hapusForm(idx As Integer)
    Select Case idx
        Case 1
            For i = 0 To 6
                Text(i) = ""
            Next
            Text(14) = ""
        Case 2
            For i = 8 To 13
                Select Case i
                    Case 9, 11
                        Text(i) = 0
                    Case Else
                        Text(i) = ""
                End Select
                chkPart = Unchecked
            Next
    End Select
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
    Text(7).Enabled = False
    For i = 8 To 17
        Text(i).Enabled = Flg
    Next
    chkPart.Enabled = Flg
End Sub

Private Sub refrGrid(nWO As String, noChassis As String, nPol As String, idx As Integer)
    Call HeadColumn
    Select Case idx
        Case 1
            ' TAMPIL DATA ORDER
            Set rsGrid = objO.FetchData("SELECT DISTINCT NOPO, no_chassis, namabarang, kodesatuan, harga, price, qty, nonota, no_Wo, keterangan, flag, KODEITEM,estwaktu FROM tempJobSublet WITH (NOLOCK) WHERE (no_Chassis = '" & noChassis & "' OR NO_POLISI = '" & nPol & "') AND no_WO = '" & nWO & "' AND KODEITEM <> 'A' ORDER BY noPO, NONOTA, namabarang")
        Case 2
            ' TAMPIL DATA BILLING
            ''Set rsGrid = objO.FetchData("SELECT DISTINCT NOPO, no_chassis, namabarang, kodesatuan, harga, price, qty, nonota, no_Wo, keterangan, flag, KODEITEM,estwaktu FROM tempJobSublet WITH (NOLOCK) WHERE no_Chassis = '" & noChassis & "' AND no_WO = '" & nWO & "' ORDER BY noPO, NONOTA, namabarang")
            ''2016-07-22 sublet diambil dari job sublet tidak dari tempjobsublet
            Set rsGrid = objO.FetchData("SELECT DISTINCT NOPO, no_chassis, namabarang, kodesatuan, harga, price, qty, nonota, no_Wo, keterangan, flag, KODEITEM,estwaktu FROM JobSublet WITH (NOLOCK) WHERE no_Chassis = '" & noChassis & "' AND no_WO = '" & nWO & "' ORDER BY noPO, NONOTA, namabarang")
        Case 3
            ' TAMPIL DATA ADD JOB
            Set rsGrid = objO.FetchData("SELECT DISTINCT NOPO, no_chassis, namabarang, kodesatuan, harga, price, qty, nonota, no_Wo, keterangan, flag, KODEITEM,estwaktu FROM tempJobSublet WITH (NOLOCK) WHERE no_Chassis = '" & noChassis & "' AND no_WO = '" & nWO & "' AND KODEITEM = 'A' ORDER BY noPO, NONOTA, namabarang")
    End Select
    
    If Not rsGrid.EOF Then
        Set GridSublet.DataSource = rsGrid
    Else
        Set GridSublet.DataSource = Nothing
    End If
End Sub

Private Sub HeadColumn()
    'no_Wo, no_chassis, namabarang, kodesatuan, qty, harga, nonota, keterangan
    With GridSublet
        .Columns(0).Caption = "No Po"
        .Columns(0).DataField = "noPO"
        .Columns(0).Width = 0
        .Columns(0).Visible = False
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Caption = "No Rangka"
        .Columns(1).DataField = "no_chassis"
        .Columns(1).Width = 0
        .Columns(1).Visible = False
        
        .Columns(2).Caption = "Pekerjaan"
        .Columns(2).DataField = "namabarang"
        .Columns(2).Width = 2500
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Caption = "Satuan"
        .Columns(3).DataField = "kodesatuan"
        .Columns(3).Width = 900
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Caption = "Beli"
        .Columns(4).DataField = "harga"
        .Columns(4).Width = 1200
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Caption = "Jual"
        .Columns(5).DataField = "price"
        .Columns(5).Width = 1200
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Caption = "Qty"
        .Columns(6).DataField = "qty"
        .Columns(6).Width = 600
        .Splits(0).Columns(6).Locked = True
        
        .Columns(7).Caption = "No Nota"
        .Columns(7).DataField = "nonota"
        .Columns(7).Width = 1200
        .Splits(0).Columns(7).Locked = True
        
        .Columns(8).Caption = "No Wo"
        .Columns(8).DataField = "no_WO"
        .Columns(8).Width = 1800
        .Columns(8).Visible = False
        .Splits(0).Columns(8).Locked = True
        
        .Columns(9).Caption = "Keterangan"
        .Columns(9).DataField = "keterangan"
        .Columns(9).Width = 2500
        .Splits(0).Columns(9).Locked = True
        
        .Columns(10).Caption = "Flag"
        .Columns(10).DataField = "flag"
        .Columns(10).Width = 0
        .Columns(10).Visible = False
        
        .Columns(11).Caption = "Add Job"
        .Columns(11).DataField = "KODEITEM"
        .Columns(11).Width = 0
        .Columns(11).Visible = False
        
        .Columns(12).Caption = "Estimasi Waktu"
        .Columns(12).DataField = "estWaktu"
        .Columns(12).Width = 1000
        .Splits(0).Columns(9).Locked = True
    End With
End Sub

Private Sub chkPart_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub cmdHapus_Click()
1        On Error GoTo cmdHapus_Click_Error

2         If MsgBox("Akan menghapus Pekerjaan : " & Text(8) & " No. Nota : " & Text(7), vbYesNo, "HAPUS SUBLET") = vbYes Then
3             Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_CHASSIS = '" & Text(0) & "' AND NOPO = '" & Text(7) & "' AND NAMABARANG = '" & Text(8) & "' AND NONOTA = '" & Text(13) & "'")
4             If gTag = "order" Then
5                 idLihat = 1
6                 frmOrder.Text(63).Text = Format(DateAdd("n", ((Left(Text(17), 2) * 60) + Right(Text(17), 2)) * -1, frmOrder.Text(63).Text), "HH:MM")
7             ElseIf gTag = "billing" Then
8                 idLihat = 2
9             ElseIf gTag = "addorder" Then
10                idLihat = 3
11            End If
12            Call refrGrid(Text(6), Text(0), "", idLihat)
13        End If

14       On Error GoTo 0
15       Exit Sub

cmdHapus_Click_Error:

16        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmOSublet")
End Sub

Private Sub CmdKeluar_Click()
Dim rsSearch As ADODB.Recordset
Dim nHarga As Double
Dim total As Double
    If gTag = "billing" Then
        On Error GoTo errBilling
        
        ''2016-05-17 || diganti begin trans
'         Call objHT.OpenTransactionBengkel
'
'         Call objHT.InsertData_BengkelOnly("DELETE JOBSUBLET WHERE NO_WO = '" & Trim(Text(6)) & "'")
'         Call objHT.InsertData_BengkelOnly("INSERT INTO JOBSUBLET SELECT * FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Text(6) & "' AND NO_CHASSIS = '" & Text(0) & "') AND NO_WO = '" & Text(6) & "' AND NO_CHASSIS = '" & Text(0) & "'")
'         Call objHT.InsertData_BengkelOnly("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & Trim(Text(6)) & "'")
'         Call objHT.CloseTransactionBengkel
         
'        Call objO.DeleteData("DELETE JOBSUBLET WHERE NO_WO = '" & Trim(Text(6)) & "'")
'        Call objO.InsertData("INSERT INTO JOBSUBLET SELECT * FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Text(6) & "' AND NO_CHASSIS = '" & Text(0) & "') AND NO_WO = '" & Text(6) & "' AND NO_CHASSIS = '" & Text(0) & "'")
'        Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & Trim(Text(6)) & "'")
        
        frmBilling.updateBilling
        
    ElseIf gTag = "order" Then
        nHarga = 0
        Set rsSearch = objO.FetchData("SELECT * FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(6)) & "'")
        If Not rsSearch.EOF Then
            Do While Not rsSearch.EOF
                nHarga = nHarga + (CDbl(rsSearch("PRICE")) * CDbl(rsSearch("QTY")))
                total = total + CDbl(hitungMenit(rsSearch("ESTWAKTU")))
                rsSearch.MoveNext
            Loop
            frmOrder.Text(60) = CDbl(nHarga)
            frmOrder.getTotalSublet = total
        End If
        rsSearch.Close
    ElseIf gTag = "addorder" Then
        nHarga = 0
        Set rsSearch = objO.FetchData("SELECT * FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(6)) & "'")
        If Not rsSearch.EOF Then
            Do While Not rsSearch.EOF
                nHarga = nHarga + (CDbl(rsSearch("PRICE")) * CDbl(rsSearch("QTY")))
                rsSearch.MoveNext
            Loop
            frmOAddJob.Text(24) = CDbl(nHarga)
        End If
        rsSearch.Close
    End If
    If gTag = "billing" Then
    Me.Hide
    Else
    Unload Me
    End If
    
    Exit Sub
errBilling:
   Call objHT.CloseTransactionBengkel_err
   MsgBox "Data gagal tersimpann", vbInformation
End Sub

Private Sub cmdSimpan_Click()
Dim nSPol As String
Dim jJ As Integer
Dim jM As Integer
Dim total As Integer
   On Error GoTo cmdSimpan_Click_Error

    If cmdSimpan.Caption = "&Tambah" Then
        newRecord = True
        Call hapusForm(2)
        Call buka_Tutup(True)
        Text(8).SetFocus
        nHPP = 0
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
    ElseIf cmdSimpan.Caption = "&Simpan" Then
        boleh = True
        pesan = "Lengkapi data " & vbCr
        
        If Text(11).Enabled Then
            If Text(8) = "" Then
                pesan = pesan + "Pekerjaan. "
                boleh = False
            End If
            If Text(9) = 0 Then
                pesan = pesan + "Qty minimal 1. "
                boleh = False
            End If
            If Text(13).Visible Then
                If Text(13) = "" Then
                    pesan = pesan + "No. Nota. "
                    boleh = False
                End If
            End If
            If Text(17) = ">>:>>" Then
                pesan = pesan + "Estimasi Pengerjaan "
                boleh = False
            End If
            If CDbl(Text(11)) < CDbl(nHPP) And CDbl(Text(11)) <> 0 Then
                pesan = pesan + "Harga jual tidak boleh lebih kecil dari HPP. "
                boleh = False
            End If
            
            If Trim(Text(17)) <> "" Then
                If Trim(Text(17)) = ">>:>>" Then
                    Text(17) = "00:00"
                End If
                
                jJ = CInt(Left(Trim(Text(17)), InStr(Trim(Text(17)), ":") - 1))
                jM = Mid(Trim(Text(17)), InStr(Trim(Text(17)), ":") + 1, Len(Trim(Text(17))) - InStr(Trim(Text(17)), ":"))
                
                total = (jJ * 60) + jM
                If total > 240 Then
                    pesan = pesan + "Estimasi pengerjaan sublet maksimal 4 Jam. "
                    boleh = False
                End If
            End If
            
            If boleh Then
                
                ClearsSql
                gsSql1(0) = "NO_WO"
                gsSql2(0) = Text(6)
                gsSql1(1) = "NO_CHASSIS"
                gsSql2(1) = Text(0)
                gsSql1(2) = "NAMABARANG"
                gsSql2(2) = Trim(Text(8))
                gsSql1(3) = "KODESATUAN"
                gsSql2(3) = Text(10)
                gsSql1(4) = "PRICE"
                gsSql2(4) = CDbl(Text(11))
                gsSql1(5) = "QTY"
                gsSql2(5) = CDbl(Text(9))
                gsSql1(6) = "NOPO"
                gsSql2(6) = Text(7)
                gsSql1(7) = "KETERANGAN"
                gsSql2(7) = Text(12)
                gsSql1(8) = "FLAG"
                gsSql2(8) = IIf(chkPart = Checked, "P", "J")
                gsSql1(9) = "NONOTA"
                gsSql2(9) = Text(13)
                gsSql1(10) = "KODEITEM"
                gsSql2(10) = IIf(chAdd.value = Checked, "A", "")
                gsSql1(11) = "NO_POLISI"
                gsSql2(11) = Trim(Text(14))
                gsSql1(12) = "TAGIH"
                gsSql2(12) = ""
                gsSql1(13) = "NOFAK"
                gsSql2(13) = Trim(Text(15))
                
                gsSql1(14) = "ESTWAKTU"
                gsSql2(14) = Trim(Text(17))
                
                If gTag = "billing" Then
                  ''2016-07-22 '' billing tidak menggunakan tempjobsublet
                  Set rsTemp = objO.FetchData("SELECT * FROM JOBsublet WITH (NOLOCK) WHERE ((no_chassis = '" & Text(0) & "' AND NO_WO = '" & Text(6) & "') OR NO_POLISI = '" & Text(14) & "') AND NOPO = '" & Text(7) & "' AND namabarang = '" & Text(8) & "'")
                Else
                  Set rsTemp = objO.FetchData("SELECT * FROM tempJOBsublet WITH (NOLOCK) WHERE ((no_chassis = '" & Text(0) & "' AND NO_WO = '" & Text(6) & "') OR NO_POLISI = '" & Text(14) & "') AND NOPO = '" & Text(7) & "' AND namabarang = '" & Text(8) & "'")
                End If
                
                If Not rsTemp.EOF Then
                    nSPol = rsTemp("NO_POLISI")
                    If newRecord Then
                        MsgBox "Data sudah ada. Jika ingin merubah tekan tombol Ubah", vbExclamation, "DUPLIKASI"
                        Exit Sub
                    Else
                        If gTag = "billing" Then
                           sSQL = GensSql2("JOBSUBLET", True) + "WHERE ((no_chassis = '" & Text(0) & "' AND NO_WO = '" & Text(6) & "') or NO_POLISI = '" & Trim(Text(14)) & "') AND NOPO = '" & Text(7) & "' AND namabarang = '" & Text(8) & "'"
                        Else
                           sSQL = GensSql2("tempJOBSUBLET", True) + "WHERE ((no_chassis = '" & Text(0) & "' AND NO_WO = '" & Text(6) & "') or NO_POLISI = '" & Trim(Text(14)) & "') AND NOPO = '" & Text(7) & "' AND namabarang = '" & Text(8) & "'"
                        End If
                        
                    End If
                Else
                    If nNotaLama = "" Then
                        If gTag = "billing" Then
                           sSQL = GensSql2("JOBSUBLET", True) + "WHERE ((no_chassis = '" & Text(0) & "' AND NO_WO = '" & Text(6) & "') OR NO_POLISI = '" & Trim(Text(14)) & "') AND NOPO = '" & Text(7) & "' AND namabarang = '" & Text(8) & "'"
                        Else
                           sSQL = GensSql2("tempJOBSUBLET", True) + "WHERE ((no_chassis = '" & Text(0) & "' AND NO_WO = '" & Text(6) & "') OR NO_POLISI = '" & Trim(Text(14)) & "') AND NOPO = '" & Text(7) & "' AND namabarang = '" & Text(8) & "'"
                        End If
                    Else
                        If gTag = "billing" Then
                           sSQL = GensSql2("JOBSUBLET", False)
                        Else
                           sSQL = GensSql2("tempJOBSUBLET", False)
                        End If
                        
                    End If
                End If
                Call objO.UpdateData(sSQL)
                If rsTemp.State <> adStateClosed Then rsTemp.Close
                
                Call hapusForm(2)
                Call buka_Tutup(False)
                If gTag = "order" Then
                    idLihat = 1
                    frmOrder.Text(63).Text = Format(DateAdd("n", ((Left(Text(17), 2) * 60) + Right(Text(17), 2)), frmOrder.Text(63).Text), "HH:MM")
                ElseIf gTag = "billing" Then
                    idLihat = 2
                ElseIf gTag = "addorder" Then
                    idLihat = 3
                End If
                Call refrGrid(Text(6), Text(0), Text(14), idLihat)
                If gTag = "billing" Then
                    cmdSimpan.Caption = "&Simpan"
                Else
                    cmdSimpan.Caption = "&Tambah"
                End If
                cmdUbah.Caption = "&Ubah"
                If gTag = "billing" Then
                    cmdUbah.SetFocus
                End If
            Else
                MsgBox pesan, vbCritical, "KELENGKAPAN"
                Exit Sub
            End If
        Else
            MsgBox "Tekan tombol ubah untuk update Harga"
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdSimpan_Click_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmOSublet")
End Sub

Private Sub cmdUbah_Click()
1        On Error GoTo cmdUbah_Click_Error

2         If cmdUbah.Caption = "&Ubah" Then
3             newRecord = False
4             Call buka_Tutup(True)
5             cmdSimpan.Caption = "&Simpan"
6             cmdUbah.Caption = "&Batal"
7             nNotaLama = Trim(Text(7))
8             If gTag = "billing" Then
9                 Text(8).Enabled = False
10                Text(9).Enabled = False
11                Text(10).Enabled = False
12                chkPart.Enabled = False
13            ElseIf gTag = "order" Then
14                frmOrder.Text(63).Text = Format(DateAdd("n", ((Left(Text(17), 2) * 60) + Right(Text(17), 2)) * -1, frmOrder.Text(63).Text), "HH:MM")
15            End If
16        ElseIf cmdUbah.Caption = "&Batal" Then
17            Call buka_Tutup(False)
18            If gTag = "billing" Then
19                cmdSimpan.Caption = "&Simpan"
20            Else
21                cmdSimpan.Caption = "&Tambah"
22            End If
23            cmdUbah.Caption = "&Ubah"
24        End If

25       On Error GoTo 0
26       Exit Sub

cmdUbah_Click_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdUbah_Click of Form frmOSublet")
End Sub

Private Sub Form_Load()
    Call hapusForm(2)
    Call buka_Tutup(False)
    cmdHapus.Visible = True
    If gTag = "order" Then
        idLihat = 1
        Call refrGrid(Trim(frmOrder.Text(64)), Trim(frmOrder.Text(15)), Trim(frmOrder.Text(14)), idLihat)
        chAdd.value = Unchecked
    ElseIf gTag = "billing" Then
        idLihat = 2
        Call refrGrid(Trim(frmBilling.Text(0)), Trim(frmBilling.Text(17)), Trim(frmBilling.Text(15)), idLihat)
        chAdd.value = Unchecked
        Label6.Visible = True
        Text(11).Visible = True
        cmdSimpan.Caption = "&Simpan"
        cmdHapus.Visible = False
    ElseIf gTag = "addorder" Then
        idLihat = 3
        Call refrGrid(frmOAddJob.Text(6), frmOAddJob.Text(0), "", idLihat)
        chAdd.value = Checked
    End If
    nNotaLama = "XXXX"
    
    If Log.isEmStall = 0 Then
        Label16.Visible = False
        Text(17) = "00:00"
        Text(17).Visible = False
    End If
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub gridSublet_FilterChange()
    On Error GoTo errHandler
    Set cols = GridSublet.Columns
    Dim C As Integer
    
    C = GridSublet.col
    GridSublet.HoldFields
    rsGrid.filter = getFilter()
    GridSublet.col = C
    GridSublet.EditActive = True
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

Private Sub GridSublet_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

    If rsGrid.State <> adStateClosed Then
        If Not rsGrid.EOF Then
            Text(7) = rsGrid.Fields(0)
            Text(8) = rsGrid.Fields(2)
            Text(9) = rsGrid.Fields(6)
            Text(10) = rsGrid.Fields(3)
            Text(11) = rsGrid.Fields(5)
            Text(12) = rsGrid.Fields(9)
            Text(13) = rsGrid.Fields(7)
            Text(17) = rsGrid.Fields(12)
            chAdd = IIf(rsGrid.Fields(11) = "A", Checked, Unchecked)
            chkPart = IIf(rsGrid.Fields(10) = "P", Checked, Unchecked)
            nHPP = CDbl(rsGrid.Fields(4))
        End If
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    Dim n As String
    
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            n = Text(Index).Mask
            Text(Index).Mask = ""
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            Text(Index).Mask = n
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
            Case 8
                gMainKeyCari = Text(Index)
                gSearchOption = "SUBLET"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 9, 11
            If Text(Index) = "" Then
                Text(Index) = 0
            End If
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 9
            If Trim(Text(Index)) <> "" Then
                If CDbl(Text(Index)) < 0 Then
                    Text(Index) = Abs(CDbl(Text(Index)))
                End If
            Else
                MsgBox "Qty harus diisi !!!"
            End If
    End Select
End Sub
