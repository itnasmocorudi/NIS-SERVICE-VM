VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmHOMOU 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MOU"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   10665
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   6615
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   10095
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   6840
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin TrueOleDBGrid70.TDBGrid GridUnit 
         Height          =   4815
         Left            =   120
         TabIndex        =   10
         Top             =   1680
         Width           =   9855
         _ExtentX        =   17383
         _ExtentY        =   8493
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
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
         _StyleDefs(73)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(74)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(75)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(76)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   6840
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   2865
         _ExtentX        =   5054
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
         Index           =   2
         Left            =   1320
         TabIndex        =   0
         Top             =   150
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Left            =   1320
         TabIndex        =   15
         Top             =   435
         Width           =   3945
         _ExtentX        =   6959
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
         Left            =   2955
         TabIndex        =   3
         Top             =   735
         Width           =   1125
         _ExtentX        =   1984
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   2040
         TabIndex        =   2
         Top             =   735
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Index           =   6
         Left            =   1320
         TabIndex        =   4
         Top             =   1035
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">??"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   5
         Top             =   1035
         Width           =   525
         _ExtentX        =   926
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
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   2400
         TabIndex        =   6
         Top             =   1035
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">???"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   1320
         TabIndex        =   1
         Top             =   735
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Index           =   17
         Left            =   1320
         TabIndex        =   20
         Top             =   1335
         Width           =   3345
         _ExtentX        =   5900
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
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4710
         Top             =   1335
         Width           =   60
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Pelanggan"
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
         TabIndex        =   21
         Top             =   1380
         Width           =   960
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   1860
         Top             =   735
         Width           =   60
      End
      Begin VB.Label Label2 
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
         Left            =   420
         TabIndex        =   18
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label23 
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
         Left            =   285
         TabIndex        =   17
         Top             =   795
         Width           =   780
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3720
         Top             =   150
         Width           =   60
      End
      Begin VB.Label Label22 
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
         Left            =   645
         TabIndex        =   16
         Top             =   195
         Width           =   420
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. MOU"
         Height          =   195
         Left            =   5910
         TabIndex        =   12
         Top             =   525
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   6915
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   12197
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "   Tabel Kendaraan Terdaftar MOU"
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
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   6120
      TabIndex        =   7
      Top             =   7200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Tambah"
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
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   7920
      TabIndex        =   13
      Top             =   7200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Hapus"
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
   Begin AIFCmp1.asxPowerButton cmdKeluar 
      Height          =   360
      Left            =   8880
      TabIndex        =   14
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
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
   Begin AIFCmp1.asxPowerButton cmdUbah 
      Height          =   360
      Left            =   7080
      TabIndex        =   19
      Top             =   7200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Ubah"
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
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   360
      Left            =   4560
      TabIndex        =   22
      Top             =   7200
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      BackColor       =   12648447
      Caption         =   "Import Excell"
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
      Height          =   360
      Left            =   720
      TabIndex        =   23
      Top             =   6960
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   635
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
   Begin AIFCmp1.asxPowerButton cmdCetak 
      Height          =   360
      Left            =   3240
      TabIndex        =   24
      Top             =   7200
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmHOMOU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 23 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsDownload  As New ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim rsGrid As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns

Private Sub hapusForm()
    For i = 2 To 9
        Text(i) = ""
    Next
End Sub

Private Sub bukaTutup(Optional Ket As Boolean = True, Optional ket2 As Boolean = True)
    Text(2).Enabled = Ket
    Text(9).Enabled = Ket
    Text(4).Enabled = Ket
    Text(5).Enabled = Ket
    Text(6).Enabled = ket2
    Text(7).Enabled = ket2
    Text(8).Enabled = ket2
End Sub

Private Sub refreshGrid(nMOU As String)
    Set rsGrid = objO.FetchDataHO("SELECT MDL, MODEL, NO_RANGKA, NO_POLISI FROM TEMPMOUDTL WITH (NOLOCK) WHERE NO_MOU = '" & nMOU & "' ORDER BY MDL, NO_RANGKA")
    If Not rsGrid.EOF Then
        GridUnit.DataSource = rsGrid
        Call headColumns
    End If
End Sub

Private Sub headColumns()
    With GridUnit
        .Columns(0).Width = 2500
        .Columns(0).Caption = "KODE MODEL"
        .Columns(0).DataField = "MDL"
        .Columns(1).Width = 4000
        .Columns(1).Caption = "MODEL"
        .Columns(1).DataField = "MODEL"
        .Columns(2).Width = 2500
        .Columns(2).Caption = "NO RANGKA"
        .Columns(2).DataField = "NO_RANGKA"
        .Columns(3).Width = 1500
        .Columns(3).Caption = "NO POLISI"
        .Columns(3).DataField = "NO_POLISI"
    End With
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
        .Sheets("Sheet1").Cells(1, 1).value = "MDL (max 50 kar - sesuai master model)"
        .Sheets("Sheet1").Cells(1, 2).value = "MODEL (OPSIONAL)"
        .Sheets("Sheet1").Cells(1, 3).value = "NO_RANGKA"
        .Sheets("Sheet1").Cells(1, 4).value = "NO_POLISI"

        .Sheets("Sheet1").Cells(2, 1).CopyFromRecordset rsDownload
        .Sheets("Sheet1").Range("1:1").Font.Bold = True
        .Sheets("Sheet1").Select
        .Close SaveChanges:=True, FileName:=onV
    End With
    Set oWB = Nothing
    oApp.Quit
    Set oApp = Nothing
End Sub

Private Sub cmdCetak_Click()
    
   On Error GoTo cmdCetak_Click_Error
    Dim Er As New ErrInfo
    If rsDownload.State Then rsDownload.Close
    Set rsDownload = Nothing
    With rsDownload
        .Fields.Append "MDL", adVarChar, 50
        .Fields.Append "MODEL", adVarChar, 100
        .Fields.Append "NO_RANGKA", adVarChar, 50
        .Fields.Append "NO_POLISI", adVarChar, 50
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

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdCetak_Click of Form frmHOProgramFOPSummary"

End Sub

Private Sub cmdHapus_Click()
    temp = MsgBox("Yakin akan menghapus Model " & Text(3) & ", Untuk No Rangka " & Trim(Text(9)) & Trim(Text(4)) & Trim(Text(5)) & " ?", vbYesNo, "HAPUS DATA")
    If temp = vbYes Then
        Call objO.DeleteDataHO("DELETE TEMPMOUDTL WHERE NO_MOU = '" & Text(0) & "' AND MDL = '" & Text(2) & "' AND NO_RANGKA = '" & Trim(Text(9)) & Trim(Text(4)) & Trim(Text(5)) & "'")
        Call refreshGrid(Text(0))
        Call hapusForm
    End If
End Sub

Private Sub cmdImport_Click()
              Dim ErrBegin As Boolean, konter As Long
1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim rsTempExcel As ADODB.Recordset
         Dim Jml As Integer
         Dim noMou As String
         
        ' Dim konter As Integer
         
2       konter = 0
3       ProgBar.Min = 0
4       CommonDialog1.FileName = ""
5       CommonDialog1.InitDir = App.Path
6       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
7       CommonDialog1.ShowOpen
8       If CommonDialog1.FileName <> "" Then
9          If CnExel.State Then CnExel.Close
10         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
11         GoSub Proses
12     Else
13        Exit Sub
14     End If
15     Exit Sub
       
Proses:
16    ErrBegin = True
         Dim VarNmFile As String
         Dim namaDealer As String
         Dim namaDealerCbg As String
         Dim namaModel As String
17       Call objHT.OpenTransactionHO
18        gBantuan = "SBE"
19       If RcExcel.State Then RcExcel.Close
20       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
21       If RcExcel.RecordCount > 0 Then
22          ProgBar.Visible = True
23          ProgBar.Max = RcExcel.RecordCount
24          RcExcel.MoveFirst
25          Do Until RcExcel.EOF
26             DoEvents
               
27             ProgBar.value = konter
28             model = ""
29             namaModel = ""
30             If Nul(Trim(RcExcel.Fields(0).value)) = "" Then Exit Do
31             noMou = Nul(Text(0))
32            If Nul(Trim(RcExcel.Fields(0).value)) = "" Then
33                Call objHT.CloseTransactionHO_err
34                ProgBar.Visible = False
35                MsgBox "No MOU baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
36                Exit Sub
37            Else
38                Set rsTempExcel = objO.FetchDataHO("select MDL,[DESC],NO_CHASSIS from model where FS_LOG_DEL = '' AND mdl = '" & Trim(RcExcel.Fields(0).value) & "'")
39                If rsTempExcel.RecordCount > 0 Then
40                    namaModel = Trim(rsTempExcel("desc"))
41                Else
42                    Call objHT.CloseTransactionHO_err
43                    ProgBar.Visible = False
44                    MsgBox "Kode model " & RcExcel.Fields(0).value & " baris ke-" & konter & " salah / tidak ada data !", vbCritical, "Error"
45                    Exit Sub
46                End If
47            End If
               
48            If Nul(Trim(RcExcel.Fields(2).value)) = "" And Nul(Trim(RcExcel.Fields(3).value)) = "" Then
49                Call objHT.CloseTransactionHO_err
50                ProgBar.Visible = False
51                MsgBox "No rangka / no polisi baris ke-" & konter & " salah / tidak ada data !", vbCritical, "Error"
52                Exit Sub
53            End If
              
              ''2015-12-01
54             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
55                ClearsSql
                  
56                gsSql1(0) = "NO_MOU"
57                gsSql2(0) = Nul(noMou)
58                gsSql1(1) = "MDL"
59                gsSql2(1) = Nul(Trim(RcExcel.Fields(0).value))
60                gsSql1(2) = "MODEL"
61                gsSql2(2) = namaModel
62                gsSql1(3) = "NO_RANGKA"
63                gsSql2(3) = Nul(Trim(RcExcel.Fields(2).value))
64                gsSql1(4) = "NO_POLISI"
65                gsSql2(4) = Nul(Trim(RcExcel.Fields(3).value))
                  
66                Set rsTemp = objO.FetchDataHO("SELECT * FROM TEMPMOUDTL WITH (NOLOCK) WHERE NO_MOU = '" & Nul(Trim(noMou)) & "' AND MDL = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' AND NO_RANGKA = '" & Nul(Trim(RcExcel.Fields(2).value)) & "' ORDER BY MDL, NO_RANGKA")
67                If Not rsTemp.EOF Then
68                    sSQL = GensSql2("TEMPMOUDTL", True) + "WHERE NO_MOU = '" & Nul(Trim(noMou)) & "' AND MDL = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' AND NO_RANGKA = '" & Nul(Trim(RcExcel.Fields(2).value)) & "'"
69                Else
70                    sSQL = GensSql2("TEMPMOUDTL", False)
71                End If
72                Call objHT.InsertData_Ho(sSQL)
73                If Nul(Trim(RcExcel.Fields(0).value)) <> "" Then
74                    Call objHT.InsertData_Ho("INSERT INTO TEMPKIRIM VALUES ('MOU', '" & Nul(Trim(RcExcel.Fields(0).value)) & "','S')")
75                End If
76             RcExcel.MoveNext
77             konter = konter + 1
78          Loop
79       End If
80       RcExcel.Close
81       CnExel.Close
82       ProgBar.Visible = False

83        Call objHT.CloseTransactionHO
84        Call refreshGrid(Text(0))
85       MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
      '   GridAFI.Refresh
86       Return
87       Exit Sub
vb_error:
88       If ErrBegin = True Then
           ' DbCon.RollbackTrans
89       End If
90       Call objHT.CloseTransactionHO_err
91       ProgBar.Visible = False
92       MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
      Dim pesan As String
      Dim boleh As Boolean
      Dim rsCari As ADODB.Recordset

1         If cmdSimpan.Caption = "&Tambah" Then
2             Call hapusForm
3             Call bukaTutup(True, True)
4             Text(2).SelLength = Len(Text(2))
5             Text(2).SetFocus
6             cmdSimpan.Caption = "&Simpan"
7             cmdUbah.Caption = "&Batal"
8         ElseIf cmdSimpan.Caption = "&Simpan" Then
              
9             boleh = True
10            pesan = "Lengkapi data "
11            If Text(2) = "" Then
12                boleh = False
13                pesan = pesan + "Model. "
14            End If
15            If Text(5) = "" Then
16                boleh = False
17                pesan = pesan + "No Rangka. "
18            End If
19            If Text(9) = "" Then
20                boleh = False
21                pesan = pesan + "No. Rangka. "
22            End If
23            If Text(4) = "" Then
24                boleh = False
25                pesan = pesan + "No. Rangka. "
26            End If
              
27            If Nul(Text(17)) = "" Then
28                boleh = False
29                pesan = pesan + "Id Pelanggan. "
30            Else
31                Set rsCari = objO.FetchDataHO("SELECT * FROM [SingleIDOri]..NewCustomer WITH (NOLOCK) WHERE idCust = '" & Nul(Text(17)) & "'")
32                If rsCari.EOF Then
33                    boleh = False
34                    pesan = pesan + "Id Pelanggan tidak terdaftar. "
35                End If
                  
36            End If
              
37            If boleh Then
38                ClearsSql
                  
39                gsSql1(0) = "NO_MOU"
40                gsSql2(0) = Text(0)
41                gsSql1(1) = "MDL"
42                gsSql2(1) = Text(2)
43                gsSql1(2) = "MODEL"
44                gsSql2(2) = Text(3)
45                gsSql1(3) = "NO_RANGKA"
46                gsSql2(3) = Trim(Text(9)) & Trim(Text(4)) & Trim(Text(5))
47                gsSql1(4) = "NO_POLISI"
48                If Text(6) = "" And Text(7) = "" And Text(8) = "" Then
49                    gsSql2(4) = ""
50                Else
51                    gsSql2(4) = Text(6) & "-" & Text(7) & "-" & Text(8)
52                End If
                  
53                gsSql1(5) = "IDCUST"
54                gsSql2(5) = Nul(Text(17))
                  
55                Set rsTemp = objO.FetchDataHO("SELECT * FROM TEMPMOUDTL WITH (NOLOCK) WHERE NO_MOU = '" & Text(0) & "' AND MDL = '" & Text(2) & "' AND NO_RANGKA = '" & gsSql2(3) & "' ORDER BY MDL, NO_RANGKA")
56                If Not rsTemp.EOF Then
57                    sSQL = GensSql2("TEMPMOUDTL", True) + "WHERE NO_MOU = '" & Text(0) & "' AND MDL = '" & Text(2) & "' AND NO_RANGKA = '" & gsSql2(3) & "'"
58                Else
59                    sSQL = GensSql2("TEMPMOUDTL", False)
60                End If
61                Call objO.UpdateDataHO(sSQL)
                  
62                Call hapusForm
63                Call bukaTutup(False, False)
64                Call refreshGrid(Text(0))
65                cmdSimpan.Caption = "&Tambah"
66                cmdUbah.Caption = "&Ubah"
67            Else
68                MsgBox pesan, vbCritical, "KELENGKAPAN"
69                Exit Sub
70            End If
71        End If
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call bukaTutup(False, True)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
        Text(6).SetFocus
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call bukaTutup(False, False)
        Call refreshGrid(Text(0))
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
    End If
End Sub

Private Sub Form_Load()
    Call hapusForm
    Call bukaTutup(False, False)
    Call refreshGrid(gBantuan)
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridUnit.Width = Frame1.Width - 200
    GridUnit.Height = Frame1.Height - Text(6).Top - Text(6).Height - 300
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdImport.Left = cmdSimpan.Left - cmdImport.Width - 100
    cmdCetak.Left = cmdImport.Left - cmdCetak.Width - 100
End Sub

Private Sub GridUnit_FilterChange()
    On Error GoTo errHandler
    Set cols = GridUnit.Columns
    Dim C As Integer
    
    C = GridUnit.col
    GridUnit.HoldFields
    rsGrid.filter = getFilter()
    GridUnit.col = C
    GridUnit.EditActive = True
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

Private Sub GridUnit_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsGrid.State <> adStateClosed Then
        If Not rsGrid.EOF Then
            Text(2) = rsGrid.Fields(0)      'GridUnit.Columns(0)
            Text(3) = rsGrid.Fields(1)      'GridUnit.Columns(1)
            If rsGrid.Fields(2) <> "" Then  'GridUnit.Columns(2)
                Text(9) = Left(rsGrid.Fields(2), 3)  'GridUnit.Columns(2)
                Text(4) = Mid(rsGrid.Fields(2), 4, 6)   'GridUnit.Columns(2)
                Text(5) = Mid(rsGrid.Fields(2), 10, Len(rsGrid.Fields(2)) - 9)
            End If
            If rsGrid.Fields(3) <> "" Then
                Dim nDpn1 As String
                Dim nTgh1 As String
                Dim nBlk1 As String
                nDpn1 = ""
                nTgh1 = ""
                nBlk1 = ""
                Call pecahNoPol(Trim(rsGrid.Fields(3)), nDpn1, nTgh1, nBlk1)
                Text(6) = nDpn1
                Text(7) = nTgh1
                Text(8) = nBlk1
            Else
                Text(6) = ""
                Text(7) = ""
                Text(8) = ""
            End If
        End If
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index).Mask = ""
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
            Case 2
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCariHO.show (vbModal)
            Case 9
                gMainKeyCari = Text(Index)
                gSearchOption = "COUNTRY"
                frmCariHO.show (vbModal)
            Case 17
                gMainKeyCari = Text(Index)
                gSearchOption = "SID"
                frmCariHO.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As New ADODB.Recordset
    
     
    Select Case Index
        Case 2
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchDataHO("SELECT * FROM MODEL WITH (NOLOCK) WHERE MDL = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCariHO.show (vbModal)
                Else
                    Text(3) = rsSearch("DESC")
'                    Text(4) = rsSearch("NO_CHASSIS")
                End If
                rsSearch.Close
            End If
        Case 9
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchDataHO("SELECT * FROM T_COUNTRY WITH (NOLOCK) WHERE KODE = '" & Trim(Text(Index)) & "'")
                If rsSearch.EOF Then
                   If MsgBox("Unit bukan Kendaraan Toyota ?", vbYesNo, "Konfirmasi No Rangka") = vbNo Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "COUNTRY"
                        frmCariHO.show (vbModal)
                   End If
                End If
                rsSearch.Close
            End If
    End Select
End Sub
