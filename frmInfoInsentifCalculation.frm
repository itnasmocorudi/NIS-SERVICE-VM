VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmInfoInsentifCalculation 
   Caption         =   ".: Perhitungan Insentif Calculation:."
   ClientHeight    =   11535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   ScaleHeight     =   11535
   ScaleWidth      =   16050
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmInfoInsentifCalculation.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   9
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   9735
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   12855
      Begin ARProgBarCtrl.ARProgressBar ProgBar 
         Height          =   615
         Left            =   240
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   12045
         _ExtentX        =   21246
         _ExtentY        =   1085
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
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4680
         MaxLength       =   60
         TabIndex        =   24
         Top             =   2040
         Width           =   2295
      End
      Begin VB.ComboBox cbAdjustment 
         Height          =   315
         ItemData        =   "frmInfoInsentifCalculation.frx":9675
         Left            =   1920
         List            =   "frmInfoInsentifCalculation.frx":9682
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2040
         Width           =   2775
      End
      Begin ucXPButton.XPButton cmdViewKPI 
         Height          =   495
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoInsentifCalculation.frx":96CF
         PICN            =   "frmInfoInsentifCalculation.frx":96EB
         PICH            =   "frmInfoInsentifCalculation.frx":D74C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdGenerateKPIMP 
         Height          =   495
         Left            =   2520
         TabIndex        =   18
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoInsentifCalculation.frx":11A3B
         PICN            =   "frmInfoInsentifCalculation.frx":11A57
         PICH            =   "frmInfoInsentifCalculation.frx":15B72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKPIMP 
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoInsentifCalculation.frx":1A1B8
         PICN            =   "frmInfoInsentifCalculation.frx":1A1D4
         PICH            =   "frmInfoInsentifCalculation.frx":1E479
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid GridSupplier 
         Height          =   5655
         Left            =   120
         TabIndex        =   1
         Top             =   2760
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9975
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "NIK"
         Columns(0).DataField=   "NIK"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama"
         Columns(1).DataField=   "Nama"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Tipe"
         Columns(2).DataField=   "Tipe"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "InsKPI"
         Columns(3).DataField=   "InsKPI"
         Columns(3).NumberFormat=   "#,#0"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "InsPart "
         Columns(4).DataField=   "InsPart"
         Columns(4).NumberFormat=   "#,#0"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "InsSublet"
         Columns(5).DataField=   "InsSublet"
         Columns(5).NumberFormat=   "#,#0"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "TotalInsentif"
         Columns(6).DataField=   "TotalInsentif"
         Columns(6).NumberFormat=   "#,#0"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "LockData"
         Columns(7).DataField=   "LockData"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "ReasonAdjKPI"
         Columns(8).DataField=   "ReasonAdjKPI"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
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
         Splits(0)._ColumnProps(16)=   "Column(3)._ColStyle=2"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(18)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(21)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(23)=   "Column(5).Width=3281"
         Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=3175"
         Splits(0)._ColumnProps(26)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(27)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(28)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(29)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(30)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(31)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(32)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(33)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(34)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(37)=   "Column(8).Width=3281"
         Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0"
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
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&H4080FF&"
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
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import Excel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1935
         MaxLength       =   50
         TabIndex        =   3
         Top             =   2400
         Width           =   5040
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3360
         MaxLength       =   60
         TabIndex        =   10
         Top             =   1680
         Width           =   3615
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1680
         Width           =   1455
      End
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   495
         Left            =   4920
         TabIndex        =   6
         Top             =   8520
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoInsentifCalculation.frx":22A3F
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
         Left            =   7320
         TabIndex        =   7
         Top             =   8520
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoInsentifCalculation.frx":22A5B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapus 
         Height          =   495
         Left            =   6120
         TabIndex        =   8
         Top             =   8520
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoInsentifCalculation.frx":22A77
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBDate6Ctl.TDBDate Periode 
         Height          =   375
         Left            =   1680
         TabIndex        =   12
         Top             =   300
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calendar        =   "frmInfoInsentifCalculation.frx":22A93
         Caption         =   "frmInfoInsentifCalculation.frx":22BBF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoInsentifCalculation.frx":22C2B
         Keys            =   "frmInfoInsentifCalculation.frx":22C49
         Spin            =   "frmInfoInsentifCalculation.frx":22CA7
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm-yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/yyyy"
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
         Text            =   "11/2024"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   45611
         CenturyMode     =   0
      End
      Begin ucXPButton.XPButton cmdUpdate 
         Height          =   495
         Left            =   7080
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   2
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
         MICON           =   "frmInfoInsentifCalculation.frx":22CCF
         PICN            =   "frmInfoInsentifCalculation.frx":22CEB
         PICH            =   "frmInfoInsentifCalculation.frx":25B72
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   10920
         TabIndex        =   20
         Top             =   8520
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
         MICON           =   "frmInfoInsentifCalculation.frx":28A90
         PICN            =   "frmInfoInsentifCalculation.frx":28AAC
         PICH            =   "frmInfoInsentifCalculation.frx":2F3DF
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
         Left            =   9720
         TabIndex        =   21
         Top             =   8520
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
         MICON           =   "frmInfoInsentifCalculation.frx":36096
         PICN            =   "frmInfoInsentifCalculation.frx":360B2
         PICH            =   "frmInfoInsentifCalculation.frx":3C6AB
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
         Left            =   8520
         TabIndex        =   22
         Top             =   8520
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
         MICON           =   "frmInfoInsentifCalculation.frx":42F5A
         PICN            =   "frmInfoInsentifCalculation.frx":42F76
         PICH            =   "frmInfoInsentifCalculation.frx":49779
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjustment"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   13
         Top             =   2085
         Width           =   1005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   12600
         X2              =   120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "NIK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1590
         TabIndex        =   11
         Top             =   1725
         Width           =   285
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   4
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1245
         TabIndex        =   5
         Top             =   2400
         Width           =   630
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmInfoInsentifCalculation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmInfoInsentifCalculation
' DateTime  : 19/09/2018
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------


Dim rsGrid      As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsTemp      As New ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset
Dim cols        As TrueOleDBGrid70.Columns
Dim Urut        As String
Dim isTambah    As Boolean
Dim isBegin     As Boolean
Dim isLoad      As Boolean
Dim iProgress   As Integer

Private Sub ber_sih()
    For i = 0 To 2
        Text(i) = ""
    Next
    isLoad = False
    cmdUpdate.Enabled = False
    
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
'    For i = 1 To 4
'        Text(i).Locked = Not Flg
'        If Flg Then
'            Text(i).BackColor = vbWhite
'        Else
'            Text(i).BackColor = &HC0E0FF
'        End If
'    Next
    cmdSimpan.Enabled = Flg
    cmdBatal.Enabled = Flg
    
   ' cmdTambah.Enabled = Not Flg
    cmdEdit.Enabled = Not Flg
    cmdHapus.Enabled = Not Flg
End Sub

Private Sub chAllDealer_Click()
     
    If chAllDealer = Unchecked Then
        frmeDealer.Visible = True
        If rsDealer.State = adStateClosed Then
            Call isiDealer
        End If
        frmeDealer.Visible = True
        GridDealer.SetFocus
    ElseIf chAllDealer = Checked Then
        frmeDealer.Visible = False
    End If
End Sub

Private Sub chModel_Click()
    If chModel.value = Checked Then
        frmeModel.Visible = False
    ElseIf chModel.value = Unchecked Then
        frmeModel.Visible = True
        Call isiModel
    End If
End Sub

Private Sub cbAdjustment_Change()
    If NulSID(cbAdjustment.Text) = NulSID("Pemusatan Training Contest") Then
        Text(3) = objO.FetchData_Double("SELECT MAX(InsKPI) from INSENTIF..grInsentifMP WHERE Tipe='" & NulSID(rsGrid!tipe) & "'")
    ElseIf NulSID(cbAdjustment.Text) = NulSID("Development / Acting") Then
        Text(3) = objO.FetchData_Double("SELECT AVG(InsKPI) from INSENTIF..grInsentifMP WHERE Tipe='" & NulSID(rsGrid!tipe) & "'")
    ElseIf NulSID(cbAdjustment.Text) = NulSID("Surat Peringatan (SP)") Then
        Text(3) = 0
    End If
End Sub

Private Sub cmdBatal_Click()
    Call buka_Tutup(False)
    Call refreshGrid
End Sub

Private Sub cmdEdit_Click()
1         isTambah = False
2         Call buka_Tutup(True)
3         If NulSID(GridSupplier.Columns(8).value) = "NR 00" Then
4             For i = 1 To 8
5                 Text(i).Locked = True
6                 Text(i).BackColor = &HC0E0FF
7             Next
8             Text(10).Locked = True
9             Text(10).BackColor = &HC0E0FF
10        ElseIf NulSID(GridSupplier.Columns(8).value) <> NulSID(gKdCbg) Then
11            For i = 1 To 8
12                Text(i).Locked = True
13                Text(i).BackColor = &HC0E0FF
14            Next
15            Text(10).Locked = True
16            Text(10).BackColor = &HC0E0FF
17        End If
End Sub

Private Sub cmdGenerateKPIMP_Click()
          Dim InsPart, InsSblt As Double
1         If Not rsGrid.EOF Then
2             isLoad = False
              Dim nilaiPersen As Double
3             rsGrid.MoveFirst
              
4             ProgBar.Visible = True
5             ProgBar.Min = 1
6             ProgBar.Max = rsTemp.RecordCount + 1
7             iProgress = 1
8             Do While Not rsGrid.EOF
9                 InsPart = 0
10                InsSblt = 0
11                sSQL = "SELECT porsimp"
12                sSQL = sSQL & " FROM  INSENTIF.dbo.GRKPIFinalScore a"
13                sSQL = sSQL & " WHERE a.NIK = '" & NulSID(rsGrid!nik) & "' and a.periode = '" & Format(Periode, "mm-yyyy") & "'"
14                rsGrid!InsKPI = SpyRound(objO.FetchData_Double(sSQL), 0)
                  
15                If InStr(1, NulSID(rsGrid!tipe), NulSID("SERVICE ADVISOR")) > 0 Then
16                    sSQL = "Select distinct round((b.TotalInsPart * d.BobotPart)/100,0) as [InsentifSparepart] from Manpower a"
17                    sSQL = sSQL & " inner join (select b1.FS_KD_Dealer, b1.Periode, b1.NIK, b1.KategoriItem,"
18                    sSQL = sSQL & " SUM(alokasiins) as TotalInsPart from [INSENTIF].dbo.GRInsProdukTersedia b1 group"
19                    sSQL = sSQL & " by b1.FS_KD_Dealer, b1.Periode, b1.NIK, b1.KategoriItem) b on a.NIK=b.NIK"
20                    sSQL = sSQL & " inner join [INSENTIF].dbo.GRTipeMP c on a.Tipe=c.Tipe"
21                    sSQL = sSQL & " inner join [INSENTIF].dbo.GRMstBobotPart d on c.KdTipe=d.KdTipe"
22                    sSQL = sSQL & " where b.KategoriItem='Part' and b.Periode='" & Format(Periode, "mm-yyyy") & "' and a.Tipe like '%SERVICE ADVISOr%' and b.FS_KD_Dealer='" & NulSID(gKdCbg) & "'"
23                    sSQL = sSQL & " and b.NIK = '" & NulSID(rsGrid!nik) & "'"
24                    rsGrid!InsPart = SpyRound(objO.FetchData_Double(sSQL), 0)
                  
25                ElseIf InStr(1, NulSID(rsGrid!tipe), NulSID("TMS TECHNICIAN ADVISOR")) > 0 Then
26                    sSQL = "Select distinct round((b.TotalInsPart * d.BobotPart)/100,0) as [InsentifSparepart] from Manpower a"
27                    sSQL = sSQL & " inner join (select b1.FS_KD_Dealer, b1.Periode, b1.NIK, b1.KategoriItem,"
28                    sSQL = sSQL & " SUM(alokasiins) as TotalInsPart from [INSENTIF].dbo.GRInsProdukTersedia b1 group"
29                    sSQL = sSQL & " by b1.FS_KD_Dealer, b1.Periode, b1.NIK, b1.KategoriItem) b on a.NIK=b.NIK"
30                    sSQL = sSQL & " inner join [INSENTIF].dbo.GRTipeMP c on a.Tipe=c.Tipe"
31                    sSQL = sSQL & " inner join [INSENTIF].dbo.GRMstBobotPart d on c.KdTipe=d.KdTipe"
32                    sSQL = sSQL & " where b.KategoriItem='Part' and b.Periode='" & Format(Periode, "mm-yyyy") & "' and a.Tipe like '%TMS TECHNICIAN ADVISOR%' and b.FS_KD_Dealer='" & NulSID(gKdCbg) & "'"
33                    sSQL = sSQL & " and b.NIK = '" & NulSID(rsGrid!nik) & "'"
34                    rsGrid!InsPart = SpyRound(objO.FetchData_Double(sSQL), 0)
                  
35                Else
36                    sSQL = "Select isnull(Sum(InsProduk),0) / (Select Count(*) as JmlMP from Manpower where tipe='" & NulSID(rsGrid!tipe) & "') from [INSENTIF].dbo.GRInsProdPosisi where tipe='" & NulSID(rsGrid!tipe) & "' and"
37                    sSQL = sSQL & " KategoriItem='0' and periode = '" & Format(Periode, "mm-yyyy") & "' and fs_kd_dealer = '" & NulSID(gKdCbg) & "'"
                      
38                    sSQL = "Select sum(a.InsProduk)/max(z.JmlMp) InsentifPartSA"
39                    sSQL = sSQL & " from insentif.dbo.GRInsProdPosisi a"
40                    sSQL = sSQL & " inner join ManPower b on a.Tipe=b.Tipe"
41                    sSQL = sSQL & " left join(select tipe, COUNT(nik) as JmlMP from ManPower group by Tipe)"
42                    sSQL = sSQL & " z on z.tipe = a.Tipe"
43                    sSQL = sSQL & " where a.Periode='" & Format(Periode, "mm-yyyy") & "' and a.FS_KD_Dealer='" & NulSID(gKdCbg) & "' and a.KategoriItem=0 and a.trn=0 and b.nik = '" & NulSID(rsGrid!nik) & "'"
44                    InsPart = InsPart + SpyRound(objO.FetchData_Double(sSQL), 0)
                      
45                    sSQL = "Select sum(a.InsProduk)/max(z.JmlMp) InsentifPartSA"
46                    sSQL = sSQL & " from insentif.dbo.GRInsProdPosisi a"
47                    sSQL = sSQL & " inner join ManPower b on a.Tipe=b.Tipe"
48                    sSQL = sSQL & " left join(select tipe, COUNT(nik) as JmlMP from ManPower group by Tipe)"
49                    sSQL = sSQL & " z on z.tipe = a.Tipe"
50                    sSQL = sSQL & " where a.Periode='" & Format(Periode, "mm-yyyy") & "' and a.FS_KD_Dealer='" & NulSID(gKdCbg) & "' and a.KategoriItem=0 and a.trn=1 and b.nik = '" & NulSID(rsGrid!nik) & "'"
51                    InsPart = InsPart + SpyRound(objO.FetchData_Double(sSQL), 0)
52                    rsGrid!InsPart = InsPart
53                End If
                  
                  
54                If InStr(1, NulSID(rsGrid!tipe), NulSID("SERVICE ADVISOR")) > 0 Then
55                    sSQL = "Select round(isnull(sum(b.TotalInsSblt * d.BobotSblt),0)/100,0) as [InsentifSublet] from Manpower a"
56                    sSQL = sSQL & " inner join (select b1.FS_KD_Dealer, b1.Periode, b1.NIK, b1.KategoriItem,"
57                    sSQL = sSQL & " SUM(alokasiins) as TotalInsSblt from [INSENTIF].dbo.GRInsProdukTersedia b1 group"
58                    sSQL = sSQL & " by b1.FS_KD_Dealer, b1.Periode, b1.NIK, b1.KategoriItem) b on a.NIK=b.NIK"
59                    sSQL = sSQL & " inner join [INSENTIF].dbo.GRTipeMP c on a.Tipe=c.Tipe"
60                    sSQL = sSQL & " inner join [INSENTIF].dbo.GRMstBobotSblt d on c.KdTipe=d.KdTipe"
61                    sSQL = sSQL & " where b.KategoriItem='Sublet' and b.Periode='" & Format(Periode, "mm-yyyy") & "' and a.Tipe like '%Service Advisor%' and b.FS_KD_Dealer='" & NulSID(gKdCbg) & "'"
62                    sSQL = sSQL & " and A.NIK = '" & NulSID(rsGrid!nik) & "'"
63                    rsGrid!InsSublet = SpyRound(objO.FetchData_Double(sSQL), 0)
64                ElseIf InStr(1, NulSID(rsGrid!tipe), NulSID("TMS TECHNICIAN ADVISOR")) > 0 Then
65                    sSQL = "Select round(isnull(sum(b.TotalInsSblt * d.BobotSblt),0)/100,0) as [InsentifSublet] from Manpower a"
66                    sSQL = sSQL & " inner join (select b1.FS_KD_Dealer, b1.Periode, b1.NIK, b1.KategoriItem,"
67                    sSQL = sSQL & " SUM(alokasiins) as TotalInsSblt from [INSENTIF].dbo.GRInsProdukTersedia b1 group"
68                    sSQL = sSQL & " by b1.FS_KD_Dealer, b1.Periode, b1.NIK, b1.KategoriItem) b on a.NIK=b.NIK"
69                    sSQL = sSQL & " inner join [INSENTIF].dbo.GRTipeMP c on a.Tipe=c.Tipe"
70                    sSQL = sSQL & " inner join [INSENTIF].dbo.GRMstBobotSblt d on c.KdTipe=d.KdTipe"
71                    sSQL = sSQL & " where b.KategoriItem='Sublet' and b.Periode='" & Format(Periode, "mm-yyyy") & "' and a.Tipe like '%TMS TECHNICIAN ADVISOR%' and b.FS_KD_Dealer='" & NulSID(gKdCbg) & "'"
72                    sSQL = sSQL & " and A.NIK = '" & NulSID(rsGrid!nik) & "'"
73                    rsGrid!InsSublet = SpyRound(objO.FetchData_Double(sSQL), 0)
                  
74                Else
75                    sSQL = "Select round(b.TotalInsSblt / c.JmlManpower,0) as [Insentif Sublet] from  Manpower a"
76                    sSQL = sSQL & " inner join (Select b1.FS_KD_Dealer, b1.Periode, b1.KategoriItem, b1.Tipe,"
77                    sSQL = sSQL & " Sum(InsProduk) as [TotalInsSblt] from [INSENTIF].dbo.GRInsProdPosisi b1 group by"
78                    sSQL = sSQL & " b1.FS_KD_Dealer, b1.Periode, b1.KategoriItem, b1.tipe) b on a.Tipe=b.Tipe"
79                    sSQL = sSQL & " left join (select tipe, COUNT(nik) as JmlManpower from  Manpower"
80                    sSQL = sSQL & " group by Tipe) c on c.Tipe=a.Tipe"
81                    sSQL = sSQL & " where b.KategoriItem='1' and a.Tipe not like '%Service Advisor%' and"
82                    sSQL = sSQL & " b.Periode='" & Format(Periode, "mm-yyyy") & "' and b.FS_KD_Dealer='" & NulSID(gKdCbg) & "'"
83                    sSQL = sSQL & " and A.NIK = '" & NulSID(rsGrid!nik) & "'"
                      
                      
84                    sSQL = "Select sum(a.InsProduk)/max(z.JmlMp) InsentifPartSA"
85                    sSQL = sSQL & " from insentif.dbo.GRInsProdPosisi a"
86                    sSQL = sSQL & " inner join ManPower b on a.Tipe=b.Tipe"
87                    sSQL = sSQL & " left join(select tipe, COUNT(nik) as JmlMP from ManPower group by Tipe)"
88                    sSQL = sSQL & " z on z.tipe = a.Tipe"
89                    sSQL = sSQL & " where a.Periode='" & Format(Periode, "mm-yyyy") & "' and a.FS_KD_Dealer='" & NulSID(gKdCbg) & "' and a.KategoriItem=1 and a.trn=0 and b.nik = '" & NulSID(rsGrid!nik) & "'"
90                    InsSblt = InsSblt + SpyRound(objO.FetchData_Double(sSQL), 0)
                      
91                    sSQL = "Select sum(a.InsProduk)/max(z.JmlMp) InsentifPartSA"
92                    sSQL = sSQL & " from insentif.dbo.GRInsProdPosisi a"
93                    sSQL = sSQL & " inner join ManPower b on a.Tipe=b.Tipe"
94                    sSQL = sSQL & " left join(select tipe, COUNT(nik) as JmlMP from ManPower group by Tipe)"
95                    sSQL = sSQL & " z on z.tipe = a.Tipe"
96                    sSQL = sSQL & " where a.Periode='" & Format(Periode, "mm-yyyy") & "' and a.FS_KD_Dealer='" & NulSID(gKdCbg) & "' and a.KategoriItem=1 and a.trn=1 and b.nik = '" & NulSID(rsGrid!nik) & "'"
97                    InsSblt = InsSblt + SpyRound(objO.FetchData_Double(sSQL), 0)
98                    rsGrid!InsSublet = InsSblt
99                End If
                  
100               rsGrid!TotalInsentif = (rsGrid!InsKPI + rsGrid!InsSublet + rsGrid!InsPart)
101               ProgBar.value = iProgress
102               DoEvents
103               iProgress = iProgress + 1
104               rsGrid.MoveNext '
105           Loop
              
106           ProgBar.Visible = False
107           isLoad = True
108       End If
End Sub

Private Sub cmdHapus_Click()
1        On Error GoTo cmdHapus_Click_Error
2         isBegin = False
3         If NulSID(GridSupplier.Columns(8).value) = "NR 00" Then
4             MsgBox "Master supplier HO tidak dapat di hapus!", vbInformation
5             Exit Sub
6         ElseIf NulSID(GridSupplier.Columns(8).value) <> NulSID(gKdCbg) Then
7             MsgBox "Master supplier Dealer lain tidak dapat di hapus!", vbInformation
8             Exit Sub
9         End If
          
10        Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE ispakai = '1' and KDSUPP = '" & NulSID(Text(0)) & "'")
11        If Not rsTemp.EOF Then
12            MsgBox "Sudah Ada Transaksi Work Order untuk sublet ini" & vbCrLf & "Kode Supplier " & NulSID(Text(0)) & " tidak dapat dihapus", vbInformation
13            Exit Sub
14        End If
15        If MsgBox("Apakah yakin menghapus data kode supplier " & NulSID(Text(0)) & " ?", vbYesNo, "Konfirmasi Hapus Data") = vbYes Then
16            Set rsTemp = objO.FetchDataHO("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..mstsublet WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND KDSUPP = '" & NulSID(Text(0)) & "'")
17            If Not rsTemp.EOF Then
18                MsgBox "Masih ada sublet untuk supplier ini." & vbCrLf & "Kode Sublet " & NulSID(rsTemp!KDSUBLET) & " masih aktif tidak dapat menghapus supplier", vbInformation
19                Exit Sub
20            End If
21            Call objHT.OpenTransactionBengkel
22            isBegin = True
23            erTranAll = True
24            sSQL = "UPDATE [" & dbReplikasiHO & "]..MSTSUPPLIER SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = GETDATE() ,STATUS = 0 WHERE KDSUPP = '" & NulSID(Text(0)) & "'"
25            Call objHT.InsertData_Bengkel(sSQL)
26            sSQL = "UPDATE T_SUPPLIER SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = GETDATE() ,STATUS = 0 WHERE KDSUPP = '" & NulSID(Text(0)) & "'"
27            Call objHT.InsertData_Bengkel(sSQL)
              
28            If erTranAll Then
29                Call objHT.CloseTransactionBengkel
30                MsgBox "Data berhasil dihapus", vbInformation
31                Call refreshGrid
32                Call buka_Tutup(False)
33            Else
34                Call objHT.CloseTransactionBengkel_err
35                MsgBox "Gagal hapus data", vbInformation
36                Exit Sub
37            End If
38        End If

39       On Error GoTo 0
40       Exit Sub

cmdHapus_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
41        varErlError = Erl
42        varErrNumber = err.Number
43        If isBegin Then
44            Call objHT.CloseTransactionBengkel_err
45        End If
46        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmInfoInsentifCalculation")
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
         Dim target, aktual, pencapaian As Double
         
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
11         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
12         GoSub Proses
13     Else
14        Exit Sub
15     End If
16     Exit Sub
       
Proses:
17    ErrBegin = True
         Dim VarNmFile As String
         Dim kodePosisi, Posisi As String
         Dim kodeKPI, nik, Nama, KPI As String
         Dim nilai As Double
         
         Dim rsModel  As ADODB.Recordset
18       Call objHT.OpenTransactionBengkel
19       isBegin = True
20       gBantuan = "SBE"
21       If RcExcel.State Then RcExcel.Close
22       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
23       If RcExcel.RecordCount > 0 Then
24          ProgBar.Visible = True
25          ProgBar.Max = RcExcel.RecordCount
26          RcExcel.MoveFirst
27          Do While Not RcExcel.EOF
28             DoEvents
29             konter = konter + 1
30             ProgBar.value = konter
               
31             nilai = 0
32             kodeKPI = ""
33             nik = ""
34             Nama = ""
35             KPI = ""
36             kodePosisi = ""
37             Posisi = ""
38             pencapaian = 0
              
39             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
               
40             nik = NulSID(RcExcel.Fields(0).value)
41             Nama = NulSID(RcExcel.Fields(1).value)
42             kodePosisi = NulSID(RcExcel.Fields(2).value)
43             Posisi = NulSID(RcExcel.Fields(3).value)
44             kodeKPI = NulSID(RcExcel.Fields(4).value)
45             KPI = NulSID(RcExcel.Fields(5).value)
46             nilai = (NulSID(RcExcel.Fields(6).value))
               
47             pencapaian = k2t(objHT.FetchDataBengkel_String("select KPIProcess from INSENTIF..GRMstKonvKPI where " & k2t(nilai) & " BETWEEN BtsBawah and BtsAtas "))
              
48             sSQL = "DELETE from INSENTIF..GRKPIProsesMP where Periode = c and FS_KD_DEALER = '" & NulSID(gKdCbg) & "' and NIK = '" & nik & "' and KdKPI = '" & kodeKPI & "' and Tipe = '" & Posisi & "'"
49             Call objHT.InsertData_Bengkel(sSQL)
50             sSQL = "insert into insentif..GRKPIProsesMP (Periode,FS_KD_DEALER,NIK,Nama,Tipe,KdKPI,NamaKPI,Nilai,AchvKPIProses,UserEntry,TglEntry) values ('" & Format(Periode.value, "MM-YYYY") & "','" & NulSID(gKdCbg) & "','" & nik & "','" & Nama & "','" & Posisi & "','" & kodeKPI & "','" & KPI & "'," & k2t(nilai) & ",'" & pencapaian & "','" & gSA & "',getdate())"
               
51             Call objHT.InsertData_Bengkel(sSQL)
                
                       
52             RcExcel.MoveNext
53          Loop
54       End If
55       RcExcel.Close
56       CnExel.Close
      '   ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
57        Call objHT.CloseTransactionBengkel
58       MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
59       ProgBar.Visible = False
60       Call cmdViewKPI_Click
61       Exit Sub
vb_error:
62       If ErrBegin = True Then
           ' DbCon.RollbackTrans
63       End If
64       If isBegin Then
65        Call objHT.CloseTransactionBengkel_err
66       End If
67       ProgBar.Visible = False
68       MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdKPIMP_Click()
          
1         Set rsTemp = objO.FetchData("Select a.Tipe, nik, nama from  Manpower a ")
2         If Not rsTemp.EOF Then
3             rsTemp.MoveFirst
4             ProgBar.Visible = True
5             ProgBar.Min = 1
6             ProgBar.Max = rsTemp.RecordCount + 1
7             iProgress = 1
8             Do While Not rsTemp.EOF
9                 With rsGrid
10                    .AddNew
                      
11                    !nik = NulSID(rsTemp!nik)
12                    !Nama = NulSID(rsTemp!Nama)
13                    !tipe = NulSID(rsTemp!tipe)
14                    .Update
15                End With
16                ProgBar.value = iProgress
17                iProgress = iProgress + 1
18                DoEvents
19                rsTemp.MoveNext
20            Loop
21            Set GridSupplier.DataSource = rsGrid
22            isLoad = True
23        End If
24        ProgBar.Visible = False
25        GridSupplier.Visible = True
26        cmdKPIMP.Enabled = False
27        cmdGenerateKPIMP.Enabled = True
          
End Sub

Private Sub cmdSimpan_Click()
   On Error GoTo cmdSimpan_Click_Error
    Dim boleh           As Boolean
    Dim pesan           As String
    Dim rsTemp          As ADODB.Recordset
    Dim namaSupplier    As String
    boleh = True
    isBegin = False
     
    If Not rsGrid.EOF Then
        rsGrid.MoveFirst
        Call objHT.OpenTransactionBengkel
        isBegin = True
        Do While Not rsGrid.EOF
            
            sSQL = "INSERT INTO [INSENTIF].[dbo].[GRInsentifMP]"
            sSQL = sSQL & " ([Periode],[FS_KD_Dealer],[NIK],[Nama],[Tipe],[InsKPI],[InsPart],[InsSublet],[TotalInsentif],[LockData],[UserEntry],[TglEntry],[ReasonAdjKPI],[ValADH],[ValKCB])"
            sSQL = sSQL & " VALUES"
            sSQL = sSQL & " ('" & Format(Periode.value, "MM-YYYY") & "','" & NulSID(gKdCbg) & "','" & NulSID(rsGrid!nik) & "','" & NulSID(rsGrid!Nama) & "','" & NulSID(rsGrid!tipe) & "','" & k2t(rsGrid!InsKPI) & "','" & k2t(rsGrid!InsPart) & "','" & k2t(rsGrid!InsSublet) & "','" & k2t(rsGrid!TotalInsentif) & "','0','" & gSA & "',getdate(),'',0,0)"

            Call objHT.InsertData_Bengkel(sSQL)
            rsGrid.MoveNext
        Loop
        Call objHT.CloseTransactionBengkel
        Call cmdViewKPI_Click
    Else
        MsgBox "Tidak ada data yang diproses.", vbInformation
        Exit Sub
    End If
   On Error GoTo 0
   Exit Sub
cmdSimpan_Click_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    If isBegin = True Then
        Call objHT.CloseTransactionBengkel_err
    End If
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmInfoInsentifCalculation")
End Sub

Private Sub cmdTambah_Click()
    isTambah = True
    Call ber_sih
    Call buka_Tutup(True)
    Call isiId
End Sub

Private Sub cmdTutup_Click()
    frmeModel.Visible = False
End Sub

Private Sub cmdTutupDealer_Click()

    frmeDealer.Visible = False
End Sub

Private Sub cmdUpdate_Click()
    Dim nilaiPersen, aktual, ach As Double
    If NulSID(Text(2).Text) = "" Then
        MsgBox "Silahkan isi alasan adjustment", vbInformation
        Exit Sub
    End If
    If NulSID(cbAdjustment.Text) = NulSID("Pemusatan Training Contest") Then
        Text(3) = SpyRound(objO.FetchData_Double("SELECT MAX(InsKPI) from INSENTIF..grInsentifMP WHERE Tipe='" & NulSID(rsGrid!tipe) & "'"))
    ElseIf NulSID(cbAdjustment.Text) = NulSID("Development / Acting") Then
        Text(3) = SpyRound(objO.FetchData_Double("SELECT AVG(InsKPI) from INSENTIF..grInsentifMP WHERE Tipe='" & NulSID(rsGrid!tipe) & "'"))
    ElseIf NulSID(cbAdjustment.Text) = NulSID("Surat Peringatan (SP)") Then
        Text(3) = 0
    End If
    rsGrid!InsKPI = Text(3)
    rsGrid!TotalInsentif = (rsGrid!InsKPI + rsGrid!InsSublet + rsGrid!InsPart)
    sSQL = "UPDATE INSENTIF..GRInsentifMP SET InsKPI = '" & k2t(rsGrid!InsKPI) & "',ReasonAdjKPI = '" & NulSID(Text(2).Text) & "',TotalInsentif = '" & k2t(rsGrid!TotalInsentif) & "'  where Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_Kd_Dealer = '" & NulSID(gKdCbg) & "' and NIK = '" & NulSID(Text(0).Text) & "'"
    
    Call objO.InsertData(sSQL)
    Call cmdViewKPI_Click
    
End Sub

Private Sub cmdViewKPI_Click()
    Call ber_sih
    Call buka_Tutup(True)
    Call refreshGrid
End Sub

Private Sub Form_Load()

    Periode.value = Format(Date, "mm-yyyy")
    Call ber_sih
    Call buka_Tutup(True)
    Call refreshGrid
End Sub


Private Sub GridSupplier_FilterChange()
    On Error GoTo errHandler
    Set cols = GridSupplier.Columns
    Dim C As Integer
    
    C = GridSupplier.col
    GridSupplier.HoldFields
    rsGrid.filter = getFilter()
    GridSupplier.col = C
    GridSupplier.EditActive = True
    Exit Sub
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub GridSupplier_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo GridSupplier_HeadClick_Error

    Urut = IIf(Urut = "DESC", "ASC", "DESC")
    rsGrid.Sort = GridSupplier.Columns(ColIndex).DataField & " " & Urut
    GridSupplier.SetFocus
    Sendkeys "{Right " & Val(ColIndex) & "}"
 
   Exit Sub

GridSupplier_HeadClick_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure GridSupplier_HeadClick of Form frmInfoInsentifCalculation")
 End Sub

Private Sub gridSupplier_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF And isLoad Then
        Text(0).Text = NulSID(rsGrid!nik)
        Text(1).Text = NulSID(rsGrid!Nama)
        'cbAdjustment.Text = NulSID(rsGrid!AdjsPoint)
        Text(2).Text = NulSID(rsGrid!ReasonAdjKPI)
        cmdUpdate.Enabled = True
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
    'KeyCode = UCase(KeyCode)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 9
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLGAP"
                frmCariAcess.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) ORDER BY RIGHT(KDSUPP,10) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KDSUPP"), 4))
        noUrut = noUrut + 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    End If
End Sub

Private Sub isiModel()
    Set rsTemp = objO.FetchData("SELECT DISTINCT a.MODEL, a.mdl, a.[desc], ISNULL(id_program, '') as PRO FROM [" & dbReplikasiHO & "]..model a WITH (NOLOCK) LEFT JOIN ho_PROmdl b WITH (NOLOCK) ON a.MDL = b.MDL AND B.ID_PROGRAM = '" & Trim(Text(0)) & "' ORDER BY a.MODEL, a.mdl")
    If Not rsTemp.EOF Then
        Call createTbl
        Do While Not rsTemp.EOF
            rsGridMdl.AddNew
            rsGridMdl!grpMDL = Trim(rsTemp("MODEL"))
            rsGridMdl!kodeMDL = Trim(rsTemp("MDL"))
            rsGridMdl!NAMAMDL = Trim(rsTemp("DESC"))
            rsGridMdl!ikut = IIf(rsTemp("PRO") = "", False, True)
            rsGridMdl.Update
            rsTemp.MoveNext
        Loop

        rsGridMdl.MoveFirst
        GridModel.Columns(0).Width = 1200
        GridModel.Columns(1).Width = 2000
        GridModel.Columns(2).Width = 4000
        GridModel.Columns(3).Alignment = dbgCenter
        GridModel.Columns(3).Width = 800
    End If
End Sub

Private Sub createTbl()
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    With rsGridMdl
        .Fields.Append "GRPMDL", adVarChar, 15
        .Fields.Append "KODEMDL", adVarChar, 20
        .Fields.Append "NAMAMDL", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridModel.DataSource = rsGridMdl
    GridModel.AlternatingRowStyle = True
End Sub

Private Sub GridModel_FilterChange()
    On Error GoTo errHandler
    Set cols = GridModel.Columns
    Dim C As Integer
    
    C = GridModel.col
    GridModel.HoldFields
    rsGridMdl.filter = getFilter()
    GridModel.col = C
    GridModel.EditActive = True
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


Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
    Set rsTmp = objO.FetchDataHO("SELECT fs_kd_dealer, fs_nm_dealer, ISNULL(ID_PROGRAM,'') AS KODE FROM t_dealer a WITH (NOLOCK) " _
                            & "LEFT JOIN HO_PROCBG b ON a.fs_kd_dealer = b.kodedealer and A.FS_LOG_DEL = '' AND b.ID_PROGRAM = '" & Text(0) & "' WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
    If Not rsTmp.EOF Then
        Call createRS
        Do While Not rsTmp.EOF
            rsDealer.AddNew
            rsDealer!KodeDealer = rsTmp("FS_KD_DEALER")
            rsDealer!namaDealer = rsTmp("FS_NM_DEALER")
            If rsTmp("KODE") <> "" Then
                rsDealer!terlibat = True
            Else
                rsDealer!terlibat = False
            End If
            rsDealer.Update
            rsTmp.MoveNext
        Loop
        
        GridDealer.DataSource = rsDealer
        GridDealer.EvenRowStyle.BackColor = vbYellow
        GridDealer.OddRowStyle.BackColor = vbWhite
        GridDealer.AlternatingRowStyle = True
    End If
End Sub

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 3500
    GridDealer.Columns(2).Width = 1000
    GridDealer.Columns(2).Alignment = dbgCenter
End Sub

 

Private Sub refreshGrid()
          ''GridSupplier.ClearFields
1         Call createRS
          ''sSQL = "SELECT KDSUPP,SUPPLIER,ALAMAT1,ALAMAT2,KOTA,TELPON,CP,KETERANGAN,KDDEALER,PKP,NONJURNAL FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY KDDEALER "
          
2         sSQL = "Select * from  INSENTIF.dbo.GRInsentifMP where Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_DEALER = '" & NulSID(gKdCbg) & "'"
          
3         Set rsTemp = objO.FetchData(sSQL)
4         If Not rsTemp.EOF Then
5             rsTemp.MoveFirst
6             iProgress = 1
7             ProgBar.Visible = True
8             DoEvents
9             ProgBar.Min = 1
10            ProgBar.Max = rsTemp.RecordCount + 1
11            Do While Not rsTemp.EOF
12                With rsGrid
13                    .AddNew
14                    !nik = NulSID(rsTemp!nik)
15                    !Nama = NulSID(rsTemp!Nama)
16                    !tipe = NulSID(rsTemp!tipe)
17                    !InsKPI = NulSID(rsTemp!InsKPI)
18                    !InsPart = NulSID(rsTemp!InsPart)
19                    !InsSublet = NulSID(rsTemp!InsSublet)
20                    !TotalInsentif = NulSID(rsTemp!TotalInsentif)
21                    !ReasonAdjKPI = NulSID(rsTemp!ReasonAdjKPI)
22                    .Update
23                End With
24                ProgBar.value = iProgress
25                iProgress = iProgress + 1
26                DoEvents
27                rsTemp.MoveNext
28            Loop
29            ProgBar.Visible = False
30            cmdSimpan.Enabled = False
31            cmdKPIMP.Enabled = False
32            cmdGenerateKPIMP.Enabled = False
33            isLoad = True
34        Else
35            cmdSimpan.Enabled = True
36            cmdKPIMP.Enabled = True
37            cmdGenerateKPIMP.Enabled = False
38        End If
39        Set GridSupplier.DataSource = rsGrid
           
End Sub

Private Sub createRS()
    If rsGrid.State Then rsGrid.Close
    Set rsGrid = Nothing
    With rsGrid
        .Fields.Append "NIK", adVarChar, 9
        .Fields.Append "Nama", adVarChar, 100
        .Fields.Append "Tipe", adVarChar, 100
        .Fields.Append "InsKPI", adDouble
        .Fields.Append "InsPart", adDouble
        .Fields.Append "InsSublet", adDouble
        .Fields.Append "TotalInsentif", adDouble
        .Fields.Append "LockData", adInteger
        .Fields.Append "ReasonAdjKPI", adVarChar, 100
        .Open
    End With
End Sub
