VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmTSubletStandarisasi 
   Caption         =   ".: MASTER SUBLET  :."
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   16320
   WindowState     =   2  'Maximized
   Begin VB.Frame frmeModel 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmTSubletStandarisasi.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   8280
      TabIndex        =   11
      Top             =   2040
      Width           =   9135
      Begin TrueOleDBGrid70.TDBGrid GridModel 
         Height          =   2895
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "GROUP"
         Columns(0).DataField=   "GRPMDL"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "MODEL"
         Columns(1).DataField=   "KODEMDL"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "KETERANGAN"
         Columns(2).DataField=   "NAMAMDL"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   4
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "OK"
         Columns(3).DataField=   "IKUT"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
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
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin AIFCmp1.asxPowerButton cmdTutup 
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   3240
         Width           =   8895
         _ExtentX        =   15690
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
   Begin VB.Frame frmeDealer 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"frmTSubletStandarisasi.frx":00A8
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   10680
      TabIndex        =   14
      Top             =   2040
      Width           =   6615
      Begin TrueOleDBGrid70.TDBGrid GridDealer 
         Height          =   3015
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5318
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE DEALER"
         Columns(0).DataField=   "KODEDEALER"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA DEALER"
         Columns(1).DataField=   "NAMADEALER"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   4
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   "TERLIBAT"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
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
         DeadAreaBackColor=   16761024
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H808080&"
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
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF00&"
         _StyleDefs(21)  =   ":id=9,.borderColor=&H80000001&"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Named:id=33:Normal"
         _StyleDefs(50)  =   ":id=33,.parent=0"
         _StyleDefs(51)  =   "Named:id=34:Heading"
         _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   ":id=34,.wraptext=-1"
         _StyleDefs(54)  =   "Named:id=35:Footing"
         _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=36:Selected"
         _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=37:Caption"
         _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin AIFCmp1.asxPowerButton cmdTutupDealer 
         Height          =   360
         Left            =   120
         TabIndex        =   16
         Top             =   3240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Tut&up"
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nama Sublet"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   360
      TabIndex        =   18
      Top             =   360
      Width           =   12975
      Begin TDBDate6Ctl.TDBDate periodeAwal 
         Height          =   255
         Left            =   7800
         TabIndex        =   6
         Top             =   1125
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         Calendar        =   "frmTSubletStandarisasi.frx":0144
         Caption         =   "frmTSubletStandarisasi.frx":025B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmTSubletStandarisasi.frx":02C0
         Keys            =   "frmTSubletStandarisasi.frx":02DE
         Spin            =   "frmTSubletStandarisasi.frx":033C
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
         Text            =   "18/10/2018"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   43391
         CenturyMode     =   0
      End
      Begin TDBNumber6Ctl.TDBNumber hargaBeli 
         Height          =   285
         Left            =   7800
         TabIndex        =   4
         Top             =   550
         Width           =   1895
         _Version        =   65536
         _ExtentX        =   3343
         _ExtentY        =   503
         Calculator      =   "frmTSubletStandarisasi.frx":0364
         Caption         =   "frmTSubletStandarisasi.frx":0384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmTSubletStandarisasi.frx":03E9
         Keys            =   "frmTSubletStandarisasi.frx":0407
         Spin            =   "frmTSubletStandarisasi.frx":0451
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "###,###.##;($ ###,###.##);Null;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########;-########"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   -65531
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
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
         Index           =   2
         Left            =   3600
         MaxLength       =   60
         TabIndex        =   35
         Top             =   550
         Width           =   2535
      End
      Begin VB.CheckBox chModel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Semua Model"
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
         Left            =   7800
         TabIndex        =   22
         Top             =   1395
         Value           =   1  'Checked
         Width           =   1455
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
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   17
         Top             =   225
         Width           =   1815
      End
      Begin VB.CheckBox chAllDealer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Semua Dealer"
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
         Left            =   10305
         TabIndex        =   21
         Top             =   1395
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1935
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
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   550
         Width           =   1815
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
         Index           =   3
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   840
         Width           =   4345
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
         Index           =   4
         Left            =   1680
         MaxLength       =   60
         TabIndex        =   2
         Top             =   1130
         Width           =   4455
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
         Height          =   645
         Index           =   5
         Left            =   1680
         MaxLength       =   30
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1395
         Width           =   4455
      End
      Begin VB.CheckBox chJurnal 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Non Jurnal"
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
         Left            =   8880
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CheckBox chPKP 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PKP"
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
         Left            =   7800
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   10920
         TabIndex        =   20
         Top             =   7440
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
         MICON           =   "frmTSubletStandarisasi.frx":0479
         PICN            =   "frmTSubletStandarisasi.frx":0495
         PICH            =   "frmTSubletStandarisasi.frx":6DC8
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
         Height          =   5295
         Left            =   120
         TabIndex        =   23
         Top             =   2040
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9340
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode Sublet"
         Columns(0).DataField=   "KDSUBLET"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama Sublet"
         Columns(1).DataField=   "NAMASUBLET"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Satuan"
         Columns(2).DataField=   "KODESATUAN"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Kode Supplier"
         Columns(3).DataField=   "KDSUPP"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Nama Supplier"
         Columns(4).DataField=   "SUPPLIER"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Keterangan"
         Columns(5).DataField=   "KETERANGAN"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Harga Beli"
         Columns(6).DataField=   "HARGABELI"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Harga Jual"
         Columns(7).DataField=   "HARGAJUAL"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Periode Awal"
         Columns(8).DataField=   "AWALPER"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Periode Akhir"
         Columns(9).DataField=   "AKHIRPER"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "ALLCABANG"
         Columns(10).DataField=   "ALLCABANG"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "ALLMODEL"
         Columns(11).DataField=   "ALLMODEL"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
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
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(10).Visible=0"
         Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(46)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(49)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H93C6FD&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
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
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=32,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=28,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=25,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=26,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=27,.parent=17"
         _StyleDefs(84)  =   "Named:id=33:Normal"
         _StyleDefs(85)  =   ":id=33,.parent=0"
         _StyleDefs(86)  =   "Named:id=34:Heading"
         _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(88)  =   ":id=34,.wraptext=-1"
         _StyleDefs(89)  =   "Named:id=35:Footing"
         _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(91)  =   "Named:id=36:Selected"
         _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=37:Caption"
         _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(95)  =   "Named:id=38:HighlightRow"
         _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(97)  =   "Named:id=39:EvenRow"
         _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&H4080FF&"
         _StyleDefs(99)  =   "Named:id=40:OddRow"
         _StyleDefs(100) =   ":id=40,.parent=33"
         _StyleDefs(101) =   "Named:id=41:RecordSelector"
         _StyleDefs(102) =   ":id=41,.parent=34"
         _StyleDefs(103) =   "Named:id=42:FilterBar"
         _StyleDefs(104) =   ":id=42,.parent=33"
         _StyleDefs(105) =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(106) =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(107) =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(108) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   495
         Left            =   4920
         TabIndex        =   24
         Top             =   7440
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
         MICON           =   "frmTSubletStandarisasi.frx":DA7F
         PICN            =   "frmTSubletStandarisasi.frx":DA9B
         PICH            =   "frmTSubletStandarisasi.frx":1423E
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
         Left            =   6120
         TabIndex        =   25
         Top             =   7440
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
         MICON           =   "frmTSubletStandarisasi.frx":1AD67
         PICN            =   "frmTSubletStandarisasi.frx":1AD83
         PICH            =   "frmTSubletStandarisasi.frx":212E5
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
         TabIndex        =   9
         Top             =   7440
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
         MICON           =   "frmTSubletStandarisasi.frx":27B08
         PICN            =   "frmTSubletStandarisasi.frx":27B24
         PICH            =   "frmTSubletStandarisasi.frx":2E11D
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
         TabIndex        =   8
         Top             =   7440
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
         MICON           =   "frmTSubletStandarisasi.frx":349CC
         PICN            =   "frmTSubletStandarisasi.frx":349E8
         PICH            =   "frmTSubletStandarisasi.frx":3B1EB
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
         Left            =   7350
         TabIndex        =   26
         Top             =   7440
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
         MICON           =   "frmTSubletStandarisasi.frx":41D91
         PICN            =   "frmTSubletStandarisasi.frx":41DAD
         PICH            =   "frmTSubletStandarisasi.frx":48920
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBNumber6Ctl.TDBNumber hargaJual 
         Height          =   285
         Left            =   7800
         TabIndex        =   5
         Top             =   840
         Width           =   1895
         _Version        =   65536
         _ExtentX        =   3343
         _ExtentY        =   503
         Calculator      =   "frmTSubletStandarisasi.frx":4F6BA
         Caption         =   "frmTSubletStandarisasi.frx":4F6DA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmTSubletStandarisasi.frx":4F73F
         Keys            =   "frmTSubletStandarisasi.frx":4F75D
         Spin            =   "frmTSubletStandarisasi.frx":4F7A7
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "###,###.##;($ ###,###.##);Null;0"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "########;-########"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99999999
         MinValue        =   0
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBDate6Ctl.TDBDate periodeAkhir 
         Height          =   255
         Left            =   9360
         TabIndex        =   7
         Top             =   1125
         Width           =   1215
         _Version        =   65536
         _ExtentX        =   2143
         _ExtentY        =   450
         Calendar        =   "frmTSubletStandarisasi.frx":4F7CF
         Caption         =   "frmTSubletStandarisasi.frx":4F8E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmTSubletStandarisasi.frx":4F94B
         Keys            =   "frmTSubletStandarisasi.frx":4F969
         Spin            =   "frmTSubletStandarisasi.frx":4F9C7
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
         Text            =   "18/10/2018"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   43391
         CenturyMode     =   0
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   6075
         Top             =   840
         Width           =   60
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Height          =   195
         Left            =   7095
         TabIndex        =   40
         Top             =   1185
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
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
         Left            =   9015
         TabIndex        =   39
         Top             =   1125
         Width           =   285
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " ~ Exclude PPn untuk Supplier PKP ~"
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
         Left            =   9690
         TabIndex        =   38
         Top             =   840
         Width           =   3045
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " ~ Exclude PPn untuk Supplier PKP ~"
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
         Left            =   9690
         TabIndex        =   37
         Top             =   555
         Width           =   4005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Jual"
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
         Left            =   6840
         TabIndex        =   36
         Top             =   930
         Width           =   900
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Satuan"
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
         Left            =   1020
         TabIndex        =   34
         Top             =   1200
         Width           =   600
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Kendaraan"
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
         Left            =   6270
         TabIndex        =   33
         Top             =   1425
         Width           =   1470
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   7920
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3500
         Top             =   550
         Width           =   60
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Harga Beli"
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
         Left            =   6885
         TabIndex        =   32
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Sublet"
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
         Left            =   555
         TabIndex        =   31
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Supplier"
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
         Left            =   465
         TabIndex        =   30
         Top             =   645
         Width           =   1155
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Sublet"
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
         TabIndex        =   29
         Top             =   255
         Width           =   1305
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Berlaku"
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
         Left            =   9480
         TabIndex        =   28
         Top             =   1425
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   600
         TabIndex        =   27
         Top             =   1485
         Width           =   990
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmTSubletStandarisasi.frx":4F9EF
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmTSubletStandarisasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmTSubletStandarisasi
' DateTime  : 05/10/2018 15:13
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------


Dim rsGrid      As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset
Dim cols        As TrueOleDBGrid70.Columns
Dim Urut        As String
Dim isTambah    As Boolean
Dim isBegin     As Boolean


Private Sub ber_sih()
    For i = 0 To 5
        Text(i) = ""
    Next
    chPKP.value = vbUnchecked
    chJurnal.value = vbUnchecked
    
    hargaJual.Text = 0
    hargaBeli.Text = 0
    
    periodeAkhir.value = Format(isnow, "DD-MM-YYYY")
    periodeAwal.value = Format(isnow, "DD-MM-YYYY")
    
    chAllDealer = Checked
    chModel = Checked
    chModel.Enabled = False
    chAllDealer.Enabled = False
    
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    frmeModel.Visible = False
    frmeDealer.Visible = False
    isTambah = True
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
    For i = 1 To 5
        Text(i).Locked = Not Flg
        If Flg And i <> 2 Then
            Text(i).BackColor = vbWhite
            Text(i).Locked = False
        Else
            Text(i).BackColor = &HC0E0FF
            Text(i).Locked = True
        End If
    Next
    
    If Flg Then
        hargaBeli.BackColor = vbWhite
        hargaJual.BackColor = vbWhite
        periodeAwal.BackColor = vbWhite
        periodeAkhir.BackColor = vbWhite
        
        hargaBeli.Enabled = True
        hargaJual.Enabled = True
        periodeAwal.Enabled = True
        periodeAkhir.Enabled = True
    Else
        hargaBeli.BackColor = &HC0E0FF
        hargaJual.BackColor = &HC0E0FF
        periodeAwal.BackColor = &HC0E0FF
        periodeAkhir.BackColor = &HC0E0FF
        
        hargaBeli.Enabled = False
        hargaJual.Enabled = False
        periodeAwal.Enabled = False
        periodeAkhir.Enabled = False
    End If
    
    cmdSimpan.Enabled = Flg
    cmdBatal.Enabled = Flg
    
    'chAllDealer.Enabled = Flg
    chAllDealer.Enabled = False
    chModel.Enabled = Flg
    
    cmdTambah.Enabled = Not Flg
    cmdEdit.Enabled = Not Flg
    cmdHapus.Enabled = Not Flg
End Sub

Private Sub chAllDealer_Click()
     
    If chAllDealer = Unchecked And cmdSimpan.Enabled Then
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
    If chModel.value = Checked And cmdSimpan.Enabled Then
        frmeModel.Visible = False
    ElseIf chModel.value = Unchecked And cmdSimpan.Enabled Then
        frmeModel.Visible = True
        Call isiModel
    End If
End Sub

Private Sub cmdBatal_Click()
    Call buka_Tutup(False)
    GridSupplier.Enabled = True
    Call refreshGrid
End Sub

Private Sub cmdEdit_Click()
1         If Left(NulSID(GridSupplier.Columns(1).value), 4) = "NR00" Then
2             MsgBox "Master Sublet HO tidak dapat diubah", vbInformation
3             Exit Sub
4         ElseIf Left(NulSID(GridSupplier.Columns(1).value), 4) <> Replace(gKdCbg, " ", "") Then
5             MsgBox "Master Sublet Cabang Lain tidak dapat diubah", vbInformation
6             Exit Sub
7         End If
8         Me.MousePointer = vbHourglass
9         isTambah = False
10        Call buka_Tutup(True)
11        If chModel.value = Unchecked And cmdSimpan.Enabled Then
12            Call isiModel
13        End If
14        If chAllDealer = Unchecked And cmdSimpan.Enabled Then
15            If rsDealer.State = adStateClosed Then
16                Call isiDealer
17            End If
18        End If
19        Me.MousePointer = vbDefault
End Sub

Private Sub cmdHapus_Click()
              
1        On Error GoTo cmdHapus_Click_Error
2         isBegin = False
3         If Left(NulSID(GridSupplier.Columns(1).value), 4) = "NR00" Then
4             MsgBox "Master Sublet HO tidak dapat diubah", vbInformation
5             Exit Sub
6         ElseIf Left(NulSID(GridSupplier.Columns(1).value), 4) <> Replace(gKdCbg, " ", "") Then
7             MsgBox "Master Sublet Cabang Lain tidak dapat diubah", vbInformation
8             Exit Sub
9         End If
10        If MsgBox("Apakah yakin menghapus data kode sublet " & NulSID(Text(0)) & " ?", vbYesNo, "Konfirmasi Hapus Data") = vbYes Then
11            Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUBLETPROMO WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND KDSUBLET = '" & NulSID(Text(0)) & "'")
12            If Not rsTemp.EOF Then
13                MsgBox "Masih ada promo untuk sublet ini" & vbCrLf & "Kode Promo " & NulSID(rsTemp!KDPROMO) & " masih aktif tidak dapat menghapus sublet", vbInformation
14                Exit Sub
15            End If
              
16            Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM JOBSUBLET WITH (NOLOCK) WHERE KODEITEM = '" & NulSID(Text(0)) & "'")
17            If Not rsTemp.EOF Then
18                MsgBox "Sudah Ada Transaksi untuk sublet ini" & vbCrLf & "Kode Sublet " & NulSID(Text(0)) & " tidak dapat menghapus sublet", vbInformation
19                Exit Sub
20            End If
              
21            Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..HIST_SUBLETnew WITH (NOLOCK) WHERE KODEITEM = '" & NulSID(Text(0)) & "'")
22            If Not rsTemp.EOF Then
23                MsgBox "Sudah Ada Transaksi Invoice untuk sublet ini" & vbCrLf & "Kode Sublet " & NulSID(Text(0)) & " tidak dapat menghapus sublet", vbInformation
24                Exit Sub
25            End If
              
26            Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUBLET WITH (NOLOCK) WHERE ispakai = '1' and KDSUBLET = '" & NulSID(Text(0)) & "'")
27            If Not rsTemp.EOF Then
28                MsgBox "Sudah Ada Transaksi Work Order untuk sublet ini" & vbCrLf & "Kode Sublet " & NulSID(Text(0)) & " tidak dapat dihapus", vbInformation
29                Exit Sub
30            End If
              
31            Call objHT.OpenTransactionBengkel
32            isBegin = True
33            erTranAll = True
34            sSQL = "UPDATE [" & dbReplikasiHO & "]..MSTSUBLET SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = GETDATE() WHERE kdsublet = '" & NulSID(Text(0)) & "'"
35            Call objHT.InsertData_Bengkel(sSQL)
               
36            If erTranAll Then
37                Call objHT.CloseTransactionBengkel
38                MsgBox "Data berhasil dihapus", vbInformation
39                Call refreshGrid
40                Call buka_Tutup(False)
41            Else
42                Call objHT.CloseTransactionBengkel_err
43                MsgBox "Gagal hapus data", vbInformation
44                Exit Sub
45            End If
46        End If

47       On Error GoTo 0
48       Exit Sub

cmdHapus_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
49        varErlError = Erl
50        varErrNumber = err.Number
51        If isBegin Then
52            Call objHT.CloseTransactionBengkel_err
53        End If
54        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmTSubletStandarisasi")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
          
1        On Error GoTo cmdSimpan_Click_Error
          Dim isBegin     As Boolean
          Dim boleh       As Boolean
          Dim pesan       As String
          Dim allModel    As Byte
          Dim allDealer   As Byte
          
2         boleh = True
3         isBegin = False
          
4         If NulSID(Text(1)) = "" Then
5             pesan = pesan & " - Kode Supplier harus diisi." & vbCrLf
6             boleh = False
7         Else
8             Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WHERE KDSUPP = '" & NulSID(Text(1).Text) & "'  AND FS_LOG_DEL = ''")
9             If rsTemp.EOF Then
10                pesan = pesan & " - Kode Supplier tidak terdaftar." & vbCrLf
11                boleh = False
12            End If
13        End If
          
14        If NulSID(Text(3)) = "" Then
15            pesan = pesan & " - Nama Sublet harus diisi." & vbCrLf
16            boleh = False
17        Else
18            Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MSTSUBLET WHERE KDSUBLET <> '" & NulSID(Text(0)) & "' AND NAMASUBLET = '" & NulSID(Text(3).Text) & "'  AND FS_LOG_DEL = ''")
19            If Not rsTemp.EOF Then
20                If Mid(NulSID(rsTemp!KDSUBLET), 1, 4) = "NR00" Then
21                    rsTemp.MoveFirst
22                    MsgBox "Sudah ada nama sublet HO " & NulSID(Text(3).Text) & " dengan Kode Sublet " & NulSID(rsTemp!KDSUBLET), vbInformation
23                    Exit Sub
24                ElseIf NulSID(rsTemp!KDSUPP) = NulSID(Text(1)) Then
25                    rsTemp.MoveFirst
26                    MsgBox "Sudah ada nama sublet " & NulSID(Text(3).Text) & " dan kode supplier " & NulSID(Text(1)) & " dengan Kode Sublet " & NulSID(rsTemp!KDSUBLET), vbInformation
27                    Exit Sub
28                End If
                  
29            End If
30        End If
          
31        If CDbl(hargaBeli.value) = 0 Then
32            pesan = pesan & " - Harga beli tidak boleh 0." & vbCrLf
33            boleh = False
34        End If
          ''20210815 rev ||nilai harga jual sublet dibatasi:
          ''1. tidak boleh lebih kecil dari hargabeli dan tidak boleh minus
          ''2. boleh 0
35        If CDbl(hargaJual.value) < 0 Then
36            pesan = pesan & " - Harga jual tidak boleh kurang dari 0." & vbCrLf
37            boleh = False
38        Else
39            If CDbl(hargaBeli.value) > CDbl(hargaJual.value) And CDbl(hargaJual.value) > 0 Then
40                pesan = pesan & " - Harga Jual tidak boleh kurang dari harga beli." & vbCrLf
41                boleh = False
42            End If
43        End If
          
44        If chModel.value = vbUnchecked Then
45            allModel = 0
46            rsGridMdl.filter = "IKUT = " & True
47            If rsGridMdl.EOF Then
48                pesan = pesan & " - Model kendaraan belum dipilih." & vbCrLf
49                boleh = False
50            End If
51            rsGridMdl.filter = ""
52        Else
53            allModel = 1
54        End If
55        allDealer = 0
          '' DEALER DIISI UNTUK DEALER YANG BERSANGKUTAN
      '    If chAllDealer.value = vbUnchecked Then
      '        allDealer = 0
      '        rsDealer.filter = "TERLIBAT = " & True
      '        If rsDealer.EOF Then
      '            pesan = pesan & " - Dealer belum dipilih." & vbCrLf
      '            boleh = False
      '        End If
      '        rsDealer.filter = ""
      '    Else
      '        allDealer = 1
      '
      '    End If

          '' ---------------------------------------------------------
          '' 20190314 | ketika ada sublet dengan nama yg sama dengan yg dibuat cabang lain (bukan HO) dipastikan kd supplier nya beda. Jika beda bisa buat sublet baru
          '' ---------------------------------------------------------
          
          
56        If boleh Then
57            Call objHT.OpenTransactionBengkel
58            isBegin = True
              
59            Call ClearsSql
        
60            gsSql1(0) = "KDSUBLET"
61            gsSql2(0) = NulSID(Text(0))
62            gsSql1(1) = "NAMASUBLET"
63            gsSql2(1) = NulSID(Text(3))
64            gsSql1(2) = "KODESATUAN"
65            gsSql2(2) = NulSID(Text(4))
66            gsSql1(3) = "keterangan"
67            gsSql2(3) = NulSID(Text(5))
68            gsSql1(4) = "AWALPER"
69            gsSql2(4) = Format(NulSID(periodeAwal.value), "MM/DD/YYYY")
70            gsSql1(5) = "AKHIRPER"
71            gsSql2(5) = Format(NulSID(periodeAkhir.value), "MM/DD/YYYY")
72            gsSql1(6) = "KDSUPP"
73            gsSql2(6) = NulSID(Text(1))
74            gsSql1(7) = "hargaBeli"
75            gsSql2(7) = NulSID(hargaBeli.value)
76            gsSql1(8) = "hargaJual"
77            gsSql2(8) = NulSID(hargaJual.value)
78            gsSql1(9) = "allModel"
79            gsSql2(9) = allModel
80            gsSql1(10) = "ALLCABANG"
81            gsSql2(10) = allDealer
                  
82            If isTambah Then
83                gsSql1(11) = "TGLENTRY"
84                gsSql2(11) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
85                gsSql1(12) = "USERENTRY"
86                gsSql2(12) = gSA
87            Else
88                gsSql1(11) = "TGLUPDATE"
89                gsSql2(11) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
90                gsSql1(12) = "USERUPDATE"
91                gsSql2(12) = gSA
92            End If

93            If isTambah = True Then
94                sSQL = GensSql2("[" & dbReplikasiHO & "]..MSTSUBLET", False)
95            Else
96                sSQL = GensSql2("[" & dbReplikasiHO & "]..MSTSUBLET", True) & " WHERE KDSUBLET = '" & NulSID(Text(0)) & "' "
97            End If
98            erTranAll = False
99            Call objHT.InsertData_Bengkel(sSQL)
              
      '        If chAllDealer.value = Unchecked Then
      '            rsDealer.filter = "TERLIBAT = " & True
      '            If Not rsDealer.EOF Then
      '                rsDealer.MoveFirst
      '                Call objHT.InsertData_Ho("DELETE FROM MSTSUBLETDEALER WHERE KDSUBLET = '" & NulSID(Text(0)) & "' ")
      '                Do While Not rsDealer.EOF
      '                    If erTranAll = True Then
      '                        GoTo errorRollback
      '                    Else
      '                        sSQL = "INSERT INTO MSTSUBLETDEALER(KDSUBLET,FS_KD_DEALER,FS_NM_DEALER,TGLENTRY) VALUES ('" & NulSID(Text(0)) & "','" & NulSID(rsDealer!KodeDealer) & "','" & NulSID(rsDealer!namaDealer) & "',GETDATE())"
      '                        Call objHT.InsertData_Ho(sSQL)
      '
      '                    End If
      '                    rsDealer.MoveNext
      '                Loop
      '            End If
      '        End If

100           Call objHT.InsertData_Bengkel("DELETE FROM [" & dbReplikasiHO & "]..MSTSUBLETDEALER WHERE KDSUBLET = '" & NulSID(Text(0)) & "' ")
               
101           sSQL = "INSERT INTO [" & dbReplikasiHO & "]..MSTSUBLETDEALER(KDSUBLET,FS_KD_DEALER,FS_NM_DEALER,TGLENTRY) VALUES ('" & NulSID(Text(0)) & "','" & NulSID(gKdCbg) & "','" & NulSID(gNmCbg) & "',GETDATE())"
102           Call objHT.InsertData_Bengkel(sSQL)
                       
103           If chModel.value = Unchecked Then
104               rsGridMdl.filter = "IKUT = " & True
105               If Not rsGridMdl.EOF Then
106                   rsGridMdl.MoveFirst
107                   Call objHT.InsertData_Bengkel("DELETE FROM [" & dbReplikasiHO & "]..MSTSUBLETMDL WHERE KDSUBLET = '" & NulSID(Text(0)) & "' ")
108                   Do While Not rsGridMdl.EOF
109                       If erTranAll = True Then
110                           GoTo errorRollback
111                       Else
112                           sSQL = "INSERT INTO [" & dbReplikasiHO & "]..MSTSUBLETMDL(KDSUBLET,MDL,DESCMDL,TGLENTRY) VALUES ('" & NulSID(Text(0)) & "','" & NulSID(rsGridMdl!kodeMDL) & "','" & NulSID(rsGridMdl!NAMAMDL) & "',GETDATE())"
113                           Call objHT.InsertData_Bengkel(sSQL)
                              
114                       End If
115                       rsGridMdl.MoveNext
116                   Loop
117               End If
118           End If
              
119           If erTranAll = False Then
120               Call objHT.CloseTransactionBengkel
121               MsgBox "Berhasil menyimpan data.", vbInformation
122               Call ber_sih
123               Call buka_Tutup(False)
124               Call refreshGrid
125           Else
errorRollback:
126               Call objHT.CloseTransactionBengkel_err
127               MsgBox "Gagal menyimpan data.", vbInformation
128               Exit Sub
129           End If
130       Else
131           MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
132       End If

133      On Error GoTo 0
134      Exit Sub

cmdSimpan_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
135       varErlError = Erl
136       varErrNumber = err.Number
137       If isBegin = True Then
138           Call objHT.CloseTransactionBengkel_err
139       End If
140       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmTSubletStandarisasi")
End Sub

Private Sub cmdTambah_Click()
    isTambah = True
    Call ber_sih
    Call isiId
    Call buka_Tutup(True)
    GridSupplier.Enabled = False
End Sub

Private Sub cmdTutup_Click()
    frmeModel.Visible = False
End Sub

Private Sub cmdTutupDealer_Click()
    frmeDealer.Visible = False
    
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False)
    Call refreshGrid
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

Private Sub GridModel_HeadClick(ByVal ColIndex As Integer)
On Error GoTo err
   
    If ColIndex = 3 Then
        If rsGridMdl.State Then
            rsGridMdl.MoveFirst
            Do While Not rsGridMdl.EOF
                rsGridMdl!ikut = IIf(rsGridMdl!ikut = True, False, True)
                rsGridMdl.MoveNext
            Loop
            rsGridMdl.MoveFirst
        End If
    Else
        Urut = IIf(Urut = "DESC", "ASC", "DESC")
        If ColIndex <= 3 Then rsGridMdl.Sort = "[" & GridModel.Columns(ColIndex).DataField & "]" & " " & Urut
        GridModel.SetFocus
        Sendkeys "{Right " & Val(ColIndex) & "}"
    End If
    Exit Sub
err:
   MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure HeadClick ")

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

1        On Error GoTo GridSupplier_HeadClick_Error
2         Urut = IIf(Urut = "DESC", "ASC", "DESC")
3         rsGrid.Sort = GridSupplier.Columns(ColIndex).DataField & " " & Urut
4         GridSupplier.SetFocus
5         Sendkeys "{Right " & Val(ColIndex) & "}"
          
6        On Error GoTo 0
7        Exit Sub

GridSupplier_HeadClick_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
8         varErlError = Erl
9         varErrNumber = err.Number
10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure GridSupplier_HeadClick of Form frmTSubletStandarisasi")
End Sub

Private Sub gridSupplier_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF Then
        Text(0).Text = NulSID(rsGrid!KDSUBLET)
        Text(1).Text = NulSID(rsGrid!KDSUPP)
        Text(2).Text = NulSID(rsGrid!supplier)
        Text(3).Text = NulSID(rsGrid!namaSublet)
        Text(4).Text = NulSID(rsGrid!KODESATUAN)
        Text(5).Text = NulSID(rsGrid!keterangan)
        hargaBeli.value = NulSID(rsGrid!hargaBeli)
        hargaJual.value = NulSID(rsGrid!hargaJual)
        periodeAwal.value = Format(NulSID(rsGrid!AWALPER), "dd/mm/yyyy")
        periodeAkhir.value = Format(NulSID(rsGrid!AKHIRPER), "dd/mm/yyyy")
        If NulSID(rsGrid!ALLCABANG) = "1" Then
            chAllDealer.value = vbChecked
        Else
            chAllDealer.value = vbUnchecked
'            Call isiDealer
'            frmeDealer.Visible = False
        End If
        
        If NulSID(rsGrid!allModel) = "1" Then
            chModel.value = vbChecked
        Else
            chModel.value = vbUnchecked
'            Call isiModel
'            frmeModel.Visible = False
        End If
    End If
End Sub


Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
    Set rsTmp = objO.FetchData("SELECT A.fs_kd_dealer, A.fs_nm_dealer, ISNULL(b.KDSUBLET,'') AS KODE FROM [" & dbReplikasiHO & "]..t_dealer a WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER b ON a.fs_kd_dealer = b.FS_KD_DEALER and A.FS_LOG_DEL = '' AND b.KDSUBLET = '" & NulSID(Text(0)) & "' WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
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

Private Sub createRS()
    
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    With rsDealer
        .Fields.Append "KODEDEALER", adVarChar, 15
        .Fields.Append "NAMADEALER", adVarChar, 25
        .Fields.Append "TERLIBAT", adBoolean
        .Open
    End With
    
    Set GridDealer.DataSource = rsDealer
End Sub


Private Sub isiModel()
    Set rsTemp = objO.FetchData("SELECT DISTINCT a.MODEL, a.mdl, a.[desc], ISNULL(kdsublet, '') as PRO FROM [" & dbReplikasiHO & "]..model a WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL b WITH (NOLOCK) ON a.MDL = b.MDL AND B.kdsublet = '" & NulSID(Text(0)) & "' ORDER BY a.MODEL, a.mdl")
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

Private Sub refreshGrid()
    GridSupplier.ClearFields
    sSQL = "SELECT ISNULL(C.KODEAP,'') AS KODEAP, A.KDSUBLET,A.NAMASUBLET,A.KODESATUAN,A.KETERANGAN,A.KDSUPP,B.SUPPLIER,A.HARGABELI,A.HARGAJUAL,A.AWALPER,A.AKHIRPER,A.ALLCABANG,A.ALLMODEL FROM [" & dbReplikasiHO & "]..MSTSUBLET A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..MSTSUPPLIER B WITH (NOLOCK) ON A.KDSUPP = B.KDSUPP LEFT JOIN T_SUPPLIER C WITH (NOLOCK) ON C.KDSUPP = B.KDSUPP WHERE A.FS_LOG_DEL = ''"
    '' 20201022DEV
    sSQL = sSQL & " AND (SUBSTRING(A.KDSUBLET,1,4) = 'NR00' OR SUBSTRING(A.KDSUBLET,1,4) = '" & Replace(gKdCbg, " ", "") & "')"
    Set rsGrid = objO.FetchData(sSQL)
    Set GridSupplier.DataSource = rsGrid
End Sub


Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            'Text(Index).Mask = ""
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            If Index = 1 Then
                Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WHERE KDSUPP = '" & NulSID(Text(1).Text) & "'  AND FS_LOG_DEL = ''")
                If rsTemp.EOF Then
                    MsgBox " - Kode Supplier tidak terdaftar.", vbInformation
                    Text(1).Text = ""
                Else
                    Set rsTemp = objO.FetchData("SELECT * FROM T_SUPPLIER WHERE KDSUPP = '" & NulSID(Text(1).Text) & "'  AND kodeap <> ''")
                    If rsTemp.EOF Then
                        MsgBox " - Kode Supplier belum diset Kode AP.", vbInformation
                        Text(1).Text = ""
                    End If
                End If
            End If
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
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "SUPPLIERSUBLETSENTRAL"
                frmCari.show (vbModal)
            Case 3
                gMainKeyCari = Text(Index)
                gSearchOption = "NAMASUBLETSENTRAL"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUBLET WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY RIGHT(KDSUBLET,10) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KDSUBLET"), 4))
        noUrut = noUrut + 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
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
    
    If KeyAscii = 13 And cmdSimpan.Enabled = True Then
        Select Case Index
            Case 1
                ''Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WHERE KDSUPP = '" & NulSID(Text(1).Text) & "'  AND FS_LOG_DEL = ''  AND (SUBSTRING(KDSUPP,1,4) = 'NR00' OR SUBSTRING(KDSUPP,1,4) = '" & Replace(gKdCbg, " ", "") & "')")
                Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WHERE KDSUPP = '" & NulSID(Text(1).Text) & "'  AND FS_LOG_DEL = ''  AND ((SUBSTRING(KDSUPP,1,4) = '" & Replace(gKdCbg, " ", "") & "') or (SUBSTRING(KDSUPP,1,4) = 'NR00' and aflflag = 1))")
                If rsTemp.EOF Then
                    MsgBox " - Kode Supplier tidak terdaftar.", vbInformation
                    Text(1).Text = ""
                    Text(2).Text = ""
                Else
                    Text(2).Text = NulSID(rsTemp!supplier)
                End If
        End Select
    End If
End Sub
 
Private Sub Text_LostFocus(Index As Integer)
    If Index = 1 Then
        Call text_KeyPress(1, 13)
    End If
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

