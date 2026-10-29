VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form FormKPI_asli 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormKPI asli.frx":0000
   ScaleHeight     =   8655
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdKPIHarian 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Harian"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      Picture         =   "FormKPI asli.frx":F6E4
      Style           =   1  'Graphical
      TabIndex        =   134
      Top             =   7440
      Width           =   735
   End
   Begin TrueOleDBGrid70.TDBGrid GridExport 
      Height          =   1095
      Left            =   1680
      TabIndex        =   122
      Top             =   7440
      Visible         =   0   'False
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   1931
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "XXX"
      Columns(0).DataField=   "XXX"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "YYY"
      Columns(1).DataField=   "YYY"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "ZZZ"
      Columns(2).DataField=   "ZZZ"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
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
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   0
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      WrapCellPointer =   -1  'True
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   2
      DeadAreaBackColor=   12577002
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
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5160
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdXls 
      BackColor       =   &H00FFFFFF&
      Caption         =   "X&LS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      Picture         =   "FormKPI asli.frx":FCDD
      Style           =   1  'Graphical
      TabIndex        =   121
      Top             =   7440
      Width           =   735
   End
   Begin VB.PictureBox PictGr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7200
      Left            =   480
      ScaleHeight     =   7200
      ScaleWidth      =   11775
      TabIndex        =   103
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin TrueOleDBGrid70.TDBGrid GridR 
         Height          =   4935
         Left            =   9000
         TabIndex        =   119
         Top             =   1620
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   8705
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KM"
         Columns(0).DataField=   "KM"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Jml"
         Columns(1).DataField=   "Jml"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "%"
         Columns(2).DataField=   "Persen"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1058"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=953"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=873"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=767"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=122,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=119,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=120,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=121,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
      Begin VB.OptionButton OptD 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ratio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   4080
         TabIndex        =   116
         Top             =   6840
         Width           =   1455
      End
      Begin VB.OptionButton OptD 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Datang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   115
         Top             =   6840
         Width           =   1335
      End
      Begin VB.OptionButton OptD 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Diundang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   114
         Top             =   6840
         Width           =   1335
      End
      Begin VB.ComboBox cboMax 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FormKPI asli.frx":10159
         Left            =   1320
         List            =   "FormKPI asli.frx":10163
         Style           =   2  'Dropdown List
         TabIndex        =   111
         Top             =   960
         Width           =   1680
      End
      Begin VB.ComboBox cboType 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "FormKPI asli.frx":1017C
         Left            =   1320
         List            =   "FormKPI asli.frx":10186
         Style           =   2  'Dropdown List
         TabIndex        =   106
         Top             =   600
         Width           =   1680
      End
      Begin TDBDate6Ctl.TDBDate Dt1 
         Height          =   315
         Left            =   10200
         TabIndex        =   108
         Top             =   960
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "FormKPI asli.frx":1019F
         Caption         =   "FormKPI asli.frx":102CB
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FormKPI asli.frx":10337
         Keys            =   "FormKPI asli.frx":10355
         Spin            =   "FormKPI asli.frx":103B3
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate Dt0 
         Height          =   315
         Left            =   8400
         TabIndex        =   109
         Top             =   960
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "FormKPI asli.frx":103DB
         Caption         =   "FormKPI asli.frx":10507
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FormKPI asli.frx":10573
         Keys            =   "FormKPI asli.frx":10591
         Spin            =   "FormKPI asli.frx":105EF
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   4935
         Left            =   9000
         TabIndex        =   118
         Top             =   1620
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   8705
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KM"
         Columns(0).DataField=   "KM"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Jumlah"
         Columns(1).DataField=   "Jml"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1667"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1826"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
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
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=86,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=83,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=84,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=85,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=122,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=119,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=120,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=121,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin MSChart20Lib.MSChart MsGraf 
         Height          =   5175
         Left            =   0
         OleObjectBlob   =   "FormKPI asli.frx":10617
         TabIndex        =   105
         Top             =   1380
         Width           =   9180
      End
      Begin VB.Label Label60 
         Caption         =   "                  .::  DATA  ::."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   9000
         TabIndex        =   120
         Top             =   1380
         Width           =   2600
      End
      Begin VB.Label Label59 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Grafik Data     :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   117
         Top             =   6840
         Width           =   1800
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   11760
         X2              =   0
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   11640
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000000FF&
         X1              =   8400
         X2              =   11640
         Y1              =   865
         Y2              =   865
      End
      Begin VB.Label Label58 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Range Data :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8400
         TabIndex        =   113
         Top             =   600
         Width           =   3240
      End
      Begin VB.Label Label57 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Max Km        :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   112
         Top             =   1020
         Width           =   1800
      End
      Begin VB.Label Label56 
         BackColor       =   &H00FFFFFF&
         Caption         =   " s/d"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   110
         Top             =   960
         Width           =   360
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   2280
         X2              =   3960
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   1440
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   240
         X2              =   3960
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "RATIO SBI AND SBE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   840
         TabIndex        =   104
         Top             =   120
         Width           =   9975
      End
      Begin VB.Shape Shape15 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00404040&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11715
      End
      Begin VB.Label Label55 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kelipatan      :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   660
         Width           =   1800
      End
   End
   Begin VB.CommandButton CmdOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Graph"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   720
      Picture         =   "FormKPI asli.frx":123F0
      Style           =   1  'Graphical
      TabIndex        =   102
      Top             =   7440
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   9000
      ScaleHeight     =   6045
      ScaleWidth      =   2775
      TabIndex        =   58
      Top             =   1080
      Width           =   2775
      Begin VB.PictureBox sKPI 
         Appearance      =   0  'Flat
         BackColor       =   &H00BFE8EA&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   14000
         Left            =   0
         ScaleHeight     =   13995
         ScaleWidth      =   2895
         TabIndex        =   62
         Top             =   0
         Width           =   2895
         Begin VB.Label KPI_3salesB 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   133
            Top             =   3615
            Width           =   1005
         End
         Begin VB.Label KPI_3salesA 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   132
            Top             =   3375
            Width           =   1005
         End
         Begin VB.Label KPI_3sales 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   131
            Top             =   3120
            Width           =   1005
         End
         Begin VB.Label KPI_2c 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   126
            Top             =   2280
            Width           =   1005
         End
         Begin VB.Label KPI_1c 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   125
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label KPI_11b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   59
            Top             =   12240
            Width           =   1005
         End
         Begin VB.Label KPI_11 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   97
            Top             =   11760
            Width           =   1005
         End
         Begin VB.Label KPI_11a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   96
            Top             =   12000
            Width           =   1005
         End
         Begin VB.Label KPI_12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   95
            Top             =   12600
            Width           =   1005
         End
         Begin VB.Label KPI_12a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   94
            Top             =   12840
            Width           =   1005
         End
         Begin VB.Label KPI_12b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   93
            Tag             =   "KPI_12b"
            Top             =   13080
            Width           =   1005
         End
         Begin VB.Label KPI_6b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   92
            Top             =   6600
            Width           =   1005
         End
         Begin VB.Label KPI_6 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   91
            Top             =   6120
            Width           =   1005
         End
         Begin VB.Label KPI_6a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   90
            Top             =   6360
            Width           =   1005
         End
         Begin VB.Label KPI_7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   89
            Top             =   7920
            Width           =   1005
         End
         Begin VB.Label KPI_7a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   88
            Top             =   8160
            Width           =   1005
         End
         Begin VB.Label KPI_7b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   87
            Top             =   8400
            Width           =   1005
         End
         Begin VB.Label KPI_8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   86
            Top             =   8760
            Width           =   1005
         End
         Begin VB.Label KPI_8a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   85
            Top             =   9000
            Width           =   1005
         End
         Begin VB.Label KPI_8b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   84
            Top             =   9240
            Width           =   1005
         End
         Begin VB.Label KPI_9 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   83
            Top             =   9600
            Width           =   1005
         End
         Begin VB.Label KPI_9a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   82
            Top             =   9840
            Width           =   1005
         End
         Begin VB.Label KPI_9b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   81
            Top             =   10080
            Width           =   1005
         End
         Begin VB.Label KPI_10 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   80
            Top             =   10920
            Width           =   1005
         End
         Begin VB.Label KPI_10a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   79
            Top             =   11160
            Width           =   1005
         End
         Begin VB.Label KPI_10b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   78
            Top             =   11400
            Width           =   1005
         End
         Begin VB.Label KPI_1b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   77
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label KPI_1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   76
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label KPI_1a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   75
            Top             =   720
            Width           =   1005
         End
         Begin VB.Label KPI_2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   74
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label KPI_2a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   73
            Top             =   1800
            Width           =   1005
         End
         Begin VB.Label KPI_2b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   72
            Top             =   2040
            Width           =   1005
         End
         Begin VB.Label KPI_3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   71
            Top             =   7005
            Width           =   1005
         End
         Begin VB.Label KPI_3a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   70
            Top             =   7260
            Width           =   1005
         End
         Begin VB.Label KPI_3b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   69
            Top             =   7500
            Width           =   1005
         End
         Begin VB.Label KPI_4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   68
            Top             =   4485
            Width           =   1005
         End
         Begin VB.Label KPI_4a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   67
            Top             =   4740
            Width           =   1005
         End
         Begin VB.Label KPI_4b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   66
            Top             =   4965
            Width           =   1005
         End
         Begin VB.Label KPI_5 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   65
            Top             =   5295
            Width           =   1005
         End
         Begin VB.Label KPI_5a 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   64
            Top             =   5540
            Width           =   1005
         End
         Begin VB.Label KPI_5b 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   840
            TabIndex        =   63
            Top             =   5790
            Width           =   1005
         End
         Begin VB.Shape Shape13 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   -360
            Shape           =   4  'Rounded Rectangle
            Top             =   10440
            Width           =   2955
         End
         Begin VB.Shape Shape11 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   -360
            Shape           =   4  'Rounded Rectangle
            Top             =   4035
            Width           =   2955
         End
         Begin VB.Shape Shape9 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   -360
            Shape           =   4  'Rounded Rectangle
            Top             =   2640
            Width           =   2955
         End
         Begin VB.Shape Shape10 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   -360
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   2955
         End
      End
   End
   Begin VB.VScrollBar VScrollKPI 
      Height          =   6090
      LargeChange     =   300
      Left            =   11760
      Max             =   7500
      SmallChange     =   250
      TabIndex        =   4
      Top             =   1080
      Width           =   430
   End
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6045
      Left            =   480
      ScaleHeight     =   6045
      ScaleWidth      =   10575
      TabIndex        =   3
      Top             =   1080
      Width           =   10575
      Begin VB.PictureBox qKPI 
         Appearance      =   0  'Flat
         BackColor       =   &H00BFE8EA&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   14000
         Left            =   0
         ScaleHeight     =   13995
         ScaleWidth      =   8880
         TabIndex        =   5
         Top             =   0
         Width           =   8880
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   130
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label ket3Sales 
            BackStyle       =   0  'Transparent
            Caption         =   ": 1K Activity Sales (=Pa/Pc)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   129
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label ket3salesa 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah pelanggan yang seharusnya diingatkan servis 1.000 (Pc)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   128
            Top             =   3375
            Width           =   7005
         End
         Begin VB.Label ket3salesb 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah pelanggan yang aktual diingatkan servis 1.000 (Pa)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   127
            Top             =   3615
            Width           =   7005
         End
         Begin VB.Label ket2c 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah carry over data pelanggan yang unitnya telah delivery ditransfer ke after sales (Pt)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   124
            Top             =   2280
            Width           =   7005
         End
         Begin VB.Label ket1c 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah carry over data pelanggan valid dari yang unitnya telah delivery (Pv)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   123
            Top             =   1200
            Width           =   7005
         End
         Begin VB.Label ket12b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah pelanggan yang servis 1.000 dari penjualan unit N-1 (Us) "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   57
            Top             =   13080
            Width           =   7005
         End
         Begin VB.Label ket12a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah unit yang telah dikirimkan kepada pelanggan N-1 (Uk)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   56
            Top             =   12840
            Width           =   7005
         End
         Begin VB.Label ket12 
            BackStyle       =   0  'Transparent
            Caption         =   ": SBI Retention (=Us/Uk) "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   55
            Top             =   12600
            Width           =   6255
         End
         Begin VB.Label Label51 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 12"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   54
            Top             =   12600
            Width           =   735
         End
         Begin VB.Label ket11b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah total unit SBE (Pt)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   53
            Top             =   12240
            Width           =   7005
         End
         Begin VB.Label ket11a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah unit SBE karena aktifitas MRS (Pd)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   52
            Top             =   12000
            Width           =   7005
         End
         Begin VB.Label ket11 
            BackStyle       =   0  'Transparent
            Caption         =   ": Kontribusi MRS terhadap SBE (=Pd/Pt)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   51
            Top             =   11760
            Width           =   6255
         End
         Begin VB.Label Label47 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 11"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   11760
            Width           =   735
         End
         Begin VB.Label ket10b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah total app. (A2)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   49
            Top             =   11400
            Width           =   7005
         End
         Begin VB.Label ket10a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah app. Karena aktifitas MRS (A1)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   48
            Top             =   11160
            Width           =   7005
         End
         Begin VB.Label ket10 
            BackStyle       =   0  'Transparent
            Caption         =   ": Kontribusi MRS terhadap appointment. (=A1/A2)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   47
            Top             =   10920
            Width           =   6255
         End
         Begin VB.Label Label43 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 10"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   10920
            Width           =   735
         End
         Begin VB.Label Label42 
            BackStyle       =   0  'Transparent
            Caption         =   ".:: MRS RESULT ::."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   45
            Top             =   10500
            Width           =   3855
         End
         Begin VB.Label ket9b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Pelanggan Yang Diundang (P2)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   44
            Top             =   10080
            Width           =   7005
         End
         Begin VB.Label ket9a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Pelanggan Yang Datang Setelah Diundang (Pd)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   43
            Top             =   9840
            Width           =   7005
         End
         Begin VB.Label ket9 
            BackStyle       =   0  'Transparent
            Caption         =   ": Persentase Pelanggan Yang Datang Setelah Diundang (=Pd/P2)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   42
            Top             =   9600
            Width           =   7935
         End
         Begin VB.Label Label38 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 9"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   9600
            Width           =   735
         End
         Begin VB.Label ket8b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Telepon Terhubung (Tt)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   40
            Top             =   9240
            Width           =   7005
         End
         Begin VB.Label ket8a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Pelanggan Yang Appoinment (A1)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   39
            Top             =   9000
            Width           =   7005
         End
         Begin VB.Label ket8 
            BackStyle       =   0  'Transparent
            Caption         =   ": Persentase Pelanggan Yang Appoinment (=A1/Tt)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   38
            Top             =   8760
            Width           =   6255
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 8"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   8760
            Width           =   735
         End
         Begin VB.Label ket7b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Pelanggan Yang Diundang (P2)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   36
            Top             =   8400
            Width           =   7005
         End
         Begin VB.Label ket7a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Pelanggan Yang Ditelepon (Tk)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   35
            Top             =   8160
            Width           =   7005
         End
         Begin VB.Label ket7 
            BackStyle       =   0  'Transparent
            Caption         =   ": Persentase Pelanggan Yang Ditelepon (=Tk/P2)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   34
            Top             =   7920
            Width           =   6255
         End
         Begin VB.Label Label30 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   7920
            Width           =   735
         End
         Begin VB.Label ket6b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Pelanggan Yang Diundang (P2= DM + SMS + E + Tlp)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   32
            Top             =   6600
            Width           =   7005
         End
         Begin VB.Label ket6a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Pelanggan Yang Harus Diundang (P1=Aktif + Pasif + Baru)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   31
            Top             =   6360
            Width           =   7005
         End
         Begin VB.Label ket6 
            BackStyle       =   0  'Transparent
            Caption         =   ": Persentase Pelanggan Yang Diundang (=P2/P1)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   30
            Top             =   6120
            Width           =   6255
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 6"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   6120
            Width           =   735
         End
         Begin VB.Label ket5b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Pelanggan Yang Ditelepon (Tk)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   28
            Top             =   5790
            Width           =   7005
         End
         Begin VB.Label ket5a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Telepon Terhubung (Tt)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   27
            Top             =   5565
            Width           =   7005
         End
         Begin VB.Label ket5 
            BackStyle       =   0  'Transparent
            Caption         =   ": Akurasi Nomor Telepon Pelanggan (=Tt/Tk)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   26
            Top             =   5295
            Width           =   6255
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 5"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   5295
            Width           =   735
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   ".:: MRS PROCESS ::."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   4095
            Width           =   3855
         End
         Begin VB.Label ket4b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Direct Mail Dikirim (Dk)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   24
            Top             =   4965
            Width           =   7005
         End
         Begin VB.Label ket4a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Direct Mail Diterima Pelanggan (Dt)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   23
            Top             =   4740
            Width           =   7005
         End
         Begin VB.Label ket4 
            BackStyle       =   0  'Transparent
            Caption         =   ": Akurasi Alamat Pelanggan (=Dt/Dk)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   22
            Top             =   4485
            Width           =   6255
         End
         Begin VB.Label Label18 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 4"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   4485
            Width           =   735
         End
         Begin VB.Shape Shape6 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   4035
            Width           =   9315
         End
         Begin VB.Label ket3b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah pelanggan yang aktual diundang servis 1.000 (Pa)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   19
            Top             =   7500
            Width           =   7005
         End
         Begin VB.Label ket3a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah pelanggan yang seharusnya diundang servis 1.000 (Pc)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   18
            Top             =   7275
            Width           =   7005
         End
         Begin VB.Label ket3 
            BackStyle       =   0  'Transparent
            Caption         =   ": 1K Activity MRS (=Pa/Pc)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   17
            Top             =   7005
            Width           =   2895
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 6b"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   7005
            Width           =   735
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   ".:: 1K SERVICE INVITATION / CR2 ::."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   2700
            Width           =   3855
         End
         Begin VB.Label ket2b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah data pelanggan yang unitnya telah delivery ditransfer ke after sales (Pt)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   14
            Top             =   2040
            Width           =   7005
         End
         Begin VB.Label ket2a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah data pelanggan yang unitnya telah delivery (Pd)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   13
            Top             =   1800
            Width           =   7005
         End
         Begin VB.Label ket2 
            BackStyle       =   0  'Transparent
            Caption         =   ": Database Transfer (=Pt/Pd)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   12
            Top             =   1560
            Width           =   2895
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 2 "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1560
            Width           =   735
         End
         Begin VB.Label ket1b 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah data pelanggan valid dari yang unitnya telah delivery (Pv)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   10
            Top             =   960
            Width           =   7000
         End
         Begin VB.Label ket1a 
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah data pelanggan yang unitnya telah delivery (Pd)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   960
            TabIndex        =   9
            Top             =   720
            Width           =   7000
         End
         Begin VB.Label ket1 
            BackStyle       =   0  'Transparent
            Caption         =   ": Database Valid (=Pv/Pd)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   840
            TabIndex        =   8
            Top             =   480
            Width           =   2895
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "KPI 1 "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   ".:: SALES DATABASE ::."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   60
            Width           =   2895
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   9315
         End
         Begin VB.Shape Shape5 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   2640
            Width           =   9315
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   10440
            Width           =   9315
         End
      End
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pr&ocess"
      Height          =   732
      Left            =   10200
      Picture         =   "FormKPI asli.frx":129E9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7440
      Width           =   732
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11280
      Picture         =   "FormKPI asli.frx":13056
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7440
      Width           =   735
   End
   Begin ucXPButton.XPButton thnMinus 
      Height          =   495
      Left            =   480
      TabIndex        =   98
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormKPI asli.frx":137E6
      PICN            =   "FormKPI asli.frx":13802
      PICH            =   "FormKPI asli.frx":13A0D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton blnMinus 
      Height          =   495
      Left            =   9000
      TabIndex        =   99
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormKPI asli.frx":13C18
      PICN            =   "FormKPI asli.frx":13C34
      PICH            =   "FormKPI asli.frx":13E3F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton thnPlus 
      Height          =   495
      Left            =   3560
      TabIndex        =   100
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormKPI asli.frx":1404A
      PICN            =   "FormKPI asli.frx":14066
      PICH            =   "FormKPI asli.frx":1425D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton blnPlus 
      Height          =   495
      Left            =   11355
      TabIndex        =   101
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormKPI asli.frx":14454
      PICN            =   "FormKPI asli.frx":14470
      PICH            =   "FormKPI asli.frx":14667
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape17 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   975
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   9240
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Shape Shape14 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label nTahun 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Periode  :  2009"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   60
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label nBulan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AGUSTUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9600
      TabIndex        =   61
      Top             =   600
      Width           =   1695
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   435
      Left            =   9480
      Shape           =   4  'Rounded Rectangle
      Top             =   510
      Width           =   1875
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   9240
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   12240
      X2              =   480
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   975
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   720
      X2              =   4440
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1920
      X2              =   4440
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2760
      X2              =   4440
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   435
      Left            =   960
      Shape           =   4  'Rounded Rectangle
      Top             =   510
      Width           =   2595
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLM KPI MONITORING"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   9975
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   11715
   End
End
Attribute VB_Name = "FormKPI_asli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'Modified KPI by Thania
'==============================================
Option Explicit
Private OnLoad As Boolean

Private rs1 As ADODB.Recordset
Private rs2 As ADODB.Recordset
Private rs3 As ADODB.Recordset
Private obj As New ClsKPI

Private rsKPI As New ADODB.Recordset

Private Bulan(12) As String
Private Bln As Integer
Private Thn As Integer

Private RsUnd As New ADODB.Recordset
Private RsDat As New ADODB.Recordset
Private RsRat As New ADODB.Recordset
Private rsKPIRepo As New ADODB.Recordset
Private rsXLS As New ADODB.Recordset

Private Sub cmdExit_Click()
   Unload Me
End Sub
Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property

Private Sub cmdKPIHarian_Click()
    FormKPIHarian.Show
    FormKPIHarian.WindowState = vbMaximized
End Sub

Private Sub CmdXls_Click()
    Dim f As New FormExplorer
    f.Show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    
    Call BackUpMyBase(MyPath & f.FileName & ".XLS")
End Sub

Private Sub Form_Load()
   If OnLoad = False Then Exit Sub
   Call FillAllCombo
   Bln = Month(gTglServer) 'Month(Now)
   Thn = Year(gTglServer)  'Year(Now)
   nBulan.Caption = Bulan(Bln)
   nTahun.Caption = "Periode  :  " & Thn
   Call cmdProcess_Click
End Sub

Private Sub FillAllCombo()
   Bulan(0) = "TOTAL"
   Bulan(1) = "JANUARI"
   Bulan(2) = "FEBRUARI"
   Bulan(3) = "MARET"
   Bulan(4) = "APRIL"
   Bulan(5) = "MEI"
   Bulan(6) = "JUNI"
   Bulan(7) = "JULI"
   Bulan(8) = "AGUSTUS"
   Bulan(9) = "SEPTEMBER"
   Bulan(10) = "OKTOBER"
   Bulan(11) = "NOVEMBER"
   Bulan(12) = "DESEMBER"
   With Me.cboMax
    .Clear
    Dim xx As Double
    For xx = 5000 To 750000 Step 5000
        .AddItem xx & " Km"
    Next xx
   End With
End Sub

Private Sub cmdProcess_Click()
On Error Resume Next
If PictGr.Visible = True Then
    Call SetRatio
    Exit Sub
End If
    Call SetKPI_1
    Call SetKPI_2
    Call SetKPI_3
    Call SetKPI_3SALES
    Call SetKPI_4
    Call SetKPI_5
    Call SetKPI_6
    Call SetKPI_7
    Call SetKPI_8
    Call SetKPI_9
    Call SetKPI_10
    Call SetKPI_11
    Call SetKPI_12
    Call SetKPI_13
End Sub

Private Sub SetKPI_1()
    On Error Resume Next
    Dim Pd As Integer 'Jumlah data pelanggan yang unitnya telah delivery (Pd)
    Dim Pv As Integer 'Jumlah data pelanggan valid dari yang unitnya telah delivery (Pv)
    Dim cO As Integer 'Jumlah carry over data pelanggan valid dari yang unitnya telah delivery (Pv)
    Set rs1 = obj.KPI_1(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then Pd = rs1.Fields(0).value
    Set rs2 = obj.KPI_1(False, Bln, Thn)
    If rs2.RecordCount <> 0 Then Pv = rs2.Fields(0).value
    Set rs3 = obj.KPI_1c(True, Bln, Thn)
    If rs3.RecordCount <> 0 Then cO = rs3.Fields(0).value
    KPI_1.Caption = "0 %"
    KPI_1a.Caption = Pd
    KPI_1b.Caption = Pv
    KPI_1c.Caption = cO
    KPI_1.Caption = CStr(Round((CInt(Pv) / CInt(Pd)) * 100, 2)) & " %"
End Sub

Private Sub SetKPI_2()
    On Error Resume Next
    Dim Pd As Integer 'Jumlah data pelanggan yang unitnya telah delivery (Pd)
    Dim Pt As Integer 'Jumlah data pelanggan yang unitnya telah delivery ditransfer ke after sales (Pt)
    Dim cO As Integer 'Jumlah carry over data pelanggan yang unitnya telah delivery ditransfer ke after sales (Pt)
    Pd = KPI_1a.Caption
    'Set rs1 = obj.KPI_1(True, Bln, Thn)
    'If rs1.RecordCount <> 0 Then Pd = rs1.Fields(0).value
    Set rs2 = obj.KPI_2(True, Bln, Thn)
    If rs2.RecordCount <> 0 Then Pt = rs2.Fields(0).value
    Set rs1 = obj.KPI_2(False, Bln, Thn)
    If rs1.RecordCount <> 0 Then cO = rs1.Fields(0).value
    KPI_2.Caption = "0 %"
    KPI_2a.Caption = Pd
    KPI_2b.Caption = Pt
    KPI_2c.Caption = cO
    KPI_2.Caption = CStr(Round((CInt(Pt) / CInt(Pd)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_3()
    On Error Resume Next
    Dim Pc As Integer 'Jumlah pelanggan yang seharusnya diundang servis 1.000 (Pc)
    Dim Pa As Integer 'Jumlah pelanggan yang aktual diundang servis 1.000 (Pa)
    Set rs1 = obj.KPI_3(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then Pc = rs1.Fields(0).value
    Set rs2 = obj.KPI_3(False, Bln, Thn)
    If rs2.RecordCount <> 0 Then Pa = rs2.Fields(0).value
    KPI_3.Caption = "0 %"
    KPI_3a.Caption = Pc
    KPI_3b.Caption = Pa
    KPI_3.Caption = CStr(Round((CInt(Pa) / CInt(Pc)) * 100, 2)) & " %"
End Sub

Private Sub SetKPI_3SALES()
    On Error Resume Next
    Dim Pc As Integer 'Jumlah pelanggan yang seharusnya diundang servis 1.000 (Pc)
    Dim Pa As Integer 'Jumlah pelanggan yang aktual diundang servis 1.000 (Pa)
    Set rs1 = obj.KPI_3sales(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then Pc = rs1.Fields(0).value
    Set rs2 = obj.KPI_3sales(False, Bln, Thn)
    If rs2.RecordCount <> 0 Then Pa = rs2.Fields(0).value
    KPI_3sales.Caption = "0 %"
    KPI_3salesA.Caption = Pc
    KPI_3salesB.Caption = Pa
    KPI_3sales.Caption = CStr(Round((CInt(Pa) / CInt(Pc)) * 100, 2)) & " %"
End Sub

Private Sub SetKPI_4()
    On Error Resume Next
    Dim Dt As Integer 'Jumlah Direct Mail Diterima Pelanggan (Dt)
    Dim Dk As Integer 'Jumlah Direct Mail Dikirim (Dk)
    Set rs1 = obj.KPI_4(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then Dt = rs1.Fields(0).value
    Set rs2 = obj.KPI_4(False, Bln, Thn)
    If rs2.RecordCount <> 0 Then Dk = rs2.Fields(0).value
    KPI_4.Caption = "0 %"
    KPI_4a.Caption = Dt
    KPI_4b.Caption = Dk
    KPI_4.Caption = CStr(Round((CInt(Dt) / CInt(Dk)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_5()
    On Error Resume Next
    Dim Tt As Integer 'Jumlah Telepon Terhubung (Tt)
    Dim Tk As Integer 'Jumlah Pelanggan Yang Ditelepon (Tk)
    Set rs1 = obj.KPI_5(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then Tt = rs1.Fields(0).value
    Set rs2 = obj.KPI_5(False, Bln, Thn)
    If rs2.RecordCount <> 0 Then Tk = rs2.Fields(0).value
    KPI_5.Caption = "0 %"
    KPI_5a.Caption = Tt
    KPI_5b.Caption = Tk
    KPI_5.Caption = CStr(Round((CInt(Tt) / CInt(Tk)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_6()
    On Error Resume Next
    Dim P1 As Integer 'Jumlah Pelanggan Yang Harus Diundang (P1=Aktif + Pasif + Baru)
    Dim P2 As Integer 'Jumlah Pelanggan Yang Diundang (P2= DM + SMS + E + Tlp)
    Set rs1 = obj.KPI_6(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then P1 = rs1.Fields(0).value
    Set rs1 = obj.KPI_6a(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then P1 = P1 + CInt(rs1.Fields(0).value)
    Set rs2 = obj.KPI_6(False, Bln, Thn)
    If rs2.RecordCount <> 0 Then P2 = rs2.Fields(0).value
    Set rs2 = obj.KPI_6b(True, Bln, Thn)
    If rs2.RecordCount <> 0 Then P2 = P2 + CInt(rs2.Fields(0).value)
    KPI_6.Caption = "0 %"
    KPI_6a.Caption = P1
    KPI_6b.Caption = P2
    KPI_6.Caption = CStr(Round((CInt(P2) / CInt(P1)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_7()
    On Error Resume Next
    'Dim Tk As Integer 'Jumlah Pelanggan Yang Ditelepon (Tk)
    'Dim P2 As Integer 'Jumlah Pelanggan Yang Diundang (P2= DM + SMS + E + Tlp)
    KPI_7a.Caption = KPI_5b.Caption
    KPI_7b.Caption = KPI_6b.Caption
    KPI_7.Caption = "0 %"
    KPI_7a.Caption = KPI_7a.Caption
    KPI_7b.Caption = KPI_7b.Caption
    KPI_7.Caption = CStr(Round((CInt(KPI_7a.Caption) / CInt(KPI_7b.Caption)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_8()
    On Error Resume Next
    Dim A1 As Integer 'Jumlah Pelanggan Yang Appoinment (A1)
    'Dim Tt As Integer 'Jumlah Telepon Terhubung (Tt)
    Set rs1 = obj.KPI_8(Bln, Thn)
    If rs1.RecordCount <> 0 Then A1 = rs1.Fields(0).value
    KPI_8.Caption = "0 %"
    KPI_8a.Caption = A1
    KPI_8b.Caption = KPI_5a.Caption
    KPI_8.Caption = CStr(Round((CInt(A1) / CInt(KPI_8b.Caption)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_9()
    On Error Resume Next
    Dim Pd As Integer 'Jumlah Pelanggan Yang Datang Setelah Diundang (Pd)
    Dim P2 As Integer 'Jumlah Pelanggan Yang Diundang (P2)
    Set rs1 = obj.KPI_9(Bln, Thn)
    If rs1.RecordCount <> 0 Then Pd = rs1.Fields(0).value
    KPI_9.Caption = "0 %"
    KPI_9a.Caption = Pd
    KPI_9b.Caption = KPI_7b.Caption
    KPI_9.Caption = CStr(Round((CInt(Pd) / CInt(KPI_9b.Caption)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_10()
    Dim A1 As Integer 'Jumlah app. Karena aktifitas MRS (A1)
    Dim A2 As Integer 'Jumlah total app. (A2)
'    Set rs1 = obj.KPI_10(True, Bln, Thn)
'    If rs1.RecordCount <> 0 Then A1 = rs1.Fields(0).value
    Set rs1 = obj.KPI_10(False, Bln, Thn)
    If rs1.RecordCount <> 0 Then A2 = rs1.Fields(0).value
    KPI_10.Caption = "0 %"
    KPI_10a.Caption = KPI_8a.Caption 'A1
    KPI_10b.Caption = A2
    If KPI_10a.Caption = "0" Or KPI_10b.Caption = "0" Then Exit Sub
    KPI_10.Caption = CStr(Round((CInt(KPI_8a.Caption) / CInt(KPI_10b.Caption)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_11()
    Dim Pd As Integer 'Jumlah unit SBE karena aktifitas MRS (Pd)
    Dim Pt As Integer 'Jumlah total unit SBE (Pt)
    Set rs1 = obj.KPI_11(Bln, Thn)
    If rs1.RecordCount <> 0 Then Pd = rs1.Fields(0).value
    Set rs2 = obj.KPI_11b(Bln, Thn)
    If rs2.RecordCount <> 0 Then Pt = rs2.Fields(0).value
    KPI_11.Caption = "0 %"
    KPI_11a.Caption = Pd
    KPI_11b.Caption = Pt
    If KPI_11a.Caption = "0" Or KPI_11b.Caption = "0" Then Exit Sub
    KPI_11.Caption = CStr(Round((CInt(Pd) / CInt(Pt)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_12()
    Dim Uk As Integer 'Jumlah unit yang telah dikirimkan kepada pelanggan (N-1) (Uk)
    Dim Us As Integer 'Jumlah pelanggan yang servis 1.000 dari penjualan unit N-1 (Us)
    Set rs1 = obj.KPI_12(Bln, Thn)
    If rs1.RecordCount <> 0 Then Uk = rs1.Fields(0).value
    Set rs2 = obj.KPI_12b(Bln, Thn)
    If rs2.RecordCount <> 0 Then Us = rs2.Fields(0).value
    KPI_12.Caption = "0 %"
    KPI_12a.Caption = Uk
    KPI_12b.Caption = Us
    If KPI_12a.Caption = "0" Or KPI_12b.Caption = "0" Then Exit Sub
    KPI_12.Caption = CStr(Round((CInt(Us) / CInt(Uk)) * 100, 2)) & " %"
End Sub
Private Sub SetKPI_13()
   '"Not Available !"
End Sub

Private Sub thnPlus_Click()
    Thn = Thn + 1
    nTahun.Caption = "Periode  :  " & Thn
    Call cmdProcess_Click
End Sub

Private Sub thnMinus_Click()
    Thn = Thn - 1
    If Thn < 2009 Then
        Thn = 2009
    Else
        nTahun.Caption = "Periode  :  " & Thn
        Call cmdProcess_Click
    End If
End Sub

Private Sub blnMinus_Click()
    Bln = Bln - 1
    If Bln = -1 Then Bln = 12
    nBulan.Caption = Bulan(Bln)
    Call cmdProcess_Click
End Sub

Private Sub blnPlus_Click()
    Bln = Bln + 1
    If Bln = 13 Then Bln = 0
    nBulan.Caption = Bulan(Bln)
    Call cmdProcess_Click
End Sub

Private Sub VScrollKPI_Change()
    Me.qKPI.Top = -(VScrollKPI.value)
    Me.sKPI.Top = -(VScrollKPI.value)
End Sub
Private Sub VScrollKPI_Scroll()
    Me.qKPI.Top = -(VScrollKPI.value)
    Me.sKPI.Top = -(VScrollKPI.value)
End Sub

Private Sub CmdOpen_Click()
    If PictGr.Visible = False Then
        PictGr.Visible = True
        Me.Dt0.Text = (IsNow)
        Me.Dt1.Text = (IsNow)
        Me.cboMax.Text = "100000 Km"
        Me.cboType.Text = "10.000 KM"
        Call SetRatio
    Else
        PictGr.Visible = False
    End If
End Sub

Private Sub SetRatio()
    Dim Kelipatan As String
    If Left(cboType.Text, 1) = "5" Then
        Kelipatan = "5000"
    Else
        Kelipatan = "0000"
    End If
    '-----------------------------------------------------------
    'UNDANG ----------------------------------------------------
    Set RsUnd = New ADODB.Recordset
    tblUnd True
    
    Dim rsU_spk As ADODB.Recordset
    Set rsU_spk = obj.GraphUndSPK(Dt0.value, Dt1.value)
        RsUnd.AddNew
        RsUnd!KM = Trim(rsU_spk.Fields(0).value)
        RsUnd!Jml = Val(rsU_spk.Fields(1).value)
    
    Dim rsU As ADODB.Recordset
    Set rsU = obj.GraphUnd(Kelipatan, Dt0.value, Dt1.value, Val(cboMax.Text) + 1)
    If rsU.RecordCount <> 0 Then
        rsU.MoveFirst
        Do Until rsU.EOF
            RsUnd.AddNew
            RsUnd!KM = Trim(Val(rsU.Fields(0).value) / 1000) & " K"
            RsUnd!Jml = Val(rsU.Fields(1).value)
            rsU.MoveNext
        Loop
    End If
    Me.OptD(0).value = True
    '-----------------------------------------------------------
    'DATANG ----------------------------------------------------
    Set RsDat = New ADODB.Recordset
    tblDat True
    
    Dim rsD As ADODB.Recordset
    Set rsD = obj.GraphDat(Kelipatan, Dt0.value, Dt1.value, Val(cboMax.Text) + 1)
    If rsD.RecordCount <> 0 Then
        rsD.MoveFirst
        Do Until rsD.EOF
            RsDat.AddNew
            RsDat!KM = Trim(Val(rsD.Fields(0).value) / 1000) & " K"
            RsDat!Jml = Val(rsD.Fields(1).value)
            rsD.MoveNext
        Loop
    End If
    
    '-----------------------------------------------------------
    'RATIO  ----------------------------------------------------
    Set RsRat = New ADODB.Recordset
    tblRat True
    
    Dim rs1U As ADODB.Recordset
    Dim JmlU As Double
    Set rs1U = obj.GraphRatio1000U(Dt0.value, Dt1.value)
    JmlU = rs1U.RecordCount
  
    Dim rs1D As ADODB.Recordset
    Dim JmlD As Double
    Set rs1D = obj.GraphRatio1000D(Dt0.value, Dt1.value)
    JmlD = 1
    JmlU = 1
    JmlD = rs1D.RecordCount
    
    Dim R1 As Double
    R1 = (JmlD * 100) / JmlU
    
        RsRat.AddNew
        RsRat!KM = "1 K"
        RsRat!Persen = Val(Round(R1, 2))
        RsRat!Jml = Val(Round(JmlD, 2))
    
    Dim rsR As ADODB.Recordset
    Set rsR = obj.GraphRat(Kelipatan, Dt0.value, Dt1.value, Val(cboMax.Text) + 1)
    If rsR.RecordCount <> 0 Then
        rsR.MoveFirst
        Do Until rsR.EOF
            RsRat.AddNew
            RsRat!KM = Trim(Val(rsR.Fields(0).value) / 1000) & " K"
            
            RsUnd.MoveFirst
            Do Until RsUnd.EOF
                If RsUnd.Fields(0).value = Trim(Val(rsR.Fields(0).value) / 1000) & " K" Then
                    RsRat!Persen = (Val(rsR.Fields(1).value) * 100) / Val(RsUnd.Fields(1).value)
                    Exit Do
                End If
                RsUnd.MoveNext
            Loop
            RsRat!Jml = Val(rsR.Fields(1).value)
            rsR.MoveNext
        Loop
    End If
    Call OptD_set
End Sub

Private Sub OptD_Click(Index As Integer)
    Call OptD_set
End Sub

Private Sub OptD_set()
    Me.GridR.Visible = False
    If Me.OptD(0).value = True Then
        Set Me.MsGraf.DataSource = RsUnd
        Set Me.Grid.DataSource = RsUnd
    ElseIf Me.OptD(1).value = True Then
        Set Me.MsGraf.DataSource = RsDat
        Set Me.Grid.DataSource = RsDat
    ElseIf Me.OptD(2).value = True Then
        Set Me.MsGraf.DataSource = RsRat
        Set Me.GridR.DataSource = RsRat
        Me.GridR.Visible = True
    End If
End Sub


Private Sub tblUnd(Optional Edit As Boolean = False)
    On Error GoTo Vb_Error
    If RsUnd.State Then RsUnd.Close
    With RsUnd
        .Fields.Append "KM", adVarChar, 15
        .Fields.Append "Jml", adDouble, 9
        .Open
    End With
    Exit Sub
Vb_Error:
    err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub tblDat(Optional Edit As Boolean = False)
    On Error GoTo Vb_Error
    If RsDat.State Then RsDat.Close
    With RsDat
        .Fields.Append "KM", adVarChar, 15
        .Fields.Append "Jml", adDouble, 9
        .Open
    End With
    Exit Sub
Vb_Error:
    err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub tblRat(Optional Edit As Boolean = False)
    On Error GoTo Vb_Error
    If RsRat.State Then RsRat.Close
    With RsRat
        .Fields.Append "KM", adVarChar, 15
        .Fields.Append "Persen", adDouble, 9
        .Fields.Append "Jml", adDouble, 9
        .Open
    End With
    Exit Sub
Vb_Error:
    err.Raise err.Number, err.Source, err.Description
End Sub

'-------------------------------------  EXCELLLLLLLLL ----------------------------------------

Private Sub BackUpMyBase(onVal As String)
On Error GoTo errHandler
    'Dim oCnn As ADODB.Connection
    'Dim oApp As Excel.Application
    'Dim oWB As Excel.Workbook
    Dim Er As New ErrInfo
    'Dim i As Integer

    
'    Set oApp = New Excel.Application
'    oApp.Visible = False
'    Set oWB = oApp.Workbooks.Add
    
    createTbl_XLS True
    Call InsertRecordset
    
    GridExport.ExportToFile onVal, True
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            
            Beep
            Er.Show (vbModal)
            
            
'    For i = 0 To rsXLS.Fields.Count - 1
'        oWB.Sheets(1).Cells(1, i + 1).value = rsXLS.Fields(i).Name
'    Next
'
'    oWB.Sheets(1).Range("1:1").Font.Bold = True
'
'    oWB.Sheets(1).Cells(2, 1).CopyFromRecordset rsXLS
'    oWB.Close SaveChanges:=True, FileName:=onVal
'    Set oWB = Nothing
'    oApp.Quit
'    Set oApp = Nothing
    Exit Sub
    
errHandler:
    Exit Sub
    err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub createTbl_XLS(Optional Edit As Boolean = False)
On Error GoTo errIsExist
    If rsXLS.State Then rsXLS.Close
    With rsXLS
        .Fields.Append "XXX", adVarChar, 100
        .Fields.Append "YYY", adVarChar, 100
        .Fields.Append "ZZZ", adVarChar, 100
        .Open
    End With
    Set GridExport.DataSource = rsXLS
errIsExist:
End Sub

Private Sub InsertRecordset()
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ""
    rsXLS!ZZZ = ""
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = nTahun.Caption
    rsXLS!YYY = ""
    rsXLS!ZZZ = nBulan.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ""
    rsXLS!ZZZ = ""
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label1.Caption
    rsXLS!YYY = ""
    rsXLS!ZZZ = ""
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label3.Caption
    rsXLS!YYY = ket1.Caption
    rsXLS!ZZZ = KPI_1.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket1a.Caption
    rsXLS!ZZZ = KPI_1a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket1b.Caption
    rsXLS!ZZZ = KPI_1b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket1c.Caption
    rsXLS!ZZZ = KPI_1c.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label8.Caption
    rsXLS!YYY = ket2.Caption
    rsXLS!ZZZ = KPI_2.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket2a.Caption
    rsXLS!ZZZ = KPI_2a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket2b.Caption
    rsXLS!ZZZ = KPI_2b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket2c.Caption
    rsXLS!ZZZ = KPI_1c.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label12.Caption
    rsXLS!YYY = ""
    rsXLS!ZZZ = ""
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label9.Caption
    rsXLS!YYY = ket3Sales.Caption
    rsXLS!ZZZ = KPI_3sales.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket3salesa.Caption
    rsXLS!ZZZ = KPI_3salesA.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket3salesb.Caption
    rsXLS!ZZZ = KPI_3salesB.Caption
    rsXLS.Update
    
    rsXLS.AddNew
    rsXLS!XXX = Label17.Caption
    rsXLS!YYY = ""
    rsXLS!ZZZ = ""
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label18.Caption
    rsXLS!YYY = ket4.Caption
    rsXLS!ZZZ = KPI_4.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket4a.Caption
    rsXLS!ZZZ = KPI_4a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket4b.Caption
    rsXLS!ZZZ = KPI_4b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label22.Caption
    rsXLS!YYY = ket5.Caption
    rsXLS!ZZZ = KPI_5.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket5a.Caption
    rsXLS!ZZZ = KPI_5a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket5b.Caption
    rsXLS!ZZZ = KPI_5b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label26.Caption
    rsXLS!YYY = ket6.Caption
    rsXLS!ZZZ = KPI_6.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket6a.Caption
    rsXLS!ZZZ = KPI_6a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket6b.Caption
    rsXLS!ZZZ = KPI_6b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label13.Caption
    rsXLS!YYY = ket3.Caption
    rsXLS!ZZZ = KPI_3.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket3a.Caption
    rsXLS!ZZZ = KPI_3a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket3b.Caption
    rsXLS!ZZZ = KPI_3b.Caption
    rsXLS.Update
    
    rsXLS.AddNew
    rsXLS!XXX = Label30.Caption
    rsXLS!YYY = ket7.Caption
    rsXLS!ZZZ = KPI_7.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket7a.Caption
    rsXLS!ZZZ = KPI_7a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket7b.Caption
    rsXLS!ZZZ = KPI_7b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label34.Caption
    rsXLS!YYY = ket8.Caption
    rsXLS!ZZZ = KPI_8.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket8a.Caption
    rsXLS!ZZZ = KPI_8a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket8b.Caption
    rsXLS!ZZZ = KPI_8b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label38.Caption
    rsXLS!YYY = ket9.Caption
    rsXLS!ZZZ = KPI_9.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket9a.Caption
    rsXLS!ZZZ = KPI_9a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket9b.Caption
    rsXLS!ZZZ = KPI_9b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label42.Caption
    rsXLS!YYY = ""
    rsXLS!ZZZ = ""
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label43.Caption
    rsXLS!YYY = ket10.Caption
    rsXLS!ZZZ = KPI_10.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket10a.Caption
    rsXLS!ZZZ = KPI_10a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket10b.Caption
    rsXLS!ZZZ = KPI_10b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label47.Caption
    rsXLS!YYY = ket11.Caption
    rsXLS!ZZZ = KPI_11.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket11a.Caption
    rsXLS!ZZZ = KPI_11a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket11b.Caption
    rsXLS!ZZZ = KPI_11b.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = Label51.Caption
    rsXLS!YYY = ket12.Caption
    rsXLS!ZZZ = KPI_12.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket12a.Caption
    rsXLS!ZZZ = KPI_12a.Caption
    rsXLS.Update
    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ket12b.Caption
    rsXLS!ZZZ = KPI_12b.Caption
    rsXLS.Update

    rsXLS.AddNew
    rsXLS!XXX = ""
    rsXLS!YYY = ""
    rsXLS!ZZZ = ""
    rsXLS.Update
End Sub

