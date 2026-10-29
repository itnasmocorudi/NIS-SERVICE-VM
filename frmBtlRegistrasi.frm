VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmBtlRegistrasi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pembatalan Registrasi"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   0
      Picture         =   "frmBtlRegistrasi.frx":0000
      ScaleHeight     =   10695
      ScaleWidth      =   495
      TabIndex        =   63
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "   SERVICE    "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6735
      Left            =   480
      TabIndex        =   34
      Top             =   2640
      Width           =   10335
      Begin TrueOleDBGrid70.TDBGrid JobGrid 
         Height          =   1695
         Left            =   120
         TabIndex        =   35
         Top             =   960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
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
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
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
         Caption         =   "JOB ORDER"
         MultipleLines   =   0
         CellTipsWidth   =   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
         _StyleDefs(58)  =   "Named:id=33:Normal"
         _StyleDefs(59)  =   ":id=33,.parent=0"
         _StyleDefs(60)  =   "Named:id=34:Heading"
         _StyleDefs(61)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(62)  =   ":id=34,.wraptext=-1"
         _StyleDefs(63)  =   "Named:id=35:Footing"
         _StyleDefs(64)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=36:Selected"
         _StyleDefs(66)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(67)  =   "Named:id=37:Caption"
         _StyleDefs(68)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(69)  =   "Named:id=38:HighlightRow"
         _StyleDefs(70)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(71)  =   "Named:id=39:EvenRow"
         _StyleDefs(72)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(73)  =   "Named:id=40:OddRow"
         _StyleDefs(74)  =   ":id=40,.parent=33"
         _StyleDefs(75)  =   "Named:id=41:RecordSelector"
         _StyleDefs(76)  =   ":id=41,.parent=34"
         _StyleDefs(77)  =   "Named:id=42:FilterBar"
         _StyleDefs(78)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid partGrid 
         Height          =   1695
         Left            =   120
         TabIndex        =   36
         Top             =   2700
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2990
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
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
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
         Caption         =   "PART ORDER"
         MultipleLines   =   0
         CellTipsWidth   =   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid subletGrid 
         Height          =   1575
         Left            =   120
         TabIndex        =   37
         Top             =   4440
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   2778
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
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "SUBLET ORDER"
         MultipleLines   =   0
         CellTipsWidth   =   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=975,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Named:id=33:Normal"
         _StyleDefs(47)  =   ":id=33,.parent=0"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   ":id=34,.wraptext=-1"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=37:Caption"
         _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(57)  =   "Named:id=38:HighlightRow"
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   28
         Left            =   840
         TabIndex        =   38
         Top             =   6255
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   27
         Left            =   5640
         TabIndex        =   39
         Top             =   525
         Width           =   4575
         _ExtentX        =   8070
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   22
         Left            =   2160
         TabIndex        =   40
         Top             =   525
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   "CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   21
         Left            =   2160
         TabIndex        =   41
         Top             =   240
         Width           =   1050
         _ExtentX        =   1852
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   23
         Left            =   3960
         TabIndex        =   42
         Top             =   240
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   24
         Left            =   3960
         TabIndex        =   43
         Top             =   525
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   25
         Left            =   5640
         TabIndex        =   44
         Top             =   240
         Width           =   1290
         _ExtentX        =   2275
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
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   26
         Left            =   6960
         TabIndex        =   45
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   20
         Left            =   120
         TabIndex        =   46
         Top             =   525
         Width           =   1200
         _ExtentX        =   2117
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
      Begin ucXPButton.XPButton TblBatal 
         Height          =   495
         Left            =   6840
         TabIndex        =   64
         Top             =   6120
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   2
         TX              =   "&b"
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
         MICON           =   "frmBtlRegistrasi.frx":2166
         PICN            =   "frmBtlRegistrasi.frx":2182
         PICH            =   "frmBtlRegistrasi.frx":26B7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton TblKeluar 
         Height          =   495
         Left            =   9000
         TabIndex        =   65
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
         MICON           =   "frmBtlRegistrasi.frx":2C6E
         PICN            =   "frmBtlRegistrasi.frx":2C8A
         PICH            =   "frmBtlRegistrasi.frx":315E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   10200
         Y1              =   6040
         Y2              =   6040
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
         Height          =   195
         Left            =   120
         TabIndex        =   54
         Top             =   6300
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Suplay Slip"
         Height          =   195
         Left            =   4695
         TabIndex        =   53
         Top             =   570
         Width           =   780
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mech"
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
         Left            =   3480
         TabIndex        =   52
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM"
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
         Left            =   1845
         TabIndex        =   51
         Top             =   285
         Width           =   210
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. No."
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
         Left            =   1485
         TabIndex        =   50
         Top             =   570
         Width           =   600
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T SA."
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
         Left            =   3480
         TabIndex        =   49
         Top             =   570
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe WO"
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
         Left            =   4800
         TabIndex        =   48
         Top             =   285
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         TabIndex        =   47
         Top             =   285
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmBtlRegistrasi.frx":36A9
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2370
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   8085
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   930
         TabIndex        =   2
         Top             =   510
         Width           =   3405
         _ExtentX        =   6006
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
         Left            =   930
         TabIndex        =   3
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   930
         TabIndex        =   4
         Top             =   1110
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   930
         TabIndex        =   5
         Top             =   1410
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   5190
         TabIndex        =   6
         Top             =   210
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   5190
         TabIndex        =   7
         Top             =   510
         Width           =   1935
         _ExtentX        =   3413
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
         Index           =   15
         Left            =   5190
         TabIndex        =   8
         Top             =   1110
         Width           =   2355
         _ExtentX        =   4154
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
         Index           =   14
         Left            =   5190
         TabIndex        =   9
         Top             =   810
         Width           =   2775
         _ExtentX        =   4895
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
         Index           =   13
         Left            =   7170
         TabIndex        =   10
         Top             =   510
         Width           =   795
         _ExtentX        =   1402
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   2520
         TabIndex        =   11
         Top             =   210
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   930
         TabIndex        =   12
         Top             =   210
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   5190
         TabIndex        =   13
         Top             =   1410
         Width           =   1410
         _ExtentX        =   2487
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   5190
         TabIndex        =   14
         Top             =   1695
         Width           =   2775
         _ExtentX        =   4895
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
         Index           =   11
         Left            =   6750
         TabIndex        =   15
         Top             =   195
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   930
         TabIndex        =   16
         Top             =   1700
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   930
         TabIndex        =   17
         Top             =   2000
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   29
         Left            =   5190
         TabIndex        =   18
         Top             =   1995
         Width           =   2775
         _ExtentX        =   4895
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
         Index           =   19
         Left            =   3645
         TabIndex        =   19
         Top             =   210
         Width           =   690
         _ExtentX        =   1217
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   18
         Left            =   5130
         TabIndex        =   61
         Top             =   2400
         Width           =   2775
         _ExtentX        =   4895
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
      Begin VB.Label Label7 
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
         Left            =   4440
         TabIndex        =   62
         Top             =   2445
         Width           =   540
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         Left            =   405
         TabIndex        =   30
         Top             =   1455
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
         Left            =   4395
         TabIndex        =   29
         Top             =   255
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
         Left            =   4620
         TabIndex        =   28
         Top             =   525
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
         Left            =   4560
         TabIndex        =   27
         Top             =   1125
         Width           =   480
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Lgn"
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
         Left            =   135
         TabIndex        =   26
         Top             =   240
         Width           =   660
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
         Left            =   4395
         TabIndex        =   25
         Top             =   1455
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
         Left            =   4530
         TabIndex        =   24
         Top             =   1740
         Width           =   510
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO TAM"
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
         Left            =   6015
         TabIndex        =   23
         Top             =   240
         Width           =   585
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
         Left            =   465
         TabIndex        =   22
         Top             =   1740
         Width           =   315
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
         Left            =   345
         TabIndex        =   21
         Top             =   2040
         Width           =   435
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOU"
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
         Left            =   4695
         TabIndex        =   20
         Top             =   2040
         Width           =   345
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2600
      TabIndex        =   55
      Top             =   120
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
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   105
      Width           =   1215
      _ExtentX        =   2143
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
      Index           =   1
      Left            =   7080
      TabIndex        =   56
      Top             =   83
      Width           =   1200
      _ExtentX        =   2117
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
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9375
      Left            =   0
      Picture         =   "frmBtlRegistrasi.frx":3749
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label26 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal"
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
      Left            =   6375
      TabIndex        =   60
      Top             =   135
      Width           =   570
   End
   Begin VB.Label LblReturn 
      BackStyle       =   0  'Transparent
      Caption         =   "Return Job"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2880
      TabIndex        =   59
      Top             =   120
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. WO"
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
      Left            =   225
      TabIndex        =   58
      Top             =   150
      Width           =   945
   End
   Begin VB.Label lblClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CLOSING"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   600
      Left            =   8640
      TabIndex        =   57
      Top             =   1080
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frmBtlRegistrasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 18 April 2009
'Created Coding  : Thania
'Tgl. Modifikasi : 16 Feb 2011
'=======================================


Dim rsTemp As New ADODB.Recordset
'Dim rsTemp2 As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsPart As New ADODB.Recordset
Dim rsSublet As New ADODB.Recordset
Dim nilaiRupiah As Double

Private Sub hapusForm()
    For i = 0 To 29
        Select Case i
            Case 10, 21
                Text(i) = 0
            Case 1, 11, 20
                Text(i) = "99-99-9999"
            Case 16
                Text(i) = "CC-9999-CCC"
            Case Else
                Text(i) = ""
        End Select
    Next
    
    If rsJob.State Then rsJob.Close
    Set rsJob = Nothing
    If rsPart.State Then rsPart.Close
    Set rsPart = Nothing
    If rsSublet.State Then rsSublet.Close
    Set rsSublet = Nothing
    lblClose.Visible = False
End Sub

Private Sub doCreatePart()
    If rsPart.State Then rsPart.Close
    Set rsPart = Nothing
    With rsPart
        'If .State Then .Close
        .Fields.Append "KODEITEM", adVarChar, 50
        .Fields.Append "NAMAITEM", adVarChar, 200
        .Fields.Append "QTY", adInteger
        .Open
    End With
End Sub

Private Sub doHeadPart()
    
    With partGrid
        .Columns(0).Width = 1800
        .Columns(0).Caption = "KODE PART"
        .Columns(0).DataField = "KODEITEM"
        .Columns(1).Width = 2500
        .Columns(1).Caption = "NAMA PART"
        .Columns(1).DataField = "NAMAITEM"
        .Columns(2).Width = 800
        .Columns(2).Caption = "QTY"
        .Columns(2).DataField = "QTY"
    End With
End Sub

Private Sub doCreateJob()
    If rsJob.State Then rsJob.Close
    'Set rsJob = Nothing
    With rsJob
        'If .State Then .Close
        .Fields.Append "JOB", adVarChar, 50
        .Fields.Append "RATE", adDouble
        .Fields.Append "GRP", adVarChar, 6
        .Fields.Append "HRG", adDouble
        .Fields.Append "fOther", adBoolean
        .Open
    End With
End Sub

Private Sub doHeadJob()
    
    With jobGrid
        .Columns(0).Width = 2500
        .Columns(0).Caption = "JOB ORDER"
        .Columns(0).DataField = "JOB"
        .Columns(1).Width = 800
        .Columns(1).Caption = "RATE"
        .Columns(1).DataField = "RATE"
        .Columns(2).Width = 0
        .Columns(2).Caption = "GROUP"
        .Columns(2).DataField = "GRP"
        .Columns(2).Visible = False
        .Columns(3).Width = 0
        .Columns(3).Caption = "HARGA"
        .Columns(3).DataField = "HRG"
        .Columns(3).Visible = False
        .Columns(4).Width = 0
        .Columns(4).Caption = "FOTHER"
        .Columns(4).DataField = "fOther"
        .Columns(4).Visible = False
    End With
End Sub

Private Sub doHeadSublet()

    With subletGrid
        .Columns(0).Width = 1000
        .Columns(0).Caption = "NO NOTA"
        .Columns(0).DataField = "NONOTA"
        .Columns(1).Width = 2000
        .Columns(1).Caption = "KODE ITEM"
        .Columns(1).DataField = "KODEITEM"
        .Columns(2).Width = 2500
        .Columns(2).Caption = "NAMA ITEM"
        .Columns(2).DataField = "NAMABARANG"
        .Columns(3).Width = 1000
        .Columns(3).Caption = "SATUAN"
        .Columns(3).DataField = "KODESATUAN"
        .Columns(4).Width = 1000
        .Columns(4).Caption = "QTY"
        .Columns(4).DataField = "QTY"
        .Columns(4).NumberFormat = "#,#0"
        .Columns(5).Width = 2000
        .Columns(5).Caption = "HARGA"
        .Columns(5).DataField = "HARGA"
        .Columns(5).NumberFormat = "#,#0"
    End With
End Sub

Private Sub panggilWO(nWO As String)
      Dim rsHlp As ADODB.Recordset

1         Set rsTemp = objO.FetchData("SELECT DISTINCT A.*, B.GSM, B.CDMA, C.SLIP FROM JOBDTA A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS " _
                                   & "INNER JOIN T_INVOICE C WITH(NOLOCK) ON A.NOINV = LEFT(C.NOINV,8) WHERE A.NO_wo = '" & nWO & "' AND A.DELIV <> 'BTL' AND A.NOINV <> ''")
2         If Not rsTemp.EOF Then
3             If rsTemp("RPR_COND") <> "K" Then
4                 lblClose.Visible = False
5             Else
6                 lblClose.Visible = True
7             End If
              
              ' TAMPIL DATA CUSTOMER
8             Text(2) = rsTemp("FS_KD_PLG")
9             Text(3) = rsTemp("FS_KD_WIL")
10            Text(4) = rsTemp("CUSTOMER")
11            Text(5) = Left(rsTemp("CUST_ADRES"), 60)
12            Text(6) = rsTemp("KOTA")
13            Text(7) = rsTemp("TELPON")
14            Text(8) = rsTemp("GSM")
15            Text(9) = rsTemp("CDMA")
16            Text(10) = rsTemp("FS_TH_MOBIL")
17            Text(11).Mask = ""
18            Text(11) = rsTemp("DOTAM")
19            Text(12) = rsTemp("MDL")
20            Text(13) = rsTemp("MODEL")
21            Text(14) = rsTemp("FMDL")
22            Text(15) = rsTemp("COLOR")
23            Text(16) = rsTemp("NO_POLISI")
24            Text(17) = rsTemp("NO_CHASSIS")
25            Text(18) = rsTemp("NO_ENGINE")
              
              ' TAMPIL DATA SERVICE
26            Text(20) = TglDDMMYYYY(rsTemp("DATE"))
27            Text(21) = rsTemp("KM")
28            Text(22) = rsTemp("NOINV")
29            Text(23) = rsTemp("MECH")
30            Text(24) = rsTemp("FS_LOG_ADD")
31            Text(25) = IIf(rsTemp("TPSVC") <> "", rsTemp("KODETIPE"), "REGULAR")
32            Text(29) = cekMOU(rsTemp("MDL"), Nul(rsTemp("NO_CHASSIS")))
              
               
33            If Trim(rsTemp("TPSVC")) = "W" Then
                  
34                If Log.RPLHistori = 0 Then
35                    Set rsHlp = objO.FetchData("SELECT * FROM HO_WARRANTY WITH(NOLOCK) WHERE ID_WARRANTY = '" & rsTemp("KODETIPE") & "' and FS_LOG_DEL = ''")
36                Else
37                    Set rsHlp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH(NOLOCK) WHERE ID_WARRANTY = '" & rsTemp("KODETIPE") & "' and FS_LOG_DEL = ''")
38                End If
39                If Not rsHlp.EOF Then
40                    Text(26) = rsHlp("WARRANTY")
41                End If
42            ElseIf Trim(rsTemp("TPSVC")) = "P" Then
                  
43                If Log.RPLHistori = 0 Then
44                    Set rsHlp = objO.FetchData("SELECT * FROM HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & rsTemp("KODETIPE") & "' and FS_LOG_DEL = ''")
45                Else
46                    Set rsHlp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_PROGRAM WITH(NOLOCK) WHERE ID_PROGRAM = '" & rsTemp("KODETIPE") & "' and FS_LOG_DEL = ''")
47                End If
48                If Not rsHlp.EOF Then
49                    Text(26) = rsHlp("PROGRAM")
50                End If
51            End If
                      
              ' TAMPIL JOB ORDER
52            myChassis = Trim(Text(17))
              Dim nDisN As Double
53            If CexChassis_Discon(Trim(Text(13))) = True Then
54                nDisN = (100 - nDiscNRM) / 100
55            Else
56                nDisN = (100 - nDiscNON) / 100
57            End If
58            If rsJob.State Then rsJob.Close
59            For i = 1 To 7
60                hlp = "REP" & i
61                hlp2 = "RTE" & i
62                hlp3 = "KLPKERJA" & i
                  
63                If rsTemp(hlp) <> "" Then
                      
64                    If Log.RPLHistori = 0 Then
65                        Set rsHlp = objO.FetchData("SELECT FS_ITEM FROM JOBOrder WITH(NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "'")
66                    Else
67                        Set rsHlp = objO.FetchData("SELECT FS_ITEM FROM [" & dbReplikasiHO & "]..JOBOrder WITH(NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "'")
68                    End If
69                    If Not rsHlp.EOF Then
70                        fOth = False
71                    Else
72                        fOth = True
73                    End If
                      
74                    myChassis = Trim(Text(17))
                      Dim HrgCar As Double
                      
75                    If Log.RPLHistori = 0 Then
76                        Set rsHlp = objO.FetchData("SELECT price FROM carclass WITH(NOLOCK) WHERE CLASSCODE = '" & rsTemp("MODEL") & "'")
77                    Else
78                        Set rsHlp = objO.FetchData("SELECT price FROM [" & dbReplikasiHO & "]..carclass WITH(NOLOCK) WHERE CLASSCODE = '" & rsTemp("MODEL") & "'")
79                    End If
80                    If Not rsHlp.EOF Then
81                        HrgCar = CDbl(rsHlp("PRICE")) * nDisN
82                    Else
83                        HrgCar = 0
84                    End If
                      
85                    With rsJob
86                    If .State = adStateClosed Then doCreateJob
87                       .AddNew
88                       !job = rsTemp(hlp)
89                       !Rate = rsTemp(hlp2)
90                       !grp = rsTemp(hlp3)
91                       !hrg = HrgCar * CDbl(rsTemp(hlp2))
92                       !fOther = fOth
93                       .Update
94                    End With
95                End If
96            Next
97            Set jobGrid.DataSource = rsJob
98            Call doHeadJob
              
              ' TAMPIL PART ORDER
99            Set rsHlp = objO.FetchData("SELECT KODE_ITEM, NAMA_ITEM, PARTBAWA, QTY, isnull(fs_kd_oli,'') as OLI, BAWA FROM jobpart a WITH(NOLOCK) LEFT JOIN T_oli b WITH(NOLOCK) ON LEFT(a.kode_item,1) = b.fs_kd_Oli AND B.TIPE = 'O' WHERE no_WO = '" & nWO & "' AND no_Chassis = '" & rsTemp("NO_CHASSIS") & "' AND GANTI = 'Y'")
100           If Not rsHlp.EOF Then
101               Do While Not rsHlp.EOF
102                   With rsPart
103                   If .State = adStateClosed Then doCreatePart
104                       .AddNew
105                       !KodeItem = IIf(rsHlp("BAWA") = "Y", "BAWA SENDIRI", rsHlp("KODE_ITEM"))
106                       !NAMAITEM = IIf(rsHlp("BAWA") = "Y", rsHlp("PARTBAWA"), rsHlp("NAMA_ITEM"))
107                       !qty = rsHlp("QTY")
108                       .Update
109                   End With
110                   rsHlp.MoveNext
111               Loop
112           End If
113           Set partGrid.DataSource = rsPart
114           Call doHeadPart
              
              ' TAMPIL SUBLET ORDER
115           Set rsSublet = objO.FetchData("SELECT NONOTA, KODEITEM, NAMABARANG, KODESATUAN, QTY, PRICE as HARGA FROM jobSublet WITH(NOLOCK) WHERE no_WO = '" & nWO & "' AND no_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
116           If Not rsSublet.EOF Then
117               Set subletGrid.DataSource = rsSublet
118               Call doHeadSublet
119           End If
              
120           nilaiRupiah = 0
121           Do While Not rsTemp.EOF
122               If Right(rsTemp("NOINV"), 2) <> "-B" Then
123                   nilaiRupiah = nilaiRupiah + CDbl(rsTemp("TOTAL"))
124               End If
125               rsTemp.MoveNext
126           Loop
127       End If
End Sub

Private Sub Form_Load()
    Call hapusForm
    Text(1) = TglDDMMYYYY(Format(isnow, "mm/dd/yyyy"))
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub tblBatal_Click()

1         On Error GoTo vbError
2         Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlRegistrasi", "tblBatal", 0)

3         If Trim(Text(22)) <> "" And Trim(Text(0)) <> "" Then
4             erTranAll = False
5             Call objHT.OpenTransactionBengkel
              
6             Call objHT.InsertData_BengkelOnly("UPDATE CHASSINGFORM SET TANGGALREGISTRASI = NULL WHERE NOWO = '" & Trim(Text(0)) & "'")
7             Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTA SET RPR_COND = 'S' WHERE NO_WO = '" & Text(0) & "' AND NOINV = '" & Text(22) & "' AND DELIV <> 'BTL'")
8             Call objHT.UpdateData_BengkelOnly("UPDATE T_INVOICE SET RPR_COND = 'S', TGLKASIR = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "', KASIRADD = '" & Trim(gSA) & "', FS_KREDIT = 1 WHERE NO_WO = '" & Text(0) & "' AND NOINV LIKE '" & Text(22) & "%' AND DELIV <> 'BTL'")
9             Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET FLAG = 'I' WHERE KETERANGAN LIKE '%|" & Text(0) & "|%' AND TANGGAL = '" & Format(Text(20), "MM/DD/YYYY") & "'")
              
10            If Trim(Text(29)) <> "" Then
                  
11                If Log.RPLHistori = 0 Then
12                    Call objHT.InsertData_BengkelOnly("INSERT INTO MOUTRN VALUES ('" & Format(isnow, "MM/DD/YYYY") & "', '" & Text(29) & "', '" & Text(0) & "', '" & Text(22) & "', '" & Text(17) & "', '" & Text(16) & "', '" & gKdCbg & "', " & (CDbl(nilaiRupiah) * -1) & ", 'B','','01/01/3000')")
13                    Call objHT.UpdateData_BengkelOnly("UPDATE MOUHDR SET BAYAR = BAYAR - " & CDbl(nilaiRupiah) & " WHERE NO_MOU = '" & Text(29) & "'")
                  
14                Else
15                    Call objHT.InsertData_BengkelOnly("INSERT INTO [" & dbReplikasiHO & "]..MOUTRN (Tanggal , NO_MOU, NO_WO, noInv, NO_CHASSIS, NO_POLISI, FS_KD_DEALER, total, Status, FS_LOG_DEL, FD_TGL_DEL) VALUES ('" & Format(isnow, "MM/DD/YYYY") & "', '" & Text(29) & "', '" & Text(0) & "', '" & Text(22) & "', '" & Text(17) & "', '" & Text(16) & "', '" & gKdCbg & "', " & (CDbl(nilaiRupiah) * -1) & ", 'B','','01/01/3000')")
16                    Call objHT.UpdateData_BengkelOnly("UPDATE [" & dbReplikasiHO & "]..MOUHDR SET BAYAR = BAYAR - " & CDbl(nilaiRupiah) & " WHERE NO_MOU = '" & Text(29) & "'")
                  
17                End If
      '                Call objHT.InsertData_BengkelOnly("INSERT INTO MOUTRN VALUES ('" & Format(isnow, "MM/DD/YYYY") & "', '" & text(29) & "', '" & text(0) & "', '" & text(22) & "', '" & text(17) & "', '" & text(16) & "', '" & gKdCbg & "', " & (CDbl(nilaiRupiah) * -1) & ", 'B','','01/01/3000')")
      '                Call objHT.UpdateData_BengkelOnly("UPDATE MOUHDR SET BAYAR = BAYAR - " & CDbl(nilaiRupiah) & " WHERE NO_MOU = '" & text(29) & "'")
18            End If
              
19                If erTranAll = False Then
20                    Call objHT.CloseTransactionBengkel
21                    onTR = False
22                    Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlRegistrasi", "tblBatal", 1)
23                    Call hapusForm
24                Else
25                    Call objHT.CloseTransactionBengkel_err
26                    onTR = False
27                    Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlRegistrasi", "tblBatal", 2)
28                    MsgBox "Proses pembatalan Registrasi gagal." & vbCr & "Ulangi proses ", vbInformation, "Gagal Simpan"
29                    Exit Sub
30                End If
31        End If
32        Exit Sub
          
vbError:
33        Call objHT.CloseTransactionBengkel_err
34        onTR = False
35        Call saveLogMenu(gSA, ComptName, gVersi, "frmBtlRegistrasi", "tblBatal", 2)
36        MsgBox "Proses pembatalan Registrasi gagal." & vbCr & "Ulangi proses ", vbInformation, "Gagal Simpan"
37        Exit Sub
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
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
                gSearchOption = "BTLREG"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As New ADODB.Recordset

     
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM jobDta WITH(NOLOCK) WHERE no_wo = '" & Text(Index) & "' AND DELIV <> 'BTL' AND RPR_COND = 'K'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "BTLREG"
                    frmCari.Show (vbModal)
                Else
                    Call panggilWO(Text(Index))
                End If
            End If
    End Select
     
End Sub
