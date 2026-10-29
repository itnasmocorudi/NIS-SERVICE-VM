VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmInjectMEstimasi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Inject Data Customer"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14850
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20160
   ShowInTaskbar   =   0   'False
   Begin TrueOleDBGrid70.TDBGrid gridCustomer 
      Height          =   6855
      Left            =   480
      TabIndex        =   30
      Top             =   1080
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   12091
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
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
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   18
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=18"
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
      Splits(0)._ColumnProps(53)=   "Column(13).Width=3281"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(69)=   "Column(17).Width=3281"
      Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=3175"
      Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=107,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=104,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=105,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=106,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=103,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=100,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=101,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=102,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=99,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=96,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=97,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=98,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=95,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=92,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=93,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=94,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=91,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=88,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=89,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=90,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=86,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=82,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=79,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=80,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=81,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=78,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=75,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=76,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=77,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=66,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=62,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=59,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=60,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=61,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=58,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=55,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=56,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=57,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=54,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=51,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=52,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=53,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=50,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=47,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=48,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=49,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=46,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=43,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=44,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=45,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=28,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=25,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=26,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=27,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=32,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=29,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=30,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=31,.parent=17"
      _StyleDefs(108) =   "Named:id=33:Normal"
      _StyleDefs(109) =   ":id=33,.parent=0"
      _StyleDefs(110) =   "Named:id=34:Heading"
      _StyleDefs(111) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(112) =   ":id=34,.wraptext=-1"
      _StyleDefs(113) =   "Named:id=35:Footing"
      _StyleDefs(114) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(115) =   "Named:id=36:Selected"
      _StyleDefs(116) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(117) =   "Named:id=37:Caption"
      _StyleDefs(118) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(119) =   "Named:id=38:HighlightRow"
      _StyleDefs(120) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(121) =   "Named:id=39:EvenRow"
      _StyleDefs(122) =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
      _StyleDefs(123) =   "Named:id=40:OddRow"
      _StyleDefs(124) =   ":id=40,.parent=33"
      _StyleDefs(125) =   "Named:id=41:RecordSelector"
      _StyleDefs(126) =   ":id=41,.parent=34"
      _StyleDefs(127) =   "Named:id=42:FilterBar"
      _StyleDefs(128) =   ":id=42,.parent=33"
      _StyleDefs(129) =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(130) =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(131) =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(132) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(133) =   "Named:id=87:SudahDitagih"
      _StyleDefs(134) =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(135) =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(136) =   ":id=87,.fontname=Tahoma"
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmInjectMEstimasi.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   22
      Top             =   0
      Width           =   450
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   14640
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
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
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin MSComCtl2.DTPicker tgl1 
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   165
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94633985
      CurrentDate     =   40142
   End
   Begin VB.Frame frmeKonfirm 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   4680
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   5120
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   503
         _Version        =   393216
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
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   1440
         TabIndex        =   27
         Top             =   1320
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
         MICON           =   "frmInjectMEstimasi.frx":8B94
         PICN            =   "frmInjectMEstimasi.frx":8BB0
         PICH            =   "frmInjectMEstimasi.frx":90C3
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
         Left            =   2640
         TabIndex        =   28
         Top             =   1320
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
         MICON           =   "frmInjectMEstimasi.frx":9665
         PICN            =   "frmInjectMEstimasi.frx":9681
         PICH            =   "frmInjectMEstimasi.frx":9BB6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   5160
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   -240
         X2              =   5160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: DATA PENAGIH ::."
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
         Left            =   1560
         TabIndex        =   29
         Top             =   60
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Penagih :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   705
         TabIndex        =   6
         Top             =   840
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Penagihan :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1260
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   5115
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   345
      Index           =   0
      Left            =   5355
      TabIndex        =   0
      Top             =   8115
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   609
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   15
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox9 
      Height          =   270
      Left            =   6195
      TabIndex        =   3
      Top             =   8145
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   300
      Index           =   3
      Left            =   6375
      TabIndex        =   7
      Top             =   8130
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   60
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSComCtl2.DTPicker tgl0 
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   165
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   99155969
      CurrentDate     =   40142
   End
   Begin ucXPButton.XPButton cmdRef 
      Height          =   495
      Left            =   5880
      TabIndex        =   21
      Top             =   105
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&r"
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
      MICON           =   "frmInjectMEstimasi.frx":A16D
      PICN            =   "frmInjectMEstimasi.frx":A189
      PICH            =   "frmInjectMEstimasi.frx":A705
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
      Left            =   13320
      TabIndex        =   23
      Top             =   8040
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
      MICON           =   "frmInjectMEstimasi.frx":AD3C
      PICN            =   "frmInjectMEstimasi.frx":AD58
      PICH            =   "frmInjectMEstimasi.frx":B22C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdKonfir 
      Height          =   495
      Left            =   10920
      TabIndex        =   24
      Top             =   8040
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
      MICON           =   "frmInjectMEstimasi.frx":B777
      PICN            =   "frmInjectMEstimasi.frx":B793
      PICH            =   "frmInjectMEstimasi.frx":BCA6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdExport 
      Height          =   495
      Left            =   12120
      TabIndex        =   26
      Top             =   8040
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&o"
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
      MICON           =   "frmInjectMEstimasi.frx":C248
      PICN            =   "frmInjectMEstimasi.frx":C264
      PICH            =   "frmInjectMEstimasi.frx":C7F9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: DAFTAR CUSTOMER ::."
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
      Left            =   6600
      TabIndex        =   25
      Top             =   780
      Width           =   2640
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SA"
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
      Index           =   1
      Left            =   4905
      TabIndex        =   20
      Top             =   8190
      Width           =   255
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
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
      Index           =   1
      Left            =   5235
      TabIndex        =   19
      Top             =   8190
      Width           =   60
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal "
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
      Index           =   0
      Left            =   600
      TabIndex        =   18
      Top             =   240
      Width           =   840
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":                              s/d"
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
      Index           =   0
      Left            =   1800
      TabIndex        =   17
      Top             =   240
      Width           =   2205
   End
   Begin VB.Label jmlRecord 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Data :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   13140
      TabIndex        =   15
      Top             =   120
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   9135
      Left            =   0
      Picture         =   "frmInjectMEstimasi.frx":CE34
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   3720
      TabIndex        =   12
      Top             =   8355
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label totTagih 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   4080
      TabIndex        =   11
      Top             =   8355
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label nTagih 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   2640
      TabIndex        =   10
      Top             =   8115
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Dipilih :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   600
      TabIndex        =   9
      Top             =   8115
      Width           =   1875
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   14055
   End
End
Attribute VB_Name = "frmInjectMEstimasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRecord As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns
Private Urut As String

Private Sub cmdBatal_Click()
    frmeKonfirm.Visible = False
End Sub

Private Sub cmdExport_Click()
On Error GoTo Err_Export
    Dim Er As New ErrInfo
    If rsRecord.RecordCount = 0 Then
        Er.HeadCap = "Export To Excel !"
        Er.InfoCap = "List Data Reminder Service Kosong !"
        Beep
        Er.show (vbModal)
        Exit Sub
    End If
    Dim MyPath As String
    Dim onVal As String
    MyPath = ""
    Dim f As New FormExplorer
    f.show (vbModal)
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    onVal = MyPath & f.FileName & ".XLS"
     gridCustomer.Columns(12).Visible = False
     gridCustomer.ExportToFile onVal, True
     gridCustomer.Columns(12).Visible = True
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            
            Beep
            Er.show (vbModal)
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdXls_Click")
Keluar:

End Sub

Private Sub CmdKeluar_Click()
   On Error GoTo cmdKeluar_Click_Error

    Unload Me

   On Error GoTo 0
   Exit Sub
cmdKeluar_Click_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdKeluar_Click of Form frmInjectMEstimasi")
End Sub

Private Sub cmdKonfir_Click()
1        On Error GoTo cmdKonfir_Click_Error
          Dim isBegin As Boolean
2         isBegin = False
3         If nTagih < 1 Then
4       MsgBox "Anda belum memilih Data yang akan di Inject !", vbExclamation + vbOKOnly
5       Exit Sub
6         ElseIf NulSID(Text(3)) = "" Then
7       MsgBox "Anda belum memilih SA yang akan di Inject !", vbExclamation + vbOKOnly
8       Exit Sub
9         End If
10      For Each col In gridCustomer.Columns
11          col.FilterText = ""
12      Next col
13      rsRecord.filter = ""
          
14        Call objHT.OpenTransactionBengkel
15        isBegin = True
          
16        With rsRecord
17            If Not .EOF Then
18                .MoveFirst
19                Do While Not .EOF
20                    If rsRecord("PILIH") Then
21                        sSQL = "INSERT INTO [M_ESTIMASIInject]([NoRangka],[RataKm]"
22                        sSQL = sSQL & " ,[JT_Tgl],[JT_Bln],[KmNext],[SBENext],[KmNow],[TglSvcNow],[NoInv],[NoWo]"
23                        sSQL = sSQL & " ,[NamaCust],[AlamatCust],[NoPolisi],[PSTN],[GSM],[CDMA],[LastSA],[SA]) VALUES "
24                        sSQL = sSQL & " ('" & NulSID(!noRangka) & "','" & NulSID(!RATAKM) & "'"
25                        sSQL = sSQL & " ,'" & Format(NulSID(!JT_Tgl), "yyyy-mm-dd") & "','" & Format(NulSID(!JT_Tgl), "yyyy-mm-dd") & "','" & NulSID(!KmNext) & "','" & NulSID(!SBNEXT) & "','" & NulSID(!KMNOW) & "','" & Format(NulSID(!TglSvcNow), "yyyy-mm-dd") & "','" & NulSID(!noInv) & "','" & NulSID(!noWO) & "'"
26                        sSQL = sSQL & " ,'" & NulSID(!namaCust) & "','" & NulSID(!AlamatCust) & "','" & NulSID(!NoPolisi) & "','" & NulSID(!PTSN) & "','" & NulSID(!GSM) & "','" & NulSID(!CDMA) & "','" & NulSID(!LastSA) & "','" & NulSID(Text(0)) & "')"
                          
27                        Call objHT.InsertData_Bengkel(sSQL)
28                    End If
29                    .MoveNext
30                Loop
31            End If
32        End With
33        Call objHT.CloseTransactionBengkel
34        MsgBox "Data berhasil disimpan", vbInformation
35        Call refreshGrid
         '' frmeKonfirm.Visible = True
          
36       On Error GoTo 0
37       Exit Sub

cmdKonfir_Click_Error:
38        If isBegin Then Call objHT.CloseTransactionBengkel_err
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
39        varErlError = Erl
40        varErrNumber = err.Number
41        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdKonfir_Click of Form frmInjectMEstimasi")
End Sub

Private Sub cmdRef_Click()
    Call refreshGrid
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
      Dim namaSupp As String
      Dim rsHlp As New ADODB.Recordset
          
1        On Error GoTo cmdSimpan_Click_Error

2         boleh = True
3         pesan = "Lengkapi data berikut .... " & vbCr & vbCr
4         If Trim(Text(1)) = "" Then
5             boleh = False
6             pesan = pesan + "No. Nota Rekap invoice dari Supplier. " & vbCr
7         End If
8         If Trim(Text(2)) = "" Then
9             boleh = False
10            pesan = pesan + "Nama Penagih. " & vbCr
11        End If
12        If MsgBox("Akan Mengkonfirmasi " & nTagih & " sublet, dengan Total " & totTagih & " ?", vbYesNo, "Konfirmasi") = vbNo Then
13            boleh = False
14            pesan = pesan + "Cek ulang rincian Sublet. "
15        End If
          
16        If boleh Then
17            With rsRecord
18                If Not .EOF Then
19                    .MoveFirst
20                    Do While Not .EOF
21                        If rsRecord("TAGIH") Then
22                            namaSupp = Trim(Text(3))
23                                Set rsHlp = objO.FetchDataACC("SELECT NAMALGN FROM APMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(rsRecord("SUPP")) & "'")
24                                If Not rsHlp.EOF Then
25                                    namaSupp = Trim(rsHlp("NAMALGN"))
26                                End If
27                            Call objHT.OpenTransactionBengkel
28                            Call objHT.InsertData_BengkelOnly("INSERT INTO KONFIRM_SUBLET VALUES ('" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Text(1)) & "','" & Trim(rsRecord("NOINVSUP")) & "','" & Trim(rsRecord("SUPP")) & "','" & Trim(namaSupp) & "','" & Trim(Text(2)) & "','" & Trim(rsRecord("NOWO")) & "', '" & Trim(rsRecord("NOINV")) & "', '" & Trim(rsRecord("NOPOLISI")) & "','" & Trim(rsRecord("SUBLET")) & "')")
29                            Call objHT.UpdateData_BengkelOnly("UPDATE JOBSUBLET SET TAGIH = 1 WHERE NO_WO = '" & Trim(rsRecord("NOWO")) & "' AND NAMABARANG = '" & Trim(rsRecord("SUBLET")) & "'")
30                            Call objHT.CloseTransactionBengkel
31                        End If
32                        .MoveNext
33                    Loop
34                End If
35            End With
36            Call cetakReport
37            Call hapusForm
38            ''Call refreshGrid(IIf(chAll.value = Checked, True, False))
39            frmeKonfirm.Visible = False
40        Else
41            MsgBox pesan
42        End If

43       On Error GoTo 0
44       Exit Sub

cmdSimpan_Click_Error:

45        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmInjectMEstimasi")
End Sub

Private Sub cetakReport()
      Dim Report As New RptKonfirmasiSublet
      Dim rsRepo As ADODB.Recordset
      Dim rsHlp As ADODB.Recordset

1        On Error GoTo cetakReport_Error

2         sSQL = "SELECT A.*, KODESATUAN, QTY, CUSTOMER, HARGA*QTY AS HPP, PRICE*QTY AS JUAL, C.TGLVALID, B.FRONT FROM KONFIRM_SUBLET A WITH(NOLOCK) " _
               & "INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NOWO = B.NO_WO AND B.FS_LOG_DEL = '' INNER JOIN JOBSUBLET C WITH(NOLOCK) ON A.NOWO = C.NO_WO AND A.SUBLET = C.NAMABARANG " _
               & "WHERE NO_TAGIHAN = '" & Trim(Text(1)) & "' AND TANGGAL BETWEEN '" & Format(isnow, "MM/DD/YYYY 00:00:00") & "' AND '" & Format(isnow, "MM/DD/YYYY 23:59:59") & "' ORDER BY NOWO "
3         Set rsRepo = objO.FetchData(sSQL)
4         Report.database.SetDataSource rsRepo
          
5         Report.txtPerusahaan.SetText gPerush
6         Report.txtAlamat.SetText gAlamat
7         Report.txtKota.SetText gKotaCbg & "," & TglDDMMMMYYYY(Format(isnow, "MM/DD/YYYY"))
          
8         Set rsHlp = objO.FetchData("SELECT FS_NM_PETUGAS FROM T_PETUGAS WITH(NOLOCK) WHERE FS_KD_PETUGAS = '" & Trim(gSA) & "'")
9         If Not rsHlp.EOF Then
10            Report.txtPetugas.SetText Trim(rsHlp("FS_NM_PETUGAS"))
11        End If
12        Report.txtPenagih.SetText Trim(Text(2))
13        Report.txtKabeng.SetText Trim(gSecmanager)

14        With CRViewer1
15            .ReportSource = Report
16            .ViewReport
17            .Visible = True
              
18            .Top = 0
19            .Left = gridSublet.Left
20            .Width = gridSublet.Width
21            .Height = cmdKonfir.Top + cmdKonfir.Height
              
22            cmdTutup.Visible = True
23            cmdTutup.Left = .Left
24            cmdTutup.Top = .Top + 520
              
25            .Zoom 100
26        End With
27        Screen.MousePointer = vbDefault
28       On Error GoTo 0
29       Exit Sub

cetakReport_Error:

30        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakReport of Form frmInjectMEstimasi")
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Load()
    Call hapusForm
End Sub

Private Sub hapusForm()
    For i = 0 To 3
        Text(i) = ""
    Next
    tgl0 = Format(isnow, "DD/MM/YYYY")
    tgl1 = Format(isnow, "DD/MM/YYYY")
    nTagih = 0
    totTagih = 0
End Sub

Private Sub createRS()
1        On Error GoTo createRS_Error
2         If rsRecord.State Then rsRecord.Close
3         Set rsRecord = Nothing
4         With rsRecord
             ' If .State <> adStateClosed Then .Close
5             .Fields.Append "NORANGKA", adVarChar, 25
6             .Fields.Append "RATAKM", adInteger
7             .Fields.Append "JT_TGL", adDate
8             .Fields.Append "JT_BLN", adDate
9             .Fields.Append "KMNEXT", adDouble
10            .Fields.Append "SBNEXT", adDouble
11            .Fields.Append "KMNOW", adDouble
12            .Fields.Append "TGLSVCNOW", adDate
13            .Fields.Append "NOINV", adVarChar, 100
14            .Fields.Append "NOWO", adVarChar, 100
15            .Fields.Append "NAMACUST", adVarChar, 100
16            .Fields.Append "ALAMATCUST", adVarChar, 200
17            .Fields.Append "NOPOLISI", adVarChar, 25
18            .Fields.Append "PTSN", adVarChar, 25
19            .Fields.Append "GSM", adVarChar, 25
20            .Fields.Append "CDMA", adVarChar, 25
21            .Fields.Append "LASTSA", adVarChar, 25
22            .Fields.Append "pilih", adBoolean
23            .Fields.Append "UDAHTAGIH", adBoolean
24            .Open
25        End With
          
26        'Set gridCustomer.DataSource = rsRecord
27        Call headColumns

28       On Error GoTo 0
29       Exit Sub

createRS_Error:

30        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRS of Form frmInjectMEstimasi")
End Sub

Private Sub headColumns()

   On Error GoTo headColumns_Error

    With gridCustomer
        .Columns(0).Width = 800
        .Columns(0).Caption = "Pilih"
        .Columns(0).DataField = "Pilih"
        .Columns(0).Alignment = dbgCenter
        '.Columns(0).value = False
        .Splits(0).Columns(0).Locked = False

        .Columns(1).Width = 2500
        .Columns(1).Caption = "No. Rangka"
        .Columns(1).DataField = "NORANGKA"
        .Splits(0).Columns(1).Locked = True

        .Columns(2).Width = 800
        .Columns(2).Caption = "Rata KM"
        .Columns(2).DataField = "RATAKM"
        .Splits(0).Columns(2).Locked = True

        .Columns(3).Width = 1500
        .Columns(3).Caption = "JT Service"
        .Columns(3).DataField = "JT_TGL"
        .Splits(0).Columns(3).Locked = True

        .Columns(4).Width = 1500
        .Columns(4).Caption = "JT_BLN"
        .Columns(4).DataField = "JT_BLN"
        .Columns(4).Visible = False
        .Splits(0).Columns(4).Locked = True

        .Columns(5).Width = 1000
        .Columns(5).Caption = "KM Next"
        .Columns(5).DataField = "KMNEXT"
        .Splits(0).Columns(5).Locked = False

        .Columns(6).Width = 1000
        .Columns(6).Caption = "SBE Next"
        .Columns(6).DataField = "SBNEXT"
        .Splits(0).Columns(6).Locked = True

        .Columns(7).Width = 1100
        .Columns(7).Caption = "Km Terakhir"
        .Columns(7).DataField = "KMNOW"
        .Splits(0).Columns(7).Locked = True

        .Columns(8).Width = 1100
        .Columns(8).Caption = "Tgl Service"
        .Columns(8).DataField = "TGLSVCNOW"
        .Splits(0).Columns(8).Locked = True

        .Columns(9).Width = 500
        .Columns(9).Caption = "NOINV"
        .Columns(9).DataField = "NOINV"
        .Columns(9).Visible = False
        .Splits(0).Columns(9).Locked = True

        .Columns(10).Width = 600
        .Columns(10).Caption = "NOWO"
        .Columns(10).DataField = "NOWO"
        .Columns(10).Visible = False
        .Splits(0).Columns(10).Locked = True

        .Columns(11).Width = 1800
        .Columns(11).Caption = "Nama Customer"
        .Columns(11).DataField = "NAMACUST"
        .Splits(0).Columns(11).Locked = True

        .Columns(12).Width = 2300
        .Columns(12).Caption = "Alamat Customer"
        .Columns(12).DataField = "ALAMATCUST"
        .Splits(0).Columns(12).Locked = True

        .Columns(13).Width = 800
        .Columns(13).Caption = "No Polisi"
        .Columns(13).DataField = "NOPOLISI"
        .Splits(0).Columns(13).Locked = True

        .Columns(14).Width = 800
        .Columns(14).Caption = "PSTN"
        .Columns(14).DataField = "PTSN"
        .Splits(0).Columns(14).Locked = True

        .Columns(15).Width = 800
        .Columns(15).Caption = "GSM"
        .Columns(15).DataField = "GSM"
        .Splits(0).Columns(15).Locked = True

        .Columns(16).Width = 800
        .Columns(16).Caption = "CDMA"
        .Columns(16).DataField = "CDMA"
        .Splits(0).Columns(16).Locked = True

        .Columns(17).Width = 800
        .Columns(17).Caption = "Last SA"
        .Columns(17).DataField = "LASTSA"
        .Splits(0).Columns(17).Locked = True
    End With

   On Error GoTo 0
   Exit Sub

headColumns_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumns of Form frmInjectMEstimasi")
End Sub

Private Sub refreshGrid()
      Dim rsSublet As ADODB.Recordset
      Dim varTagih As String
          
1        On Error GoTo refreshGrid_Error

2         varTagih = " AND TAGIH = ''"
3         If nTagih Then
4             varTagih = ""
5         End If
          
6         sSQL = "select * from ("
7         sSQL = sSQL & " select A.*,(SELECT TOP 1 FRONT FROM T_INVOICE WHERE NO_CHASSIS = A.NoRangka AND FS_LOG_DEL = '' ORDER BY INVDATE desc) lastSa"
8         sSQL = sSQL & " from M_ESTIMASI A"
9         sSQL = sSQL & " WHERE A.SBENEXT > isnull("
10        sSQL = sSQL & "     (select MAX(BB.FS_NILAI)"
11        sSQL = sSQL & "     from [ReplBengkel99]..T_INVOICEnew AA"
12        sSQL = sSQL & "     left join [ReplBengkel99]..T_REGULAR BB on BB.FS_RC = AA.FS_RC"
13        sSQL = sSQL & " where AA.NO_CHASSIS = A.NoRangka and FS_LOG_DEL = '')"
14        sSQL = sSQL & "     ,0)  and"
15        sSQL = sSQL & " cast(A.SBENEXT AS varchaR(50))+A.NORANGKA NOT IN (SELECT CAST(SbeNext AS VArchar(50))+norangka FROM [M_ESTIMASIInject] cc WITH (NOLOCK) WHERE cc.norangka = A.NORANGKA ) and "
16        sSQL = sSQL & " (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", -30, Format(tgl0.value, "DD/MM/YYYY")), "yyyy-mm-dd") & "' AND '" & Format(DateAdd("D", -30, Format(tgl1.value, "DD/MM/YYYY")), "yyyy-mm-dd") & "'"
17        sSQL = sSQL & "  and A.NoRangka NOT IN (select No_chassis from BOOK where DATE_BOOK between '" & Format(DateAdd("D", -30, Format(tgl0.value, "DD/MM/YYYY")), "yyyy-mm-dd") & "' and '" & Format(DateAdd("D", -30, Format(tgl1.value, "DD/MM/YYYY")), "yyyy-mm-dd") & "' and FS_LOG_DEL = '')"
18        sSQL = sSQL & "  ) a where a.lastSa not in (select FS_KD_PETUGAS from T_PETUGAS WHERE TIPE = 'SA')"
19        sSQL = sSQL & "  ORDER BY JT_Tgl"
20        Set rsSublet = objO.FetchData(sSQL)
21        Call createRS
22        If Not rsSublet.EOF Then
              
23            Do While Not rsSublet.EOF
24                With rsRecord
25                    .AddNew
26                    !noRangka = NulSID(rsSublet("noRangka"))
27                    !RATAKM = NulSID(rsSublet("RATAKM"))
28                    !JT_Tgl = TglDDMMYYYY(NulSID(rsSublet("JT_TGL")))
29                    !JT_BLN = TglDDMMYYYY(NulSID(rsSublet("JT_BLN")))
30                    !KmNext = NulSID(rsSublet("KmNext"))
31                    !SBNEXT = NulSID(rsSublet("SBENEXT"))
32                    !KMNOW = NulSID(rsSublet("KMNOW"))
33                    !TglSvcNow = TglDDMMYYYY(IIf(NulSID(rsSublet("TGLSVCNOW")) = "", "1900-01-01", NulSID(rsSublet("TGLSVCNOW"))))
34                    !noInv = NulSID(rsSublet("noInv"))
35                    !noWO = NulSID(rsSublet("noWO"))
36                    !namaCust = NulSID(rsSublet("namaCust"))
37                    !AlamatCust = NulSID(rsSublet("ALAMATCUST"))
38                    !NoPolisi = NulSID(rsSublet("NoPolisi"))
39                    !PTSN = NulSID(rsSublet("PSTN"))
40                    !GSM = NulSID(rsSublet("GSM"))
41                    !CDMA = NulSID(rsSublet("CDMA"))
42                    !LastSA = NulSID(rsSublet("LASTSA"))
43                    !pilih = False
44                    .Update
45                End With
46                rsSublet.MoveNext
47            Loop
48
49            gridCustomer.DataSource = rsRecord
50            jmlRecord.Caption = "Jumlah Data : " & rsRecord.RecordCount
51           '' gridSublet.FetchRowStyle = True
              ''gridCustomer.ReBind
              ''Call headColumns
52            'rsRecord.MoveFirst
53        End If
54        nTagih = 0 'rsRecord.RecordCount
55        totTagih = 0

56       On Error GoTo 0
57       Exit Sub

refreshGrid_Error:

58        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form frmInjectMEstimasi")
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        Image1.Height = Me.Height - 500
    End If
End Sub

Private Sub gridCustomer_AfterColUpdate(ByVal ColIndex As Integer)
1        On Error GoTo gridSublet_AfterColUpdate_Error

2         With gridCustomer
3             Select Case ColIndex
                  Case 0
4                     If .Columns(ColIndex) Then
5                         nTagih = nTagih + 1
6                        ' totTagih = totTagih + (CDbl(.Columns(9)) * CDbl(.Columns(10)))
7                     Else
8                         nTagih = nTagih - 1
9                        ' totTagih = totTagih - (CDbl(.Columns(9)) * CDbl(.Columns(10)))
10                    End If
'11                    If chAll.value = True Then
'12                        MsgBox "Cekbox tampilkan semua tidak boleh dicetang !", vbInformation + vbOKOnly, "Sublet"
'13                    End If
14            End Select
              
15            '''nTagih = Format(CDbl(nTagih), "##,0")
16           ' totTagih = Format(CDbl(totTagih), "##,0")
17        End With

18       On Error GoTo 0
19       Exit Sub

gridSublet_AfterColUpdate_Error:

20        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure gridSublet_AfterColUpdate of Form frmInjectMEstimasi")
End Sub

'Private Sub gridSublet_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
'    Dim rs As New ADODB.Recordset
'    Set rs = rsRecord.Clone
'    If Bookmark > 1 Then
'        rs.Bookmark = Bookmark
'    End If
'   '' If Nul(rs.Fields(14).value) = True Then RowStyle = "SudahDitagih"
'End Sub

Private Sub gridSublet_FilterChange()
    On Error GoTo errHandler
    Set cols = gridSublet.Columns
    Dim C As Integer
    
    C = gridSublet.col
    gridSublet.HoldFields
    rsRecord.filter = getFilter()
    If rsRecord.EOF Then
        For Each col In gridSublet.Columns
            col.FilterText = ""
        Next col
        rsRecord.filter = ""
    End If
    gridSublet.col = C
    gridSublet.EditActive = True
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

Private Sub gridSublet_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 12 Then rsRecord.Sort = gridSublet.Columns(ColIndex).DataField & " " & Urut
4        gridSublet.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
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
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.show (vbModal)
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

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsRepo As ADODB.Recordset
    If Index = 0 Then
        If Text(Index) <> "" Then
            Set rsRepo = objO.FetchData("SELECT FS_NM_PETUGAS FROM T_PETUGAS WITH (NOLOCK) WHERE FS_KD_PETUGAS = '" & Trim(Text(Index)) & "'") ''objO.FetchDataACC("SELECT KODELGN, NAMALGN FROM APMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(Text(Index)) & "'")
            If rsRepo.EOF Then
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.show (vbModal)
            Else
                Text(3) = Trim(rsRepo("FS_NM_PETUGAS"))
            End If
            If rsRepo.State <> adStateClosed Then rsRepo.Close
            'Call refreshGrid '(IIf(chAll.value = Checked, True, False))
        End If
    End If
    If Index = 1 Then
        If Text(1) <> "" Then
            If objO.FetchData_Boolean("Select No_Tagihan FROM KONFIRM_SUBLET WITH(NOLOCK) WHERE NO_Tagihan = '" & Text(1) & "'") = True Then
                MsgBox "Nomor Tagihan : " & Text(1) & ", Sudah terdaftar !"
                Text(1) = ""
            End If
        End If
    End If
End Sub

Private Sub tgl1_CloseUp()
    ''Call refreshGrid(IIf(chAll.value = Checked, True, False))
End Sub


