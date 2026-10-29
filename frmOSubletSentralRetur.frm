VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmOSubletSentralRetur 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Retur Sublet"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   11760
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7890
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   11205
      Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
         Bindings        =   "frmOSubletSentralRetur.frx":0000
         Height          =   1695
         Left            =   2400
         TabIndex        =   4
         Top             =   4920
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2990
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode Supplier"
         Columns(0).DataField=   "KODEAP"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama Supplier"
         Columns(1).DataField=   "SUPPLIER"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Alamat"
         Columns(2).DataField=   "ALAMAT"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).ExtendRightColumn=   -1  'True
         Splits(0).MarqueeStyle=   3
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2381"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5292"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5212"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         AllowRowSizing  =   0   'False
         Appearance      =   1
         BorderStyle     =   1
         ColumnHeaders   =   -1  'True
         DataMode        =   0
         DefColWidth     =   0
         Enabled         =   -1  'True
         HeadLines       =   1
         RowDividerStyle =   2
         LayoutName      =   ""
         LayoutFileName  =   ""
         LayoutURL       =   ""
         EmptyRows       =   0   'False
         ListField       =   ""
         DataField       =   ""
         IntegralHeight  =   0   'False
         FetchRowStyle   =   0   'False
         AlternatingRowStyle=   0   'False
         DataMember      =   ""
         ColumnFooters   =   0   'False
         FootLines       =   1
         DeadAreaBackColor=   12577002
         ValueTranslate  =   0   'False
         RowDividerColor =   15790320
         RowSubDividerColor=   15790320
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid gridSublet 
         Height          =   4815
         Left            =   0
         TabIndex        =   5
         Top             =   3000
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   8493
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No. PO"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   1
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Supplier"
         Columns(1).DataField=   ""
         Columns(1).DropDown=   "TDBDropDown1"
         Columns(1).DropDown.vt=   8
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama Sublet"
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Satuan"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Qty"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Harga Satuan"
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Subtotal"
         Columns(6).DataField=   ""
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   7
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=7"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1).Button=1"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(30)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=212,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.locked=-1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(64)  =   "Named:id=33:Normal"
         _StyleDefs(65)  =   ":id=33,.parent=0"
         _StyleDefs(66)  =   "Named:id=34:Heading"
         _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(68)  =   ":id=34,.wraptext=-1"
         _StyleDefs(69)  =   "Named:id=35:Footing"
         _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(71)  =   "Named:id=36:Selected"
         _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(73)  =   "Named:id=37:Caption"
         _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(75)  =   "Named:id=38:HighlightRow"
         _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=39:EvenRow"
         _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(79)  =   "Named:id=40:OddRow"
         _StyleDefs(80)  =   ":id=40,.parent=33"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   930
         TabIndex        =   6
         Top             =   750
         Width           =   4815
         _ExtentX        =   8493
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
         Index           =   4
         Left            =   930
         TabIndex        =   7
         Top             =   1035
         Width           =   4815
         _ExtentX        =   8493
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
         TabIndex        =   8
         Top             =   1350
         Width           =   4815
         _ExtentX        =   8493
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
         Index           =   6
         Left            =   930
         TabIndex        =   9
         Top             =   1650
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
         Left            =   6750
         TabIndex        =   10
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
         Index           =   10
         Left            =   6750
         TabIndex        =   11
         Top             =   750
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   13
         Left            =   6750
         TabIndex        =   12
         Top             =   1350
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   12
         Left            =   6750
         TabIndex        =   13
         Top             =   1035
         Width           =   4215
         _ExtentX        =   7435
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
         Index           =   11
         Left            =   10050
         TabIndex        =   14
         Top             =   750
         Width           =   915
         _ExtentX        =   1614
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
         Index           =   2
         Left            =   3000
         TabIndex        =   15
         Top             =   450
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   1
         Left            =   930
         TabIndex        =   16
         Top             =   450
         Width           =   1845
         _ExtentX        =   3254
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   14
         Left            =   6750
         TabIndex        =   17
         Top             =   1650
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   15
         Left            =   6750
         TabIndex        =   18
         Top             =   1935
         Width           =   4215
         _ExtentX        =   7435
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
         Index           =   17
         Left            =   9750
         TabIndex        =   19
         Top             =   435
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
         Index           =   7
         Left            =   930
         TabIndex        =   20
         Top             =   1935
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
         Index           =   8
         Left            =   930
         TabIndex        =   21
         Top             =   2235
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
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   2805
         TabIndex        =   22
         Top             =   450
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   16
         Left            =   6750
         TabIndex        =   23
         Top             =   2235
         Width           =   4215
         _ExtentX        =   7435
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   8760
         Top             =   6720
         Visible         =   0   'False
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
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
      Begin VB.Line Line8 
         BorderColor     =   &H0000FFFF&
         X1              =   7680
         X2              =   10920
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line7 
         BorderColor     =   &H0000FFFF&
         X1              =   7680
         X2              =   10920
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line6 
         BorderColor     =   &H0000FFFF&
         X1              =   120
         X2              =   3360
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H0000FFFF&
         X1              =   120
         X2              =   3360
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retur Pemakaian Sublet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Left            =   4080
         TabIndex        =   44
         Top             =   2685
         Width           =   3075
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   10920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   10920
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3360
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3360
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
         Left            =   3720
         TabIndex        =   43
         Top             =   45
         Width           =   3675
      End
      Begin VB.Label Label6 
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
         TabIndex        =   38
         Top             =   795
         Width           =   405
      End
      Begin VB.Label Label3 
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
         TabIndex        =   37
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label4 
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
         TabIndex        =   36
         Top             =   1395
         Width           =   330
      End
      Begin VB.Label Label5 
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
         TabIndex        =   35
         Top             =   1695
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
         Left            =   5955
         TabIndex        =   34
         Top             =   465
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
         Left            =   6180
         TabIndex        =   33
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Label19 
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
         Left            =   6120
         TabIndex        =   32
         Top             =   1365
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
         TabIndex        =   31
         Top             =   480
         Width           =   660
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
         Left            =   5955
         TabIndex        =   30
         Top             =   1695
         Width           =   645
      End
      Begin VB.Label Label21 
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
         Left            =   5820
         TabIndex        =   29
         Top             =   1980
         Width           =   780
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delivery Date"
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
         Left            =   8640
         TabIndex        =   28
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label8 
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
         TabIndex        =   27
         Top             =   1980
         Width           =   315
      End
      Begin VB.Label Label9 
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
         TabIndex        =   26
         Top             =   2280
         Width           =   435
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Mesin"
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
         Left            =   5955
         TabIndex        =   25
         Top             =   2280
         Width           =   645
      End
      Begin VB.Label SgnRekom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Rekom"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   8280
         TabIndex        =   24
         Top             =   1642
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11055
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   2640
         Width           =   11055
      End
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali Ke Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8400
      Width           =   2175
   End
   Begin VB.CommandButton cmdSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8400
      Width           =   1575
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2805
      TabIndex        =   39
      Top             =   120
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   0
      Left            =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1320
      _ExtentX        =   2328
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
   Begin VB.Image Image1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmOSubletSentralRetur.frx":0015
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   195
      Left            =   735
      TabIndex        =   42
      Top             =   165
      Width           =   570
   End
   Begin VB.Label grandTotal 
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
      Left            =   2520
      TabIndex        =   41
      Top             =   8400
      Width           =   180
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Grand Total :"
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
      Left            =   480
      TabIndex        =   40
      Top             =   8400
      Width           =   1875
   End
End
Attribute VB_Name = "frmOSubletSentralRetur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSublet As New ADODB.Recordset
Dim rsSearch As ADODB.Recordset
'Dim rsHelp As New ADODB.Recordset
Dim vPDPGL As String
Dim vHutangGL As String
Dim vPpnGL As String

Private Function createSublet() As Boolean
1         On Error GoTo vb_error
          
2         If rsSublet.State Then rsSublet.Close '
3         Set rsSublet = Nothing
4         With rsSublet
5             .Fields.Append "NOPO", adVarChar, 25
6             .Fields.Append "SUPP", adVarChar, 15
7             .Fields.Append "SUBLET", adVarChar, 50 '40
8             .Fields.Append "SAT", adVarChar, 10
9             .Fields.Append "QTY", adDouble
10            .Fields.Append "HARGASAT", adDouble
11            .Fields.Append "SUBTOTAL", adDouble
            ' .Fields.Append "RETUR", adBoolean
12            .Open
13        End With
14        If rsSublet.State Then createSublet = True
15        Exit Function
          
vb_error:
16        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub headSublet()
    
    With gridSublet
        .Columns(0).Caption = "No. PO"
        .Columns(0).DataField = "NOPO"
        .Columns(0).Width = 1400
        
        .Columns(1).Caption = "Supplier"
        .Columns(1).DataField = "SUPP"
        .Columns(1).Width = 1200
        .Columns(1).Locked = True
        
        .Columns(2).Caption = "Pekerjaan Sublet"
        .Columns(2).DataField = "SUBLET"
        .Columns(2).Width = 3000
        
        .Columns(3).Caption = "Satuan"
        .Columns(3).DataField = "SAT"
        .Columns(3).Width = 1000
        
        .Columns(4).Caption = "Qty"
        .Columns(4).DataField = "QTY"
        .Columns(4).Width = 800
        
        .Columns(5).Caption = "Harga Satuan"
        .Columns(5).DataField = "HARGASAT"
        .Columns(5).Width = 1200
        .Columns(5).Locked = True
        
        .Columns(6).Caption = "Sub Total"
        .Columns(6).DataField = "SUBTOTAL"
        .Columns(6).Width = 1200
        .Columns(6).Locked = True
        
        '.Columns(7).Caption = "Retur"
        '.Columns(7).DataField = "RETUR"
        '.Columns(7).Width = 400
        '.Columns(7).Alignment = dbgCenter
    End With
End Sub

Private Sub hapusForm()
    For i = 1 To 17
        Select Case i
            Case 9
                Text(i) = 0
            Case 17
                Text(i) = "99-99-9999"
            Case Else
                Text(i) = ""
        End Select
    Next
    If rsSublet.State Then rsSublet.Close
    grandTotal = 0
    
'    With Adodc1
'        .ConnectionString = SQLConACC
'        .CommandType = adCmdText
'        .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST WITH(NOLOCK)) ORDER BY KODEAP"
'        .Refresh
'    End With
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
      Dim pesanIT As String
      Dim noJurnal As Boolean

1         On Error GoTo vbError
2         boleh = True
3         pesan = ""
4         pesanIT = ""
          Dim sum_bengkel As Double
5         If rsSublet.State <> adStateClosed Then
6             rsSublet.MoveFirst
7             Do While Not rsSublet.EOF
8                 If rsSublet("SUPP") = "" Then
9                     boleh = False
10                    pesan = pesan + "Kode Supplier untuk " & rsSublet("SUBLET") & ". " & vbCr
11                End If
12                If Abs(rsSublet("HARGASAT")) = "0" Then
13                    If MsgBox("Yakin harga beli " & rsSublet("SUBLET") & " 0 (Nol) ?", vbYesNo, "Validasi Harga") = vbNo Then
14                        boleh = False
15                        pesan = pesan + "Harga Beli Satuan " & rsSublet("SUBLET") & ". " & vbCr
16                    End If
17                End If
18                sum_bengkel = sum_bengkel + (Abs(CDbl(rsSublet("HARGASAT"))) * CDbl(rsSublet("QTY")))
19                rsSublet.MoveNext
20            Loop
21        End If
          Dim rsTemp As ADODB.Recordset
          
          
22        If CDbl(Right(NisDb, 2)) <> CDbl(Month(isnow)) Then
23            boleh = False
24            pesan = pesan + "Beda Fiskal Accounting tidak dapat dibuat invoice. " & vbCr & "Pindah Fiskal Accounting pada bulan bersangkutan. "
25        End If
          
26        If boleh Then
27             Set rsTemp = objO.FetchDataACC("select isnull(SUM(jumlahtrn),0) total from aptrn where nobukti like '" & "S-" & Trim(Text(0)) & "%'")
28             If Not rsTemp.EOF Then
29               rsTemp.MoveFirst
                 ''If rsTemp!total <> sum_bengkel Then '' And rsTemp!total >= 0 Then
30               If Not (rsTemp!total = sum_bengkel And rsTemp!total <> 0) Then
                    ''insert data ke selisih sublet
                    ''2016-07-26 || nilai bengkel dibuat negatif
31                  Call objO.InsertData("INSERT INTO [SelisihSublet] ([no_wo] ,[no_chassis] ,[sum_bengkel] ,[sum_ap] ,[tglEntry],KodeUser,menu) VALUES ('" & Trim(Text(0)) & "','" & Trim(Text(15)) & "'," & k2t(sum_bengkel) * -1 & "," & k2t(CDbl(rsTemp!total)) & ",GETDATE(),'" & gSA & "','Retur Sublet While Insert')")
32               End If
33             End If
              Dim jobSub(100) As String
              Dim iSub As Integer
34            iSub = 0
              
35            Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletSentralRetur", "cmdSimpan", 0)
36            erTranAll = False
37            onTR = True
38            Call objHT.OpenTransactionAll
                  
              ' SIMPAN SUBLET
39            If rsSublet.State Then
40                rsSublet.MoveFirst
41                If Not rsSublet.EOF Then
      '''                ' CARI AKUN UNTUK PDP DAN HUTANG
                      Dim isDyna As Boolean
                      
                      'PENGECEKAN PENYERAGAMAN AKUN
                      'TANGGAL 04-09-2014
                      'RUDI KRISTANTO
                      
42                    If Log.isAkun = 1 Then
43                        isDyna = IIf(InStr(1, UCase(Text(12).Text), "DYNA") > 0, True, False)
44                        Set rsGL = objO.Get_AkunGL_NEW_Sublet("OTHERS", isDyna, True)
45                        If isDyna = True Then
46                            vPDPGL = cekKodeAkun(Trim(rsGL.Fields(0).value))
47                            vPpnGL = cekKodeAkun(Trim(rsGL.Fields(1).value))
48                            vHutangGL = cekKodeAkun(Trim(rsGL.Fields(2).value))
49                        Else
50                            vPDPGL = cekKodeAkun(Trim(rsGL.Fields(0).value))
51                            vPpnGL = cekKodeAkun(Trim(rsGL.Fields(1).value))
52                            vHutangGL = cekKodeAkun(Trim(rsGL.Fields(2).value))
53                        End If
54                    Else
55                        Set rsGL = objO.FetchData("SELECT FS_KD_PDP, FS_KD_PPN, FS_KD_HUTANG FROM T_AKUN WITH(NOLOCK) ")
56                        If Not rsGL.EOF Then
57                            vPDPGL = cekKodeAkun(Trim(rsGL("FS_KD_PDP")))
58                            vPpnGL = cekKodeAkun(Trim(rsGL("FS_KD_PPN")))
59                            vHutangGL = cekKodeAkun(Trim(rsGL("FS_KD_HUTANG")))
60                        End If
61                    End If
                      
                      
62                    Do While Not rsSublet.EOF
                          Dim sudahValid As Boolean
63                        sudahValid = False
64                         Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'")
65                        If Not rsSearch.EOF Then
66                            If rsSearch("VALID") = 0 Then
67                                sudahValid = True
68                            End If
69                        End If
                          
                         ' If rsSublet("RETUR") Then
70                            If Not sudahValid Then
                                  'update status sublet jadi valid
                                  'Call objHT.UpdateData_Bengkel("UPDATE JOBSUBLET SET HARGA = '0', SUPPLIER = '" & Trim(rsSublet("SUPP")) & "', VALID = 0, TGLRETUR = GETDATE() WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'")
71                                jobSub(iSub) = "UPDATE JOBSUBLET SET HARGA = '0', SUPPLIER = '" & Trim(rsSublet("SUPP")) & "', VALID = 0, TGLRETUR = GETDATE() WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'"
72                                iSub = iSub + 1
                                  
                                  ' INPUT PDP TERHADAP HUTANG SUPPLIER
                                  Dim vDelIndex As String
                                  Dim vKodeLgn As String
                                  Dim vTglTempo As Date
                                  Dim vTglSkr As Date
                                  Dim vTBeli As Double
                                  Dim vTPPN As Double
                                  Dim vTotal As Double
                                  Dim vKeterangan As String
                                  
73                                vDelIndex = gKode & Left(McDelIndex(7), 8)
74                                vKodeLgn = Trim(rsSublet("SUPP"))
75                                vTglSkr = Format(isnow, "dd/mm/yyyy")
76                                vTglTempo = vTglSkr
77                                vTotal = Abs(CDbl(rsSublet("HARGASAT"))) * CDbl(rsSublet("QTY"))
78                                vTPPN = Round(vTotal * (profilACC.nilaippn / profilACC.nilaitotal)) ''(10 / 110))
79                                vTBeli = vTotal - vTPPN
                                  
80                                If vPDPGL = "" Or vPpnGL = "" Or vHutangGL = "" Then
81                                    pesanIT = pesanIT + "Kode GL Ada yang kosong !"
82                                    GoTo vbError
83                                End If
                                  
84                                noJurnal = False
85                                Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM T_SUPPLIER WITH(NOLOCK) WHERE KODEAP = '" & Trim(rsSublet("SUPP")) & "'")
86                                If Not rsSearch.EOF Then
87                                    vKeterangan = rsSearch("SUPPLIER") & " (" & rsSublet("SUBLET") & ") No. WO :" & Trim(Text(0)) & gKodeRetur
88                                    If Trim(rsSearch("PKP")) = "0" Then
                                          'vPpnGL = ""
89                                        vTBeli = vTBeli + vTPPN
90                                        vTPPN = 0
91                                        vTotal = vTBeli
92                                    End If
93                                    If Trim(rsSearch("NONJURNAL")) = "1" Then
94                                        noJurnal = True
95                                    End If
96                                End If
                                  
97                                If gConnecAAS = True Then
98                                    If Not noJurnal Then
99                                        If Abs(CDbl(rsSublet("HARGASAT"))) <> 0 Then
                                              ' INPUT DI APTRN
                                              ' PDP
100                                           sSQL1 = "insert into APtrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                                                        & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, tglEntry,kodejurnal) " _
                                                        & "values " _
                                                        & " ('" & Trim(vKodeLgn) & "', '" & "S-" & Trim(Text(0)) & "', 'Q'," _
                                                        & " '" & "S-" & Trim(Text(0)) & gKodeRetur & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', 'Sublet " & rsSublet("SUBLET") & " No. WO : " & Trim(Text(0)) & gKodeRetur & "'," _
                                                        & " " & SimpanKoma2Titik(vTBeli * -1) & ", '" & gSA & "'," _
                                                        & " '" & vDelIndex & "','" & gKode & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vHutangGL & "')"
101                                           Call objHT.InsertData_ACC(sSQL1)
                                              
                                              ' PPN
102                                           sSQL1 = "insert into APtrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                                                        & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, tglEntry,kodejurnal) " _
                                                        & "values " _
                                                        & " ('" & Trim(vKodeLgn) & "', '" & "S-" & Trim(Text(0)) & "', 'S'," _
                                                        & " '" & "S-" & Trim(Text(0)) & gKodeRetur & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', 'Sublet " & rsSublet("SUBLET") & " No. WO : " & Trim(Text(0)) & gKodeRetur & "'," _
                                                        & " " & SimpanKoma2Titik(vTPPN * -1) & ", '" & gSA & "'," _
                                                        & " '" & vDelIndex & "','" & gKode & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vHutangGL & "')"
103                                           Call objHT.InsertData_ACC(sSQL1)
                                              
                                              ' INPUT DI APMST
104                                           sSQL1 = "Update APMST set SALDOJALAN = SALDOJALAN + " & _
                                              (vTBeli * -1) & ", TGLBELIAHR = '" & Format(isnow, "MM/DD/YYYY") & "' where KodeLgn= '" & Trim(vKodeLgn) & "'"
105                                           Call objHT.UpdateData_ACC(sSQL1)
                                              
106                                           sSQL1 = "Update APMST set SALDOJALAN = SALDOJALAN + " & SimpanKoma2Titik(vTPPN * -1) & ", TGLBELIAHR = '" & Format(isnow, "MM/DD/YYYY") & "' where KodeLgn= '" & Trim(vKodeLgn) & "'"
107                                           Call objHT.UpdateData_ACC(sSQL1)
                      
                                              ' INPUT DI GLTRN
108                                           If vTotal > 0 Then
109                                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(Text(0)) & gKodeRetur & "', '" & vPDPGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTBeli) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
110                                               Call objHT.InsertData_ACC(sSQL1)
                                                  
111                                               If vTPPN > 0 Then
112                                                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(Text(0)) & gKodeRetur & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
113                                                   Call objHT.InsertData_ACC(sSQL1)
114                                               End If
                                                  
115                                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(Text(0)) & gKodeRetur & "', '" & vHutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTotal) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
116                                               Call objHT.InsertData_ACC(sSQL1)
                                                  
                                                  ' INPUT DI GLMST
117                                               sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTBeli * -1) & ", " _
                                                        & " JKreditBi = JKreditBi + " & SimpanKoma2Titik(vTBeli) & " " _
                                                        & " where KodeGl = '" & vPDPGL & "'"
118                                               Call objHT.UpdateData_ACC(sSQL1)
                                                  
119                                               If vTPPN > 0 Then
120                                                   sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTPPN * -1) & ", " _
                                                        & " JKreditBi = JKreditBi + " & SimpanKoma2Titik(vTPPN) & " " _
                                                        & " where KodeGl = '" & vPpnGL & "'"
121                                                   Call objHT.UpdateData_ACC(sSQL1)
122                                               End If
                                                  
123                                               sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTotal * -1) & ", " _
                                                        & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(vTotal) & " " _
                                                        & " where KodeGl = '" & vHutangGL & "'"
124                                               Call objHT.UpdateData_ACC(sSQL1)
125                                           End If
126                                       End If
127                                   End If
128                               End If
129                           End If
                          'End If
130                       rsSublet.MoveNext
131                   Loop
                      Dim xI As Integer
132                   For xI = 0 To 100
133                       If jobSub(xI) <> "" Then
134                           Call objHT.UpdateData_Bengkel(jobSub(xI))
135                       Else
136                           Exit For
137                       End If
138                   Next xI
139               End If
140           End If
                  
141           If erTranAll = False Then
142               Call objHT.CloseTransactionAll
                  ''2016-07-22 || cek setelah commit
143               Set rsTemp = objO.FetchDataACC("select isnull(SUM(jumlahtrn),0) total from aptrn where nobukti like '" & "S-" & Trim(Text(0)) & "%'")
144               If Not rsTemp.EOF Then
145                 rsTemp.MoveFirst
      ''              If rsTemp!total <> sum_bengkel And rsTemp!total <> 0 Then
146                 If Not (rsTemp!total <> sum_bengkel And rsTemp!total = 0) Then
                       ''insert data ke selisih sublet
                       ''2016-07-26 || nilai bengkel dibuat negatif
147                    Call objO.InsertData("INSERT INTO [SelisihSublet] ([no_wo] ,[no_chassis] ,[sum_bengkel] ,[sum_ap] ,[tglEntry],KodeUser,menu) VALUES ('" & Trim(Text(0)) & "','" & Trim(Text(15)) & "'," & k2t(sum_bengkel) * -1 & "," & k2t(CDbl(rsTemp!total)) & ",GETDATE(),'" & gSA & "','Retur Sublet After Insert')")
148                 End If
149               End If
150               onTR = False
151               Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletSentralRetur", "cmdSimpan", 1)
152               MsgBox "Data berhasil disimpan "
153               cmdSimpan.Enabled = False
154           Else
155               Call objHT.CloseTransactionAll_Err
156               onTR = False
157               Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletSentralRetur", "cmdSimpan", 2)
158               MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
159           End If


160       Else
161           MsgBox "Data belum lengkap !!" & vbCr & pesan
162       End If
163       Exit Sub
          
vbError:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
164       varErlError = Erl
165       varErrNumber = err.Number
          ''varErlError,varErrNumber
166       Call objHT.CloseTransactionAll_Err
167       onTR = False
168       Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletSentralRetur", "cmdSimpan", 2)
169       MsgBox ErrMessage(varErlError, varErrNumber, "") & vbCr & "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
170       Exit Sub
End Sub

Private Sub Form_Activate()
'    With Adodc1
'        .ConnectionString = SQLConACC
'        .CommandType = adCmdText
'        .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST) ORDER BY KODEAP"
'        .Refresh
'    End With
    If orderAktif Or validAktif Or subletAktif Then
        MsgBox "Tutup Menu sebelumnya terlebih Dahulu !!", vbCritical
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    subRetAktif = True
    grandTotal = "0"
    TDBDropDown1.Visible = False
    TDBDropDown1.Width = 10000
    Set TDBDropDown1.DataSource = Adodc1
    Call hapusForm
End Sub

Private Sub Form_Resize()
    Image1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    subRetAktif = False
End Sub

Private Sub gridSublet_AfterColUpdate(ByVal ColIndex As Integer)
Dim rsHlp As New ADODB.Recordset

    Select Case ColIndex
        Case 5
            With gridSublet
                .Columns(5) = Abs(CDbl(.Columns(5)))
                rsSublet("HARGASAT") = Abs(CDbl(.Columns(5)))
                .Columns(6) = CDbl(.Columns(4)) * Abs(CDbl(.Columns(5)))
                rsSublet.MoveFirst
                grandTotal = 0
                Do While Not rsSublet.EOF
                    grandTotal = Format(grandTotal + CDbl(rsSublet("SUBTOTAL")), "##,0")
                    rsSublet.MoveNext
                Loop
            End With
        Case 1
            With gridSublet
                Set rsHlp = objO.FetchData("SELECT * FROM T_SUPPLIER WITH(NOLOCK) WHERE KODEAP = '" & Trim(.Columns(1)) & "' AND FS_LOG_DEL = ''")
                If rsHlp.EOF Then
                    rsSublet.Find "SUBLET = '" & Trim(.Columns(2)) & "'", , adSearchForward, 1
                    If Not rsSublet.EOF Then
                        rsSublet("SUPP") = ""
                    End If
                End If
            End With
    End Select
    Set gridSublet.DataSource = rsSublet
End Sub

Private Sub GridSublet_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 18
            With Adodc1
                .ConnectionString = adoCon
                .CommandType = adCmdText
                .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST WITH(NOLOCK)) ORDER BY KODEAP"
                .Refresh
            End With
    End Select
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
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "RETURSUBLET"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    Select Case KeyAscii
        Case 13
    End Select
End Sub

Private Function cariValid(nWO As String, nChassis As String) As Integer
Dim rsCari As ADODB.Recordset
    Set rsCari = objO.FetchData("SELECT SUM(VALID) AS OK FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(nWO) & "' AND NO_CHASSIS = '" & Trim(nChassis) & "'")
    If Not rsCari.EOF Then
        cariValid = CDbl(rsCari("OK"))
    End If
End Function

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            ' NO WO VALIDASI
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT B.FS_KD_PLG, B.FS_KD_WIL, B.CUSTOMER, B.CUST_ADRES, B.KOTA, B.TELPON, GSM, CDMA, B.FS_TH_MOBIL, B.MDL, " _
                                           & "B.MODEL, B.FMDL, B.COLOR, B.NO_POLISI, B.NO_CHASSIS, C.NO_ENGINE, C.DDATE, NOPO, NAMABARANG, KODESATUAN, QTY, HARGA, SUPPLIER, NOINV FROM JOBSUBLET A WITH(NOLOCK) " _
                                           & "INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS " _
                                           & "INNER JOIN CUSTDAT C WITH(NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS WHERE B.NOINV = '' AND A.NO_WO = '" & Trim(Text(Index)) & "' AND A.NO_WO NOT IN (SELECT NO_WO FROM VALIDORDER WITH(NOLOCK) WHERE USERDEL = '') ORDER BY A.NO_WO") ' AND VALID = 1
                If Not rsSearch.EOF Then
                    Call hapusForm
                    If Trim(rsSearch("NOINV")) = "" Then
                        If CDbl(rsSearch.RecordCount) >= cariValid(Trim(Text(Index)), Trim(rsSearch("NO_CHASSIS"))) And cariValid(Trim(Text(Index)), Trim(rsSearch("NO_CHASSIS"))) > 0 Then
                            Text(1) = Trim(rsSearch("FS_KD_PLG"))
                            Text(2) = Trim(rsSearch("FS_KD_WIL"))
                            Text(3) = Trim(rsSearch("CUSTOMER"))
                            Text(4) = Left(Trim(rsSearch("CUST_ADRES")), 60)
                            Text(5) = Trim(rsSearch("KOTA"))
                            Text(6) = Trim(rsSearch("TELPON"))
                            Text(7) = Trim(rsSearch("GSM"))
                            Text(8) = Trim(rsSearch("CDMA"))
                            Text(9) = Trim(rsSearch("FS_TH_MOBIL"))
                            Text(10) = Trim(rsSearch("MDL"))
                            Text(11) = Trim(rsSearch("MODEL"))
                            Text(12) = Trim(rsSearch("FMDL"))
                            Text(13) = Trim(rsSearch("COLOR"))
                            Text(14) = Trim(rsSearch("NO_POLISI"))
                            Text(15) = Trim(rsSearch("NO_CHASSIS"))
                            Text(16) = Trim(rsSearch("NO_ENGINE"))
                            Text(17).Mask = ""
                            Text(17) = Nul(IIf(Replace(rsSearch("DDATE"), "/", "-") = "01-01-3000", "", IIf(Replace(rsSearch("DDATE"), "/", "-") = "01-01-1900", "", Replace(rsSearch("DDATE"), "/", "-"))))
                            'Trim(rsSearch("DDATE"))
                            
                            If rsSublet.State <> adStateClosed Then rsSublet.Close
                            Call createSublet
                            Do While Not rsSearch.EOF
                                If rsSublet.State Then
                                    With rsSublet
                                        .AddNew
                                        !NOPO = Trim(rsSearch("NOPO"))
                                        !SUPP = Trim(rsSearch("SUPPLIER"))
                                        !SUBLET = Trim(rsSearch("NAMABARANG"))
                                        !SAT = Trim(rsSearch("KODESATUAN"))
                                        !qty = CDbl(Trim(rsSearch("QTY")))
                                        !HARGASAT = SpyRound(CDbl(Trim(rsSearch("HARGA"))), 0)
                                        !Subtotal = SpyRound(CDbl(Trim(rsSearch("QTY"))) * CDbl(Trim(rsSearch("HARGA"))), 0)
                                      '  !Retur = False
                                        .Update
                                    End With
                                    grandTotal = Format(grandTotal + SpyRound(CDbl(Trim(rsSearch("QTY"))) * CDbl(Trim(rsSearch("HARGA"))), 0), "##,0")
                                End If
                                rsSearch.MoveNext
                            Loop
                            gridSublet.DataSource = rsSublet
                            Call headSublet
                        Else
                            MsgBox "WO tidak memiliki pemakaian sublet "
                        End If
                    Else
                        MsgBox "WO telah dibuat invoice."
                    End If
                Else
                    MsgBox "WO tersebut tidak ada !!"
                End If
            End If
    End Select
End Sub


