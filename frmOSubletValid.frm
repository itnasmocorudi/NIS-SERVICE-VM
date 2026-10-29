VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmOSubletValid 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Pemakaian Sublet"
   ClientHeight    =   9330
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   9330
   ScaleWidth      =   11715
   Begin VB.CommandButton cmdSimpan 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Simpan"
      Height          =   495
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali Ke Menu"
      Height          =   495
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8520
      Width           =   1935
   End
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
      Height          =   8130
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   11085
      Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
         Bindings        =   "frmOSubletValid.frx":0000
         Height          =   1695
         Left            =   2400
         TabIndex        =   42
         Top             =   4800
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
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Sublet"
         Columns(3).DataField=   "PEKERJAAN"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
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
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid gridSublet 
         Height          =   4815
         Left            =   0
         TabIndex        =   39
         Top             =   3120
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
         Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=8196"
         Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
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
         TabIndex        =   2
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
         TabIndex        =   3
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
         TabIndex        =   4
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
         TabIndex        =   5
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
         TabIndex        =   6
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
         TabIndex        =   7
         Top             =   750
         Width           =   3375
         _ExtentX        =   5953
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         Left            =   10170
         TabIndex        =   10
         Top             =   750
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
         Index           =   2
         Left            =   3000
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         TabIndex        =   15
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
         TabIndex        =   16
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
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   10920
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   10920
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3360
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3360
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pemakaian Sublet"
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
         TabIndex        =   44
         Top             =   2685
         Width           =   2265
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
      Begin VB.Label SgnRekom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Rekom"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   210
         Left            =   8280
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   1290
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
         TabIndex        =   33
         Top             =   2280
         Width           =   645
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
         TabIndex        =   32
         Top             =   2280
         Width           =   435
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
         TabIndex        =   31
         Top             =   1980
         Width           =   315
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
         TabIndex        =   30
         Top             =   480
         Width           =   975
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
         TabIndex        =   28
         Top             =   1695
         Width           =   645
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
         TabIndex        =   27
         Top             =   480
         Width           =   660
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
         TabIndex        =   26
         Top             =   1365
         Width           =   480
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
         TabIndex        =   25
         Top             =   765
         Width           =   420
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
         TabIndex        =   24
         Top             =   465
         Width           =   645
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
         TabIndex        =   23
         Top             =   1695
         Width           =   375
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
         TabIndex        =   22
         Top             =   1395
         Width           =   330
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
         TabIndex        =   21
         Top             =   1095
         Width           =   495
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
         TabIndex        =   20
         Top             =   795
         Width           =   405
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
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   2640
         Width           =   11055
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2805
      TabIndex        =   36
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
      Height          =   9015
      Left            =   0
      Picture         =   "frmOSubletValid.frx":0015
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
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
      TabIndex        =   41
      Top             =   8520
      Width           =   1875
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
      TabIndex        =   40
      Top             =   8520
      Width           =   180
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
      Left            =   615
      TabIndex        =   37
      Top             =   165
      Width           =   570
   End
End
Attribute VB_Name = "frmOSubletValid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSublet As New ADODB.Recordset
Dim rsSearch As ADODB.Recordset
Dim rsHelp As ADODB.Recordset
Dim vPDPGL As String
Dim vHutangGL As String
Dim vPpnGL As String

Private Function createSublet() As Boolean
1         On Error GoTo vb_error
          
2         If rsSublet.State Then rsSublet.Close
3         Set rsSublet = Nothing
4         With rsSublet
5             .Fields.Append "NOPO", adVarChar, 25
6             .Fields.Append "SUPP", adVarChar, 15
7             .Fields.Append "SUBLET", adVarChar, 40
8             .Fields.Append "SAT", adVarChar, 10
9             .Fields.Append "QTY", adDouble
10            .Fields.Append "HARGASAT", adDouble
11            .Fields.Append "SUBTOTAL", adDouble
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
        
        .Columns(6).Caption = "Sub Total"
        .Columns(6).DataField = "SUBTOTAL"
        .Columns(6).Width = 1200
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

    With Adodc1
        .ConnectionString = SQLConACC
        .CommandType = adCmdText
        .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST WITH(NOLOCK)) ORDER BY KODEAP"
        .Refresh
    End With
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
12                If rsSublet("HARGASAT") = "0" Then
13                    If MsgBox("Yakin harga beli " & rsSublet("SUBLET") & " 0 (Nol) ?", vbYesNo, "Validasi Harga") = vbNo Then
14                        boleh = False
15                        pesan = pesan + "Harga Beli Satuan " & rsSublet("SUBLET") & ". " & vbCr
16                    End If
17                End If
                  
                  ''2016-07-22 || cek total bengkel dengan aptrn
18                sum_bengkel = sum_bengkel + (Abs(CDbl(rsSublet("HARGASAT"))) * CDbl(rsSublet("QTY")))
19                rsSublet.MoveNext
20            Loop
21        End If
          Dim rsTemp As ADODB.Recordset
          
          
22        rsSublet.MoveFirst
23        If CDbl(Right(NisDb, 2)) <> CDbl(Month(isnow)) Then
24            boleh = False
25            pesan = pesan + "Beda Fiskal Accounting tidak dapat dibuat invoice. " & vbCr & "Pindah Fiskal Accounting pada bulan bersangkutan. "
26        End If
          
          ''2015-11-25
          ''tambah validasi tempjobsublet untuk wo tsb harus kosong
          ''Masih ada user yang membuka menu order dengan no order = ''
27        Set rsHelp = objO.FetchData("SELECT * FROM TEMPJOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "'")
28        If Not rsHelp.EOF Then
29            rsHelp.MoveFirst
30            boleh = False
31            pesan = pesan + "Masih ada user yang membuka menu order dengan no order = " & rsHelp!no_wo
32        End If
33        If boleh Then
34             Set rsTemp = objO.FetchDataACC("select isnull(SUM(jumlahtrn),0) total from aptrn where nobukti like '" & "S-" & Trim(Text(0)) & "%'")
35             If Not rsTemp.EOF Then
36               rsTemp.MoveFirst
                 
                 ''If rsTemp!total <> sum_bengkel And rsTemp!total <> 0 Then
                 
37               If Not (rsTemp!total <> sum_bengkel And rsTemp!total = 0) Then
                    ''insert data ke selisih sublet
38                  Call objO.InsertData("INSERT INTO [SelisihSublet] ([no_wo] ,[no_chassis] ,[sum_bengkel] ,[sum_ap] ,[tglEntry],KodeUser,menu) VALUES ('" & Trim(Text(0)) & "','" & Trim(Text(15)) & "'," & sum_bengkel & "," & CDbl(rsTemp!total) & ",GETDATE(),'" & gSA & "','Validasi Sublet While Insert')")
39               End If
40             End If
              Dim jobSub(100) As String
              Dim iSub As Integer
41            iSub = 0
42            Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletValid", "cmdSimpan", 0)
43            erTranAll = False
44            onTR = True
45            Call objHT.OpenTransactionAll


              ' SIMPAN SUBLET
              
46            If rsSublet.State Then
47                If Not rsSublet.EOF Then
                      Dim isDyna As Boolean
                      'PENAMBAHAN PENGECEKAN PENYERAGAMAN AKUN
                      'TANGGAL 04-09-2014
                      'RUDI KRISTANTO
                      
48                    If Log.isAkun = 1 Then
49                        isDyna = IIf(InStr(1, UCase(Text(12).Text), "DYNA") > 0, True, False)
50                        Set rsGL = objO.Get_AkunGL_NEW_Sublet("OTHERS", isDyna, False)
51                        If isDyna = True Then
52                            vPDPGL = cekKodeAkun(Trim(rsGL.Fields(0).value))
53                            vPpnGL = cekKodeAkun(Trim(rsGL.Fields(1).value))
54                            vHutangGL = cekKodeAkun(Trim(rsGL.Fields(2).value))
55                        Else
56                            vPDPGL = cekKodeAkun(Trim(rsGL.Fields(0).value))
57                            vPpnGL = cekKodeAkun(Trim(rsGL.Fields(1).value))
58                            vHutangGL = cekKodeAkun(Trim(rsGL.Fields(2).value))
59                        End If
60                    Else
61                        Set rsGL = objHT.FetchDataBengkel("SELECT FS_KD_PDP, FS_KD_PPN, FS_KD_HUTANG FROM T_AKUN WITH(NOLOCK)")
62                        If Not rsGL.EOF Then
63                            vPDPGL = cekKodeAkun(Trim(rsGL("FS_KD_PDP")))
64                            vPpnGL = cekKodeAkun(Trim(rsGL("FS_KD_PPN")))
65                            vHutangGL = cekKodeAkun(Trim(rsGL("FS_KD_HUTANG")))
66                        End If
67                    End If
                      
                      
68                    Do While Not rsSublet.EOF
                          Dim sudahValid As Boolean
69                        sudahValid = False
70                        Set rsHelp = objHT.FetchDataBengkel("SELECT Valid FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'")
71                        If Not rsHelp.EOF Then
72                            If rsHelp("VALID") = 1 Then sudahValid = True
73                        End If
                          
74                        If Not sudahValid Then
                              ' INPUT PDP TERHADAP HUTANG SUPPLIER
                                  Dim vDelIndex As String
                                  Dim vKodeLgn As String
                                  Dim vTglTempo As Date
                                  Dim vTglSkr As Date
                                  Dim vTBeli As Double
                                  Dim vTPPN As Double
                                  Dim vTotal As Double
                                  Dim vKeterangan As String
                                  
75                                vDelIndex = gKode & Left(McDelIndex(7), 8)
76                                vKodeLgn = Trim(rsSublet("SUPP"))
77                                vTglSkr = Format(isnow, "dd/mm/yyyy hh:nn")
78                                vTglTempo = vTglSkr
79                                vTotal = Abs(CDbl(rsSublet("HARGASAT"))) * CDbl(rsSublet("QTY"))
80                                vTPPN = Round(vTotal * (10 / 110))
81                                vTBeli = vTotal - vTPPN
                                  
82                                If vPDPGL = "" Or vPpnGL = "" Or vHutangGL = "" Then
83                                    pesanIT = pesanIT + "Kode GL ada yang kosong !"
84                                    GoTo vbError
85                                End If
                                  
86                                noJurnal = False
87                                Set rsSearch = objHT.FetchDataBengkel("SELECT * FROM T_SUPPLIER WITH(NOLOCK) WHERE KODEAP = '" & Trim(rsSublet("SUPP")) & "'")
88                                If Not rsSearch.EOF Then
89                                    vKeterangan = rsSearch("SUPPLIER") & " (" & rsSublet("SUBLET") & ") No. WO :" & Trim(Text(0))
90                                    If Trim(rsSearch("PKP")) = "0" Then
                                          'vPpnGL = ""
91                                        vTBeli = vTBeli + vTPPN
92                                        vTPPN = 0
93                                        vTotal = vTBeli
94                                    End If
95                                    If Trim(rsSearch("NONJURNAL")) = "1" Then
96                                        noJurnal = True
97                                    End If
98                                End If
                                  
                                  ' update status sublet jadi valid
                                  'Call objHT.UpdateData_Bengkel("UPDATE JOBSUBLET SET HARGA = '" & Round(CDbl(vTotal) / CDbl(rsSublet("QTY"))) & "', SUPPLIER = '" & Trim(rsSublet("SUPP")) & "', VALID = 1, TGLVALID = GETDATE() WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'")
                                  'jobSub(iSub) = "UPDATE JOBSUBLET SET HARGA = '" & Round(CDbl(vTotal) / CDbl(rsSublet("QTY"))) & "', SUPPLIER = '" & Trim(rsSublet("SUPP")) & "', VALID = 1, TGLVALID = '" & Format(vTglSkr, "MM/DD/YYYY HH:NN") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'"
99                                jobSub(iSub) = "UPDATE JOBSUBLET SET HARGA = '" & Round(CDbl(vTotal) / CDbl(rsSublet("QTY"))) & "', SUPPLIER = '" & Trim(rsSublet("SUPP")) & "', VALID = 1, TGLVALID = '" & Format(isnow, "MM/DD/YYYY HH:NN") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(15)) & "' AND NAMABARANG = '" & Trim(rsSublet("SUBLET")) & "'"
100                               iSub = iSub + 1
                                  
101                               If gConnecAAS = True Then
102                                   If Not noJurnal Then
103                                       If Abs(CDbl(rsSublet("HARGASAT"))) <> 0 Then
                                              ' INPUT DI APTRN
                                              ' PDP
104                                           sSQL1 = "insert into APtrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                                                        & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, tglEntry,kodejurnal) " _
                                                        & "values " _
                                                        & " ('" & Trim(vKodeLgn) & "', '" & "S-" & Trim(Text(0)) & "', 'C'," _
                                                        & " '" & "S-" & Trim(Text(0)) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', 'PDP Sublet " & rsSublet("SUBLET") & " No. WO : " & Trim(Text(0)) & "'," _
                                                        & " " & SimpanKoma2Titik(vTBeli) & ", '" & gSA & "'," _
                                                        & " '" & vDelIndex & "','" & gKode & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vHutangGL & "')"
105                                           Call objHT.InsertData_ACC(sSQL1)
                                              
                                              ' PPN
106                                           sSQL1 = "insert into APtrn(KodeLgn, NoFaktur, TypeTrn, NoBukti, TglTrn, TglJthTmp, " _
                                                        & "TglTrnFaktur, TglJtpFaktur, Keterangan, JumlahTrn, Kodeuser, DelIndex, Kodesumber, tglEntry,kodejurnal) " _
                                                        & "values " _
                                                        & " ('" & Trim(vKodeLgn) & "', '" & "S-" & Trim(Text(0)) & "', 'F'," _
                                                        & " '" & "S-" & Trim(Text(0)) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', '" & FormatTgl(vTglSkr) & "'," _
                                                        & " '" & FormatTgl(vTglTempo) & "', 'PPN Sublet " & rsSublet("SUBLET") & " No. WO : " & Trim(Text(0)) & "'," _
                                                        & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "'," _
                                                        & " '" & vDelIndex & "','" & gKode & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & vHutangGL & "')"
107                                           Call objHT.InsertData_ACC(sSQL1)
                                              
                                              ' INPUT DI APMST
108                                           sSQL1 = "Update APMST set SALDOJALAN = SALDOJALAN + " & SimpanKoma2Titik(vTBeli) & ", TGLBELIAHR = '" & Format(isnow, "MM/DD/YYYY") & "' where KodeLgn= '" & Trim(vKodeLgn) & "'"
109                                           Call objHT.UpdateData_ACC(sSQL1)
                                              
110                                           sSQL1 = "Update APMST set SALDOJALAN = SALDOJALAN + " & SimpanKoma2Titik(vTPPN) & ", TGLBELIAHR = '" & Format(isnow, "MM/DD/YYYY") & "' where KodeLgn= '" & Trim(vKodeLgn) & "'"
111                                           Call objHT.UpdateData_ACC(sSQL1)
                      
                                              ' INPUT DI GLTRN
112                                           If vTotal > 0 Then
113                                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(Text(0)) & "', '" & vPDPGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTBeli) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
114                                               Call objHT.InsertData_ACC(sSQL1)
                                                  
115                                               If vTPPN > 0 Then
116                                                   sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhDebit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(Text(0)) & "', '" & vPpnGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTPPN) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
117                                                   Call objHT.InsertData_ACC(sSQL1)
118                                               End If
                                                  
119                                               sSQL1 = "insert into gltrn(NoBukti, KodeGl, TglTrn, KodeSumber, TypeTrn, " _
                                                        & " JlhKredit, KodeUser, KodeLgn, DelIndex, Keterangan, lastUpdated) values " _
                                                        & " ('" & "S-" & Trim(Text(0)) & "', '" & vHutangGL & "', '" & FormatTgl(vTglSkr) & "', '" & gKode & "', 'C', " _
                                                        & " " & SimpanKoma2Titik(vTotal) & ", '" & gSA & "', '" & vKodeLgn & "', " _
                                                        & " '" & vDelIndex & "','" & vKeterangan & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "')"
120                                               Call objHT.InsertData_ACC(sSQL1)
                                                  
                                                  ' INPUT DI GLMST
121                                               sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTBeli) & ", " _
                                                        & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(vTBeli) & " " _
                                                        & " where KodeGl = '" & vPDPGL & "'"
122                                               Call objHT.UpdateData_ACC(sSQL1)
                                                  
123                                               If vTPPN > 0 Then
124                                                   sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTPPN) & ", " _
                                                        & " JDebitBi = JDebitBi + " & SimpanKoma2Titik(vTPPN) & " " _
                                                        & " where KodeGl = '" & vPpnGL & "'"
125                                                   Call objHT.UpdateData_ACC(sSQL1)
126                                               End If
                                                  
127                                               sSQL1 = "update glmst set JbulanIni = JbulanIni + " & SimpanKoma2Titik(vTotal) & ", " _
                                                        & " JKreditBi = JKreditBi + " & SimpanKoma2Titik(vTotal) & " " _
                                                        & " where KodeGl = '" & vHutangGL & "'"
128                                               Call objHT.UpdateData_ACC(sSQL1)
129                                           End If
130                                       End If
131                                   End If
132                               End If
133                           End If
134                       rsSublet.MoveNext
135                   Loop
                      Dim xI As Integer
136                   For xI = 0 To 100
137                       If jobSub(xI) <> "" Then
138                           Call objHT.UpdateData_Bengkel(jobSub(xI))
139                       Else
140                           Exit For
141                       End If
142                   Next xI
143               End If
                  
144               If erTranAll = False Then
145                   Call objHT.CloseTransactionAll
                      ''2016-07-22
146                   Set rsTemp = objO.FetchDataACC("select isnull(SUM(jumlahtrn),0) total from aptrn where nobukti like '" & "S-" & Trim(Text(0)) & "%'")
147                     If Not rsTemp.EOF Then
148                       rsTemp.MoveFirst
                          ''If rsTemp!total <> sum_bengkel And rsTemp!total <> 0 Then
149                       If Not (rsTemp!total = sum_bengkel And rsTemp!total <> 0) Then
                             ''insert data ke selisih sublet
150                          Call objO.InsertData("INSERT INTO [SelisihSublet] ([no_wo] ,[no_chassis] ,[sum_bengkel] ,[sum_ap] ,[tglEntry],KodeUser,menu) VALUES ('" & Trim(Text(0)) & "','" & Trim(Text(15)) & "'," & sum_bengkel & "," & CDbl(rsTemp!total) & ",GETDATE(),'" & gSA & "','Validasi Sublet After Insert')")
151                       End If
152                     End If
153                   onTR = False
154                   Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletValid", "cmdSimpan", 1)
155                   MsgBox "Data berhasil disimpan "
156                   cmdSimpan.Enabled = False
157               Else
158                   Call objHT.CloseTransactionAll_Err
159                   onTR = False
160                   Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletValid", "cmdSimpan", 2)
161                   MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
162               End If

163           End If
164       Else
165           MsgBox "Data belum lengkap !!" & vbCr & pesan
166       End If
167       Exit Sub

vbError:
168       Call objHT.CloseTransactionAll_Err
169       onTR = False
170       Call saveLogMenu(gSA, ComptName, gVersi, "frmOSubletValid", "cmdSimpan", 2)
171       MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
172       Exit Sub
End Sub

Private Sub Form_Activate()
    With Adodc1
        .ConnectionString = SQLConACC
        .CommandType = adCmdText
        .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST WITH(NOLOCK)) ORDER BY KODEAP"
        .Refresh
    End With
    If orderAktif Or validAktif Or subRetAktif Then
        MsgBox "Tutup Menu sebelumnya terlebih Dahulu !!", vbCritical
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    subletAktif = True
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
    subletAktif = False
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
                Set rsHelp = objO.FetchData("SELECT * FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE KODEAP = '" & Trim(.Columns(1)) & "' AND FS_LOG_DEL = ''")
                If rsHelp.EOF Then
                    rsSublet.Find "SUBLET = '" & Trim(.Columns(2)) & "'", , adSearchForward, 1
                    If Not rsSublet.EOF Then
                        rsSublet("SUPP") = ""
                    End If
                Else
                    Set rsHlp = objO.FetchDataACC("SELECT * FROM APMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(.Columns(1)) & "'")
                    If Not rsHlp.EOF Then
                        rsSublet("SUPP") = UCase(Trim(.Columns(1)))
                    Else
                        MsgBox "Kode Supplier tidak terdaftar di Accounting !"
                        rsSublet.Find "SUBLET = '" & Trim(.Columns(2)) & "'", , adSearchForward, 1
                        If Not rsSublet.EOF Then
                            rsSublet("SUPP") = ""
                        End If
                    End If
                End If
            End With
    End Select
    Set gridSublet.DataSource = rsSublet
End Sub

Private Sub GridSublet_ButtonClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 1
            With Adodc1
                .ConnectionString = SQLConACC
                .CommandType = adCmdText
                .RecordSource = "SELECT KODEAP, SUPPLIER, ALAMAT1 + '' + ALAMAT2 AS ALAMAT, PEKERJAAN FROM " & gDb & "..T_SUPPLIER WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND KODEAP IN (SELECT KODELGN FROM [" & NisDb & "]..APMST WITH(NOLOCK)) ORDER BY KODEAP"
                .Refresh
            End With
    End Select
End Sub

Private Sub GridSublet_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 18
            With Adodc1
                .ConnectionString = SQLConACC
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
                gSearchOption = "VALIDSUBLET"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Function cariValid(nWO As String, nChassis As String) As Integer
Dim rsCari As ADODB.Recordset

    Set rsCari = objO.FetchData("SELECT SUM(VALID) AS OK FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = '" & Trim(nWO) & "' AND NO_CHASSIS = '" & Trim(nChassis) & "'")
    If Not rsCari.EOF Then
        cariValid = CDbl(rsCari("OK"))
    End If
End Function

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
1         Select Case Index
              Case 0
                  ' NO WO VALIDASI
2                 If Text(Index) <> "" Then
3                     Set rsSearch = objO.FetchData("SELECT B.FS_KD_PLG, B.FS_KD_WIL, B.CUSTOMER, B.CUST_ADRES, B.KOTA, B.TELPON, GSM, CDMA, B.FS_TH_MOBIL, B.MDL, " _
                                                 & "B.MODEL, B.FMDL, B.COLOR, B.NO_POLISI, B.NO_CHASSIS, C.NO_ENGINE, C.DDATE, NOPO, NAMABARANG, KODESATUAN, QTY, HARGA, SUPPLIER, NOINV, VALID FROM JOBSUBLET A WITH(NOLOCK) " _
                                                 & "INNER JOIN JOBDTA B WITH(NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS " _
                                                 & "INNER JOIN CUSTDAT C WITH(NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS WHERE B.NOINV = '' AND A.NO_WO = '" & Trim(Text(Index)) & "'ORDER BY A.NO_WO")     '
4                     If Not rsSearch.EOF Then
5                         Call hapusForm
6                         If Trim(rsSearch("NOINV")) = "" Then
7                             If CDbl(rsSearch.RecordCount) > cariValid(Trim(Text(Index)), Trim(rsSearch("NO_CHASSIS"))) Then
8                                 Text(1) = Trim(rsSearch("FS_KD_PLG"))
9                                 Text(2) = Trim(rsSearch("FS_KD_WIL"))
10                                Text(3) = Trim(rsSearch("CUSTOMER"))
11                                Text(4) = Left(Trim(rsSearch("CUST_ADRES")), 60)
12                                Text(5) = Trim(rsSearch("KOTA"))
13                                Text(6) = Trim(rsSearch("TELPON"))
14                                Text(7) = Trim(rsSearch("GSM"))
15                                Text(8) = Trim(rsSearch("CDMA"))
16                                Text(9) = Trim(rsSearch("FS_TH_MOBIL"))
17                                Text(10) = Trim(rsSearch("MDL"))
18                                Text(11) = Trim(rsSearch("MODEL"))
19                                Text(12) = Trim(rsSearch("FMDL"))
20                                Text(13) = Trim(rsSearch("COLOR"))
21                                Text(14) = Trim(rsSearch("NO_POLISI"))
22                                Text(15) = Trim(rsSearch("NO_CHASSIS"))
23                                Text(16) = Trim(rsSearch("NO_ENGINE"))
24                                Text(17).Mask = ""
25                                Text(17) = Nul(IIf(Replace(rsSearch("DDATE"), "/", "-") = "01-01-3000", "", IIf(Replace(rsSearch("DDATE"), "/", "-") = "01-01-1900", "", Replace(rsSearch("DDATE"), "/", "-"))))
                                  'Trim(rsSearch("DDATE"))
                                  
26                                Call createSublet
27                                Do While Not rsSearch.EOF
28                                    If rsSearch("VALID") = 0 Then
29                                        If rsSublet.State Then
30                                            With rsSublet
31                                                .AddNew
32                                                !NOPO = Trim(rsSearch("NOPO"))
33                                                !SUPP = Trim(rsSearch("SUPPLIER"))
34                                                !SUBLET = Trim(rsSearch("NAMABARANG"))
35                                                !SAT = Trim(rsSearch("KODESATUAN"))
36                                                !qty = CDbl(Trim(rsSearch("QTY")))
37                                                !HARGASAT = CDbl(Trim(rsSearch("HARGA")))
38                                                !Subtotal = CDbl(Trim(rsSearch("QTY"))) * CDbl(Trim(rsSearch("HARGA")))
39                                                .Update
40                                            End With
41                                            grandTotal = Format(grandTotal + CDbl(Trim(rsSearch("QTY"))) * CDbl(Trim(rsSearch("HARGA"))), "##,0")
42                                        End If
43                                    End If
44                                    rsSearch.MoveNext
45                                Loop
46                                cmdSimpan.Enabled = True
47                                Set gridSublet.DataSource = rsSublet
48                                Call headSublet
49                            Else
50                                MsgBox "WO telah divalidasi "
51                            End If
52                        Else
53                            MsgBox "WO telah dibuat invoice. Retur Invoice"
54                        End If
55                    Else
56                        MsgBox "WO tersebut tidak ada !!"
57                    End If
58                End If
59        End Select
End Sub
