VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoHarian 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Laporan Harian"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H000000FF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   165
      Style           =   1  'Graphical
      TabIndex        =   75
      Top             =   6015
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1275
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   6015
      Width           =   1440
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Cancel          =   -1  'True
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2760
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   6015
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdCetak 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8205
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   6015
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5085
      TabIndex        =   71
      Top             =   6045
      Visible         =   0   'False
      Width           =   3015
   End
   Begin ucXPButton.XPButton XPButton1 
      Height          =   300
      Left            =   3975
      TabIndex        =   70
      Top             =   6030
      Visible         =   0   'False
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   529
      BTYPE           =   8
      TX              =   "Printer :"
      ENAB            =   0   'False
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
      MICON           =   "frmRepoHarian.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   24
      Top             =   5760
      Visible         =   0   'False
      Width           =   12375
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame FrRate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9255
      Left            =   5400
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
      Begin TrueOleDBGrid70.TDBGrid GridFR2 
         Height          =   7245
         Left            =   3360
         TabIndex        =   68
         Top             =   360
         Visible         =   0   'False
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   12779
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tanggal"
         Columns(0).DataField=   "Tgl"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Jml WO"
         Columns(1).DataField=   "NoWO"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Full Sup"
         Columns(2).DataField=   "FullSup"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Order SA"
         Columns(3).DataField=   "QtyOrderSA"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "FR"
         Columns(4).DataField=   "FR"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Li Supply"
         Columns(5).DataField=   "Li_Supply"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Li BO"
         Columns(6).DataField=   "Li_BO"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Li Cancel"
         Columns(7).DataField=   "Li_Cancel"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "SR"
         Columns(8).DataField=   "SR"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1667"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1667"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1667"
         Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1667"
         Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(24)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=1667"
         Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(29)=   "Column(6).Width=1773"
         Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=1667"
         Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(34)=   "Column(7).Width=1773"
         Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=1667"
         Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(39)=   "Column(8).Width=1773"
         Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=1667"
         Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=1"
         Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         ColumnFooters   =   -1  'True
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=104,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=101,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=102,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=103,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=78,.parent=13,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=86,.parent=13,.alignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=83,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=84,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=85,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=90,.parent=13,.alignment=2"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=87,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=88,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=89,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=98,.parent=13,.alignment=2"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
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
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
         _StyleDefs(93)  =   "Named:id=99:Normal0"
         _StyleDefs(94)  =   ":id=99,.parent=42,.bgcolor=&HFFFFCE&,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(95)  =   ":id=99,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(96)  =   ":id=99,.fontname=MS Sans Serif"
         _StyleDefs(97)  =   "Named:id=100:Normal1"
         _StyleDefs(98)  =   ":id=100,.parent=99,.bgcolor=&H55FFFF&"
         _StyleDefs(99)  =   "Named:id=0:"
         _StyleDefs(100) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(101) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(102) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(103) =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(104) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(105) =   ":id=0,.fontname=MS Sans Serif"
      End
      Begin TrueOleDBGrid70.TDBGrid GridFR3 
         Height          =   7245
         Left            =   2040
         TabIndex        =   62
         Top             =   360
         Visible         =   0   'False
         Width           =   10770
         _ExtentX        =   18997
         _ExtentY        =   12779
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No"
         Columns(0).DataField=   "No"
         Columns(0).DataWidth=   10
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No Wo"
         Columns(1).DataField=   "NoWo"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Part Nomor"
         Columns(2).DataField=   "PartNomor"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Part Name"
         Columns(3).DataField=   "PartName"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Kode SCC"
         Columns(4).DataField=   "KodeSCC"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "QTY"
         Columns(5).DataField=   "qty"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Unit Price"
         Columns(6).DataField=   "UnitPrice"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "T / Price"
         Columns(7).DataField=   "TPrice"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Remark"
         Columns(8).DataField=   "Remark"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1402"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1296"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4419"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4313"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=5292"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=5186"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=0"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3545"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3440"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=0"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(26)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=1667"
         Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=0"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(31)=   "Column(6).Width=3545"
         Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=3440"
         Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(36)=   "Column(7).Width=3889"
         Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=3784"
         Splits(0)._ColumnProps(39)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(40)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(41)=   "Column(8).Width=1773"
         Splits(0)._ColumnProps(42)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(8)._WidthInPix=1667"
         Splits(0)._ColumnProps(44)=   "Column(8)._ColStyle=1"
         Splits(0)._ColumnProps(45)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         ColumnFooters   =   -1  'True
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.alignment=0"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=0"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=70,.parent=13,.alignment=0"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=67,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=68,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=69,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=82,.parent=13,.alignment=1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=79,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=80,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=81,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=86,.parent=13,.alignment=2"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=17"
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
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
         _StyleDefs(93)  =   "Named:id=99:Normal0"
         _StyleDefs(94)  =   ":id=99,.parent=42,.bgcolor=&HFFFFCE&,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(95)  =   ":id=99,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(96)  =   ":id=99,.fontname=MS Sans Serif"
         _StyleDefs(97)  =   "Named:id=100:Normal1"
         _StyleDefs(98)  =   ":id=100,.parent=99,.bgcolor=&H55FFFF&"
         _StyleDefs(99)  =   "Named:id=0:"
         _StyleDefs(100) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(101) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(102) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(103) =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(104) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(105) =   ":id=0,.fontname=MS Sans Serif"
      End
      Begin TrueOleDBGrid70.TDBGrid GridFR1 
         Height          =   7245
         Left            =   840
         TabIndex        =   57
         Top             =   360
         Visible         =   0   'False
         Width           =   10410
         _ExtentX        =   18362
         _ExtentY        =   12779
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tanggal"
         Columns(0).DataField=   "Tgl"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Jml WO"
         Columns(1).DataField=   "NoWO"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Full Sup"
         Columns(2).DataField=   "FullSup"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Order SA"
         Columns(3).DataField=   "QtyOrderSA"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "FR"
         Columns(4).DataField=   "FR"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Li Supply"
         Columns(5).DataField=   "Li_Supply"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Li BO"
         Columns(6).DataField=   "Li_BO"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Li Cancel"
         Columns(7).DataField=   "Li_Cancel"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "SR"
         Columns(8).DataField=   "SR"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1667"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1667"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1667"
         Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1667"
         Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(24)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=1667"
         Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(29)=   "Column(6).Width=1773"
         Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=1667"
         Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(34)=   "Column(7).Width=1773"
         Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=1667"
         Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(39)=   "Column(8).Width=1773"
         Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=1667"
         Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=1"
         Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         ColumnFooters   =   -1  'True
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=104,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=101,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=102,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=103,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=78,.parent=13,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=86,.parent=13,.alignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=83,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=84,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=85,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=90,.parent=13,.alignment=2"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=87,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=88,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=89,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=98,.parent=13,.alignment=2"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
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
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
         _StyleDefs(93)  =   "Named:id=99:Normal0"
         _StyleDefs(94)  =   ":id=99,.parent=42,.bgcolor=&HFFFFCE&,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(95)  =   ":id=99,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(96)  =   ":id=99,.fontname=MS Sans Serif"
         _StyleDefs(97)  =   "Named:id=100:Normal1"
         _StyleDefs(98)  =   ":id=100,.parent=99,.bgcolor=&H55FFFF&"
         _StyleDefs(99)  =   "Named:id=0:"
         _StyleDefs(100) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(101) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(102) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(103) =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(104) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(105) =   ":id=0,.fontname=MS Sans Serif"
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1575
         Left            =   120
         TabIndex        =   45
         Top             =   7680
         Width           =   11055
         Begin VB.OptionButton Opt 
            Caption         =   "Lost Sales"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   8400
            TabIndex        =   67
            Top             =   720
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker dtSelesai 
            Height          =   375
            Left            =   3960
            TabIndex        =   66
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   98828289
            CurrentDate     =   41802
         End
         Begin MSComCtl2.DTPicker dtMulai 
            Height          =   375
            Left            =   1920
            TabIndex        =   64
            Top             =   120
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
            _Version        =   393216
            Format          =   98828289
            CurrentDate     =   41802
         End
         Begin VB.ComboBox cboJob 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Rekap Bulanan"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   6600
            TabIndex        =   61
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Rekap Harian"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   5040
            TabIndex        =   60
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Opt 
            Caption         =   "Detail Harian"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   3480
            TabIndex        =   59
            Top             =   720
            Value           =   -1  'True
            Width           =   1815
         End
         Begin ucXPButton.XPButton cmdRef 
            Height          =   495
            Left            =   9360
            TabIndex        =   46
            Top             =   120
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
            MICON           =   "frmRepoHarian.frx":001C
            PICN            =   "frmRepoHarian.frx":0038
            PICH            =   "frmRepoHarian.frx":05B4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label9 
            Caption         =   " s/d "
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3360
            TabIndex        =   69
            Top             =   120
            Width           =   495
         End
         Begin VB.Label Label3 
            Caption         =   "Range Tanggal"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   65
            Top             =   120
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "job Desc  :                            Filter  :"
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
            Left            =   0
            TabIndex        =   58
            Top             =   720
            Width           =   3405
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000080FF&
            X1              =   0
            X2              =   10800
            Y1              =   0
            Y2              =   0
         End
      End
      Begin TrueOleDBGrid70.TDBGrid GridFR 
         Height          =   7245
         Left            =   120
         TabIndex        =   47
         Top             =   360
         Width           =   14250
         _ExtentX        =   25135
         _ExtentY        =   12779
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tanggal"
         Columns(0).DataField=   "Tgl"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No WO"
         Columns(1).DataField=   "NoWO"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Job Desc"
         Columns(2).DataField=   "Job"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "No Polisi"
         Columns(3).DataField=   "NoPol"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "No Rangka"
         Columns(4).DataField=   "NoRangka"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Model"
         Columns(5).DataField=   "Model"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Kode Part"
         Columns(6).DataField=   "Kode"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Nama Part"
         Columns(7).DataField=   "Nama"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Order SA"
         Columns(8).DataField=   "QtyOrderSA"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Supply"
         Columns(9).DataField=   "QtySupply"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "FR_Li"
         Columns(10).DataField=   "FR_Li"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "FR"
         Columns(11).DataField=   "FR"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Li Supply"
         Columns(12).DataField=   "Li_Supply"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Li BO"
         Columns(13).DataField=   "Li_BO"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Li Cancel"
         Columns(14).DataField=   "Li_Cancel"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "SR_Li"
         Columns(15).DataField=   "SR_Li"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "SR"
         Columns(16).DataField=   "SR"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   17
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=17"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2461"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2355"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1667"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=2302"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2196"
         Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=0"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(4).Width=4048"
         Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=3942"
         Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=0"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(24)=   "Column(5).Width=5292"
         Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=5186"
         Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=0"
         Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(29)=   "Column(6).Width=3175"
         Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=3069"
         Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=0"
         Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(34)=   "Column(7).Width=3519"
         Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=3413"
         Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=0"
         Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(39)=   "Column(8).Width=1773"
         Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=1667"
         Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=1"
         Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(44)=   "Column(9).Width=1773"
         Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=1667"
         Splits(0)._ColumnProps(47)=   "Column(9)._ColStyle=1"
         Splits(0)._ColumnProps(48)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(49)=   "Column(10).Width=1773"
         Splits(0)._ColumnProps(50)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(10)._WidthInPix=1667"
         Splits(0)._ColumnProps(52)=   "Column(10)._ColStyle=1"
         Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(54)=   "Column(11).Width=1773"
         Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=1667"
         Splits(0)._ColumnProps(57)=   "Column(11)._ColStyle=1"
         Splits(0)._ColumnProps(58)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(59)=   "Column(12).Width=1773"
         Splits(0)._ColumnProps(60)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(61)=   "Column(12)._WidthInPix=1667"
         Splits(0)._ColumnProps(62)=   "Column(12)._ColStyle=1"
         Splits(0)._ColumnProps(63)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(64)=   "Column(13).Width=1773"
         Splits(0)._ColumnProps(65)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(13)._WidthInPix=1667"
         Splits(0)._ColumnProps(67)=   "Column(13)._ColStyle=1"
         Splits(0)._ColumnProps(68)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(69)=   "Column(14).Width=1773"
         Splits(0)._ColumnProps(70)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(14)._WidthInPix=1667"
         Splits(0)._ColumnProps(72)=   "Column(14)._ColStyle=1"
         Splits(0)._ColumnProps(73)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(74)=   "Column(15).Width=1773"
         Splits(0)._ColumnProps(75)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(76)=   "Column(15)._WidthInPix=1667"
         Splits(0)._ColumnProps(77)=   "Column(15)._ColStyle=1"
         Splits(0)._ColumnProps(78)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(79)=   "Column(16).Width=1773"
         Splits(0)._ColumnProps(80)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(16)._WidthInPix=1667"
         Splits(0)._ColumnProps(82)=   "Column(16)._ColStyle=1"
         Splits(0)._ColumnProps(83)=   "Column(16).Order=17"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         ColumnFooters   =   -1  'True
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HC1EDFF&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=104,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=101,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=102,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=103,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=0"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=2"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=82,.parent=13,.alignment=2"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=86,.parent=13,.alignment=2"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=90,.parent=13,.alignment=2"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
         _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=94,.parent=13,.alignment=2"
         _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
         _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
         _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
         _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=98,.parent=13,.alignment=2"
         _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14"
         _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
         _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
         _StyleDefs(104) =   "Named:id=33:Normal"
         _StyleDefs(105) =   ":id=33,.parent=0"
         _StyleDefs(106) =   "Named:id=34:Heading"
         _StyleDefs(107) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(108) =   ":id=34,.wraptext=-1"
         _StyleDefs(109) =   "Named:id=35:Footing"
         _StyleDefs(110) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(111) =   "Named:id=36:Selected"
         _StyleDefs(112) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(113) =   "Named:id=37:Caption"
         _StyleDefs(114) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(115) =   "Named:id=38:HighlightRow"
         _StyleDefs(116) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(117) =   "Named:id=39:EvenRow"
         _StyleDefs(118) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(119) =   "Named:id=40:OddRow"
         _StyleDefs(120) =   ":id=40,.parent=33"
         _StyleDefs(121) =   "Named:id=41:RecordSelector"
         _StyleDefs(122) =   ":id=41,.parent=34"
         _StyleDefs(123) =   "Named:id=42:FilterBar"
         _StyleDefs(124) =   ":id=42,.parent=33"
         _StyleDefs(125) =   "Named:id=99:Normal0"
         _StyleDefs(126) =   ":id=99,.parent=42,.bgcolor=&HFFFFCE&,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(127) =   ":id=99,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(128) =   ":id=99,.fontname=MS Sans Serif"
         _StyleDefs(129) =   "Named:id=100:Normal1"
         _StyleDefs(130) =   ":id=100,.parent=99,.bgcolor=&H55FFFF&"
         _StyleDefs(131) =   "Named:id=105:Normal3"
         _StyleDefs(132) =   ":id=105,.parent=100,.bgcolor=&HFFCACA&"
      End
      Begin ucXPButton.XPButton cmdClose 
         Height          =   495
         Left            =   13200
         TabIndex        =   48
         Top             =   7800
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
         MICON           =   "frmRepoHarian.frx":0BEB
         PICN            =   "frmRepoHarian.frx":0C07
         PICH            =   "frmRepoHarian.frx":10DB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdExportFill 
         Height          =   495
         Left            =   12000
         TabIndex        =   49
         Top             =   7800
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
         MICON           =   "frmRepoHarian.frx":1626
         PICN            =   "frmRepoHarian.frx":1642
         PICH            =   "frmRepoHarian.frx":1BD7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Ket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".::  SERVICE RATE  DAN FILL RATE ::."
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
         Index           =   9
         Left            =   5295
         TabIndex        =   50
         Top             =   60
         Width           =   3885
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000080FF&
         X1              =   14400
         X2              =   120
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   14295
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoHarian.frx":2212
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   23
      Top             =   0
      Width           =   445
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   3495
      Left            =   10080
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1425
         TabIndex        =   1
         Top             =   360
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   2295
         TabIndex        =   2
         Top             =   360
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   5295
         TabIndex        =   3
         Top             =   720
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
         Index           =   1
         Left            =   2415
         TabIndex        =   4
         Top             =   360
         Width           =   2955
         _ExtentX        =   5212
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
         Index           =   3
         Left            =   1425
         TabIndex        =   5
         Top             =   1800
         Width           =   795
         _ExtentX        =   1402
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
      Begin MSMask.MaskEdBox MaskEdBox5 
         Height          =   270
         Left            =   2295
         TabIndex        =   6
         Top             =   1800
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
         Index           =   4
         Left            =   2415
         TabIndex        =   7
         Top             =   1800
         Width           =   3195
         _ExtentX        =   5636
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
         Left            =   1425
         TabIndex        =   8
         Top             =   1080
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   2775
         TabIndex        =   9
         Top             =   1080
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   1410
         TabIndex        =   10
         Top             =   1440
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   10
         Left            =   2280
         TabIndex        =   11
         Top             =   1440
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   503
         _Version        =   393216
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
         Index           =   2
         Left            =   1440
         TabIndex        =   12
         Top             =   720
         Width           =   3795
         _ExtentX        =   6694
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
      Begin ComctlLib.ProgressBar oLoad 
         Height          =   480
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   847
         _Version        =   327682
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   14
         Top             =   2160
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   3240
         TabIndex        =   15
         Top             =   2160
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
         Index           =   6
         Left            =   3360
         TabIndex        =   16
         Top             =   2160
         Width           =   3915
         _ExtentX        =   6906
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Penjual"
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
         Left            =   300
         TabIndex        =   22
         Top             =   1845
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Order"
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
         Index           =   0
         Left            =   615
         TabIndex        =   21
         Top             =   765
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Job Order"
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
         TabIndex        =   20
         Top             =   405
         Width           =   1200
      End
      Begin VB.Label Label2 
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
         Left            =   750
         TabIndex        =   19
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Delivery"
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
         TabIndex        =   18
         Top             =   1485
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Kendaraan :"
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
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   2160
         Width           =   1350
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   0
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   661
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   480
      TabIndex        =   25
      Top             =   480
      Width           =   6735
      Begin VB.OptionButton opLap 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rinci SA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   27
         Top             =   1080
         Width           =   1215
      End
      Begin VB.OptionButton opLap 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Global"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2040
         TabIndex        =   26
         Top             =   1080
         Value           =   -1  'True
         Width           =   975
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93585409
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   375
         Left            =   4200
         TabIndex        =   29
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   93585409
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   405
         Index           =   11
         Left            =   2040
         TabIndex        =   30
         Top             =   600
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   714
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox3 
         Height          =   375
         Left            =   3000
         TabIndex        =   31
         Top             =   600
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   405
         Index           =   12
         Left            =   3120
         TabIndex        =   32
         Top             =   600
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   714
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Laporan"
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
         Left            =   360
         TabIndex        =   38
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   6240
         X2              =   120
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Label Label8 
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
         Left            =   1920
         TabIndex        =   37
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label8 
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
         Index           =   2
         Left            =   1920
         TabIndex        =   36
         Top             =   120
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode SA"
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
         Index           =   2
         Left            =   360
         TabIndex        =   35
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label8 
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
         Index           =   3
         Left            =   1920
         TabIndex        =   34
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Range Data                                     s/d"
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
         Index           =   4
         Left            =   360
         TabIndex        =   33
         Top             =   120
         Width           =   3675
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         X1              =   6240
         X2              =   120
         Y1              =   1560
         Y2              =   1560
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Stop"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5400
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   1680
      Width           =   1050
   End
   Begin ucXPButton.XPButton tblSimpan 
      Height          =   495
      Left            =   2520
      TabIndex        =   41
      Top             =   2160
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&p"
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
      MICON           =   "frmRepoHarian.frx":4222
      PICN            =   "frmRepoHarian.frx":423E
      PICH            =   "frmRepoHarian.frx":4695
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton tblRefresh 
      Height          =   495
      Left            =   3720
      TabIndex        =   40
      Top             =   2160
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
      MICON           =   "frmRepoHarian.frx":4B7F
      PICN            =   "frmRepoHarian.frx":4B9B
      PICH            =   "frmRepoHarian.frx":5117
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton tblKeluar 
      Height          =   495
      Left            =   5520
      TabIndex        =   42
      Top             =   4560
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
      MICON           =   "frmRepoHarian.frx":574E
      PICN            =   "frmRepoHarian.frx":576A
      PICH            =   "frmRepoHarian.frx":5C3E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1575
      Left            =   600
      TabIndex        =   51
      Top             =   3000
      Width           =   7335
      Begin VB.ComboBox cboBlnSR 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   480
         Width           =   855
      End
      Begin VB.ComboBox cboThnSR 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   480
         Width           =   1095
      End
      Begin ucXPButton.XPButton cmdPrintSR 
         Height          =   495
         Left            =   4080
         TabIndex        =   54
         Top             =   480
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
         MICON           =   "frmRepoHarian.frx":6189
         PICN            =   "frmRepoHarian.frx":61A5
         PICH            =   "frmRepoHarian.frx":65FC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODE  :"
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
         Index           =   4
         Left            =   720
         TabIndex        =   56
         Top             =   600
         Width           =   1020
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H000080FF&
         FillColor       =   &H00FF0000&
         Height          =   735
         Left            =   0
         Top             =   360
         Width           =   6135
      End
      Begin VB.Label Ket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".::  SERVICE RATE DAN FILL RATE  ::."
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
         Index           =   7
         Left            =   1080
         TabIndex        =   55
         Top             =   120
         Width           =   3885
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   6135
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000080FF&
         X1              =   -7440
         X2              =   6120
         Y1              =   1200
         Y2              =   1200
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   6720
      X2              =   600
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmRepoHarian.frx":6AE6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label nLabel 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "\(^_^)/  Please Wait"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   43
      Top             =   3000
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmRepoHarian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 13 November 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

' Outstanding Part
'''Select A.NoFK, A.tglFK, A.NO_WO, A.Customer, A.No_Polisi, A.Front as SA,
'''B.KodeUser as User_Part, B.tglEntry as Tgl_Entry_Part
'''from Bengkel05..Jobdta A Inner Join [acc05-201204]..SIHDR B On A.NoFK = B.NoFaktur Where
'''A.NOFK <> '' and A.NoInv ='' and A.NoFKRetur ='' and (A.[date] between '04/01/2012' and '05/01/2012')



Dim nStop As Boolean
Dim vNamaTabel As String
Dim Report As Object
Private rsSR  As New ADODB.Recordset
Private rsSR1 As New ADODB.Recordset
Private rsSR2 As New ADODB.Recordset
Private rsFR  As New ADODB.Recordset
Private rsGetData As New ADODB.Recordset
Private rsSR3 As New ADODB.Recordset

Private getKodeItem(1000) As String
Private iKodeItem As Integer
''@@@ 2015-04-25 || Penambahan untuk lost sales @@@@
Private getNosoBO(1000) As String
Private iNosoBO As Integer

Private Sub hapusForm()
    For i = 0 To 12
        Select Case i
            Case 7, 8
                Text(i) = ">>->>->>>>"
            Case 9, 10
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    Tgl1 = isnow
    Tgl2 = isnow
    Command1.Visible = False
    cboJob.Clear
    cboJob.AddItem "ALL"
    cboJob.AddItem "SBI"
    cboJob.AddItem "SBE"
    cboJob.AddItem "WAR"
    cboJob.AddItem "GRP"
    
    cboThnSR.Clear
    For i = 2014 To year(isnow)
        cboThnSR.AddItem i
    Next i
    cboThnSR.Text = year(isnow)
    
    cboBlnSR.Clear
    For i = 1 To 12
        cboBlnSR.AddItem i
    Next i
    cboBlnSR.Text = Month(isnow)
       
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
        
End Sub

Private Sub CetakTotal()
Dim oRs As Object
Screen.MousePointer = vbHourglass
Dim Report As Object

If opLap(0).value Then
    Set Report = New rptRepoHarian
    sSQL = "Select Tanggal as [DATE], Sum(SO) as SO, Sum(INV) as INV, Sum(Jasa) as jasa ,Sum(Part) as Part, Sum(Oli) as Oli, " & _
           "Sum(Sublet) as Sublet, Sum(Other) as Other, Sum(PDS) as PDS, Sum(SBI) as SBI, Sum(TWC) as TWC, " & _
           "Sum(RTJ) as RTJ, Sum(SBE) as SBE, Sum(GRP) as GRP, Sum(Lub) as LUB " & _
           "FROm REPOHSA WITH(NOLOCK) Where Tanggal BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' Group By Tanggal"
Else
    Set Report = New rptRepoHarianSA
    
    sSQL = "SELECT SA, Tanggal as DATE, SO, INV, Jasa,PART, OLI, SUBLET, OTHER, PDS, SBI, TWC, RTJ, SBE, GRP, LUB FROM RepoHSA WITH(NOLOCK) " & _
           "WHERE Tanggal BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' ORDER BY SA, TANGGAL"
           
    Dim sHlp As Object
    Dim crxSubReportHistory As CRAXDRT.Report
        
    sSQL1 = "SELECT SA, Tanggal as DATE, SO, INV,Jasa, PART, OLI, SUBLET, OTHER, PDS, SBI, TWC, RTJ, SBE, GRP, LUB FROM RepoHSA WITH(NOLOCK) " & _
           "WHERE Tanggal BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' ORDER BY SA, TANGGAL"
           
    Set sHlp = objO.FetchData(sSQL1)
    Set crxSubReportHistory = Report.Subreport1.OpenSubreport
    crxSubReportHistory.database.SetDataSource sHlp

End If
    

    Set oRs = objO.FetchData(sSQL)
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
    If opLap(0).value Then
        If Trim(Text(12)) <> "" Then
            Report.txtSA.SetText "oleh SA : " & Trim(Text(12)) & " [" & Trim(Text(11)) & "] "
        End If
    End If
    oLoad.Visible = False
    DoEvents
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClose_Click()
    FrRate.Visible = False
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdCetak.Visible = False
End Sub

 

Private Sub Command1_Click()
    nStop = True
    TblSimpan.Enabled = True
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Sub Form_Load()
    Call hapusForm
    vNamaTabel = "MapCust" & DriveSerialNumber("c:")
End Sub

Private Sub TblKeluar_Click()
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call hapusForm
End Sub

Private Sub tblSimpan_Click()
nLabel.Visible = True
DoEvents
TblSimpan.Enabled = False
'Dim txtTemp As String
'Dim txt2 As String
'    nStop = False
'    txtTemp = ""
'    Command1.Visible = True
'    If Text(0) <> "" Then
'        txtTemp = txtTemp & " AND X.KLPKERJA1+X.KLPKERJA2+X.KLPKERJA3+X.KLPKERJA4+X.KLPKERJA5+X.KLPKERJA6+X.KLPKERJA7+X.KLPKERJA8 LIKE '%" & Text(0) & "%' "
'        txt2 = txt2 & " AND Y.KLPKERJA1+Y.KLPKERJA2+Y.KLPKERJA3+Y.KLPKERJA4+Y.KLPKERJA5+Y.KLPKERJA6+Y.KLPKERJA7 LIKE '%" & Text(0) & "%' "
'    End If
'    If Text(2) <> "" Then
'        txtTemp = txtTemp & " AND X.REP1+X.REP2+X.REP3+X.REP4+X.REP5+X.REP6+X.REP7+X.REP8 LIKE '%" & Text(2) & "%' "
'        txt2 = txt2 & " AND Y.REP1+Y.REP2+Y.REP3+Y.REP4+Y.REP5+Y.REP6+Y.REP7 LIKE '%" & Text(2) & "%' "
'    End If
'    If Text(5) <> "" Then
'        txtTemp = txtTemp & " AND X.MDL = '" & Text(5) & "' "
'        txt2 = txt2 & " AND Y.MDL = '" & Text(5) & "' "
'    End If
'    If Text(3) <> "" Then
'        txtTemp = txtTemp & " AND y.FS_KD_DEALER = '" & Text(3) & "' "
'        txt2 = ""
'    End If
'    If Text(7) <> ">>->>->>>>" And Text(8) <> ">>->>->>>>" Then
'        txtTemp = txtTemp & " AND Y.DOTAM BETWEEN '" & Format(Trim(Text(7)), "MM/DD/YYYY") & "' AND '" & Format(Trim(Text(8)), "MM/DD/YYYY") & "'"
'        txt2 = ""
'    End If
'    If Text(9) > 0 And Text(10) > 0 Then
'        txtTemp = txtTemp & " AND RIGHT(Y.DDATE, 4) >= '" & CDbl(Trim(Text(9))) & "' AND RIGHT(Y.DDATE, 4) <= '" & CDbl(Trim(Text(10))) & "'"
'        txt2 = ""
'    End If
'    If Text(11) <> "" Then
'        txtTemp = txtTemp & " AND x.FRONT = '" & Trim(Text(11)) & "'"
'        txt2 = txt2 & " AND Y.FRONT = '" & Trim(Text(11)) & "'"
'    End If
'
'    If txt2 = "" Then
        Call LoadDataLaporan2
'    Else
'        Call Cetak(vNamaTabel, txtTemp, txt2)
'    End If
    TblSimpan.Enabled = True
    nLabel.Visible = False
End Sub

Private Sub LoadDataLaporan2()
          Dim oRs As Object
1         Screen.MousePointer = vbHourglass
          
          Dim strSql As String
2         If opLap(1).value = True Then
          '------------------------------------------------------------------------------------
          '-=========== 2015-07-14 | RUDI | REVISI LAPORAN GR TANPA SBE DAN SBI ===------------
          '------------------------------------------------------------------------------------
3         strSql = "Select G.FRONT, G.Tanggal, sum(G.SO)AS SO " & _
                  ", sum(G.INV) as INV, sum(G.JASA) as Jasa, Sum(G.Genuine) as Genuine, Sum(G.NonGenuine) as NonGenuine, " & _
                  "Sum(G.Sublet) as Sublet,Sum(G.Others) as Others, sum(G.PDS) as PDS, sum(G.SBI) as SBI, sum(G.TWC) as TWC, sum(G.RTJ) as RTJ,  sum(G.SBE) as SBE, " & _
                  "sum(G.GRP) as GRP, sum(G.LUB) as LUB FROM ( "
4         strSql = strSql & "Select X.FRONT, X.Tanggal , sum(isnull(SO,0))AS SO " & _
                  ", sum(jmlInv) as INV, sum(X.Labor) as Jasa, Sum(X.Genuine) as Genuine, Sum(X.NonGenuine) as NonGenuine, " & _
                  "Sum(X.Sublet) as Sublet,Sum(X.Others) as Others, sum(PDS) as PDS, sum(SBI) as SBI, sum(TWC) as TWC, sum(RTJ) as RTJ,  sum(SBE) as SBE, " & _
                  "sum(GRP) as GRP, sum(LUB) as LUB from " & _
                  "( " & _
                  "Select A.FRONT, InvDate as Tanggal , isnull(B.WO,0) as SO, " & _
                  " sUM(A.Labor) as Labor, cOUNT(DISTINCT(rEPLACE(REpLACE(A.nOiNV,'-A',''),'-B',''))) as JmlInv, " & _
                  "SUM(ISNULL(A.PART_G+A.OIL_G,0))  as Genuine, SUM(ISNULL(A.PART_NG+A.OIL_NG,0)) as NonGenuine, SUM(ISNULL(A.SBLT,0)) as Sublet, " & _
                  "SUM(ISNULL(A.MTRL+A.MET+A.BODY,0)) as Others, " & _
                  "Sum(Case When RIGHT(A.NOINV,2) <> '-B' AND A.PDS = 1 AND A.FN_RETURN = 0 AND A.TPSVC <> 'W' AND A.FS_RC = '' then 1 else 0 End) as PDS, " & _
                  "Sum(Case When RIGHT(A.NOINV,2) <> '-B' AND A.FS_RC = 'R001' AND A.PDS = 0 AND A.FN_RETURN = 0 AND A.TPSVC <> 'W' Then 1 Else 0 End) as SBI, " & _
                  "sum(case When RIGHT(A.NOINV,2) <> '-B' AND A.TPSVC = 'W' AND A.PDS = 0 AND A.FN_RETURN = 0 AND A.FS_RC = '' then 1 else 0 end) as TWC, " & _
                  "Sum(case When RIGHT(A.NOINV,2) <> '-B' AND A.FN_RETURN = 1 then 1 else 0 end) as RTJ, " & _
                  "Sum(case When RIGHT(A.NOINV,2) <> '-B' AND A.FS_RC <> 'R001' AND A.FS_RC <> '' AND A.PDS = 0 AND A.FN_RETURN = 0 AND A.TPSVC <> 'W' Then 1 Else 0 End) as SBE, " & _
                  "Sum(case When RIGHT(A.NOINV,2) <> '-B' AND A.FS_RC = '' AND A.PDS = 0 AND A.TPSVC <> 'W' AND A.FN_RETURN = 0 AND ((A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7) not like '%SBE%' and (A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7) not like '%SBI%') AND (LEN(REPLACE(A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7,'LUB','')) > 0 or A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7 = '') then 1 else 0 End) as GRP, " & _
                  "Sum(case when RIGHT(A.NOINV,2) <> '-B' AND A.FS_RC = '' AND LEN(REPLACE(A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7,'LUB','')) = 0 AND A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7 <> '' AND A.PDS = 0 AND A.TPSVC <> 'W' AND A.FN_RETURN = 0 Then 1 else 0 End) as LUB " & _
                  " from " & _
                  "T_invoice A FULL OUTER Join " & _
                  "(Select [date] as TglWO, FRONT, Count(DISTINCT No_WO) as WO from Jobdta Where FS_LOG_DEL = '' and (date between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') group by FRONT, Date) B " & _
                  "On A.InvDate = B.TglWO and A.FRONT = B.FRONT " & _
                  "Where (A.InvDATe between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') group by A.FRONT, A.Invdate, B.WO " & _
                  "Union "
                  ''strSql = strSql & "Select FRONT, Convert(varchar(10),FD_Retur,101) as Tanggal, isnull(count(distinct No_WO),0)*-1 as SO, "
5                 strSql = strSql & "Select FRONT, Convert(varchar(10),FD_Retur,101) as Tanggal, isnull(count(distinct No_WO),0)* 0 as SO, " & _
                  "Sum(Labor)*-1 as Labor , cOUNT(DISTINCT(rEPLACE(REpLACE(nOiNV,'-A',''),'-B','')))*-1 as JmlInv, " & _
                  "SUM(ISNULL(PART_G+OIL_G,0))*-1  as Genuine, SUM(ISNULL(PART_NG+OIL_NG,0))*-1 as NonGenuine, SUM(ISNULL(SBLT,0))*-1 as Sublet, " & _
                  "SUM(ISNULL(MTRL+MET+BODY,0))*-1 as Others, " & _
                  "Sum(Case When RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' then -1 else 0 End) as PDS, " & _
                  "Sum(Case When RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' Then -1 Else 0 End) as SBI, " & _
                  "sum(case When RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' then -1 else 0 end) as TWC, " & _
                  "Sum(case When RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 then -1 else 0 end) as RTJ, " & _
                  "Sum(case When RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' Then -1 Else 0 End) as SBE, " & _
                  "Sum(case When RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND ((KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7) not like '%SBE%' and (KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7) not like '%SBI%') AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 or KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') then -1 else 0 End) as GRP, " & _
                  "Sum(case when RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 Then -1 else 0 End) as LUB " & _
                  " from " & _
                  "T_invoice Where (FD_Retur between '" & TglNul2(Tgl1.value) & " 00:00:00' and '" & TglNul2(Tgl2.value) & " 23:59:59') group by FRONT, Convert(varchar(10),FD_Retur,101)) X "
                  
6                 If Text(11) <> "" Then
7                     strSql = strSql & " Where X.FRONT  = '" & Text(11) & "' Group by X.FRONT, X.Tanggal "
8                 Else
9                     strSql = strSql & " Group by X.FRONT, X.Tanggal "
10                End If
11                strSql = strSql & "Union " & _
                  "select Q.FRONT, Q.DATE,Q.JumlahWO, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 FROM " & _
                  "(select A.* FROm " & _
                  "(Select X.XX , Sum(X.DD) as Jumlah, X.FRONT FROM " & _
                  "(select date as XX ,count(distinct No_WO) as DD, FRONT from Jobdta Where (date between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " And INVDATE <> DATE  and DELIv <> 'BTL' group by date, FRONT " & _
                  "UNION Select date as XX, -1*(count(distinct No_WO)) as DD, FRONT FROM JOBDTA Where (date  between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " and DELIv <> 'BTL'  group by date, FRONT " & _
                  "UNION Select INVDATE as XX, -1*(count(distinct No_WO)) as DD, FRONT FROM JOBDTA Where (INVDATE  between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " and DELIv <> 'BTL'  group by INVDATE, FRONT ) x " & _
                  "Group by x.XX, X.FRONT) A Where A.Jumlah = 0) W INNER JOIN " & _
                  "(Select DATE, count(distinct No_WO) as JumlahWO, FRONT FROM Jobdta  Where (date  between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " And DELIv <> 'BTL' group by date, FRONT) Q " & _
                  "On W.XX = Q.DATE and W.FRONT = Q.FRONT) G " & _
                  "Group by G.FRONT, G.Tanggal"
12                Set Report = New rptRepoHarianSA
          
          Dim crxSubReportHistory As CRAXDRT.Report
              
13                Set oRs = objO.FetchData(strSql)
                 
14        Set crxSubReportHistory = Report.Subreport1.OpenSubreport
15        crxSubReportHistory.database.SetDataSource oRs
          
16                Report.database.SetDataSource oRs
17                Report.txtPerusahaan.SetText gPerush
18                Report.txtAlamat.SetText gAlamat
19                Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")

20                With CRViewer1
21                    .ReportSource = Report
22                    .ViewReport
23                    .Zoom 100
24                    .Visible = True
25                End With
26                cmdTutup.Visible = True
27                Screen.MousePointer = vbDefault
28            Call DisableExport
29            Call setLokasiTombol

30        Else
                  

31        strSql = "Select G.Tanggal, sum(G.SO)AS SO " & _
                  ", sum(G.INV) as INV, sum(G.JASA) as Jasa, Sum(G.Genuine) as Genuine, Sum(G.NonGenuine) as NonGenuine, " & _
                  "Sum(G.Sublet) as Sublet,Sum(G.Others) as Others, sum(G.PDS) as PDS, sum(G.SBI) as SBI, sum(G.TWC) as TWC, sum(G.RTJ) as RTJ,  sum(G.SBE) as SBE, " & _
                  "sum(G.GRP) as GRP, sum(G.LUB) as LUB FROM ( "
32        strSql = strSql & "Select X.Tanggal , sum(SO)AS SO " & _
                  ", sum(jmlInv) as INV, sum(X.Labor) as Jasa, Sum(X.Genuine) as Genuine, Sum(X.NonGenuine) as NonGenuine, " & _
                  "Sum(X.Sublet) as Sublet,Sum(X.Others) as Others, sum(PDS) as PDS, sum(SBI) as SBI, sum(TWC) as TWC, sum(RTJ) as RTJ,  sum(SBE) as SBE, " & _
                  "sum(GRP) as GRP, sum(LUB) as LUB from " & _
                  "( " & _
                  "Select InvDate as Tanggal , isnull(B.WO,0) as SO, " & _
                  " sUM(A.Labor) as Labor, cOUNT(DISTINCT(rEPLACE(REpLACE(A.nOiNV,'-A',''),'-B',''))) as JmlInv, " & _
                  "SUM(ISNULL(A.PART_G+A.OIL_G,0))  as Genuine, SUM(ISNULL(A.PART_NG+A.OIL_NG,0)) as NonGenuine, SUM(ISNULL(A.SBLT,0)) as Sublet, " & _
                  "SUM(ISNULL(A.MTRL+A.MET+A.BODY,0)) as Others, " & _
                  "Sum(Case When RIGHT(A.NOINV,2) <> '-B' AND A.PDS = 1 AND A.FN_RETURN = 0 AND A.TPSVC <> 'W' AND A.FS_RC = '' then 1 else 0 End) as PDS, " & _
                  "Sum(Case When RIGHT(A.NOINV,2) <> '-B' AND A.FS_RC = 'R001' AND A.PDS = 0 AND A.FN_RETURN = 0 AND A.TPSVC <> 'W' Then 1 Else 0 End) as SBI, " & _
                  "sum(case When RIGHT(A.NOINV,2) <> '-B' AND A.TPSVC = 'W' AND A.PDS = 0 AND A.FN_RETURN = 0 AND A.FS_RC = '' then 1 else 0 end) as TWC, " & _
                  "Sum(case When RIGHT(A.NOINV,2) <> '-B' AND A.FN_RETURN = 1 then 1 else 0 end) as RTJ, " & _
                  "Sum(case When RIGHT(A.NOINV,2) <> '-B' AND A.FS_RC <> 'R001' AND A.FS_RC <> '' AND A.PDS = 0 AND A.FN_RETURN = 0 AND A.TPSVC <> 'W' Then 1 Else 0 End) as SBE, " & _
                  "Sum(case When RIGHT(A.NOINV,2) <> '-B' AND A.FS_RC = '' AND A.PDS = 0 AND A.TPSVC <> 'W' AND A.FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7,'LUB','')) > 0 or A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7 = '') then 1 else 0 End) as GRP, " & _
                  "Sum(case when RIGHT(A.NOINV,2) <> '-B' AND A.FS_RC = '' AND LEN(REPLACE(A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7,'LUB','')) = 0 AND A.KLPKERJA1+A.KLPKERJA2+A.KLPKERJA3+A.KLPKERJA4+A.KLPKERJA5+A.KLPKERJA6+A.KLPKERJA7 <> '' AND A.PDS = 0 AND A.TPSVC <> 'W' AND A.FN_RETURN = 0 Then 1 else 0 End) as LUB " & _
                  " from " & _
                  "T_invoice A FULL OUTER Join " & _
                  "(Select [date] as TglWO, Count(DISTINCT No_WO) as WO from Jobdta Where FS_LOG_DEL = '' and (date between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " group by  Date) B " & _
                  "On A.InvDate = B.TglWO " & _
                  "Where (A.InvDATe between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND A.FRONT = '" & Text(11) & "' ", "") & " group by A.Invdate, B.WO " & _
                  "Union "
                  ''strSql = strSql & "Select Convert(varchar(10),FD_Retur,101) as Tanggal, isnull(count(distinct No_WO),0)*-1 as SO, "
33                strSql = strSql & "Select Convert(varchar(10),FD_Retur,101) as Tanggal, isnull(count(distinct No_WO),0) * 0 as SO, " & _
                  "Sum(Labor)*-1 as Labor , cOUNT(DISTINCT(rEPLACE(REpLACE(nOiNV,'-A',''),'-B','')))*-1 as JmlInv, " & _
                  "SUM(ISNULL(PART_G+OIL_G,0))*-1  as Genuine, SUM(ISNULL(PART_NG+OIL_NG,0))*-1 as NonGenuine, SUM(ISNULL(SBLT,0))*-1 as Sublet, " & _
                  "SUM(ISNULL(MTRL+MET+BODY,0))*-1 as Others, " & _
                  "Sum(Case When RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' then -1 else 0 End) as PDS, " & _
                  "Sum(Case When RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' Then -1 Else 0 End) as SBI, " & _
                  "sum(case When RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' then -1 else 0 end) as TWC, " & _
                  "Sum(case When RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 then -1 else 0 end) as RTJ, " & _
                  "Sum(case When RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' Then -1 Else 0 End) as SBE, " & _
                  "Sum(case When RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 or KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') then -1 else 0 End) as GRP, " & _
                  "Sum(case when RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 Then -1 else 0 End) as LUB " & _
                  " from " & _
                  "T_invoice Where (FD_Retur between '" & TglNul2(Tgl1.value) & " 00:00:00' and '" & TglNul2(Tgl2.value) & " 23:59:59') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " group by Convert(varchar(10),FD_Retur,101)) X "
34                    strSql = strSql & " Group by X.Tanggal "
35                strSql = strSql & "Union " & _
                  "select Q.DATE, " & IIf(Text(11) = "", "0", "Q.JumlahWO") & " , 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 FROM " & _
                  "(select A.* FROm " & _
                  "(Select X.XX , Sum(X.DD) as Jumlah, X.FRONT FROM " & _
                  "(select date as XX ,count(distinct No_WO) as DD, FRONT from Jobdta Where (date between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " And INVDATE <> DATE  and DELIv <> 'BTL' group by date, FRONT " & _
                  "UNION Select date as XX, -1*(count(distinct No_WO)) as DD, FRONT FROM JOBDTA Where (date  between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " and DELIv <> 'BTL'  group by date, FRONT ) x " & _
                  "Group by x.XX, X.FRONT) A Where A.Jumlah = 0) W INNER JOIN " & _
                  "(Select DATE, count(distinct No_WO) as JumlahWO, FRONT FROM Jobdta  Where (date  between '" & TglNul2(Tgl1.value) & "' and '" & TglNul2(Tgl2.value) & "') " & IIf(Text(11) <> "", " AND FRONT = '" & Text(11) & "' ", "") & " And DELIv <> 'BTL' group by date, FRONT) Q " & _
                  "On W.XX = Q.DATE and W.FRONT = Q.FRONT) G " & _
                  "Group by G.Tanggal"
                      
                      
36                Set Report = New rptRepoHarian
37                Set oRs = objO.FetchData(strSql)
38                Report.database.SetDataSource oRs
39                Report.txtPerusahaan.SetText gPerush
40                Report.txtAlamat.SetText gAlamat
41                Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
                  'If opLap(0).value Then
                  '    If Trim(Text(12)) <> "" Then
                  '        Report.txtSa.SetText "oleh SA : " & Trim(Text(12)) & " [" & Trim(Text(11)) & "] "
                  '    End If
                  'End If
                  'oLoad.Visible = False
                  'DoEvents
42                With CRViewer1
43                    .ReportSource = Report
44                    .ViewReport
45                    .Zoom 100
46                    .Visible = True
47                End With
48                cmdTutup.Visible = True
49                Screen.MousePointer = vbDefault
50                Call DisableExport
51                Call setLokasiTombol
52        End If
End Sub

Private Sub LoadDataLaporan()
          Dim nDate1 As Date
          Dim nDate2 As Date
          Dim iSA As Integer
1         nDate1 = Format(Tgl1, "DD/MM/YYYY")
2         nDate2 = Format(Tgl2, "DD/MM/YYYY")
3         oLoad.Visible = True
4         DoEvents
5         oLoad.Min = 0
6         If nDate1 = nDate2 Then
7             oLoad.Max = 1
8         Else
9             oLoad.Max = DateDiff("D", nDate1, nDate2)
10        End If
          Dim rsLast As ADODB.Recordset
11        Set rsLast = objO.FetchData("Select top 1 tanggal from RepoHSA Where tanggal >= '" & Format(DateAdd("d", -6, isnow), "MM/DD/YYYY") & "' order By tanggal desc")
12        If rsLast.RecordCount > 0 Then Call objO.DeleteData("Delete FROM RepoHSA Where Tanggal = '" & Format(rsLast.Fields(0).value, "MM/DD/YYYY") & "'")
          Dim rsSA As ADODB.Recordset
13        Set rsSA = objO.FetchData("Select FS_KD_Petugas from T_Petugas Where Tipe in ('SA','SA_X')")
14        oLoad.Max = (oLoad.Max * rsSA.RecordCount) + 12
15        oLoad.value = 0
16        DoEvents
          Dim iEx As Boolean
          Dim ivEx As Boolean
17        Do Until nDate1 > nDate2
18            rsSA.MoveFirst
19            iSA = 1
20            Do Until rsSA.EOF
21            If oLoad.value < oLoad.Max Then oLoad.value = oLoad.value + 1
22            nLabel.Caption = "Proses ... " & Round((oLoad.value * 100) / oLoad.Max, 0) & " % - Kalkulasi Tanggal : " & nDate1 & ", SA : " & rsSA.Fields(0).value
23            DoEvents
24                iEx = objO.FetchData_Boolean("Select SA FROM RepoHSA Where SA = '" & rsSA.Fields(0).value & "' And Tanggal = '" & Format(nDate1, "MM/DD/YYYY") & "'")
25                If iEx = False Then
26                    ivEx = objO.FetchData_Boolean("Select top 1 front FROM T_Invoice Where front = '" & rsSA.Fields(0).value & "' And invDate = '" & Format(nDate1, "MM/DD/YYYY") & "'")
27                    If ivEx = True Then
                      
28                        sSQL = "SELECT front as SA, DATE, (SELECT COUNT(DISTINCT NO_WO) FROM JOBDTA x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE x.FS_LOG_DEL = '' AND DATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & ") AS SO, " & _
                                 "(SELECT COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") + (SELECT (-1) * COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") AS INV, " & _
                                 "ISNULL((SELECT SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) AS JASA, " & _
                                 "ISNULL((SELECT SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) AS PART, " & _
                                 "ISNULL((SELECT SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) AS OLI, " & _
                                 "ISNULL((SELECT SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) AS SUBLET, " & _
                                 "ISNULL((SELECT SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' " & Trim(txt) & "),0) AS OTHER, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") AS PDS, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBI, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") AS TWC, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") AS RTJ, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBE, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 or KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' " & _
                                 "AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 OR KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") AS GRP, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = '" & Format(nDate1, "MM/DD/YYYY") & "'  AND FRONT = '" & rsSA.Fields(0).value & "' AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' " & _
                                 "AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") AS LUB " & _
                                 "FROM (SELECT FRONT, DATE From JOBDTA Union SELECT FRONT, INVDATE From T_INVOICE Union SELECT FRONT, CONVERT(varchar(15),FD_RETUR,10) From T_INVOICE) AS D " & _
                                 "WHERE D.DATE BETWEEN '" & Format(nDate1, "MM/DD/YYYY") & "' AND '" & Format(nDate1, "MM/DD/YYYY") & "' AND D.Front = '" & Trim(rsSA.Fields(0).value) & "'"
                      
                      
                          'sSQL = "SELECT front as SA, DATE, (SELECT COUNT(DISTINCT NO_WO) FROM JOBDTA x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE x.FS_LOG_DEL = '' AND DATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & ") AS SO, " & _
                                 "(SELECT COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") + (SELECT (-1) * COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") AS INV, " & _
                                 "ISNULL((SELECT SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS JASA, " & _
                                 "ISNULL((SELECT SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS PART, " & _
                                 "ISNULL((SELECT SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS OLI, " & _
                                 "ISNULL((SELECT SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS SUBLET, " & _
                                 "ISNULL((SELECT SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS OTHER, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") AS PDS, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBI, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") AS TWC, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") AS RTJ, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBE, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 or KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 OR KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") AS GRP, " & _
                                 "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") AS LUB " & _
                                 "FROM (SELECT FRONT, DATE From JOBDTA Union SELECT FRONT, INVDATE From T_INVOICE Union SELECT FRONT, CONVERT(varchar(15),FD_RETUR,10) From T_INVOICE) AS D " & _
                                 "WHERE D.DATE BETWEEN '" & Format(nDate1, "MM/DD/YYYY") & "' AND '" & Format(nDate1, "MM/DD/YYYY") & "' AND D.Front = '" & Trim(rsSA.Fields(0).value) & "'"
                          
29                        Call objO.InsertData("Insert Into RepoHSA " & sSQL & "")
30                        If nStop = True Then
31                            Call objO.DeleteData("Delete FROM RepoHSA Where Tanggal = '" & Format(nDate1, "MM/DD/YYYY") & "'")
32                            nLabel.Caption = ""
33                            oLoad.Visible = False
34                            Command1.Visible = False
35                            Exit Sub
36                        End If
37                    End If
38                Else
                      'rsSA.MoveNext
39                    On Error GoTo errVal
                      'oLoad.value = (oLoad.value + rsSA.RecordCount) - iSA
errVal:
40                End If
41                rsSA.MoveNext
42                iSA = iSA + 1
43                If nStop = True Then
44                    nLabel.Caption = ""
45                    oLoad.Visible = False
46                    Command1.Visible = False
47                    Exit Sub
48                End If
49            Loop
50            nDate1 = DateAdd("D", 1, nDate1)
51        Loop
52        If oLoad.value < oLoad.Max Then oLoad.value = oLoad.value + 1
53        DoEvents
54        Command1.Visible = False
55        Call CetakTotal
56        nLabel.Caption = ""
End Sub


Private Sub Cetak(vNamaTabel As String, txt As String, txtx As String)

    Dim nDate1 As Date
    Dim nDate2 As Date
    Dim iSA As Integer
    nDate1 = Format(Tgl1, "DD/MM/YYYY")
    nDate2 = Format(Tgl2, "DD/MM/YYYY")
    oLoad.Visible = True
    DoEvents
    oLoad.Min = 0
    If nDate1 = nDate2 Then
        oLoad.Max = 1
    Else
        oLoad.Max = DateDiff("D", nDate1, nDate2)
    End If
    oLoad.value = 0
    DoEvents

If opLap(0).value Then
    Call objO.DeleteData("Delete from x10DayGlobal")
Else
    Call objO.DeleteData("Delete from x10DayRinci")
End If

Dim oRs As Object
Screen.MousePointer = vbHourglass
Dim Report As Object
    

Dim nSO As Double

    Do Until nDate1 > nDate2

        If opLap(0).value Then
            sSQL = "INSERT INTO x10DayGlobal SELECT distinct DATE, (SELECT COUNT(DISTINCT NO_WO) FROM JOBDTA x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE x.FS_LOG_DEL = '' AND DATE = D.DATE " & Trim(txt) & ") AS SO, " & _
                   "(SELECT COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") + (SELECT (-1) * COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") AS INV, " & _
                   "ISNULL((SELECT SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE " & Trim(txt) & "),0) AS JASA, " & _
                   "ISNULL((SELECT SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE " & Trim(txt) & "),0) AS PART, " & _
                   "ISNULL((SELECT SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE " & Trim(txt) & "),0) AS OLI, " & _
                   "ISNULL((SELECT SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE " & Trim(txt) & "),0) AS SUBLET, " & _
                   "ISNULL((SELECT SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE " & Trim(txt) & "),0) AS OTHER, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") AS PDS, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBI, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") AS TWC, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") AS RTJ, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBE, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 or KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 OR KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") AS GRP, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") AS LUB " & _
                   "FROM (SELECT FRONT, DATE From JOBDTA Union SELECT FRONT, INVDATE From T_INVOICE Union SELECT FRONT, CONVERT(varchar(15),FD_RETUR,10) From T_INVOICE) AS D " & _
                   "WHERE D.DATE BETWEEN '" & Format(nDate1, "MM/DD/YYYY") & "' AND '" & Format(nDate1, "MM/DD/YYYY") & "' ORDER BY D.DATE"
                   
                   Call objO.InsertData(sSQL)
        Else
            
'            nSO = objO.FetchData_Double("SELECT COUNT(DISTINCT NO_WO) FROM JOBDTA x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE x.FS_LOG_DEL = '' AND DATE = '" & Format(nDate1, "MM/DD/YYYY") & "' AND FRONT = D.FRONT " & Trim(txt))
'
'            sSQL = "INSERT INTO x10DayRinci SELECT front as SA, DATE, " & nSO & " AS SO, " & _
'                   "(SELECT COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") + (SELECT (-1) * COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") AS INV, " & _
'                   "ISNULL((SELECT SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS JASA, " & _
'                   "ISNULL((SELECT SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS PART, " & _
'                   "ISNULL((SELECT SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS OLI, " & _
'                   "ISNULL((SELECT SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS SUBLET, " & _
'                   "ISNULL((SELECT SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS OTHER, " & _
'                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") AS PDS, " & _
'                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBI, " & _
'                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") AS TWC, " & _
'                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") AS RTJ, " & _
'                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBE, " & _
'                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 or KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 OR KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") AS GRP, " & _
'                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") AS LUB " & _
'                   "FROM (SELECT FRONT, DATE From JOBDTA Union SELECT FRONT, INVDATE From T_INVOICE Union SELECT FRONT, CONVERT(varchar(15),FD_RETUR,10) From T_INVOICE) AS D " & _
'                   "WHERE D.DATE BETWEEN '" & Format(nDate1, "MM/DD/YYYY") & "' AND '" & Format(nDate1, "MM/DD/YYYY") & "' ORDER BY d.FRONT, D.DATE"

            
            sSQL = "INSERT INTO x10DayRinci SELECT front as SA, DATE, (SELECT COUNT(DISTINCT NO_WO) FROM JOBDTA x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE x.FS_LOG_DEL = '' AND DATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & ") AS SO, " & _
                   "(SELECT COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") + (SELECT (-1) * COUNT(DISTINCT NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' " & Trim(txt) & ") AS INV, " & _
                   "ISNULL((SELECT SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(LABOR,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS JASA, " & _
                   "ISNULL((SELECT SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_G+OIL_G,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS PART, " & _
                   "ISNULL((SELECT SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(PART_NG+OIL_NG,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS OLI, " & _
                   "ISNULL((SELECT SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(SBLT,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS SUBLET, " & _
                   "ISNULL((SELECT SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) + ISNULL((SELECT (-1)*SUM(ISNULL(MTRL+MET+BODY,0)) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT " & Trim(txt) & "),0) AS OTHER, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND PDS = 1 AND FN_RETURN = 0 AND TPSVC <> 'W' AND FS_RC = '' " & Trim(txt) & ") AS PDS, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1) * COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = 'R001' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBI, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND TPSVC = 'W' AND PDS = 0 AND FN_RETURN = 0 AND FS_RC = '' " & Trim(txt) & ") AS TWC, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FN_RETURN = 1 " & Trim(txt) & ") AS RTJ, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC <> 'R001' AND FS_RC <> '' AND PDS = 0 AND FN_RETURN = 0 AND TPSVC <> 'W' " & Trim(txt) & ") AS SBE, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 or KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 AND (LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) > 0 OR KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 = '') " & Trim(txt) & ") AS GRP, " & _
                   "(SELECT COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE INVDATE = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") + (SELECT (-1)*COUNT(NOINV) FROM T_INVOICE x INNER JOIN CUSTDAT y ON x.NO_CHASSIS = y.NO_CHASSIS WHERE CONVERT(varchar(15),FD_RETUR,10) = D.DATE AND FRONT = D.FRONT AND RIGHT(NOINV,2) <> '-B' AND FS_RC = '' AND LEN(REPLACE(KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7,'LUB','')) = 0 AND KLPKERJA1+KLPKERJA2+KLPKERJA3+KLPKERJA4+KLPKERJA5+KLPKERJA6+KLPKERJA7 <> '' AND PDS = 0 AND TPSVC <> 'W' AND FN_RETURN = 0 " & Trim(txt) & ") AS LUB " & _
                   "FROM (SELECT FRONT, DATE From JOBDTA Union SELECT FRONT, INVDATE From T_INVOICE Union SELECT FRONT, CONVERT(varchar(15),FD_RETUR,10) From T_INVOICE) AS D " & _
                   "WHERE D.DATE BETWEEN '" & Format(nDate1, "MM/DD/YYYY") & "' AND '" & Format(nDate1, "MM/DD/YYYY") & "' ORDER BY d.FRONT, D.DATE"
                   
                   Call objO.InsertData(sSQL)
                   
        End If
        
        nDate1 = DateAdd("D", 1, nDate1)
        If oLoad.value < oLoad.Max Then oLoad.value = oLoad.value + 1
        DoEvents
    Loop
    
  
    If opLap(0).value = True Then
        Set Report = New rptRepoHarian
        Set oRs = objO.FetchData("Select * from x10DayGlobal")
    Else
        Set Report = New rptRepoHarianSA
        Set oRs = objO.FetchData("Select * from x10DayRinci")
    End If
    
        
        If opLap(0).value = False Then

            Dim sHlp As Object
            Dim crxSubReportHistory As CRAXDRT.Report
            Set sHlp = objO.FetchData("Select * from x10DayRinci")
            Set crxSubReportHistory = Report.Subreport1.OpenSubreport
            crxSubReportHistory.database.SetDataSource sHlp
        End If
        
    Report.database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.txtTanggal.SetText "TANGGAL : " & Format(Tgl1, "DD-MM-YYYY") & " s/d " & Format(Tgl2, "DD-MM-YYYY")
    If opLap(0).value Then
        If Trim(Text(12)) <> "" Then
            Report.txtSA.SetText "oleh SA : " & Trim(Text(12)) & " [" & Trim(Text(11)) & "] "
        End If
    End If
    oLoad.Visible = False
    DoEvents
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
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
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCari.show (vbModal)
            Case 2
                If Text(0) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = Text(0)
                    frmCari.show (vbModal)
                Else
                    MsgBox "Group Job Order Harus diisi"
                    Text(0).SetFocus
                    Exit Sub
                End If
            Case 5
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
            Case 3
                gMainKeyCari = Text(Index)
                gBantuan = "ALL"
                gSearchOption = "DEALER"
                frmCari.show (vbModal)
            Case 11
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT DISTINCT JOB, REMARK FROM HO_FLATE WHERE JOB = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = oRs("REMARK")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 2
            If Text(0) <> "" Then
                If Text(Index) <> "" Then
                    Set oRs = objO.FetchData("SELECT DISTINCT FS_ITEM, KETERANGAN FROM HO_FLATE WHERE JOB = '" & Text(0) & "' AND FS_ITEM = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                    If oRs.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "JOB_ORDER"
                        gBantuan = Text(0)
                        frmCari.show (vbModal)
                    Else
                        Text(Index + 1) = oRs("KETERANGAN")
                    End If
                    oRs.Close
                End If
            Else
                MsgBox "Group Job Order Harus diisi"
                Text(0).SetFocus
                Exit Sub
            End If
        Case 5
            If Text(Index) <> "" Then
                If Log.RPLHistori = 0 Then
                    Set oRs = objO.FetchData("SELECT DISTINCT MDL, [DESC] FROM MODEL WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                Else
                    Set oRs = objO.FetchData("SELECT DISTINCT MDL, [DESC] FROM [" & dbReplikasiHO & "]..MODEL WHERE MDL = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                End If
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = oRs("DESC")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 3
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT * FROM T_DEALER WHERE FS_KD_DEALER = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gBantuan = "ALL"
                    gSearchOption = "DEALER"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = oRs("FS_NM_DEALER")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
        Case 7
            If Not CekTgl(Text(Index)) Then
                If Text(Index) <> ">>->>->>>>" Then
                    MsgBox "Format tanggal salah"
                End If
                Text(Index) = ">>->>->>>>"
            End If
        Case 8
            If Not CekTgl(Text(Index)) Then
                If Text(Index) <> ">>->>->>>>" Then
                    MsgBox "Format tanggal salah"
                End If
                Text(Index) = ">>->>->>>>"
            End If
        Case 9
            If Text(Index) = "" Then
                Text(Index) = 0
            ElseIf CDbl(Text(Index)) <> 0 Then
                If CDbl(Text(Index)) > CDbl(year(isnow)) Or CDbl(Text(Index)) < CDbl(year(isnow) - 30) Then
                    MsgBox "Tahun salah !!"
                    Text(Index) = 0
                End If
            End If
        Case 10
            If Text(Index) = "" Then
                Text(Index) = 0
            ElseIf CDbl(Text(Index)) > CDbl(year(isnow)) Or CDbl(Text(Index)) < CDbl(year(isnow) - 30) Then
                MsgBox "Tahun salah !!"
                Text(Index) = 0
            End If
        Case 11
            If Text(Index) <> "" Then
                Set oRs = objO.FetchData("SELECT * FROM T_PETUGAS WHERE TIPE in ('SA','SA_X') AND FS_KD_PETUGAS = '" & Text(Index) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SA"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = oRs("FS_NM_PETUGAS")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
    End Select
End Sub



Private Sub createTbl_ServRate(Optional Edit As Boolean = False)
1     On Error GoTo errIsExist
2         If rsSR.State Then rsSR.Close
3         Set rsSR = Nothing
4         With rsSR
5             .Fields.Append "Tgl", adDate
6             .Fields.Append "NoWO", adVarChar, 15
7             .Fields.Append "NoPol", adVarChar, 15
8             .Fields.Append "Job", adVarChar, 10
9             .Fields.Append "NoRangka", adVarChar, 30
10            .Fields.Append "Model", adVarChar, 100
11            .Fields.Append "Kode", adVarChar, 30
12            .Fields.Append "Nama", adVarChar, 100
13            .Fields.Append "QtyOrderSA", adDouble
14            .Fields.Append "QtySupply", adDouble
15            .Fields.Append "FR_Li", adDouble
16            .Fields.Append "FR", adVarChar, 10
17            .Fields.Append "li_Supply", adDouble
18            .Fields.Append "Li_BO", adDouble
19            .Fields.Append "Li_Cancel", adDouble
20            .Fields.Append "SR_Li", adDouble
21            .Fields.Append "SR", adVarChar, 10
22            .Fields.Append "FullSup", adDouble
23            .Fields.Append "JobFilter", adVarChar, 10
24            .Fields.Append "NoWOFilter", adVarChar, 15
25            .Open
26        End With
27        Set GridFR.DataSource = rsSR
errIsExist:
End Sub

Private Sub cmdPrintSR_Click()
    dtMulai = DateSerial(cboThnSR, cboBlnSR, 1)
    dtSelesai = DateSerial(cboThnSR, cboBlnSR + 1, 1 - 1)
   
    If Log.isFill = 0 Then Exit Sub
    GridFR.Visible = True
    GridFR1.Visible = False
    GridFR2.Visible = False
    Opt(0).value = True
    Opt(1).value = False
    Opt(2).value = False
    If cboJob.Text = "" Then cboJob.Text = "ALL"
    
    FrRate.Visible = True
    FrRate.Height = 9200 '8415
    FrRate.Width = 14535
    FrRate.Left = 480
    FrRate.Top = 0
    DoEvents
    createTbl_ServRate True
    Dim rsR As ADODB.Recordset
    
    Dim rsGetAppRate As New ADODB.Recordset
    ''2015-10-06 || ditambah menit
    Set rsGetAppRate = objO.FetchData("SELECT CASE LEFT(apprate,1) WHEN 'N' THEN 'MINUTE' WHEN 'H' THEN 'HOUR' WHEN 'D' THEN 'DAY' END as inter, SUBSTRING(apprate,2,len(apprate)) as jangka from profilbengkel..profil")
    
    Dim interval, jangka As String
    interval = rsGetAppRate.Fields(0).value
    jangka = rsGetAppRate.Fields(1).value
    
    Set rsR = objO.FetchData("select A.[DAte] as Tanggal , A.No_WO as NOWO, A.No_polisi, A.No_Chassis, A.MDL + ' - ' + A.FMDL as Model , " & _
                             "B.KodeItem, B.Keterangan as NamaItem, B.Qty as Qty_OrderSA, (IsNull(B.QtyKirimBI,0) + IsNull(B.QtyKirimBL,0) + IsNull(B.QtyKeluar,0)) as Qty_Supply, " & _
                             "case B.Qty - (IsNull(B.QtyKirimBI,0) + IsNull(B.QtyKirimBL,0) + IsNull(B.QtyKeluar,0)) When 0 Then 100 Else Round((((IsNull(B.QtyKirimBI,0) + IsNull(B.QtyKirimBL,0) + IsNull(B.QtyKeluar,0))/B.Qty) * 100),2) END as FillRate_LineItem, " & _
                             "case (IsNull(B.QtyKirimBI,0) + IsNull(B.QtyKirimBL,0) + IsNull(B.QtyKeluar,0)) When 0 Then 0 Else 1 End as Li_supply, 0 as Li_BO, " & _
                             "case B.BO When 'CANCEL' Then 1 else 0 end as Li_Cancel,0,0, B.Nomor,A.FS_RC,A.PDS,A.TPSVC,datediff(" & interval & ",TglPerlu,TglSupply)  " & _
                             "from T_invoice A Inner Join [ACC" & Trim(Right(gDb, 2)) & "-" & cboThnSR.Text & IIf(CInt(cboBlnSR.Text) < 10, "0" & cboBlnSR.Text, cboBlnSR.Text) & "]..SOORD B On A.No_WO = B.NoSO " & _
                             "Where Right(A.noInv,1) <> 'B' and A.Fd_retur = '01/01/1900' " & _
                             "and Month(A.[Date]) = " & CInt(cboBlnSR.Text) & " and Year(A.[Date]) = " & CInt(cboThnSR.Text) & " Order By A.[date] asc")
    
    Dim nWO1 As String
    Dim xWO As Double
    Dim nTgl1 As Date
    Dim nFrLi As Integer
    Dim isNewWO1 As Boolean
    Dim isNewTgl1 As Boolean
    nWO1 = ""
    xWO = 0
    nTgl1 = "01/01/1900"
    
    Dim sumOrdSA As Double
    Dim sumSuply As Double
    Dim sumLiSup As Double
    Dim sumLiBO As Double
    Dim sumLiCan As Double
    Dim SumFullSup As Double
    sumOrdSA = 0
    sumSuply = 0
    sumLiSup = 0
    sumLiBO = 0
    sumLiCan = 0
    SumFullSup = 0
    
    Dim nSupply As Double
    Dim nBO As Double
    Dim nCancel As Double
    Dim nFR1 As Double
    nSupply = 0
    nBO = 0
    nCancel = 0
    nFR1 = 100
    
    Dim fR100 As Double
    Dim fRALL As Double
    fR100 = 0
    fRALL = 0
    
    Dim sRSP As Double
    Dim sRBC As Double
    sRSP = 0
    sRBC = 0

    Dim fR100i As Double
    Dim fRALLi As Double
    fR100i = 0
    fRALLi = 0
    
    Dim sRSPi As Double
    Dim sRBCi As Double
    sRSPi = 0
    sRBCi = 0
    Dim nJob As String
    nJob = ""
    
    Dim aSumOrdSA As Double
    Dim asumSuply As Double
    Dim asumLiSup As Double
    Dim asumLiBO As Double
    Dim asumLiCan As Double
    aSumOrdSA = 0
    asumSuply = 0
    asumLiSup = 0
    asumLiBO = 0
    asumLiCan = 0
    
    Dim isFRi As String
    Dim isSRi As String
    isFRi = ""
    isSRi = ""
    
    isNewWO1 = False
    isNewTgl1 = False
    
    Dim selisih As Double
    
    If rsR.RecordCount > 0 Then
        rsR.MoveFirst
        iKodeItem = 0
        
        Do Until rsR.EOF
        selisih = 0
        selisih = IIf(IsNull(rsR.Fields(19).value), jangka + 1, rsR.Fields(19).value)
        
            With rsSR
                .AddNew
                rsSR!Tgl = rsR.Fields(0).value
                If rsSR!Tgl <> nTgl1 Then
                    isNewTgl1 = True
                    If nTgl1 <> "01/01/1900" Then

                        'If rsSR!Tgl = nTgl1 Then
                            rsSR!Tgl = nTgl1
                            rsSR!noWO = nWO1
                            rsSR!NoWoFilter = nWO1
                            rsSR!nopol = ""
                            rsSR!job = ""
                            rsSR!JobFilter = nJob
                            rsSR!noRangka = ""
                            rsSR!model = ""
                            rsSR!kode = ""
                            rsSR!Nama = ""
                            rsSR!QtyOrderSA = aSumOrdSA
                            rsSR!QtySupply = asumSuply
                            rsSR!FR_LI = 0
                            rsSR!li_Supply = asumLiSup
                            rsSR!Li_BO = asumLiBO
                            rsSR!Li_Cancel = asumLiCan
                            rsSR!SR_Li = 0
    
                            'If fR100i <> 0 Then
                                rsSR!fr = isFRi 'Trim(Round((fR100i / fRALLi), 3) * 100)
                            'Else
                            '    rsSR!fr = ""
                            'End If
                            'If sRSPi + sRBCi <> 0 Then
                                rsSR!SR = isSRi 'Trim(Round((sRSPi / (sRSPi + sRBCi)), 3) * 100)
                            'Else
                            '    rsSR!SR = ""
                            'End If
                            rsSR!FullSup = 0 'SumFullSup
                            
                            isFRi = rsSR!fr
                            isSRi = rsSR!SR
                            .Update
                            aSumOrdSA = 0
                            asumSuply = 0
                            asumLiBO = 0
                            asumLiCan = 0
                            asumLiSup = 0
                            .AddNew
                        'End If
                        
                        rsSR!Tgl = nTgl1
                        rsSR!noWO = xWO
                        rsSR!NoWoFilter = ""
                        rsSR!nopol = ""
                        rsSR!job = ""
                        rsSR!JobFilter = nJob
                        rsSR!noRangka = ""
                        rsSR!model = ""
                        rsSR!kode = ""
                        rsSR!Nama = ""
                        rsSR!QtyOrderSA = sumOrdSA
                        rsSR!QtySupply = sumSuply
                        rsSR!FR_LI = 0
                        rsSR!li_Supply = sumLiSup
                        rsSR!Li_BO = sumLiBO
                        rsSR!Li_Cancel = sumLiCan
                        rsSR!SR_Li = 0
        
                        'If fR100i <> 0 Then
                        If selisih >= 0 And selisih <= CInt(jangka) Then
                            rsSR!fr = Trim(Round((fR100i / fRALLi), 3) * 100)
                        Else
                            rsSR!fr = "0"
                        End If
                        'Else
                        '    rsSR!fr = ""
                        'End If
                        If sRSPi + sRBCi <> 0 Then
                            rsSR!SR = Trim(Round((sRSPi / (sRSPi + sRBCi)), 3) * 100)
                        Else
                            rsSR!SR = "0"
                        End If
                        rsSR!FullSup = SumFullSup
                        isFRi = rsSR!fr
                        isSRi = rsSR!SR
                        .Update
                        xWO = 0
                        sumOrdSA = 0
                        sumSuply = 0
                        sumLiSup = 0
                        sumLiBO = 0
                        sumLiCan = 0
                        fR100i = 0
                        fRALLi = 0
                        sRSPi = 0
                        sRBCi = 0
                        SumFullSup = 0
                        
                        aSumOrdSA = 0
                        asumSuply = 0
                        asumLiCan = 0
                        asumLiCan = 0
                        asumLiBO = 0
                        
                        .AddNew
                        rsSR!Tgl = rsR.Fields(0).value
                    End If
                    nTgl1 = rsSR!Tgl
                Else
                    isNewTgl1 = False
                End If
                
                rsSR!noWO = rsR.Fields(1).value
                rsSR!NoWoFilter = rsR.Fields(1).value
                If rsSR!noWO <> nWO1 Then
                    If nWO1 <> "" And isNewTgl1 = False Then
                        rsSR!noWO = nWO1
                        rsSR!NoWoFilter = nWO1
                        rsSR!nopol = ""
                        rsSR!job = ""
                        rsSR!JobFilter = nJob
                        rsSR!noRangka = ""
                        rsSR!model = ""
                        rsSR!kode = ""
                        rsSR!Nama = ""
                        rsSR!QtyOrderSA = aSumOrdSA
                        rsSR!QtySupply = asumSuply
                        rsSR!FR_LI = 0
                        rsSR!li_Supply = asumLiSup
                        rsSR!Li_BO = asumLiBO
                        rsSR!Li_Cancel = asumLiCan
                        rsSR!SR_Li = 0

                        'If fR100i <> 0 Then
                            rsSR!fr = isFRi 'Trim(Round((fR100i / fRALLi), 3) * 100)
                        'Else
                        '    rsSR!fr = ""
                        'End If
                        'If sRSPi + sRBCi <> 0 Then
                            rsSR!SR = isSRi 'Trim(Round((sRSPi / (sRSPi + sRBCi)), 3) * 100)
                        'Else
                        '    rsSR!SR = ""
                        'End If
                        rsSR!FullSup = 0 'SumFullSup
                        isFRi = rsSR!fr
                        isSRi = rsSR!SR
                        .Update
                        
                        aSumOrdSA = 0
                        asumSuply = 0
                        asumLiSup = 0
                        asumLiBO = 0
                        asumLiCan = 0
                        
                        .AddNew
                        rsSR!noWO = rsR.Fields(1).value
                        rsSR!NoWoFilter = rsR.Fields(1).value
                        rsSR!Tgl = rsR.Fields(0).value
                    End If
                
                    xWO = xWO + 1
                    isNewWO1 = True
                    nWO1 = rsSR!noWO
                    rsSR!nopol = rsR.Fields(2).value
                    If Trim(rsR!FS_RC) = "R001" Then
                        rsSR!job = "SBI"
                    ElseIf Trim(rsR!FS_RC) = "" Then
                        If CInt(rsR!pds) = 1 Then
                            rsSR!job = "PDS"
                        ElseIf Trim(rsR!TPSVC) = "W" Then
                            rsSR!job = "WAR"
                        Else
                            rsSR!job = "GRP"
                        End If
                    Else
                        If Left(Trim(rsR!FS_RC), 1) = "R" Then
                            rsSR!job = "SBE"
                        Else
                            rsSR!job = ""
                        End If
                    End If
                    rsSR!JobFilter = rsSR!job
                            
                    rsSR!noRangka = rsR.Fields(3).value
                    rsSR!model = rsR.Fields(4).value
                Else
                    isNewWO1 = False
                    rsSR!nopol = ""
                    rsSR!job = ""
                    rsSR!JobFilter = nJob
                    rsSR!noRangka = ""
                    rsSR!model = ""
                    rsSR!noWO = ""
                    rsSR!NoWoFilter = rsR.Fields(1).value
                End If
                rsSR!kode = rsR.Fields(5).value
                rsSR!Nama = rsR.Fields(6).value
                rsSR!QtyOrderSA = rsR.Fields(7).value
                sumOrdSA = sumOrdSA + 1 'rsR.Fields(7).value
                aSumOrdSA = aSumOrdSA + 1
                Dim rsBO1 As ADODB.Recordset
                Set rsBO1 = objO.FetchDataACC("Select Qty,kodeitem from BOFILE Where  KodeItem = '" & rsSR!kode & "' and (NoSO = '" & nWO1 & "' OR NoSO =(Select top 1 (NoSPK) From SOORD Where NOSO ='" & nWO1 & "' and KodeItem = '" & rsSR!kode & "')) and Nomor = '" & rsR!nomor & "' ")
                If rsBO1.RecordCount > 0 Then
                    rsSR!Li_BO = 1
                    asumLiBO = asumLiBO + 1
                    getKodeItem(iKodeItem) = rsBO1.Fields(1).value
                    getNosoBO(iKodeItem) = nWO1
                    
                    
                    iKodeItem = iKodeItem + 1
                    rsSR!QtySupply = rsR.Fields(8).value - Val(rsBO1.Fields(0).value)
                Else
                    rsSR!Li_BO = 0
                    rsSR!QtySupply = rsR.Fields(8).value
                End If
                If rsSR!QtySupply <> 0 Then asumSuply = asumSuply + 1
                sumLiBO = sumLiBO + rsSR!Li_BO
                If selisih >= 0 And selisih <= jangka Then
                    rsSR!FR_LI = IIf(rsSR!QtyOrderSA = rsSR!QtySupply, 100, 0)
                Else
                    rsSR!FR_LI = "0"
                End If
                
                sumSuply = sumSuply + IIf(rsSR!QtySupply = 0, 0, 1) 'rsSR!QtySupply
                
                rsR.MoveNext
                If rsR.EOF = True Then
                    If nFR1 = 0 Then
                        rsSR!fr = "0"
                    Else
                        If rsSR!FR_LI = 100 And nFR1 = 100 Then
                        'If rsSR!FR_Li = 100 And nFR1 = 100 And (selisih >= 0 And selisih <= CInt(jangka)) Then
                        
                            nFR1 = 100
                        Else
                            nFR1 = 0
                        End If
                    End If
                    rsSR!fr = Trim(nFR1)
                    nFR1 = 100
                Else
                    If rsR.Fields(1).value = nWO1 Then
                        If rsSR!FR_LI <> 100 Then nFR1 = 0
                    Else
                        If nFR1 = 0 Then
                            rsSR!fr = "0"
                        Else
                            'If rsSR!fr_li = 100 Then
                            'If rsSR!FR_LI = 100 And (rsR.Fields(19).value >= 0 And rsR.Fields(19).value <= CInt(jangka)) Then
                            If rsSR!FR_LI = 100 And (selisih >= 0 And selisih <= CInt(jangka)) Then
                                nFR1 = 100
                            Else
                                nFR1 = 0
                            End If
                        End If
                        rsSR!fr = Trim(nFR1)
                        nFR1 = 100
                    End If
                End If
                rsR.MovePrevious
                
                
                SumFullSup = SumFullSup + IIf(rsSR!fr = 0, 0, 1)
                If rsSR!QtySupply = 0 Then
                rsSR!li_Supply = 0
                Else
                rsSR!li_Supply = rsR.Fields(10).value
                End If
                
                sRSP = sRSP + rsSR!li_Supply
                sumLiSup = sumLiSup + rsSR!li_Supply
                asumLiSup = asumLiSup + rsSR!li_Supply
                sRSPi = sRSPi + rsSR!li_Supply
                rsSR!Li_Cancel = rsR.Fields(12).value
                sRBC = sRBC + rsSR!Li_Cancel
                sumLiCan = sumLiCan + rsSR!Li_Cancel
                asumLiCan = asumLiCan + rsSR!Li_Cancel
                sRBCi = sRBCi + rsSR!Li_Cancel
                If isNewWO1 = False Then
                    nSupply = nSupply + rsSR!li_Supply
                    nBO = nBO + rsSR!Li_BO
                    nCancel = nCancel + rsSR!Li_Cancel
                Else
                    nSupply = rsSR!li_Supply
                    nBO = rsSR!Li_BO
                    nCancel = rsSR!Li_Cancel
                End If
                sRBC = sRBC + rsSR!Li_BO
                sRBCi = sRBCi + rsSR!Li_BO
                
                rsR.MoveNext
                If rsR.EOF = True Then
                    If nSupply = 0 Then
                        rsSR!SR = 0
                    Else
                        rsSR!SR = Trim(Round(((nSupply / (nSupply + nBO + nCancel)) * 100), 2))
                    End If
                Else
                    If rsR.Fields(1).value = nWO1 Then
                        rsSR!SR = ""
                    Else
                        If nSupply = 0 Then
                            rsSR!SR = "0"
                        Else
                            rsSR!SR = Trim(Round(((nSupply / (nSupply + nBO + nCancel)) * 100), 2))
                        End If
                        nSupply = 0
                        nBO = 0
                        nCancel = 0
                    End If
                End If
                rsR.MovePrevious
                               
                If rsSR!li_Supply = 1 And rsSR!Li_BO = 0 And rsSR!Li_Cancel = 0 Then
                    rsSR!SR_Li = 100
                ElseIf rsSR!li_Supply = 1 Then
                    rsSR!SR_Li = Round(100 / (rsSR!li_Supply + rsSR!Li_BO + rsSR!Li_Cancel), 2)
                Else
                    rsSR!SR_Li = 0
                End If
                
                'Select Case rsSR!fr
                 '   Case "100"
                  '      fR100 = fR100 + 1
                   '     fRALL = fRALL + 1
                    '    fR100i = fR100i + 1
                     '   fRALLi = fRALLi + 1
                   ' Case "0"
                    ';    fRALL = fRALL + 1
                      '  fRALLi = fRALLi + 1
               ' End Select
                
                If rsSR!fr = "100" And (rsR.Fields(19).value >= 0 And rsR.Fields(19).value <= CInt(jangka)) Then
                fR100 = fR100 + 1
                        fRALL = fRALL + 1
                        fR100i = fR100i + 1
                        fRALLi = fRALLi + 1
                ElseIf rsSR!fr = "0" Then
                fRALL = fRALL + 1
                        fRALLi = fRALLi + 1
                End If
                rsSR!FullSup = 0
                isFRi = rsSR!fr
                isSRi = rsSR!SR
                rsSR!fr = ""
                rsSR!SR = ""
               ' If rsR.Fields(19).value < 0 Or rsR.Fields(19).value <= jangka Then
               ' rsSR!fr = 0
               ' End If
                .Update
            End With
            rsR.MoveNext
        Loop
        
        rsR.MoveLast
        rsSR.AddNew
        rsSR!noWO = rsR.Fields(1).value
        rsSR!Tgl = rsR.Fields(0).value
        rsSR!NoWoFilter = rsR.Fields(1).value
        rsSR!nopol = ""
        rsSR!job = ""
        rsSR!JobFilter = nJob
        rsSR!noRangka = ""
        rsSR!model = ""
        rsSR!kode = ""
        rsSR!Nama = ""
        rsSR!QtyOrderSA = aSumOrdSA
        rsSR!QtySupply = asumSuply
        rsSR!FR_LI = 0
        rsSR!li_Supply = asumLiSup
        rsSR!Li_BO = asumLiBO
        rsSR!Li_Cancel = asumLiCan
        rsSR!SR_Li = 0

        'If fR100i <> 0 Then
            rsSR!fr = isFRi 'Trim(Round((fR100i / fRALLi), 3) * 100)
        'Else
        '    rsSR!fr = ""
        'End If
        'If sRSPi + sRBCi <> 0 Then
            rsSR!SR = isSRi 'Trim(Round((sRSPi / (sRSPi + sRBCi)), 3) * 100)
        'Else
        '    rsSR!SR = ""
        'End If
        rsSR!FullSup = 0 'SumFullSup
        isFRi = rsSR!fr
        isSRi = rsSR!SR
        rsSR.Update

        rsSR.AddNew
        rsSR!Tgl = rsR.Fields(0).value
        rsSR!noWO = xWO
        rsSR!NoWoFilter = ""
        rsSR!nopol = ""
        rsSR!job = ""
        rsSR!JobFilter = nJob
        rsSR!noRangka = ""
        rsSR!model = ""
        rsSR!kode = ""
        rsSR!Nama = ""
        rsSR!QtyOrderSA = sumOrdSA
        rsSR!QtySupply = sumSuply
        rsSR!FR_LI = 0
        rsSR!li_Supply = sumLiSup
        rsSR!Li_BO = sumLiBO
        rsSR!Li_Cancel = sumLiCan
        rsSR!SR_Li = 0

        If fR100i <> 0 Then
            rsSR!fr = Trim(Round(((fR100i / fRALLi) * 100), 1))
        Else
            rsSR!fr = ""
        End If
        If sRSPi + sRBCi <> 0 Then
            rsSR!SR = Trim(Round(((sRSPi / (sRSPi + sRBCi)) * 100), 1))
        Else
            rsSR!SR = ""
        End If
        rsSR!FullSup = SumFullSup
        isFRi = rsSR!fr
        isSRi = rsSR!SR
        rsSR.Update
        
        Set rsSR1 = rsSR.Clone
        Set GridFR1.DataSource = rsSR1
        rsSR1.filter = "Kode = '' and Nama = ''"
        
        Set rsSR2 = rsSR.Clone
        Set GridFR2.DataSource = rsSR2
        rsSR2.filter = "NOWOFilter = '' and Kode = '' and Nama = ''"

        If fR100 <> 0 Then
            GridFR.Columns(11).FooterText = Round(((fR100 / fRALL) * 100), 1)
            GridFR1.Columns(4).FooterText = Round(((fR100 / fRALL) * 100), 1)
            GridFR2.Columns(4).FooterText = Round(((fR100 / fRALL) * 100), 1)
        Else
            GridFR.Columns(11).FooterText = "0"
            GridFR1.Columns(4).FooterText = "0"
            GridFR2.Columns(4).FooterText = "0"
        End If
        If sRSP + sRBC <> 0 Then
            GridFR.Columns(16).FooterText = Round(((sRSP / (sRSP + sRBC)) * 100), 1)
            GridFR1.Columns(8).FooterText = Round(((sRSP / (sRSP + sRBC)) * 100), 1)
            GridFR2.Columns(8).FooterText = Round(((sRSP / (sRSP + sRBC)) * 100), 1)
        Else
            GridFR.Columns(16).FooterText = "0"
            GridFR1.Columns(8).FooterText = "0"
            GridFR2.Columns(8).FooterText = "0"
        End If
        
    End If
     
End Sub

Private Sub GridFR_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    On Error Resume Next
    Dim RsNav   As New ADODB.Recordset
    Set RsNav = rsSR.Clone
    If Bookmark > 1 Then RsNav.Bookmark = Bookmark
    If RsNav.Fields(6).value <> "" And RsNav.Fields(7).value <> "" Then
            RowStyle = "Normal0"
    Else
        If RsNav.Fields(19).value <> "" Then
            RowStyle = "Normal1"
        Else
            RowStyle = "Normal2"
        End If
    End If

End Sub

Private Sub GridFR2_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    On Error Resume Next
    Dim RsNav   As New ADODB.Recordset
    Set RsNav = rsSR2.Clone
    If Bookmark > 1 Then RsNav.Bookmark = Bookmark
    If RsNav.Fields(19).value <> "" Then
            RowStyle = "Normal0"
    Else
            RowStyle = "Normal1"
    End If

End Sub


Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
    XPButton1.Top = CRViewer1.Top + 345
    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = XPButton1.Left + XPButton1.Width + 30
    
    cmdCetak.Top = CRViewer1.Top + 345
    cmdCetak.Left = cboPrinter.Left + cboPrinter.Width + 30
End Sub

Private Sub DisableExport()
    Dim objPrinter As Printer
    Set objPrinter = Printer
    If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
        cboPrinter = ""
    Else
        cboPrinter = objPrinter.DeviceName
    End If
    cboPrinter.Clear
    
    For Each pr In Printers
        If InStr(UCase(pr.DeviceName), "PDF") = 0 Then cboPrinter.AddItem pr.DeviceName
    Next pr
    If Log.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        CRViewer1.EnablePrintButton = True
        cmdExport.Visible = False
        cboPrinter.Visible = False
        cmdCetak.Visible = False
        XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdCetak.Visible = True
        XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub

Private Sub cmdCetak_Click()
    For Each pr In Printers
        If UCase(pr.DeviceName) = UCase(cboPrinter) Then
            GoTo Cetak
        End If
    Next pr
    MsgBox "Printer tidak ditemukan Pilih Printer dahulu", vbInformation, "Proses Cetak Gagal"
    cboPrinter.SetFocus
    Exit Sub
    
Cetak:
    Dim Psize As Integer
    Dim POr As Integer
    
    Psize = Report.PaperSize
    POr = Report.PaperOrientation
    
    Set Printer = Printers(cboPrinter.ListIndex)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Printer.Port
    Report.PaperOrientation = POr
    Report.PaperSize = Psize
    Report.PrintOut
    
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Salah
    Dim NamaFile As String
    NamaFile = ""
    
    Dim f As New FormExplorer
    f.show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    
    If Log.PdfOnly = 1 Then
         
        NamaFile = MyPath & f.FileName
        
        If UCase(Right(Trim(f.FileName), 5)) = ".XLSX" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 5)
        If UCase(Right(Trim(f.FileName), 4)) = ".XLS" Then NamaFile = Mid(Trim(f.FileName), 1, Len(Trim(f.FileName)) - 4)
         
        
        Report.ExportOptions.DestinationType = crEDTDiskFile
        Report.ExportOptions.DiskFileName = NamaFile & ".Doc"
        Report.ExportOptions.RTFExportAllPages = True
        Report.ExportOptions.FormatType = crEFTExactRichText
        Report.Export False
        MsgBox NamaFile & ".Doc" & "  ... Data sudah di Export !", vbInformation
        Exit Sub
        With msg
            .PassingMsg = NamaFile & ".Doc" & "  ... Data sudah di Export !"
            .passingTipe = 2
            .show vbModal
        End With
        Exit Sub
    ElseIf onExpo = 0 Then
        Report.Export
    End If
    Exit Sub
Salah:
    If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
    MsgBox "Err.Line : " & Erl & "Err.Numb : " & err.Number & "Err.Desc. : " & err.Description & vbCr & "Procedure : " & Me.Name & " export"
Keluar:
End Sub

Private Sub cmdExportFill_Click()
    On Error GoTo Err_Export
    If rsSR.RecordCount = 0 Then Exit Sub
    
    Dim Er As New ErrInfo
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
    
    If Opt(0) Then
        GridFR.ExportToFile onVal, True
    ElseIf Opt(1) Then
        GridFR2.ExportToFile onVal, True
    ElseIf Opt(2) Then
        GridFR1.ExportToFile onVal, True
    ElseIf Opt(3) Then
        GridFR3.ExportToFile onVal, True
    End If
    'If cexAll.value = 0 Then
        
    'Else
     '   GridFR1.ExportToFile onVal, True
    'End If
    Er.HeadCap = "Export to excel !"
    Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                 "File : " & onVal
    Beep
    Er.show (vbModal)
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub cmdRef_Click()
    If Opt(0).value = True Then
        GridFR.Visible = True
        GridFR1.Visible = False
        GridFR2.Visible = False
        GridFR3.Visible = False
        If cboJob.Text <> "ALL" Then
            rsSR.filter = "JobFilter = '" & cboJob.Text & "' AND Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
        Else
'            rsSR.filter = "JobFilter = ''" ' AND Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
            rsSR.filter = "Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
        End If
    ElseIf Opt(1).value = True Then
        GridFR.Visible = False
        GridFR1.Visible = False
        GridFR2.Visible = True
        GridFR3.Visible = False
        If cboJob.Text <> "ALL" Then
            rsSR2.filter = "Kode = '' and Nama = '' AND JobFilter = '" & cboJob.Text & "' AND Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
        Else
            rsSR2.filter = "Kode = '' and Nama = '' AND Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
        End If
    ElseIf Opt(2).value = True Then
        GridFR.Visible = False
        GridFR1.Visible = True
        GridFR2.Visible = False
        GridFR3.Visible = False
        If cboJob.Text <> "ALL" Then
            rsSR1.filter = "Kode = '' and Nama = '' AND NoWOFilter = '' AND JobFilter = '" & cboJob.Text & "'  AND Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
        Else
            rsSR1.filter = "Kode = '' and Nama = '' AND NoWOFilter = '' AND Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
        End If
    ElseIf Opt(3).value = True Then
        GridFR.Visible = False
        GridFR1.Visible = False
        GridFR2.Visible = False
        GridFR3.Visible = True
        Dim i As Integer
        Dim inArray, inArray2, SQL As String
        i = 0
        Do While (i <= iKodeItem)
            If Trim(getKodeItem(i)) <> "" Then
               inArray = inArray & "'" & getKodeItem(i) & "',"
            End If
            
            If Trim(getNosoBO(i)) <> "" Then
               inArray2 = inArray2 & "'" & getNosoBO(i) & "',"
            End If
            
            i = i + 1
        Loop
        
        If inArray <> "" Then
            inArray = Left(inArray, Len(inArray) - 1)
        End If
        
        If inArray2 <> "" Then
            inArray2 = Left(inArray2, Len(inArray2) - 1)
        End If
        
        SQL = " select  NoSo,KodeItem,Keterangan,SUM(qty) as qty,max(unitprice) unitprice,SCC_Code,isnull(Tgl,'')  from ( select a.[DAte] as Tgl,s.noso,s.nomor,s.kodeitem,s.keterangan,case s.BO WHEN 'cancel' then s.Qty - (IsNull(s.QtyKirimBI,0) + IsNull(s.QtyKirimBL,0) + IsNull(s.QtyKeluar,0)) else b.Qty end as qty,i.HargaJualA as UnitPrice,i.SCC_Code "
        SQL = SQL + "from [ACC" & Trim(Right(gDb, 2)) & "-" & cboThnSR.Text & IIf(CInt(cboBlnSR.Text) < 10, "0" & cboBlnSR.Text, cboBlnSR.Text) & "]..Soord s "
        SQL = SQL + "left outer join [ACC" & Trim(Right(gDb, 2)) & "-" & cboThnSR.Text & IIf(CInt(cboBlnSR.Text) < 10, "0" & cboBlnSR.Text, cboBlnSR.Text) & "]..Inmst i on i.KodeItem = s.KodeItem "
        SQL = SQL + "left outer join T_invoice A on s.NoSo = A.No_WO "
        SQL = SQL + "left outer join [ACC" & Trim(Right(gDb, 2)) & "-" & cboThnSR.Text & IIf(CInt(cboBlnSR.Text) < 10, "0" & cboBlnSR.Text, cboBlnSR.Text) & "]..BoFile b on b.kodeitem = s.kodeitem "
        ''SQL = SQL + "where s.BO='cancel'  "
'        If inArray <> "" Then
'            SQL = SQL + "OR i.KodeItem in (" & inArray & ") "
'        End If
         SQL = SQL + "where "
        ''@ 2016-04-25 || ''
        If inArray <> "" Then
            SQL = SQL + "( s.BO='cancel' OR (i.KodeItem in (" & inArray & ") and s.noso in (" & inArray2 & ")))"
        Else
            SQL = SQL + " s.BO='cancel' "
        End If

        
        SQL = SQL + "and Right(A.noInv,1) <> 'B' and A.Fd_retur = '01/01/1900' "
        'SQL = SQL + "where s.BO='cancel' or i.KodeItem in (" & inArray & ") and Right(A.noInv,1) <> 'B' and A.Fd_retur = '01/01/1900' "
        'sql = sql + "where s.BO='cancel' and Right(A.noInv,1) <> 'B' and A.Fd_retur = '01/01/1900' "
        'sql = sql + "and Month(A.[Date]) = " & CInt(cboBlnSR.Text) & " and Year(A.[Date]) = " & CInt(cboThnSR.Text) & " ) fl where fl.qty > 0 group by noso,KodeItem,keterangan,SCC_Code"
        SQL = SQL + "and Month(A.[Date]) = " & CInt(cboBlnSR.Text) & " and Year(A.[Date]) = " & CInt(cboThnSR.Text) & " ) fl group by noso,KodeItem,keterangan,SCC_Code,Tgl"
  
  
  Set rsGetData = objO.FetchData(SQL)
   If rsSR3.State Then rsSR3.Close
    Set rsSR3 = Nothing
    With rsSR3
        .Fields.Append "No", adInteger
        .Fields.Append "NoWo", adVarChar, 100
        .Fields.Append "PartNomor", adVarChar, 50
        .Fields.Append "PartName", adVarChar, 50
        .Fields.Append "KodeSCC", adVarChar, 50
        .Fields.Append "qty", adVarChar, 30
        .Fields.Append "UnitPrice", adVarChar, 100
        .Fields.Append "Tprice", adVarChar, 100
        .Fields.Append "Remark", adVarChar, 100
       .Fields.Append "Tgl", adDate
        .Open
    End With
        If rsGetData.RecordCount > 0 Then
            rsGetData.MoveFirst
            i = 1
            
            Do Until rsGetData.EOF
                With rsSR3
                .AddNew
                rsSR3!No = i
                rsSR3!noWO = CStr(rsGetData.Fields(0).value)
                rsSR3!PartNomor = rsGetData.Fields(1).value
                rsSR3!PartName = rsGetData.Fields(2).value
                rsSR3!KodeSCC = rsGetData.Fields(5).value
                rsSR3!qty = CInt(rsGetData.Fields(3).value)
                rsSR3!UnitPrice = CDbl(rsGetData.Fields(4).value)
                rsSR3!TPrice = rsGetData.Fields(3).value * rsGetData.Fields(4).value
                rsSR3!Remark = ""
                rsSR3!Tgl = IIf(IsNull(rsGetData.Fields(6).value), "", rsGetData.Fields(6).value)
                End With
                i = i + 1
                 rsGetData.MoveNext
            Loop
        End If
           
        Set GridFR3.DataSource = rsSR3
        If cboJob.Text <> "ALL" Then
            rsSR3.filter = "Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
        Else
            rsSR3.filter = "Tgl >= '" & dtMulai & "' AND Tgl <= '" & dtSelesai & "'"
        End If
    End If
End Sub


