VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmFakturPajak 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cetak Faktur Pajak"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   18960
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   18720
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   720
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdTutup3 
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
      Height          =   300
      Left            =   17640
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   720
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer3 
      Height          =   8055
      Left            =   17520
      TabIndex        =   40
      Top             =   240
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
   Begin VB.CommandButton cmdSettingPrinter2 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   16800
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   1440
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdTutup4 
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
      Height          =   300
      Left            =   15720
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1440
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer4 
      Height          =   8055
      Left            =   15600
      TabIndex        =   47
      Top             =   960
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   9735
      Left            =   13920
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   15255
      Begin VB.CommandButton Command3 
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
         Left            =   12240
         Picture         =   "frmFakturPajak.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   8880
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Kembali ke &Faktur"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   8880
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Kembali ke Menu"
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
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   9240
         Width           =   2175
      End
      Begin TrueOleDBGrid70.TDBGrid GridR 
         Height          =   8175
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   14420
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No Invoice"
         Columns(0).DataField=   "NOINV"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Customer"
         Columns(1).DataField=   "CUSTOMER"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Alamat"
         Columns(2).DataField=   "Cust_Adres"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Kota"
         Columns(3).DataField=   "KOTA"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Jasa"
         Columns(4).DataField=   "JASA"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Part"
         Columns(5).DataField=   "PART"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Sublet"
         Columns(6).DataField=   "SUBLET"
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Total"
         Columns(7).DataField=   "TOTAL"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "No Faktur R"
         Columns(8).DataField=   "NOFAKTUR"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Harga Jual"
         Columns(9).DataField=   "HARGAJUAL"
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "PPN"
         Columns(10).DataField=   "PPN"
         Columns(10).NumberFormat=   "Standard"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Tanggal"
         Columns(11).DataField=   "Tanggal"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=5636"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5530"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=5636"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=5530"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=2646"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2540"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2302"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2196"
         Splits(0)._ColumnProps(20)=   "Column(4)._ColStyle=2"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(22)=   "Column(5).Width=2302"
         Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2196"
         Splits(0)._ColumnProps(25)=   "Column(5)._ColStyle=2"
         Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(27)=   "Column(6).Width=2302"
         Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2196"
         Splits(0)._ColumnProps(30)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(32)=   "Column(7).Width=2302"
         Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=2196"
         Splits(0)._ColumnProps(35)=   "Column(7)._ColStyle=2"
         Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(37)=   "Column(8).Width=4048"
         Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=3942"
         Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(41)=   "Column(9).Width=3281"
         Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=3175"
         Splits(0)._ColumnProps(44)=   "Column(9)._ColStyle=2"
         Splits(0)._ColumnProps(45)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(46)=   "Column(10).Width=3281"
         Splits(0)._ColumnProps(47)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(48)=   "Column(10)._WidthInPix=3175"
         Splits(0)._ColumnProps(49)=   "Column(10)._ColStyle=2"
         Splits(0)._ColumnProps(50)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(51)=   "Column(11).Width=2831"
         Splits(0)._ColumnProps(52)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(53)=   "Column(11)._WidthInPix=2725"
         Splits(0)._ColumnProps(54)=   "Column(11)._ColStyle=1"
         Splits(0)._ColumnProps(55)=   "Column(11).Order=12"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFCC&"
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
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.alignment=1"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13,.alignment=1"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=82,.parent=13,.alignment=2"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
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
         _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
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
         _StyleDefs(109) =   "Named:id=87:invoiceRetur"
         _StyleDefs(110) =   ":id=87,.parent=33,.fgcolor=&HFF&"
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: DATA FAKTUR PAJAK - INVOICE RETUR ::."
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
         Left            =   5280
         TabIndex        =   70
         Top             =   60
         Width           =   4665
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   15135
      End
   End
   Begin VB.CommandButton cmdReturCetak 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Data Invoice &Retur"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10920
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   9240
      Width           =   2175
   End
   Begin VB.Frame frmPrinter 
      BackColor       =   &H00FFFFC0&
      Caption         =   "               Pilih Printer               "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   6000
      TabIndex        =   56
      Top             =   3360
      Visible         =   0   'False
      Width           =   3255
      Begin VB.OptionButton prtrn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LQ2190"
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   62
         Top             =   600
         Width           =   1095
      End
      Begin VB.OptionButton prtrn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LX"
         Height          =   255
         Index           =   2
         Left            =   1800
         TabIndex        =   60
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cmdLanjut 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Cetak Fa&ktur"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   1320
         Width           =   1815
      End
      Begin VB.OptionButton prtrn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LQ 300"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   58
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton prtrn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "LQ 2180"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   57
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C00000&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00C00000&
         Height          =   1455
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame frmFaktur 
      BackColor       =   &H00FFC0C0&
      Height          =   4815
      Left            =   3960
      TabIndex        =   25
      Top             =   2280
      Visible         =   0   'False
      Width           =   6855
      Begin VB.CheckBox chGanti 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Penggantian Dicoret"
         Height          =   195
         Left            =   4740
         TabIndex        =   65
         Top             =   3970
         Width           =   1935
      End
      Begin VB.CheckBox chHarga 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Harga Jual Dicoret"
         Height          =   195
         Left            =   4740
         TabIndex        =   64
         Top             =   3750
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   1
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1680
         Width           =   5340
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   0
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   7
         Top             =   2040
         Width           =   5340
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   8
         Top             =   2400
         Width           =   5340
      End
      Begin VB.CheckBox chJabatan 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Display Jabatan"
         Height          =   195
         Left            =   4740
         TabIndex        =   55
         Top             =   1150
         Width           =   1695
      End
      Begin VB.CheckBox chSilang 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Keterangan Dicoret"
         Height          =   195
         Left            =   4740
         TabIndex        =   53
         Top             =   3533
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.ComboBox cboTtd 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4740
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CheckBox chInPPn 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Harga Jual Include PPn"
         Height          =   195
         Left            =   4740
         TabIndex        =   43
         Top             =   1380
         Width           =   2055
      End
      Begin VB.ComboBox cboType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton cmdCtkFaktur 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Cetak &Faktur"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4320
         Width           =   1815
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   300
         Index           =   7
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   45
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin TDBDate6Ctl.TDBDate tglFaktur 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   960
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   582
         Calendar        =   "frmFakturPajak.frx":047C
         Caption         =   "frmFakturPajak.frx":05A8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmFakturPajak.frx":0614
         Keys            =   "frmFakturPajak.frx":0632
         Spin            =   "frmFakturPajak.frx":0690
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   300
         Index           =   4
         Left            =   1320
         TabIndex        =   29
         Top             =   240
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   529
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
         Height          =   300
         Index           =   0
         Left            =   1320
         TabIndex        =   15
         Top             =   2040
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
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
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   14
         Top             =   1680
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox Text 
         Height          =   300
         Index           =   2
         Left            =   1320
         TabIndex        =   9
         Top             =   2760
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   529
         _Version        =   393216
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
      Begin MSMask.MaskEdBox Text 
         Height          =   300
         Index           =   3
         Left            =   1320
         TabIndex        =   13
         Top             =   3840
         Visible         =   0   'False
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   529
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   320
         Left            =   6200
         TabIndex        =   44
         Top             =   2760
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   300
         Index           =   5
         Left            =   1320
         TabIndex        =   16
         Top             =   2400
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   16777215
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
         Height          =   300
         Index           =   6
         Left            =   1320
         TabIndex        =   10
         Top             =   3120
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   529
         _Version        =   393216
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
      Begin MSMask.MaskEdBox Text 
         Height          =   300
         Index           =   8
         Left            =   1320
         TabIndex        =   11
         Top             =   3480
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   529
         _Version        =   393216
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
      Begin ucXPButton.XPButton cmd1 
         Height          =   330
         Left            =   6320
         TabIndex        =   75
         ToolTipText     =   "00.000.000.0-000.000"
         Top             =   2760
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   582
         BTYPE           =   7
         TX              =   "&v"
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
         MICON           =   "frmFakturPajak.frx":06B8
         PICN            =   "frmFakturPajak.frx":06D4
         PICH            =   "frmFakturPajak.frx":09A9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Ttd"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   360
         TabIndex        =   50
         Top             =   3165
         Width           =   810
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jabatan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   540
         TabIndex        =   49
         Top             =   3525
         Width           =   630
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   780
         TabIndex        =   45
         Top             =   2445
         Width           =   375
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPPKP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   615
         TabIndex        =   34
         Top             =   3885
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Npwp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   675
         TabIndex        =   33
         Top             =   2805
         Width           =   480
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   375
         TabIndex        =   32
         Top             =   1725
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   31
         Top             =   2085
         Width           =   555
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Faktur"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   30
         Top             =   285
         Width           =   990
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   795
         TabIndex        =   28
         Top             =   1372
         Width           =   360
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   510
         TabIndex        =   27
         Top             =   1020
         Width           =   645
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   720
         TabIndex        =   26
         Top             =   645
         Width           =   435
      End
      Begin VB.Image btnBatal 
         Height          =   375
         Index           =   1
         Left            =   6480
         Picture         =   "frmFakturPajak.frx":0C7E
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
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
      Left            =   13485
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   9240
      Width           =   1770
   End
   Begin VB.ComboBox cboNoFP 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      TabIndex        =   22
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdCreateFP 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Buat Faktur Pajak"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9240
      Width           =   2175
   End
   Begin VB.CommandButton cmdCetakUlang 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cetak Faktur Pajak"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9240
      Width           =   2175
   End
   Begin VB.ComboBox cboFaktur 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   150
      Width           =   2655
   End
   Begin TDBDate6Ctl.TDBDate DtDOAwal 
      Height          =   330
      Left            =   8520
      TabIndex        =   1
      Top             =   165
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmFakturPajak.frx":4599
      Caption         =   "frmFakturPajak.frx":46C5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFakturPajak.frx":4731
      Keys            =   "frmFakturPajak.frx":474F
      Spin            =   "frmFakturPajak.frx":47AD
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate DtDOAkhir 
      Height          =   330
      Left            =   10560
      TabIndex        =   2
      Top             =   165
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmFakturPajak.frx":47D5
      Caption         =   "frmFakturPajak.frx":4901
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmFakturPajak.frx":496D
      Keys            =   "frmFakturPajak.frx":498B
      Spin            =   "frmFakturPajak.frx":49E9
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TrueOleDBGrid70.TDBGrid GridProgram 
      Height          =   7695
      Left            =   120
      TabIndex        =   42
      Top             =   1200
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   13573
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ST"
      Columns(0).DataField=   "ST"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NFP"
      Columns(1).DataField=   "NFP"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "TGLFP"
      Columns(2).DataField=   "TGLFP"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "INVO"
      Columns(3).DataField=   "INVO"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TGLINVO"
      Columns(4).DataField=   "TGLINVO"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "NOPOL"
      Columns(5).DataField=   "NOPOL"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "WO"
      Columns(6).DataField=   "WO"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "RGK"
      Columns(7).DataField=   "RGK"
      Columns(7).NumberFormat=   "Standard"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "CUST"
      Columns(8).DataField=   "CUST"
      Columns(8).NumberFormat=   "Standard"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "TTL"
      Columns(9).DataField=   "TTL"
      Columns(9).NumberFormat=   "Standard"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "W"
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "PLG"
      Columns(11).DataField=   "PLG"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "R"
      Columns(12).DataField=   "R"
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
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&H0&"
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
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFCC&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=86,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=83,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=84,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=85,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=50,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=74,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=71,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=72,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=73,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=82,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
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
      _StyleDefs(113) =   "Named:id=87:invoiceRetur"
      _StyleDefs(114) =   ":id=87,.parent=33,.fgcolor=&HFF&"
   End
   Begin VB.CommandButton cmdTtd 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Default &Tanda Tangan"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   9240
      Width           =   2175
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   120
      Width           =   1695
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
      Left            =   14400
      Picture         =   "frmFakturPajak.frx":4A11
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton cmdCetakRekap 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Cetak &Khusus"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   67
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "* Warna Merah : Invoice diretur"
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
      Left            =   12330
      TabIndex        =   63
      Top             =   8950
      Width           =   2925
   End
   Begin VB.Label lblJumlah 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Invoice :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10440
      TabIndex        =   41
      Top             =   720
      Width           =   1560
   End
   Begin VB.Label Label13 
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
      Left            =   120
      TabIndex        =   38
      Top             =   9120
      Width           =   1875
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
      Left            =   2160
      TabIndex        =   37
      Top             =   9120
      Width           =   180
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
      Left            =   3600
      TabIndex        =   36
      Top             =   9120
      Width           =   180
   End
   Begin VB.Label Label12 
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
      Left            =   3240
      TabIndex        =   35
      Top             =   9120
      Width           =   330
   End
   Begin VB.Label lblFaktur 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Look Up No. Faktur Pajak :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   23
      Top             =   780
      Width           =   2505
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   14040
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipe Faktur"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   19
      Top             =   210
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   10080
      TabIndex        =   18
      Top             =   210
      Width           =   345
   End
   Begin VB.Label label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tgl Invoice dari :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6735
      TabIndex        =   17
      Top             =   210
      Width           =   1605
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00A99778&
      Height          =   975
      Left            =   14280
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmFakturPajak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsInvoice As New ADODB.Recordset
Dim vNamaTabel As String
Dim kdDealer As String
Dim kdType As String
Dim objBilling As New BClsBilling
Dim namaPT As String
Dim alamatPT As String
Dim npwpPT As String
Dim tglKukuhPT As String
Dim thnInv As String
Dim NoBaru As Boolean
Dim countInv As Integer
Dim rpt As Object
Dim rpt2 As Object
Dim cols As TrueOleDBGrid70.Columns
Dim noFakturCetak As String

Private Sub refreshList(Optional nFP As String)
    If Nul(DtDOAkhir) <> "" And Nul(DtDOAkhir) <> "" Then
        nTagih = 0
        totTagih = 0
        vNamaTabel = "LISTFAKTUR" & DriveSerialNumber("c:")
        Select Case cboFaktur
            Case "Faktur Pajak"
                Call panggilINV(DtDOAwal, DtDOAkhir, 1, nFP)
                'Call displayFaktur(DtDOAwal, DtDOAkhir, 1, nFP)
            Case "Faktur Pajak Gabungan"
                Call panggilINV(DtDOAwal, DtDOAkhir, 2, nFP)
                'Call displayFaktur(DtDOAwal, DtDOAkhir, 2, nFP)
            Case "Faktur Pajak Retur"
                vNamaTabel = "LISTRETUR" & DriveSerialNumber("c:")
                'Call panggilINV(DtDOAwal, DtDOAkhir, 2, nFP)
        End Select
    End If
End Sub

Private Sub displayFaktur(tglAwal As Date, tglAkhir As Date, idx As Integer, noFP As String)
      Dim rsSearch As ADODB.Recordset

1        On Error GoTo displayFaktur_Error

2         Select Case idx
              Case 1
3                 If Trim(cboNoFP) <> "" Then
4                     sSQL = "SELECT A.NOFPS, TGLFPS, NPWP,NPPKP,KETERANGAN, HARGAJUAL+PPN AS HARGAJUAL, PPN, NAMAFP, ALAMATFP1, ALAMATFP2, KODELGN, MODUL, KODEUSER, TGLENTRY, TGLINVOICE, '' AS NO_POLISI, NOINVOICE FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..SIFPSTANDAR A WITH(NOLOCK) INNER JOIN [FAKTUR" & Trim(Right(gDb, 2)) & "]..INVOICE B WITH(NOLOCK) ON A.NOFPS = B.NOFPS WHERE A.NOFPS LIKE '%" & Trim(noFP) & "%' AND MODUL = 'SVC' AND A.NOFPS NOT IN (SELECT ISNULL(NOFPS,'') AS NOFPS FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..FPGABUNGAN WITH(NOLOCK) WHERE MODUL = 'SVC')"
5                 Else
6                     sSQL = "SELECT A.NOFPS, TGLFPS, NPWP,NPPKP,KETERANGAN, HARGAJUAL+PPN AS HARGAJUAL, PPN, NAMAFP, ALAMATFP1, ALAMATFP2, KODELGN, MODUL, KODEUSER, TGLENTRY, TGLINVOICE, '' AS NO_POLISI, NOINVOICE FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..SIFPSTANDAR A WITH(NOLOCK) INNER JOIN [FAKTUR" & Trim(Right(gDb, 2)) & "]..INVOICE B WITH(NOLOCK) ON A.NOFPS = B.NOFPS WHERE TGLFPS BETWEEN '" & Format(tglAwal, "MM/DD/YYYY") & "' AND '" & Format(tglAkhir, "MM/DD/YYYY") & "' AND MODUL = 'SVC' AND A.NOFPS NOT IN (SELECT ISNULL(NOFPS,'') as NOFPS FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..FPGABUNGAN WITH(NOLOCK) WHERE MODUL = 'SVC')"
7                 End If
8             Case 2
9                 If Trim(cboNoFP) <> "" Then
10                    sSQL = "SELECT NOFPS, (SELECT TGLFPS FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..SIFPSTANDAR WITH(NOLOCK) WHERE NOFPS = A.NOFPS) AS TGLFPS, NPWP, NPPKP, 'No Invoice ' + NOINVOICE + ', ' + KET2 + ', No Polisi '+ KET1 AS KETERANGAN, DPP+PPN AS HARGAJUAL, PPN,NAMAFP, ALAMAT1 AS ALAMATFP1, ALAMAT2 AS ALAMATFP2, KODELGN, MODUL, KODEENTRY AS KODEUSER, TGLENTRY, TGLINVOICE, KET1 AS NO_POLISI, NOINVOICE FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..FPGABUNGAN A WITH(NOLOCK) WHERE NOFPS LIKE '%" & Trim(noFP) & "%' AND MODUL = 'SVC'"
11                Else
12                    sSQL = "SELECT NOFPS, (SELECT TGLFPS FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..SIFPSTANDAR WITH(NOLOCK) WHERE NOFPS = A.NOFPS) AS TGLFPS, NPWP, NPPKP, 'No Invoice ' + NOINVOICE + ', ' + KET2 + ', No Polisi '+ KET1 AS KETERANGAN, DPP+PPN AS HARGAJUAL, PPN,NAMAFP, ALAMAT1 AS ALAMATFP1, ALAMAT2 AS ALAMATFP2, KODELGN, MODUL, KODEENTRY AS KODEUSER, TGLENTRY, TGLINVOICE, KET1 AS NO_POLISI, NOINVOICE FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..FPGABUNGAN A WITH(NOLOCK) WHERE TGLINVOICE BETWEEN '" & Format(tglAwal, "MM/DD/YYYY") & "' AND '" & Format(tglAkhir, "MM/DD/YYYY") & "' AND MODUL = 'SVC'"
13                End If
14        End Select
15        Set rsSearch = objO.FetchData(sSQL)
16        If Not rsSearch.EOF Then
17            If rsInvoice.State <> adStateClosed Then rsInvoice.Close
18            If rsInvoice.State = adStateClosed Then Call createRsInv(idx)
19            Do While Not rsSearch.EOF
20                With rsInvoice
21                    .AddNew
22                    !ST = False
23                    !nFP = Trim(rsSearch("NOFPS"))
24                    !TGLFP = Nul(Trim(rsSearch("TGLFPS")))
25                    !INVO = Trim(rsSearch("NOINVOICE"))
26                    !TGLINVO = Trim(rsSearch("TGLINVOICE"))
27                    !noPol = Trim(rsSearch("NO_POLISI"))
28                    !wo = ""
29                    !RGK = ""
30                    !cust = Trim(rsSearch("NAMAFP"))
31                    !TTL = RepRp(CDbl(rsSearch("HARGAJUAL")))
32                    !PLG = Nul(Trim(rsSearch("KODELGN")))
                      '!W = Trim(rsSearch("WAR"))
33                    .Update
34                End With
35                rsSearch.MoveNext
36            Loop
37            GridProgram.DataSource = rsInvoice
38            Call headGrid(idx)
39            lblJumlah = "Jumlah Invoice = " & rsInvoice.RecordCount
40            rsInvoice.MoveFirst
41        End If

42       On Error GoTo 0
43       Exit Sub

displayFaktur_Error:

44        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure displayFaktur of Form frmFakturPajak")
End Sub

Private Sub BuatTabelInv(vNamaTabel As String)
1        On Error GoTo BuatTabelInv_Error

2         sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
               & " drop table [tempdb]..[" & vNamaTabel & "]"
3         Call objO.UpdateDataACC(sSQL)
4         sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
               & "    [SA][char](5) NOT NULL DEFAULT '', " _
               & "    [NOINV][varchar](15) NOT NULL DEFAULT '', " _
               & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
               & "    [NO_POLISI][varchar](12) NOT NULL DEFAULT '', " _
               & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
               & "    [JOR][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [JRTE][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [RTJ][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [THS][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [SBI][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [SBE][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [GRP][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [TWC][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [LUB][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [PDI][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [WAS][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [BOR][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [MATERIAL][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [PART][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [OLI][decimal](18,2) NOT NULL DEFAULT 0, "
               
5         sSQL = sSQL + "[SUBLET][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [SUBTOT][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [METERAI][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [TOTAL][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [TIPEINV][varchar](20) NOT NULL DEFAULT '', " _
               & "    [CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
               & "    [LABOR][decimal](18,2) NOT NULL DEFAULT 0, " _
               & "    [NOFAKTUR][varchar](30) NOT NULL DEFAULT ''," _
               & "    [HPP][decimal](18,2) NOT NULL DEFAULT 0," _
               & "    [INVDATE][datetime] NOT NULL DEFAULT '01/01/3000'," _
               & "    [NOFP][varchar](30) NOT NULL DEFAULT ''," _
               & "    [NOMB][varchar](30) NOT NULL DEFAULT '', " _
               & "    [WAR][char](1) NOT NULL DEFAULT '0', " _
               & "    [KDPLG][varchar](50) NOT NULL DEFAULT '', " _
               & "    [DIRETUR][char](1) NOT NULL DEFAULT '0') ON [PRIMARY]"
6         Call objO.UpdateDataACC(sSQL)

7        On Error GoTo 0
8        Exit Sub

BuatTabelInv_Error:

9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure BuatTabelInv of Form frmFakturPajak")
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset
    createTemp = False
    Set rsHlp = objO.FetchDataACC("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelInv(vNamaTabel)
        createTemp = True
    End If
End Function

Private Sub panggilINV(DtAwal As String, DtAkhir As String, idx As Integer, Optional nFP As String)
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim hlp6 As String
      Dim hlp7 As String
      Dim rsHlp As ADODB.Recordset
      Dim rsAdd As ADODB.Recordset
      Dim rsTemp As ADODB.Recordset
      Dim nNilai As Double
      Dim nDuoInv As Boolean
      Dim subletOnly As Boolean
      Dim noInv As String
      Dim nmCust As String
      Dim nStatus As String
      Dim kdCust As String

1        On Error GoTo panggilINV_Error
          
2         If Log.RPLHistori = 0 Then
3             Set rsTemp = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, NOFP, NOFP2, NOMB, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP, FD_RETUR " _
                                   & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE A.NO_WO IN (SELECT NO_WO FROM JOBDTA WITH(NOLOCK) WHERE DELIV <> 'BTL') AND INVDATE BETWEEN '" & Format(DtAwal, "MM/DD/YYYY") & "' AND '" & Format(DtAkhir, "MM/DD/YYYY") & "' ORDER BY NOINV")
4         Else
5             Set rsTemp = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, NOFP, NOFP2, NOMB, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP, FD_RETUR " _
                                   & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE A.NO_WO IN (SELECT NO_WO FROM JOBDTA WITH(NOLOCK) WHERE DELIV <> 'BTL') AND INVDATE BETWEEN '" & Format(DtAwal, "MM/DD/YYYY") & "' AND '" & Format(DtAkhir, "MM/DD/YYYY") & "' ORDER BY NOINV")
6         End If
7         If Not rsTemp.EOF Then
8             Call BuatTabelInv(vNamaTabel)
9             Do While Not rsTemp.EOF
10                subletOnly = True
11                For i = 1 To 8
12                    hlp = "REP" & i
13                    If rsTemp(hlp) <> "" Then
14                        subletOnly = False
15                        hlp2 = "KLPKERJA" & i
16                        hlp3 = "RTE" & i    '"ACTRATE" & i
17                        hlp4 = "MEKANIK" & i
18                        hlp5 = "OFF" & i
19                        hlp6 = "DISC" & i
20                        hlp7 = "FREEC" & i
21                        If createTemp(vNamaTabel) Then
22                            nDuoInv = False
23                            If rsTemp("NDUOINV") = 1 Then
24                                nDuoInv = True
25                            End If
                              
26                            nStatus = rsTemp("STATUS")
27                            nNilai = Round(CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100)))
28                            noInv = Trim(rsTemp("NOINV"))
29                            nmCust = Trim(rsTemp("CUSTOMER"))
30                            kdCust = Trim(rsTemp("FS_KD_PLG"))
                              
31                            If nStatus = "AFFNONPPN" Then
32                                nNilai = Round(nNilai * (100 / 110))
33                            End If
                              
34                            Set rsHlp = objO.FetchDataACC("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
35                            If Not rsHlp.EOF Then
36                                Call objO.UpdateDataACC("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                               & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & RepRp(IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0)) & "', " _
                                               & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & RepRp(IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0)) & "', " _
                                               & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & RepRp(IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0)) & "', " _
                                               & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & RepRp(IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0)) & "', " _
                                               & "BOR = '" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', MATERIAL = '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', PART = '" & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', " _
                                               & "OLI = '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', SUBLET = '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', SUBTOT = SUBTOT + '" & RepRp(setKoma((IIf(rsTemp(hlp2) = "SBI", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "SBE", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", (CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "LUB", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "PDS", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "WAS", (CDbl(nNilai)), 0)))) & "', NOFP = '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', " _
                                               & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', HPP = '" & RepRp(cekDuo(CDbl(rsTemp("HPP")), noInv)) & "', TOTAL = TOTAL + '" & setKoma((IIf(rsTemp(hlp2) = "SBI", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "SBE", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", (CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "LUB", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "PDS", (CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "WAS", (CDbl(nNilai)), 0))) & "', " _
                                               & "WAR = '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "', DIRETUR = '" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "' " _
                                               & ",KDPLG = '" & Trim(kdCust) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
37                            Else
38                                Call objO.InsertDataACC("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                               & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "','" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "'" & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', " _
                                               & "'" & setKoma((IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) + IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "LUB", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "PDS", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "WAS", CDbl(nNilai), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                               & " '" & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & "', '" & setKoma((IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) + IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "LUB", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "PDS", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "WAS", CDbl(nNilai), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                               & "'" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', '" & setKoma(cekDuo(CDbl(rsTemp("HPP")), noInv)) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "','" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "','" & Trim(rsTemp("NOMB")) & "','" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "','" & Trim(kdCust) & "','" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "')")
39                            End If
40                        End If
41                    End If
42                Next
43                If subletOnly Then
44                    nDuoInv = False
45                    If rsTemp("NDUOINV") = 1 Then nDuoInv = True
46                    noInv = Trim(rsTemp("NOINV"))
47                    kdCust = Trim(rsTemp("FS_KD_PLG"))
                  
48                    Set rsHlp = objO.FetchDataACC("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & rsTemp("NOINV") & "'")
49                    If Not rsHlp.EOF Then
50                        Call objO.UpdateDataACC("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                       & "JOR = '0', TIPEINV = '" & Trim(rsTemp("STATUS")) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', BOR = '" & setKoma(CDbl(rsTemp("BODY"))) & "', MATERIAL = '" & setKoma(CDbl(rsTemp("MTRL"))) & "', " _
                                       & "PART = '" & setKoma(CDbl(rsTemp("PART_G") + rsTemp("PART_NG"))) & "', OLI = '" & setKoma(CDbl(rsTemp("OIL_G") + rsTemp("OIL_NG"))) & "', SUBLET = '" & setKoma(CDbl(rsTemp("SBLT"))) & "', NOFP = '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', " _
                                       & "LABOR = '" & setKoma(CDbl(rsTemp("LABOR"))) & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', HPP = '" & setKoma(CDbl(rsTemp("HPP"))) & "', WAR = '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "' " _
                                       & ",KDPLG = '" & Trim(kdCust) & "', DIRETUR = '" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & rsTemp("NOINV") & "'")
51                    Else
52                        Call objO.InsertDataACC("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                       & "'0','0', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                       & "'" & setKoma(CDbl(rsTemp("BODY"))) & "', '" & setKoma(CDbl(rsTemp("MTRL"))) & "', '" & setKoma(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))) & "', '" & setKoma(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))) & "', '" & setKoma(CDbl(rsTemp("SBLT"))) & "', '" & setKoma(CDbl(cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                       & " '" & setKoma(CDbl(rsTemp("MET"))) & "', '" & setKoma(CDbl(cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                       & " '" & Trim(rsTemp("STATUS")) & "', '" & Trim(rsTemp("CUSTOMER")) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & Trim(rsTemp("NOFAKTUR")) & "','" & CDbl(rsTemp("HPP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', '" & Trim(rsTemp("NOMB")) & "', '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "','" & Trim(kdCust) & "','" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "')")
53                    End If
54                End If
                  
                  ' DARI ADD JOB
55                Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND FS_LOG_DEL = '" & Trim(rsTemp("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
56                If Not rsAdd.EOF Then
57                    Do While Not rsAdd.EOF
58                        hlp2 = "KLPKERJA"
59                        hlp3 = "RTE"
60                        hlp4 = "MEKANIK"
61                        hlp5 = "OFFN"
62                        hlp6 = "FREEC"
                          
63                        nDuoInv = False
64                        If rsTemp("NDUOINV") = 1 Then
65                            nDuoInv = True
66                        End If
                          
67                        nStatus = rsTemp("STATUS")
                          'nNilai = Round(CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100)))
68                        nNilai = CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsTemp("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
69                        noInv = Trim(rsTemp("NOINV"))
70                        nmCust = Trim(rsTemp("CUSTOMER"))
71                        kdCust = Trim(rsTemp("FS_KD_PLG"))
                          
72                        If Trim(nStatus) = "AFFNONPPN" Then
73                            nNilai = Round(nNilai * (100 / 110))
74                        End If
                          
                           
75                        Set rsHlp = objO.FetchDataACC("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
76                        If Not rsHlp.EOF Then
77                            Call objO.UpdateDataACC("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                           & "JOR = 0, JRTE = JRTE + " & Replace(CDbl(rsAdd(hlp3)), ",", ".") & ", SBI = SBI + " & RepRp(IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0)) & ", " _
                                           & "SBE = SBE + " & RepRp(IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0)) & ", GRP = GRP + " & RepRp(IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0)) & ", " _
                                           & "TWC = TWC + " & RepRp(IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0)) & ", LUB = LUB + " & RepRp(IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0)) & ", " _
                                           & "PDI = PDI + " & RepRp(IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0)) & ", WAS = WAS + " & RepRp(IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0)) & ", " _
                                           & "BOR = " & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & ", MATERIAL = " & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & ", PART = " & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & ", " _
                                           & "OLI = " & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & ", SUBLET = " & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & ", SUBTOT = SUBTOT + " & RepRp(setKoma((IIf(rsAdd(hlp2) = "SBI", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", (CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", (CDbl(nNilai)), 0)))) & ", HPP = " & RepRp(cekDuo(CDbl(rsTemp("HPP")), noInv)) & ", " _
                                           & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = " & Replace(CDbl(rsTemp("LABOR")), ",", ".") & ", NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', NOFP = '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', " _
                                           & "TOTAL = TOTAL + " & RepRp(setKoma((IIf(rsAdd(hlp2) = "SBI", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", (CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", (CDbl(nNilai)), 0)))) & ", WAR = '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "'" _
                                           & ",KDPLG = '" & Trim(kdCust) & "', DIRETUR = '" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
78                        Else
79                            Call objO.InsertDataACC("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                           & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & "', " _
                                           & "'" & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', " _
                                           & "'" & setKoma(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', '" & setKoma((IIf(rsAdd(hlp2) = "SBI", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", (CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", (CDbl(nNilai)), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                           & "'" & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & "', '" & setKoma((IIf(rsAdd(hlp2) = "SBI", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", (CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", (CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", (CDbl(nNilai)), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                           & "'" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "','" & cekDuo(CDbl(rsTemp("HPP")), noInv) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "','" & Trim(rsTemp("NOMB")) & "','" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "','" & Trim(kdCust) & "', '" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "')")
80                        End If
81                        If rsHlp.State <> adStateClosed Then rsHlp.Close
82                        rsAdd.MoveNext
83                    Loop
84                End If
85                If rsAdd.State <> adStateClosed Then rsAdd.Close
86                rsTemp.MoveNext
87            Loop
88            Call tampilGrid(idx, nFP)
89        End If

90       On Error GoTo 0
91       Exit Sub

panggilINV_Error:

92        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilINV of Form frmFakturPajak")
End Sub

Private Sub setCetakLampiran(idx As Integer, dNamaTabel As String, Optional nFP As String)
      Dim hlp As String
      Dim hlp2 As String
      Dim hlp3 As String
      Dim hlp4 As String
      Dim hlp5 As String
      Dim hlp6 As String
      Dim hlp7 As String
      Dim rsHlp As ADODB.Recordset
      Dim rsAdd As ADODB.Recordset
      Dim rsTemp As ADODB.Recordset
      Dim nNilai As Double
      Dim nDuoInv As Boolean
      Dim subletOnly As Boolean
      Dim noInv As String
      Dim nmCust As String
      Dim nStatus As String
      Dim kdCust As String

1        On Error GoTo setCetakLampiran_Error
          
2         If Log.RPLHistori = 0 Then
3             Set rsTemp = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, NOFP, NOFP2, NOMB, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE A.NO_WO IN (SELECT NO_WO FROM JOBDTA WITH(NOLOCK) WHERE DELIV <> 'BTL') AND NOFP = '" & Trim(nFP) & "' OR NOFP2 = '" & Trim(nFP) & "' ORDER BY NOINV")
4         Else
5             Set rsTemp = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                                   & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, NOFP, NOFP2, NOMB, " _
                                   & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WITH(NOLOCK) WHERE NO_WO = A.NO_WO) AS HPP " _
                                   & "FROM T_INVOICE A WITH(NOLOCK) INNER JOIN CUSTDAT B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z WITH(NOLOCK) ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM C WITH(NOLOCK) ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                                   & "WHERE A.NO_WO IN (SELECT NO_WO FROM JOBDTA WITH(NOLOCK) WHERE DELIV <> 'BTL') AND NOFP = '" & Trim(nFP) & "' OR NOFP2 = '" & Trim(nFP) & "' ORDER BY NOINV")
6         End If
7         If Not rsTemp.EOF Then
8             Call BuatTabelInv(dNamaTabel)
9             Do While Not rsTemp.EOF
10                subletOnly = True
11                For i = 1 To 8
12                    hlp = "REP" & i
13                    If rsTemp(hlp) <> "" Then
14                        subletOnly = False
15                        hlp2 = "KLPKERJA" & i
16                        hlp3 = "RTE" & i    '"ACTRATE" & i
17                        hlp4 = "MEKANIK" & i
18                        hlp5 = "OFF" & i
19                        hlp6 = "DISC" & i
20                        hlp7 = "FREEC" & i
21                        If createTemp(dNamaTabel) Then
22                            nDuoInv = False
23                            If rsTemp("NDUOINV") = 1 Then
24                                nDuoInv = True
25                            End If
26                            If nDuoInv Then
27                                If rsTemp(hlp7) = 0 Then
28                                    nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
29                                    nNilai = Round(CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And Trim(rsTemp(hlp2)) = "SBE") Or (Trim(rsTemp(hlp2) = "SBI")), 0, CDbl(rsTemp(hlp6)) / 100)))
30                                    noInv = Trim(rsTemp("NOINV")) & "-B"
31                                    nmCust = Trim(rsTemp("NMPENJUAL"))
32                                    kdCust = objBilling.checkKdDealer(rsTemp("KDPENJUAL"))
33                                Else
34                                    nStatus = rsTemp("STATUS")
35                                    nNilai = Round(CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100)))
36                                    noInv = Trim(rsTemp("NOINV")) & "-A"
37                                    nmCust = Trim(rsTemp("CUSTOMER"))
38                                    kdCust = Trim(rsTemp("FS_KD_PLG"))
39                                End If
40                            Else
41                                If rsTemp("FS_RC") = "R001" Then
42                                    nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
43                                    nNilai = Round(CDbl(rsTemp(hlp5)) * (1 - IIf((nDuoInv And Trim(rsTemp(hlp2)) = "SBE") Or (Trim(rsTemp(hlp2) = "SBI")), 0, CDbl(rsTemp(hlp6)) / 100)))
44                                    noInv = Trim(rsTemp("NOINV"))
45                                    nmCust = Trim(rsTemp("NMPENJUAL"))
46                                    kdCust = objBilling.checkKdDealer(rsTemp("KDPENJUAL"))
47                                Else
48                                    nStatus = rsTemp("STATUS")
49                                    nNilai = Round(CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100)))
50                                    noInv = Trim(rsTemp("NOINV"))
51                                    nmCust = Trim(rsTemp("CUSTOMER"))
52                                    kdCust = Trim(rsTemp("FS_KD_PLG"))
53                                End If
54                            End If
                              
55                            If nStatus = "AFFNONPPN" Then
56                                nNilai = Round(nNilai * (100 / 110))
57                            End If
                              
                               
58                            Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & dNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
59                            If Not rsHlp.EOF Then
60                                Call objO.UpdateData("UPDATE [TEMPDB]..[" & dNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                               & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "BOR = '" & cekDuo(CDbl(rsTemp("BODY")), noInv) & "', MATERIAL = '" & cekDuo(CDbl(rsTemp("MTRL")), noInv) & "', PART = '" & cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) & "', " _
                                               & "OLI = '" & cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) & "', SUBLET = '" & cekDuo(CDbl(rsTemp("SBLT")), noInv) & "', SUBTOT = SUBTOT + '" & CDbl(IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0)) & "', NOFP = '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', " _
                                               & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', HPP = '" & cekDuo(CDbl(rsTemp("HPP")), noInv) & "', TOTAL = TOTAL + '" & CDbl(IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0)) & "', " _
                                               & "WAR = '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "' " _
                                               & ",KDPLG = '" & Trim(kdCust) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
61                            Else
62                                Call objO.InsertData("INSERT INTO [TEMPDB]..[" & dNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                               & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "','" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "'" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "'" & setKoma(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', '" & setKoma(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', " _
                                               & "'" & setKoma(CDbl(IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) + IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "LUB", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "PDS", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "WAS", CDbl(nNilai), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                               & " '" & setKoma(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & "', '" & setKoma(CDbl(IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "SBE", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) + IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "LUB", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "PDS", CDbl(nNilai), 0) + IIf(rsTemp(hlp2) = "WAS", CDbl(nNilai), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                               & "'" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & setKoma(CDbl(rsTemp("LABOR"))) & "', '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', '" & setKoma(cekDuo(CDbl(rsTemp("HPP")), noInv)) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "','" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "','" & Trim(rsTemp("NOMB")) & "','" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "','" & Trim(kdCust) & "','" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "')")
63                            End If
                              
64                            If nDuoInv Then
65                                nStatus = rsTemp("STATUS")
66                                nNilai = Round(CDbl(rsTemp(hlp5)) * (1 - (CDbl(rsTemp(hlp6)) / 100)))
67                                noInv = Trim(rsTemp("NOINV")) & "-A"
68                                nmCust = Trim(rsTemp("CUSTOMER"))
69                                kdCust = Trim(rsTemp("FS_KD_PLG"))
70                                Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & dNamaTabel & "] WHERE NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsTemp("NO_WO")) & "' AND NOINV = '" & Trim(noInv) & "'")
71                                If Not rsHlp.EOF Then
72                                    Call objO.UpdateData("UPDATE [TEMPDB]..[" & dNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                                   & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) & "', " _
                                                   & "SBE = SBE + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                                   & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                                   & "PDI = PDI + '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', " _
                                                   & "BOR = '" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', MATERIAL = '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', PART = '" & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', " _
                                                   & "OLI = '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', SUBLET = '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', SUBTOT = SUBTOT + '" & CDbl(CDbl(IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0))) & "', NOFP = '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', " _
                                                   & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', HPP = '" & CDbl(rsTemp("HPP")) & "', TOTAL = TOTAL + '" & CDbl(CDbl(IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0))) & "', WAR = '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "' " _
                                                   & ",KDPLG = '" & Trim(kdCust) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
73                                Else
74                                    Call objO.InsertData("INSERT INTO [TEMPDB]..[" & dNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                                   & "'0','" & Replace(CDbl(rsTemp(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsTemp(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                                   & "'" & IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "','" & IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                                   & "'" & IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', '" & IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) & "', " _
                                                   & "'" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', " & _
                                                   " '" & CDbl(CDbl(IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                                   & " '" & RepRp(cekDuo(CDbl(rsTemp("MET")), noInv, 2)) & "', '" & CDbl(CDbl(IIf(rsTemp(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                                   & " '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & Trim(rsTemp("NOFAKTUR")) & "', '" & CDbl(rsTemp("HPP")) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', '" & Trim(rsTemp("NOMB")) & "', '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "','" & Trim(kdCust) & "','" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "')")
75                                End If
76                            End If
77                        End If
78                    End If
79                Next
80                If subletOnly Then
81                    nDuoInv = False
82                    If rsTemp("NDUOINV") = 1 Then nDuoInv = True
83                    If nDuoInv Then
84                        noInv = Trim(rsTemp("NOINV")) & "-A"
85                    Else
86                        noInv = Trim(rsTemp("NOINV"))
87                    End If
                      
88                    kdCust = Trim(rsTemp("FS_KD_PLG"))
89                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & dNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
90                    If Not rsHlp.EOF Then
91                        Call objO.UpdateData("UPDATE [TEMPDB]..[" & dNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                       & "JOR = '0', TIPEINV = '" & Trim(rsTemp("STATUS")) & "', CUSTOMER = '" & Trim(rsTemp("CUSTOMER")) & "', BOR = '" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', MATERIAL = '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', " _
                                       & "PART = '" & RepRp(cekDuo(CDbl(rsTemp("PART_G") + rsTemp("PART_NG")), noInv)) & "', OLI = '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G") + rsTemp("OIL_NG")), noInv)) & "', SUBLET = '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', NOFP = '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', " _
                                       & "LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', HPP = '" & RepRp(CDbl(rsTemp("HPP"))) & "', WAR = '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "' " _
                                       & ",KDPLG = '" & Trim(kdCust) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
92                    Else
93                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & dNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                       & "'0','0', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                       & "'" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', '" & RepRp(CDbl(cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                       & " '" & RepRp(cekDuo(CDbl(rsTemp("MET")), noInv, 2)) & "', '" & RepRp(CDbl(cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv))) & "', " _
                                       & " '" & Trim(rsTemp("STATUS")) & "', '" & Trim(rsTemp("CUSTOMER")) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & Trim(rsTemp("NOFAKTUR")) & "','" & RepRp(CDbl(rsTemp("HPP"))) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', '" & Trim(rsTemp("NOMB")) & "', '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "','" & Trim(kdCust) & "','" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "')")
94                    End If
95                End If
                  
                  ' DARI ADD JOB
96                Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WITH(NOLOCK) WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND FS_LOG_DEL = '" & Trim(rsTemp("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsTemp("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
97                If Not rsAdd.EOF Then
98                    Do While Not rsAdd.EOF
99                        hlp2 = "KLPKERJA"
100                       hlp3 = "RTE"
101                       hlp4 = "MEKANIK"
102                       hlp5 = "OFFN"
103                       hlp6 = "FREEC"
                          
104                       nDuoInv = False
105                       If rsTemp("NDUOINV") = 1 Then
106                           nDuoInv = True
107                       End If
                          
108                       If nDuoInv Then
109                           If rsAdd(hlp6) = 0 Then
110                               nStatus = statusDealer(objBilling.checkDealerStatus(rsTemp("KDPENJUAL")))
111                               nNilai = Round(CDbl(rsAdd(hlp5)) * (1 - IIf((nDuoInv And rsAdd(hlp2) = "SBE"), 0, CDbl(rsAdd("DISC")) / 100)))
112                               noInv = Trim(rsTemp("NOINV")) & "-B"
113                               nmCust = Trim(rsTemp("NMPENJUAL"))
114                               kdCust = objBilling.checkKdDealer(rsTemp("KDPENJUAL"))
115                           Else
116                               nStatus = rsTemp("STATUS")
117                               nNilai = Round(CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100)))
118                               noInv = Trim(rsTemp("NOINV")) & "-A"
119                               nmCust = Trim(rsTemp("CUSTOMER"))
120                               kdCust = Trim(rsTemp("FS_KD_PLG"))
121                           End If
122                       Else
123                           nStatus = rsTemp("STATUS")
124                           nNilai = Round(CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100)))
125                           noInv = Trim(rsTemp("NOINV"))
126                           nmCust = Trim(rsTemp("CUSTOMER"))
127                           kdCust = Trim(rsTemp("FS_KD_PLG"))
128                       End If
                          
129                       If Trim(nStatus) = "AFFNONPPN" Then
130                           nNilai = Round(nNilai * (100 / 110))
131                       End If
                          
132                       Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & dNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
133                       If Not rsHlp.EOF Then
134                           Call objO.UpdateData("UPDATE [TEMPDB].." & dNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                           & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                           & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                           & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                           & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', " _
                                           & "BOR = '" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', MATERIAL = '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', PART = '" & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', " _
                                           & "OLI = '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', SUBLET = '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', SUBTOT = SUBTOT + '" & CDbl(IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0)) & "', HPP = '" & cekDuo(CDbl(rsTemp("HPP")), noInv) & "', " _
                                           & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "', NOFP = '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', " _
                                           & "TOTAL = TOTAL + '" & CDbl(IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0)) & "', WAR = '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "'" _
                                           & ",KDPLG = '" & Trim(kdCust) & "' WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
135                       Else
136                           Call objO.InsertData("INSERT INTO [TEMPDB]..[" & dNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                           & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "SBE", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                           & "'" & IIf(rsAdd(hlp2) = "LUB", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS", CDbl(nNilai), 0) & "', " _
                                           & "'" & RepRp(cekDuo(CDbl(rsTemp("BODY")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("MTRL")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv)) & "', '" & RepRp(cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv)) & "', " & _
                                           " '" & RepRp(cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', '" & CDbl(IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', " _
                                           & " '" & RepRp(cekDuo(CDbl(rsTemp("MET")), noInv, 1)) & "', '" & CDbl(IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', " _
                                           & "'" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsTemp("NOFAKTUR"))) & "','" & cekDuo(CDbl(rsTemp("HPP")), noInv) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "','" & Trim(rsTemp("NOMB")) & "','" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "','" & Trim(kdCust) & "','" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "')")
137                       End If
138                       If nDuoInv Then
139                           nStatus = rsTemp("STATUS")
140                           nNilai = Round(CDbl(rsAdd(hlp5)) * (1 - (CDbl(rsAdd("DISC")) / 100)))
141                           noInv = Trim(rsTemp("NOINV")) & "-A"
142                           nmCust = Trim(rsTemp("CUSTOMER"))
143                           kdCust = Trim(rsTemp("FS_KD_PLG"))
144                           Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & dNamaTabel & "] WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
145                           If Not rsHlp.EOF Then
146                               Call objO.UpdateData("UPDATE [TEMPDB]..[" & dNamaTabel & "] SET SA = '" & rsTemp("FRONT") & "', NOINV = '" & Trim(noInv) & "', NO_POLISI = '" & rsTemp("NO_POLISI") & "', " _
                                               & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                               & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "TWC = TWC + '" & IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) & "', " _
                                               & "BOR = '" & RepRp(CDbl(rsTemp("BODY"))) & "', MATERIAL = '" & RepRp(CDbl(rsTemp("MTRL"))) & "', PART = '" & RepRp(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))) & "', " _
                                               & "OLI = '" & RepRp(CDbl(rsTemp("OIL_G"))) + RepRp(CDbl(rsTemp("OIL_NG"))) & "', SUBLET = '" & RepRp(CDbl(rsTemp("SBLT"))) & "', SUBTOT =  SUBTOT + '" & CDbl(IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0)) & "', HPP = '" & RepRp(CDbl(rsTemp("HPP"))) & "', " _
                                               & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsTemp("NOFAKTUR")) & "', NOFP = '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', TOTAL =  TOTAL + '" & CDbl(IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0)) & "', " _
                                               & " WAR = '" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "', KDPLG = '" & Trim(kdCust) & "' " _
                                               & "WHERE NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "' AND NO_WO = '" & rsTemp("NO_WO") & "' AND NOINV = '" & Trim(noInv) & "'")
147                           Else
148                               Call objO.InsertData("INSERT INTO [TEMPDB]..[" & dNamaTabel & "] VALUES ('" & rsTemp("FRONT") & "', '" & Trim(noInv) & "','" & rsTemp("NO_WO") & "', '" & rsTemp("NO_POLISI") & "','" & rsTemp("NO_CHASSIS") & "', " _
                                               & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsTemp("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsTemp("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI", CDbl(nNilai), 0) & "', " _
                                               & "'" & IIf(rsAdd(hlp2) = "SBE", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsTemp("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                               & "'" & IIf(rsAdd(hlp2) = "LUB", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS", CDbl(nNilai), 0) & "', " _
                                               & "'" & RepRp(CDbl(rsTemp("BODY"))) & "', '" & RepRp(CDbl(rsTemp("MTRL"))) & "', '" & RepRp(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG"))) & "', '" & RepRp(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG"))) & "', '" & RepRp(CDbl(rsTemp("SBLT"))) & "', '" & CDbl(IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', " _
                                               & " '" & RepRp(cekDuo(CDbl(rsTemp("MET")), noInv, 2)) & "', '" & CDbl(IIf(rsAdd(hlp2) = "SBI", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "SBE", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "GRP" And rsTemp("TPSVC") <> "W", setKoma(CDbl(nNilai)), 0) + IIf(rsTemp("TPSVC") = "W", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "LUB", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "PDS", setKoma(CDbl(nNilai)), 0) + IIf(rsAdd(hlp2) = "WAS", setKoma(CDbl(nNilai)), 0) + cekDuo(CDbl(rsTemp("BODY")), noInv) + cekDuo(CDbl(rsTemp("MTRL")), noInv) + cekDuo(CDbl(rsTemp("PART_G")) + CDbl(rsTemp("PART_NG")), noInv) + cekDuo(CDbl(rsTemp("OIL_G")) + CDbl(rsTemp("OIL_NG")), noInv) + cekDuo(CDbl(rsTemp("SBLT")), noInv)) & "', " _
                                               & "'" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsTemp("LABOR")), ",", ".") & "', '" & Trim(rsTemp("NOFAKTUR")) & "','" & RepRp(CDbl(rsTemp("HPP"))) & "','" & Format(rsTemp("INVDATE"), "MM/DD/YYYY") & "', '" & IIf(Right(noInv, 2) = "-B", Trim(rsTemp("NOFP2")), Trim(rsTemp("NOFP"))) & "', NOMB = '" & Trim(rsTemp("NOMB")) & "', " _
                                               & "'" & IIf(Trim(rsTemp("TPSVC")) <> "W", 0, 1) & "', '" & Trim(kdCust) & "','" & IIf(Trim(rsTemp("DELIV")) = "BTL", 1, 0) & "')")
149                           End If
150                       End If
151                       rsAdd.MoveNext
152                   Loop
153               End If
154               If rsAdd.State <> adStateClosed Then rsAdd.Close
155               rsTemp.MoveNext
156           Loop
157           Call tampilGrid(idx, nFP)
158       End If

159      On Error GoTo 0
160      Exit Sub

setCetakLampiran_Error:

161       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure setCetakLampiran of Form frmFakturPajak")
End Sub

Private Sub createRsInv(idxx As Integer)
1        On Error GoTo createRsInv_Error
2         If rsInvoice.State Then rsInvoice.Close
3         Set rsInvoice = Nothing
4         With rsInvoice
              'If .State Then .Close
5             .Fields.Append "ST", adBoolean
6             .Fields.Append "NFP", adVarChar, 30
7             .Fields.Append "TGLFP", adVarChar, 20
8             .Fields.Append "INVO", adVarChar, 15
9             .Fields.Append "TGLINVO", adDate
10            .Fields.Append "NOPOL", adVarChar, 15
11            .Fields.Append "WO", adVarChar, 15
12            .Fields.Append "RGK", adVarChar, 25
13            .Fields.Append "CUST", adVarChar, 300
14            .Fields.Append "TTL", adDouble
15            .Fields.Append "W", adVarChar, 1
16            .Fields.Append "PLG", adVarChar, 50
17            .Fields.Append "R", adVarChar, 1
18            .Open
19        End With
20        Set GridProgram.DataSource = rsInvoice
21        Call headGrid(idxx)

22       On Error GoTo 0
23       Exit Sub

createRsInv_Error:

24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRsInv of Form frmFakturPajak")
End Sub

Private Sub tampilGrid(idx As Integer, Optional nnFP As String)
      Dim rsSearch As New ADODB.Recordset
          
1        On Error GoTo tampilGrid_Error

2         nTagih = 0
3         totTagih = 0
4         Select Case idx
              Case 1
5                 Set rsSearch = objO.FetchDataACC("SELECT (CASE ISNULL(C.NOFPS,'') WHEN '' THEN ISNULL(NOFP,'') ELSE ISNULL(C.NOFPS,'') END) AS NO, TGLFPS AS TGLF, '" & gKodeFaktur & "'+NOINV AS INV, INVDATE, NO_POLISI, NO_WO, NO_CHASSIS, NAMAFP AS CUSTOMER, TOTAL, WAR, KDPLG, DIRETUR " & _
                                                "FROM [TEMPDB]..[" & vNamaTabel & "] A LEFT JOIN [FAKTUR" & Trim(Right(gDb, 2)) & "]..INVOICE C ON 'E-'+A.NOINV = C.NOINVOICE LEFT JOIN [FAKTUR" & Trim(Right(gDb, 2)) & "]..SIFPSTANDAR B ON B.NOFPS = C.NOFPS " & _
                                                "WHERE NOFP LIKE '%" & Trim(nnFP) & "%' AND NOFP <> '' AND TIPEINV <> 'AFFNONPPN' AND TOTAL > 0 ORDER BY NOINV") '
              
6             Case 2
7                 Set rsSearch = objO.FetchDataACC("SELECT (CASE ISNULL(C.NOFPS,'') WHEN '' THEN ISNULL(NOFP,'') ELSE ISNULL(C.NOFPS,'') END) AS NO, ISNULL((SELECT ISNULL(TGLFPS,'') FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..SIFPSTANDAR WITH(NOLOCK) WHERE NOFPS = CASE ISNULL(B.NOFPS,'') WHEN '' THEN A.NOFP ELSE B.NOFPS END),'') AS TGLF, '" & gKodeFaktur & "'+NOINV AS INV, INVDATE, " & _
                                                "NO_POLISI, NO_WO, NO_CHASSIS, CASE ISNULL(NAMAFP,'') WHEN '' THEN CUSTOMER ELSE ISNULL(NAMAFP,'') END AS CUSTOMER, TOTAL, WAR, KDPLG, DIRETUR " & _
                                                "FROM [TEMPDB]..[" & vNamaTabel & "] A LEFT JOIN [FAKTUR" & Trim(Right(gDb, 2)) & "]..INVOICE C ON 'E-'+A.NOINV = C.NOINVOICE  LEFT JOIN [FAKTUR" & Trim(Right(gDb, 2)) & "]..FPGABUNGAN B ON B.NOINVOICE = C.NOINVOICE  " & _
                                                "WHERE A.NOFP LIKE '%" & Trim(nnFP) & "%' AND (A.NOFP IN (SELECT NOFPS FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..FPGABUNGAN WHERE MODUL = 'SVC') " & _
                                                "OR '" & gKodeFaktur & "'+A.NOINV NOT IN (SELECT NOINVOICE FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..INVOICE WHERE NOFPS NOT IN (SELECT NOFPS FROM [FAKTUR" & Trim(Right(gDb, 2)) & "]..FPGABUNGAN WHERE MODUL = 'SVC')))" & _
                                                "AND TIPEINV <> 'AFFNONPPN' AND TOTAL > 0 ORDER BY NOINV") '
8             Case 3
9                 Set rsSearch = objO.FetchDataACC("SELECT NOMB AS NO, TGLMBS AS TGLF, '" & gKodeFaktur & "'+NOINV AS INV, INVDATE, NO_POLISI, NO_WO, NO_CHASSIS, CUSTOMER, TOTAL, WAR " & _
                                                "FROM [TEMPDB]..[" & vNamaTabel & "] A LEFT JOIN [FAKTUR" & Trim(Right(gDb, 2)) & "]..SIFPSTANDAR B ON A.NOFP = B.NOFPS " & _
                                                "WHERE NOMB LIKE '%" & Trim(nnFP) & "%' AND TIPEINV <> 'AFFNONPPN' AND TOTAL > 0 ORDER BY NOINV") '
10        End Select
11        If rsInvoice.State <> adStateClosed Then rsInvoice.Close
12        If rsInvoice.State = adStateClosed Then Call createRsInv(idx)
13        If Not rsSearch.EOF Then
14            Do While Not rsSearch.EOF
15                With rsInvoice
16                    .AddNew
17                    !ST = False
18                    !nFP = Trim(rsSearch("NO"))
19                    !TGLFP = IIf(Nul(Trim(rsSearch("TGLF"))) = "01/01/1900", "", Nul(Trim(rsSearch("TGLF"))))
20                    !INVO = Trim(rsSearch("INV"))
21                    !TGLINVO = Trim(rsSearch("INVDATE"))
22                    !noPol = Trim(rsSearch("NO_POLISI"))
23                    !wo = Trim(rsSearch("NO_WO"))
24                    !RGK = Trim(rsSearch("NO_CHASSIS"))
25                    !cust = Nul(Trim(rsSearch("CUSTOMER")))
26                    !TTL = CDbl(rsSearch("TOTAL"))
27                    !W = Trim(rsSearch("WAR"))
28                    !PLG = Trim(rsSearch("KDPLG"))
29                    !R = Trim(rsSearch("DIRETUR"))
30                    .Update
31                End With
32                rsSearch.MoveNext
33            Loop
34            rsInvoice.MoveFirst
              
35        End If
36        GridProgram.DataSource = rsInvoice
37        Call headGrid(idx)
38        GridProgram.FetchRowStyle = True
39        lblJumlah = "Jumlah Invoice = " & rsInvoice.RecordCount

40       On Error GoTo 0
41       Exit Sub

tampilGrid_Error:

42        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tampilGrid of Form frmFakturPajak")
End Sub

Private Sub headGrid(idxx As Integer)

1        On Error GoTo headGrid_Error

2         With GridProgram
3             Select Case idxx
                  Case 1
4                     .Columns(0).Width = 400
5                     .Columns(0).Caption = "Y"
6                     .Columns(0).Locked = False
7                     .Columns(0).Alignment = dbgCenter
8                     .Columns(1).Width = 2000
9                     .Columns(1).Caption = "No Faktur"
10                    .Columns(1).Locked = True
11                    .Columns(2).Width = 1200
12                    .Columns(2).Caption = "Tgl Faktur"
13                    .Columns(2).Locked = True
14                    .Columns(3).Width = 1300
15                    .Columns(3).Caption = "No Invoice"
16                    .Columns(3).Locked = True
17                    .Columns(4).Width = 1200
18                    .Columns(4).Caption = "Tgl Inv"
19                    .Columns(4).Locked = True
20                    .Columns(5).Width = 1200
21                    .Columns(5).Caption = "No Polisi"
22                    .Columns(5).Locked = True
23                    .Columns(6).Width = 900
24                    .Columns(6).Caption = "No WO"
25                    .Columns(6).Locked = True
26                    .Columns(6).Visible = False
27                    .Columns(7).Width = 2000
28                    .Columns(7).Caption = "No Rangka"
29                    .Columns(7).Locked = True
30                    .Columns(7).Visible = False
31                    .Columns(8).Width = 4000
32                    .Columns(8).Caption = "Customer"
33                    .Columns(8).Locked = True
34                    .Columns(9).Width = 1500
35                    .Columns(9).Caption = "Nominal"
36                    .Columns(9).Alignment = dbgRight
37                    .Columns(9).Locked = True
38                    .Columns(10).Width = 0
39                    .Columns(10).Caption = "W"
40                    .Columns(10).Visible = False
41                    .Columns(10).Locked = True
42                    .Columns(11).Width = 1500
43                    .Columns(11).Caption = "Kode"
44                    .Columns(11).Locked = True
45                    .Columns(12).Width = 0
46                    .Columns(12).Caption = "R"
47                    .Columns(12).Visible = False
48                    .Columns(12).Locked = True
49                Case 2
50                    .Columns(0).Width = 400
51                    .Columns(0).Caption = "Y"
52                    .Columns(0).Locked = False
53                    .Columns(0).Alignment = dbgCenter
54                    .Columns(1).Width = 2000
55                    .Columns(1).Caption = "No Faktur"
56                    .Columns(1).Locked = True
57                    .Columns(2).Width = 1200
58                    .Columns(2).Caption = "Tgl Faktur"
59                    .Columns(2).Locked = True
60                    .Columns(3).Width = 1300
61                    .Columns(3).Caption = "No Invoice"
62                    .Columns(3).Locked = True
63                    .Columns(4).Width = 1200
64                    .Columns(4).Caption = "Tgl Inv"
65                    .Columns(4).Locked = True
66                    .Columns(5).Width = 1200
67                    .Columns(5).Caption = "No Polisi"
68                    .Columns(5).Locked = True
69                    .Columns(5).Visible = True
70                    .Columns(6).Width = 900
71                    .Columns(6).Caption = "No WO"
72                    .Columns(6).Locked = True
73                    .Columns(6).Visible = False
74                    .Columns(7).Width = 2000
75                    .Columns(7).Caption = "No Rangka"
76                    .Columns(7).Locked = True
77                    .Columns(7).Visible = False
78                    .Columns(8).Width = 4000
79                    .Columns(8).Caption = "Customer"
80                    .Columns(8).Locked = True
81                    .Columns(9).Width = 1500
82                    .Columns(9).Caption = "Nominal"
83                    .Columns(9).Alignment = dbgRight
84                    .Columns(9).Locked = True
85                    .Columns(10).Width = 0
86                    .Columns(10).Caption = "W"
87                    .Columns(10).Visible = False
88                    .Columns(10).Locked = True
89                    .Columns(11).Width = 1500
90                    .Columns(11).Caption = "Kode"
91                    .Columns(11).Locked = True
92                    .Columns(12).Width = 0
93                    .Columns(12).Caption = "R"
94                    .Columns(12).Visible = False
95                    .Columns(12).Locked = True
96                Case 3
97                    .Columns(0).Width = 400
98                    .Columns(0).Caption = "Y"
99                    .Columns(0).Locked = False
100                   .Columns(0).Alignment = dbgCenter
101                   .Columns(1).Width = 2000
102                   .Columns(1).Caption = "No Retur"
103                   .Columns(1).Locked = True
104                   .Columns(2).Width = 1000
105                   .Columns(2).Caption = "Tgl Retur"
106                   .Columns(2).Locked = True
107                   .Columns(3).Width = 1300
108                   .Columns(3).Caption = "No Invoice"
109                   .Columns(3).Locked = True
110                   .Columns(4).Width = 1000
111                   .Columns(4).Caption = "Tgl Inv"
112                   .Columns(4).Locked = True
113                   .Columns(5).Width = 1000
114                   .Columns(5).Caption = "No Polisi"
115                   .Columns(5).Locked = True
116                   .Columns(5).Visible = True
117                   .Columns(6).Width = 900
118                   .Columns(6).Caption = "No WO"
119                   .Columns(6).Locked = True
120                   .Columns(6).Visible = False
121                   .Columns(7).Width = 2000
122                   .Columns(7).Caption = "No Rangka"
123                   .Columns(7).Locked = True
124                   .Columns(7).Visible = False
125                   .Columns(8).Width = 3000
126                   .Columns(8).Caption = "Customer"
127                   .Columns(8).Locked = True
128                   .Columns(9).Width = 1500
129                   .Columns(9).Caption = "Nominal"
130                   .Columns(9).Alignment = dbgRight
131                   .Columns(9).Locked = True
132                   .Columns(10).Width = 0
133                   .Columns(10).Caption = "W"
134                   .Columns(10).Visible = False
135                   .Columns(10).Locked = True
                  
136           End Select
137       End With

138      On Error GoTo 0
139      Exit Sub

headGrid_Error:

140       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headGrid of Form frmFakturPajak")
End Sub

Private Sub btnBatal_Click(Index As Integer)
    frmFaktur.Visible = False
End Sub

Private Sub cboFaktur_Change()
1        On Error GoTo cboFaktur_Change_Error

2         cboNoFP = ""
3         If cboFaktur = "Faktur Pajak" Then
4             lblFaktur.Caption = "Look Up No. Faktur Pajak :"
5             Call NoFaktur(1)
6         ElseIf cboFaktur = "Faktur Pajak Gabungan" Then
7             lblFaktur.Caption = "Look Up No. Faktur Pajak Gabungan :"
8             Call NoFaktur(2)
9         ElseIf cboFaktur = "Faktur Pajak Retur" Then
10            lblFaktur.Caption = "Look Up No. Retur Faktur :"
11            Call NoFaktur(3)
12        End If

13       On Error GoTo 0
14       Exit Sub

cboFaktur_Change_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cboFaktur_Change of Form frmFakturPajak")
End Sub

Private Sub setNoFaktur(pilType As String)
      Dim rsSearch As ADODB.Recordset
1        On Error GoTo setNoFaktur_Error

2         Set rsSearch = objO.FetchDataACC("SELECT * FROM [FAKTUR" & Right(gDb, 2) & "]..DEALER WITH(NOLOCK)")
3         If Not rsSearch.EOF Then
              'kdDealer = Trim(rsSearch("KODEDEALER"))
4             namaPT = Trim(rsSearch("NAMA"))
5             alamatPT = Trim(rsSearch("alamat"))
6             npwpPT = Trim(rsSearch("npwp"))
7             tglKukuhPT = Trim(rsSearch("tglPENGUKUHAN"))
8             gKotaFaktur = Trim(rsSearch("KOTATTD"))
9         End If
          
10        kdType = ""
11        If pilType <> "" Then
12            Set rsSearch = objO.FetchDataACC("SELECT * FROM [FAKTUR" & Right(gDb, 2) & "]..TYPECUST WITH(NOLOCK) WHERE NAMA = '" & Trim(pilType) & "'")
13            If Not rsSearch.EOF Then
14                kdType = Trim(rsSearch("KODETYPE"))
15            End If
16        End If
17        If OnNewFP = 0 Then
18            text(4) = Trim(kdType) & "." & Trim(kdDealer) & "." & IIf(Trim(thnInv) <> "", Trim(thnInv), Right(year(isnow), 2)) & "."
19        Else
20            text(4) = Trim(kdType) & "." & Trim(kdDealer) & "-" & IIf(Trim(thnInv) <> "", Trim(thnInv), Right(year(isnow), 2)) & "."
21        End If

22       On Error GoTo 0
23       Exit Sub

setNoFaktur_Error:

24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure setNoFaktur of Form frmFakturPajak")
End Sub

Private Sub cboFaktur_Click()
1        On Error GoTo cboFaktur_Click_Error

2         cboNoFP = ""
3         If cboFaktur = "Faktur Pajak" Then
4             lblFaktur.Caption = "Look Up No. Faktur Pajak :"
5             Call NoFaktur(1)
6         ElseIf cboFaktur = "Faktur Pajak Gabungan" Then
7             lblFaktur.Caption = "Look Up No. Faktur Pajak Gabungan :"
8             Call NoFaktur(2)
9         ElseIf cboFaktur = "Faktur Pajak Retur" Then
10            lblFaktur.Caption = "Look Up No. Retur Faktur :"
11            Call NoFaktur(3)
12        End If

13       On Error GoTo 0
14       Exit Sub

cboFaktur_Click_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cboFaktur_Click of Form frmFakturPajak")
End Sub

Private Sub cboFaktur_DropDown()
1        On Error GoTo cboFaktur_DropDown_Error

2         cboNoFP = ""
3         If cboFaktur = "Faktur Pajak" Then
4             lblFaktur.Caption = "Look Up No. Faktur Pajak :"
5             Call NoFaktur(1)
6         ElseIf cboFaktur = "Faktur Pajak Gabungan" Then
7             lblFaktur.Caption = "Look Up No. Faktur Pajak Gabungan :"
8             Call NoFaktur(2)
9         ElseIf cboFaktur = "Faktur Pajak Retur" Then
10            lblFaktur.Caption = "Look Up No. Retur Faktur :"
11            Call NoFaktur(3)
12        End If

13       On Error GoTo 0
14       Exit Sub

cboFaktur_DropDown_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cboFaktur_DropDown of Form frmFakturPajak")
End Sub

Private Sub cboNoFP_Change()
    If Trim(cboNoFP) <> "" Then
        Call refreshList(Trim(cboNoFP))
    End If
End Sub

Private Sub cboNoFP_Click()
    If Trim(cboNoFP) <> "" Then
        Call refreshList(Trim(cboNoFP))
    End If
End Sub

Private Sub cboNoFP_DropDown()
    If Trim(cboNoFP) <> "" Then
        Call refreshList(Trim(cboNoFP))
    End If
End Sub

Private Sub cboTtd_Change()
      Dim rsSearch As ADODB.Recordset
          ' USER OTORITAS
1        On Error GoTo cboTtd_Change_Error

2         Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..T_PROFIL WITH(NOLOCK)")
3         If Not rsSearch.EOF Then
4             Select Case cboTtd
                  Case "KACAB"
5                     text(6) = Trim(rsSearch("FS_KA_CABANG"))
6                     text(8) = "Kepala Cabang"
7                 Case "ADH"
8                     text(6) = Trim(rsSearch("FS_KA_ADH"))
9                     text(8) = "ADH"
10                Case "KABENG"
11                    text(6) = Trim(rsSearch("FS_KA_BENGKEL"))
12                    text(8) = "Kepala Bengkel"
13            End Select
14        End If

15       On Error GoTo 0
16       Exit Sub

cboTtd_Change_Error:

17        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cboTtd_Change of Form frmFakturPajak")
End Sub

Private Sub cboTtd_Click()
      Dim rsSearch As ADODB.Recordset
          ' USER OTORITAS
1        On Error GoTo cboTtd_Click_Error

2         Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..T_PROFIL WITH(NOLOCK)")
3         If Not rsSearch.EOF Then
4             Select Case cboTtd
                  Case "KACAB"
5                     text(6) = Trim(rsSearch("FS_KA_CABANG"))
6                     text(8) = "Kepala Cabang"
7                 Case "ADH"
8                     text(6) = Trim(rsSearch("FS_KA_ADH"))
9                     text(8) = "ADH"
10                Case "KABENG"
11                    text(6) = Trim(rsSearch("FS_KA_BENGKEL"))
12                    text(8) = "Kepala Bengkel"
13            End Select
14        End If

15       On Error GoTo 0
16       Exit Sub

cboTtd_Click_Error:

17        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cboTtd_Click of Form frmFakturPajak")
           
End Sub

Private Sub cboTtd_DropDown()
      Dim rsSearch As ADODB.Recordset
          ' USER OTORITAS
1        On Error GoTo cboTtd_DropDown_Error

2         Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..T_PROFIL WITH(NOLOCK)")
3         If Not rsSearch.EOF Then
4             Select Case cboTtd
                  Case "KACAB"
5                     text(6) = Trim(rsSearch("FS_KA_CABANG"))
6                     text(8) = "Kepala Cabang"
7                 Case "ADH"
8                     text(6) = Trim(rsSearch("FS_KA_ADH"))
9                     text(8) = "ADH"
10                Case "KABENG"
11                    text(6) = Trim(rsSearch("FS_KA_BENGKEL"))
12                    text(8) = "Kepala Bengkel"
13            End Select
14        End If

15       On Error GoTo 0
16       Exit Sub

cboTtd_DropDown_Error:

17        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cboTtd_DropDown of Form frmFakturPajak")
           
End Sub

Private Sub cboType_Change()
    Call setNoFaktur(cboType.text)
End Sub

Private Sub cboType_Click()
    Call setNoFaktur(cboType.text)
End Sub

Private Sub cboType_DropDown()
    Call setNoFaktur(cboType.text)
End Sub

Private Sub bersih()
    For i = 0 To 8
        Select Case i
            Case Else
                text(i) = ""
        End Select
    Next
    tglFaktur = ""
    chInPPn = Unchecked
    chSilang = Checked
    chHarga = Checked
    chGanti = Unchecked
    noFakturCetak = ""
End Sub

Private Sub chSilang_Click()
    If chSilang = Unchecked Then
        chHarga = Checked
        chGanti = Unchecked
    ElseIf chSilang = Checked Then
        chHarga = Unchecked
        chGanti = Unchecked
    End If
End Sub

Private Sub createTempTable(vNamaTabel As String)
1        On Error GoTo createTempTable_Error

2         sSQL = "if exists (select * from sysobjects where id = object_id(N'[dbo].[" & vNamaTabel & "]') and OBJECTPROPERTY(id, N'IsUserTable') = 1) " _
               & " drop table [tempdb]..[" & vNamaTabel & "]"
3         oConn.Execute sSQL
4         sSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
               & "    [NOFP][varchar](30) NOT NULL DEFAULT '', " _
               & "    [NAMAFP][varchar](100) NOT NULL DEFAULT '', " _
               & "    [TGLFP][datetime] NOT NULL DEFAULT '', " _
               & "    [NO_INV][varchar](15) NOT NULL DEFAULT '', " _
               & "    [NOMINAL][decimal](15) NOT NULL DEFAULT '', " _
               & "    [R][char](1) NOT NULL DEFAULT '0') ON [PRIMARY]"
5         oConn.Execute sSQL

6        On Error GoTo 0
7        Exit Sub

createTempTable_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createTempTable of Form frmFakturPajak")
End Sub

Private Sub cmd1_Click()
 text(2).text = "00.000.000.0-000.000"
End Sub

Private Sub cmdCetakRekap_Click()
      Dim sHlp2 As Object
      Dim rsFaktur As New ADODB.Recordset
      Dim rsSearch As ADODB.Recordset
      Dim tempTable As String
1        On Error GoTo cmdCetakRekap_Click_Error

2     Set rpt = New RptFakturRekap
              
3         Call refreshList(Trim(cboNoFP))
4         If rsInvoice.State <> adStateClosed Then
5             tempTable = "TABLE_CETAK" & DriveSerialNumber("c:")
6             If rsFaktur.State Then rsFaktur.Close
7             Set rsFaktur = rsInvoice.Clone
8             If Not rsFaktur.EOF Then
9                 rsFaktur.MoveFirst
10                Call createTempTable(tempTable)
11                Do While Not rsFaktur.EOF
12                    Call objO.InsertData("INSERT INTO [tempdb]..[" & tempTable & "] VALUES ('" & Trim(rsFaktur("NFP")) & "','" & Trim(rsFaktur("CUST")) & "','" & Format(rsFaktur("TGLFP"), "MM/DD/YYYY") & "','" & Trim(rsFaktur("INVO")) & "','" & Round(CDbl(rsFaktur("TTL")) * 100 / 110, 0) & "','" & CDbl(rsFaktur("R")) & "')")
13                    rsFaktur.MoveNext
14                Loop
15            End If
              
16            sSQL = "SELECT * FROM [tempdb]..[" & tempTable & "] ORDER BY NOFP"
17            Set rsRepo = objO.FetchData(sSQL)
18            rpt.database.SetDataSource rsRepo
19            rpt.txtPerusahaan.SetText gPerush
20            rpt.txtAlamat.SetText gAlamat
21            rpt.txtTanggal.SetText "Tanggal " & TglDDMMMMYYYY(Format(DtDOAwal, "DD/MM/YYYY"), 1) & "    ~    " & TglDDMMMMYYYY(Format(DtDOAkhir, "DD/MM/YYYY"), 1)
              
22            With CRViewer3
23                .ReportSource = rpt
24                .ViewReport
25                .Zoom 100
26                .Top = 0
27                .Left = 0
28                .Width = GridProgram.Width + 100
29                .Height = cmdCreateFP.Top + cmdCreateFP.Height
                  
30                cmdTutup3.Top = .Top + 430
31                cmdTutup3.Left = .Left
32                cmdTutup3.Visible = True
                  
33                cmdSettingPrinter.Top = cmdTutup3.Top
34                cmdSettingPrinter.Left = cmdTutup3.Left + cmdTutup3.Width + 50
35                cmdSettingPrinter.Visible = True
36                .Visible = True
37            End With
38            cmdTutup3.Visible = True
39            Screen.MousePointer = vbDefault
40        End If

41       On Error GoTo 0
42       Exit Sub

cmdCetakRekap_Click_Error:

43        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetakRekap_Click of Form frmFakturPajak")
End Sub

Private Sub cmdCetakUlang_Click()
      Dim rsSearch As ADODB.Recordset
      Dim boleh As Boolean
      Dim nFakturLama As String

1        On Error GoTo cmdCetakUlang_Click_Error

2         boleh = False
3         If rsInvoice.State <> adStateClosed Then
4             rsInvoice.MoveFirst
5             rsInvoice.Find "ST = true", , adSearchForward, 1
6             If Not rsInvoice.EOF Then
7                 If rsInvoice("ST") Then
8                     boleh = True
9                 End If
10            End If
11        End If
          
12        If rsInvoice.State <> adStateClosed Then
13            rsInvoice.MoveFirst
14            rsInvoice.Find "ST = true", , adSearchForward, 1
15            If Not rsInvoice.EOF Then
16                Do While Not rsInvoice.EOF
17                    If rsInvoice("ST") Then
18                        If rsInvoice("NFP") = "expired" Then
19                            boleh = False
20                        End If
21                    End If
22                    rsInvoice.MoveNext
23                Loop
24            End If
25        End If
          
26        nFakturLama = ""
27        If rsInvoice.State <> adStateClosed Then
28            rsInvoice.MoveFirst
29            If Not rsInvoice.EOF Then
30                Do While Not rsInvoice.EOF
31                    If rsInvoice("ST") Then
32                        If Trim(rsInvoice("TGLFP")) = "" Then
33                            boleh = False
34                            MsgBox "Faktur tidak terdaftar !"
35                            GoTo Proses
36                        End If
37                        If Trim(rsInvoice("NFP")) = "" Then
38                            boleh = False
39                            MsgBox "Invoice tersebut belum dibuat faktur"
40                            GoTo Proses
41                        Else
42                            If Trim(nFakturLama) = "" Then
43                                nFakturLama = Trim(rsInvoice("NFP"))
44                            Else
45                                If Trim(rsInvoice("NFP")) <> Trim(nFakturLama) Then
46                                    boleh = False
47                                    GoTo Proses
48                                Else
49                                    boleh = True
50                                End If
51                            End If
52                        End If
53                    End If
54                    rsInvoice.MoveNext
55                Loop
56            End If
57        End If
              
Proses:
58        If boleh Then
59            rsInvoice.MoveFirst
60            If rsInvoice.State <> adStateClosed Then
61                rsInvoice.Find "ST = true", , adSearchForward, 1
62            End If
63            If MsgBox("Hanya lihat ?", vbYesNo, "Cetak Ulang") = vbNo Then
64                NoBaru = False
65                If rsInvoice.State <> adStateClosed Then
66                    rsInvoice.MoveFirst
67                    rsInvoice.Find "ST = true", , adSearchForward, 1
68                    If Not rsInvoice.EOF Then
69                        Call bersih
70                        text(7) = Trim(gKotaCbg)
                          ' TIPE CUSTOMER
71                        Set rsSearch = objO.FetchDataACC("SELECT * FROM [FAKTUR" & Right(gDb, 2) & "]..TYPECUST WITH(NOLOCK) ORDER BY KODETYPE")
72                        If Not rsSearch.EOF Then
73                            With cboType
74                                .Clear
75                                Do While Not rsSearch.EOF
76                                    .AddItem (rsSearch("NAMA"))
77                                    rsSearch.MoveNext
78                                Loop
79                            End With
80                        End If
                           
                          
81                        With cboTtd
82                            .Clear
83                            .AddItem ("KACAB")
84                            .AddItem ("ADH")
85                            .AddItem ("KABENG")
86                        End With
                          
87                        Set rsSearch = objO.FetchData("SELECT * FROM SETTANDA WITH(NOLOCK)")
88                        If Not rsSearch.EOF Then
89                            cboTtd = Trim(rsSearch("OTORITAS")) '"KABENG"
90                            chJabatan = IIf(Trim(rsSearch("DISPLAY")) = 0, Unchecked, Checked)
91                            Select Case CDbl(rsSearch("TPPRINT"))
                                  Case 0
92                                    prtrn(0).value = True
93                                Case 1
94                                    prtrn(1).value = True
95                                Case 2
96                                    prtrn(2).value = True
97                                Case 3
98                                    prtrn(3).value = True
99                            End Select
100                       Else
101                           cboTtd = "KABENG"
102                           chJabatan = Checked
103                       End If
                           
                          
                          ' USER OTORITAS
104                       Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..T_PROFIL WITH(NOLOCK)")
105                       If Not rsSearch.EOF Then
106                           Select Case cboTtd
                                  Case "KACAB"
107                                   text(6) = Trim(rsSearch("FS_KA_CABANG"))
108                                   text(8) = "Kepala Cabang"
109                               Case "ADH"
110                                   text(6) = Trim(rsSearch("FS_KA_ADH"))
111                                   text(8) = "ADH"
112                               Case "KABENG"
113                                   text(6) = Trim(rsSearch("FS_KA_BENGKEL"))
114                                   text(8) = "Kepala Bengkel"
115                           End Select
116                       End If
                           
                          
                          ' DEFAULT CUSTOMER
117                       Set rsSearch = objO.FetchDataACC("SELECT * FROM [FAKTUR" & Right(gDb, 2) & "]..SIFPSTANDAR WITH(NOLOCK) WHERE NOFPS = '" & Trim(rsInvoice("NFP")) & "'")
118                       If Not rsSearch.EOF Then
119                           tglFaktur = Format(rsSearch("TGLFPS"), "DD/MM/YYYY")
120                           tglFaktur.Enabled = False
121                           Text1(1) = Trim(rsSearch("NAMAFP"))
122                           Text1(0) = Trim(rsSearch("ALAMATFP1"))
123                           Text1(5) = Trim(rsSearch("ALAMATFP2"))
124                           text(2) = IIf(Trim(rsSearch("NPWP")) <> "", Trim(rsSearch("NPWP")), "")
125                       End If
                           
                          
126                       If Trim(text(2)) = "0" Then
127                           Set rsSearch = objO.FetchData("SELECT NPWP FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsInvoice("RGK")) & "'")
128                           If Not rsSearch.EOF Then
129                               text(2) = IIf(Trim(rsSearch("NPWP")) <> "", Trim(rsSearch("NPWP")), "")
130                           End If
                               
131                       End If
                              
                          ' CEK TAHUN INVOICE
                          Dim thnStat As Boolean
132                       thnInv = ""
133                       thnStat = False
134                       If Not rsInvoice.EOF Then rsInvoice.MoveFirst
135                       rsInvoice.Find "ST = true", , adSearchForward, 1
136                       If Not rsInvoice.EOF Then
137                           thnInv = Right(year(rsInvoice("TGLINVO")), 2)
138                       End If
139                       If cekAdaPart(Trim(rsInvoice("NFP"))) Then
140                           chHarga = Unchecked
141                       End If
142                       frmFaktur.Visible = True
143                   End If
144               End If
145           Else
146               Set rsSearch = objO.FetchData("SELECT * FROM SETTANDA WITH(NOLOCK)")
147               If Not rsSearch.EOF Then
148                   Select Case CDbl(rsSearch("TPPRINT"))
                          Case 0
149                           prtrn(0).value = True
150                       Case 1
151                           prtrn(1).value = True
152                       Case 2
153                           prtrn(2).value = True
154                       Case 3
155                           prtrn(3).value = True
156                   End Select
157               End If
                   
              
                  'Call cetakFaktur(rsInvoice("NFP"), 2)
158               noFakturCetak = Trim(rsInvoice("NFP"))
159               If cekAdaPart(Trim(rsInvoice("NFP"))) Then
160                   chHarga = Unchecked
161               End If
162               frmPrinter.Visible = True
163               cmdLanjut.SetFocus
164           End If
165       Else
166           MsgBox "Pilih yang akan dicetak ", vbInformation
167       End If

168      On Error GoTo 0
169      Exit Sub

cmdCetakUlang_Click_Error:

170       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetakUlang_Click of Form frmFakturPajak")
End Sub

Private Function cekAdaPart(NoFak As String) As Boolean
      Dim rsSearch As ADODB.Recordset
      Dim rsSearch2 As ADODB.Recordset
          
1        On Error GoTo cekAdaPart_Error

2         cekAdaPart = False
3         Set rsSearch = objO.FetchDataACC("SELECT * FROM [FAKTUR" & Right(gDb, 2) & "]..INVOICE WITH(NOLOCK) WHERE NOFPS = '" & Trim(NoFak) & "'")
4         If Not rsSearch.EOF Then
5             rsSearch.MoveFirst
6             Do While Not rsSearch.EOF
7                 Set rsSearch2 = objO.FetchData("SELECT NOFAKTUR FROM T_INVOICE WITH(NOLOCK) WHERE NOINV = '" & onlyNum(Trim(rsSearch("NOINVOICE"))) & "'")
8                 If Not rsSearch2.EOF Then
9                     If Trim(rsSearch2("NOFAKTUR")) <> "" Then
10                        cekAdaPart = True
11                        Exit Function
12                    End If
13                End If
14                If rsSearch2.State <> adStateClosed Then rsSearch2.Close
15                rsSearch.MoveNext
16            Loop
17        End If

18       On Error GoTo 0
19       Exit Function

cekAdaPart_Error:

20        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekAdaPart of Form frmFakturPajak")
           
End Function

Private Sub cmdCreateFP_Click()
          Dim objIn As New BClsBilling
1        On Error GoTo cmdCreateFP_Click_Error

2         isFP_OK = True
3         If objIn.FetchAvailable_FPNew = False Then
4             isFP_OK = False
5         End If
6         isNoFP_OK = True
7         isNoFP_New = objIn.FetchRange_FPNew
8         If objIn.FetchMinRange_FPNew >= isNoFP_New Then
9             isNoFP_OK = False
10        End If
          
11        If isFP_OK = False Then
12            MsgBox "Anda tidak punya Nomor Faktur" & Chr(13) & "Hubungi Administrasi (ADH) - FP" & Chr(13) & "Invoice tidak Bisa Dibuka", vbCritical + vbOKOnly
13            Exit Sub
14        End If
15        If isNoFP_OK = False Then
16             MsgBox "Sisa Nomor Faktur : " & isNoFP_New & Chr(13) & "Segera Hubungi Administrasi (ADH) - FP", vbInformation + vbOKOnly
17        End If
18        If isNoFP_New < 2 Then
19            MsgBox "Anda tidak punya Nomor Faktur" & Chr(13) & "Hubungi Administrasi (ADH) - FP" & Chr(13) & "Invoice tidak Bisa Dibuka", vbCritical + vbOKOnly
20            Exit Sub
21        End If

          Dim rsSearch As ADODB.Recordset
          Dim boleh As Boolean
          Dim pesan As String
          Dim Pesan1 As String

22        boleh = False
23        NoBaru = True
24        pesan = ""
25        Pesan1 = ""
26        If rsInvoice.State <> adStateClosed Then
27            Do While Not rsInvoice.EOF
28                If rsInvoice("ST") Then
29                    boleh = True
30                End If
31                rsInvoice.MoveNext
32            Loop
33            rsInvoice.MoveFirst
              
34            Do While Not rsInvoice.EOF
35                If Trim(rsInvoice("NFP")) <> "" And rsInvoice("ST") Then
36                    boleh = False
37                End If
38                rsInvoice.MoveNext
39            Loop
40            If Not boleh Then
41                GoTo Cetak
42            End If
43            rsInvoice.MoveFirst
              
              ' PEMBUATAN FP DENGAN 1 INVOICE UTK FPGABUNGAN
44            countInv = 0
45            Do While Not rsInvoice.EOF
46                If rsInvoice("ST") And Trim(rsInvoice("NFP")) = "" Then
47                    countInv = countInv + 1
48                End If
49                rsInvoice.MoveNext
50            Loop
51            rsInvoice.MoveFirst
              
52            Pesan1 = ""
53            rsInvoice.Find "ST = true", , adSearchForward, 1
54            If Not rsInvoice.EOF Then
55                kdLgn = Trim(rsInvoice("PLG"))
56                rsInvoice.MoveNext
57                Do While Not rsInvoice.EOF
58                    If rsInvoice("ST") Then
59                        If Trim(kdLgn) <> Trim(rsInvoice("PLG")) Then
60                            boleh = False
61                            Pesan1 = "Faktur gabungan dengan customer berbeda tidak bisa dibuat. "
62                        End If
63                    End If
64                    rsInvoice.MoveNext
65                Loop
66            End If
67            rsInvoice.MoveFirst
68            rsInvoice.Find "ST = true", , adSearchForward, 1
69            If Trim(pesan) <> "" Then
70                pesan = pesan & vbCr & Pesan1
71            Else
72                pesan = Pesan1
73            End If
74            Pesan1 = ""
              
              Dim tglInv As Date
75            rsInvoice.Find "ST = true", , adSearchForward, 1
76            If Not rsInvoice.EOF Then
77                tglInv = Format(rsInvoice("TGLINVO"), "DD/MM/YYYY")
78                rsInvoice.MoveNext
79                Do While Not rsInvoice.EOF
80                    If rsInvoice("ST") Then
81                        If DateDiff("m", tglInv, Format(rsInvoice("TGLINVO"), "DD/MM/YYYY")) <> 0 Then
82                            boleh = False
83                            Pesan1 = "Faktur gabungan dengan bulan berbeda tidak bisa dibuat. "
84                        End If
85                    End If
86                    rsInvoice.MoveNext
87                Loop
88            End If
89            rsInvoice.MoveFirst
90            rsInvoice.Find "ST = true", , adSearchForward, 1
91            If Trim(pesan) <> "" Then
92                pesan = pesan & vbCr & Pesan1
93            Else
94                pesan = Pesan1
95            End If
96            Pesan1 = ""
              
              Dim notExistFPGab As Boolean
97            notExistFPGab = True
98            rsInvoice.MoveFirst
99            rsInvoice.Find "ST = true", , adSearchForward, 1
100           Do While Not rsInvoice.EOF
101               If rsInvoice("ST") Then
102                   Set rsSearch = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN WITH(NOLOCK) WHERE NOINVOICE = '" & Trim(rsInvoice("INVO")) & "'")
103                   If Not rsSearch.EOF Then
104                       If Not Trim(rsSearch("PEMERINTAH")) Then
105                           If DateDiff("m", rsInvoice("TGLINVO"), isnow) <> 0 Or (year(Format(rsInvoice("TGLINVO"), "DD/MM/YYYY")) <> year(Format(isnow, "DD/MM/YYYY"))) Then
106                               boleh = False
107                               Pesan1 = "Untuk perusahaan Non Pemerintah Tidak boleh buat Faktur Gabungan beda fiskal. "
108                           End If
109                       Else
110                           If DateDiff("m", rsInvoice("TGLINVO"), isnow) > 3 Or (year(Format(rsInvoice("TGLINVO"), "DD/MM/YYYY")) <> year(Format(isnow, "DD/MM/YYYY"))) Then
111                               boleh = False
112                               Pesan1 = "Untuk perusahaan Pemerintah Tidak boleh buat Faktur Gabungan Lebih dari 3 bulan. "
113                           End If
114                       End If
115                   Else
116                       notExistFPGab = False
117                   End If
                       
118               End If
119               rsInvoice.MoveNext
120           Loop
121           If Trim(pesan) <> "" Then
122               pesan = pesan & vbCr & Pesan1
123           Else
124               pesan = Pesan1
125           End If
126           Pesan1 = ""
              
127           rsInvoice.MoveFirst
128           rsInvoice.Find "ST = true", , adSearchForward, 1
129           If Not notExistFPGab Then
130               Set rsSearch = objO.FetchDataACC("SELECT KODELGN FROM ARMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(rsInvoice("PLG")) & "'")
131               If rsSearch.EOF Then
132                   boleh = False
133                   pesan = pesan & vbCr & "Kode Pelanggan tersebut tidak terdaftar pada fiskal bulan ini. "
134               End If
                   
              
135               Set rsSearch = objO.FetchDataACC("SELECT * FROM ARMST WITH(NOLOCK) WHERE KODELGN = '" & Trim(rsInvoice("PLG")) & "'")
136               If Not rsSearch.EOF Then
137                   If Trim(rsSearch("STATUSPERSH")) <> "P" Then
138                       If DateDiff("m", rsInvoice("TGLINVO"), isnow) <> 0 Or (year(Format(rsInvoice("TGLINVO"), "DD/MM/YYYY")) <> year(Format(isnow, "DD/MM/YYYY"))) Then
139                           boleh = False
140                           pesan = pesan & vbCr & "Untuk perusahaan Non Pemerintah Tidak boleh buat Faktur Gabungan beda fiskal. "
141                       End If
142                   Else
143                       If DateDiff("m", rsInvoice("TGLINVO"), isnow) > 3 Or (year(Format(rsInvoice("TGLINVO"), "DD/MM/YYYY")) <> year(Format(isnow, "DD/MM/YYYY"))) Then
144                           boleh = False
145                           pesan = pesan & vbCr & "Untuk perusahaan Pemerintah Tidak boleh buat Faktur Gabungan Lebih dari 3 bulan. "
146                       End If
147                   End If
148               End If
                   
149           End If
              
150           rsInvoice.MoveFirst
151           rsInvoice.Find "ST = true", , adSearchForward, 1
              
152           If boleh Then
153               If rsInvoice.State <> adStateClosed Then
154                   rsInvoice.MoveFirst
155                   rsInvoice.Find "ST = true", , adSearchForward, 1
156                   If Not rsInvoice.EOF Then
157                       Call bersih
158                       text(7) = Trim(gKotaCbg)
                          ' TIPE CUSTOMER
159                       Set rsSearch = objO.FetchDataACC("SELECT * FROM [FAKTUR" & Right(gDb, 2) & "]..TYPECUST WITH(NOLOCK) ORDER BY KODETYPE")
160                       If Not rsSearch.EOF Then
161                           With cboType
162                               .Clear
163                               Do While Not rsSearch.EOF
164                                   .AddItem (rsSearch("NAMA"))
165                                   rsSearch.MoveNext
166                               Loop
167                           End With
168                       End If
                           
                          
                          ' USER OTORITAS
169                       With cboTtd
170                           .Clear
171                           .AddItem ("KACAB")
172                           .AddItem ("ADH")
173                           .AddItem ("KABENG")
174                       End With
                          
                           
175                       Set rsSearch = objO.FetchData("SELECT * FROM SETTANDA WITH(NOLOCK)")
176                       If Not rsSearch.EOF Then
177                           cboTtd = Trim(rsSearch("OTORITAS")) '"KABENG"
178                           chJabatan = IIf(Trim(rsSearch("DISPLAY")) = 0, Unchecked, Checked)
179                           Select Case CDbl(rsSearch("TPPRINT"))
                                  Case 0
180                                   prtrn(0).value = True
181                               Case 1
182                                   prtrn(1).value = True
183                               Case 2
184                                   prtrn(2).value = True
185                               Case 3
186                                   prtrn(3).value = True
187                           End Select
188                       Else
189                           cboTtd = "KABENG"
190                           chJabatan = Checked
191                       End If
                           
                          
                          ' USER OTORITAS
                           
192                       Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..T_PROFIL WITH(NOLOCK)")
193                       If Not rsSearch.EOF Then
194                           Select Case cboTtd
                                  Case "KACAB"
195                                   text(6) = Trim(rsSearch("FS_KA_CABANG"))
196                                   text(8) = "Kepala Cabang"
197                               Case "ADH"
198                                   text(6) = Trim(rsSearch("FS_KA_ADH"))
199                                   text(8) = "ADH"
200                               Case "KABENG"
201                                   text(6) = Trim(rsSearch("FS_KA_BENGKEL"))
202                                   text(8) = "Kepala Bengkel"
203                           End Select
204                       End If
                           
                          
205                       tglFaktur = Format(isnow, "DD/MM/YYYY") 'GetTglAkhir(Month(rsInvoice("TGLINVO")), Year(rsInvoice("TGLINVO")), 0) '
206                       tglFaktur.Enabled = False
                          
                          
                          Dim rsSearch2 As ADODB.Recordset
207                       Set rsSearch2 = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN WITH(NOLOCK) WHERE NOINVOICE = '" & Trim(rsInvoice("INVO")) & "'")
208                       If Not rsSearch2.EOF Then
209                           Text1(1) = Trim(rsSearch2("NAMAFP"))
210                           Text1(0) = Trim(rsSearch2("ALAMAT1"))
211                           Text1(5) = Trim(rsSearch2("ALAMAT2"))
212                           text(2) = Trim(rsSearch2("NPWP"))
213                       Else
                          ' DEFAULT CUSTOMER
214                           If Right(rsInvoice("INVO"), 2) <> "-B" Then
215                               If Trim(rsInvoice("W")) = 0 Then
216                                   If Trim(rsInvoice("CUST")) = "PT. NASMOCO" Then
217                                       Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..T_DEALER WITH(NOLOCK) WHERE KDPLG = '" & Trim(rsInvoice("PLG")) & "'")
218                                       If Not rsSearch.EOF Then
219                                           Text1(1) = Trim(rsSearch("FS_NM_TAGIH"))
220                                           Text1(0) = Trim(rsSearch("FS_ALAMAT"))
221                                           Text1(5) = Trim(rsSearch("FS_KOTA"))
222                                           text(2) = IIf(Trim(rsSearch("FS_NPWP")) <> "", Trim(rsSearch("FS_NPWP")), "")
223                                       End If
                                           
224                                   Else
225                                       Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsInvoice("RGK")) & "'")
226                                       If Not rsSearch.EOF Then
227                                           Text1(1) = Trim(rsSearch("CUSTOMER"))
228                                           Text1(0) = Trim(rsSearch("CUST_ADRES"))
229                                           Text1(5) = Trim(rsSearch("KOTA"))
230                                           text(2) = IIf(Trim(rsSearch("NPWP")) <> "", Trim(rsSearch("NPWP")), "")
231                                       End If
                                           
232                                   End If
233                               Else
234                                   Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..T_PROFIL WITH(NOLOCK)")
235                                   If Not rsSearch.EOF Then
236                                       Text1(1) = Trim(rsSearch("FN_HO"))
237                                       Text1(0) = Trim(rsSearch("FN_HOADRES"))
238                                       Text1(5) = Trim(rsSearch("FN_HOKOTA"))
239                                       text(2) = IIf(Trim(rsSearch("FN_HONPWP")) <> "", Trim(rsSearch("FN_HONPWP")), "")
240                                   End If
                                       
241                               End If
242                           ElseIf Right(rsInvoice("INVO"), 2) = "-B" Then
243                               Set rsSearch = objO.FetchData("SELECT * FROM " & gDb & "..T_DEALER WITH(NOLOCK) WHERE FS_KD_DEALER IN (SELECT FS_KD_DEALER FROM CUSTDAT WITH(NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsInvoice("RGK")) & "')")
244                               If Not rsSearch.EOF Then
245                                   Text1(1) = Trim(rsSearch("FS_NM_TAGIH"))
246                                   Text1(0) = Trim(rsSearch("FS_ALAMAT"))
247                                   Text1(5) = Trim(rsSearch("FS_KOTA"))
248                                   text(2) = IIf(Trim(rsSearch("FS_NPWP")) <> "", Trim(rsSearch("FS_NPWP")), "")
249                               End If
                                   
250                           End If
251                       End If
252                       If rsSearch2.State <> adStateClosed Then rsSearch2.Close
                          
                          ' CEK TAHUN INVOICE
                          Dim thnStat As Boolean
253                       thnInv = ""
254                       thnStat = False
255                       If Not rsInvoice.EOF Then rsInvoice.MoveFirst
256                       rsInvoice.Find "ST = true", , adSearchForward, 1
257                       If Not rsInvoice.EOF Then
258                           thnInv = Right(year(rsInvoice("TGLINVO")), 2)
259                       End If
260                       frmFaktur.Visible = True
261                   End If
262               End If
263           ElseIf Trim(pesan) <> "" Then
264               MsgBox pesan
265           Else
Cetak:
266               If MsgBox("Cetak Ulang ?", vbYesNo, "") = vbYes Then
267                   cmdCetakUlang_Click
268               End If
269           End If
270       Else
271           MsgBox "Pilih Invoice yang akan dibuat Faktur Pajak"
272       End If

273      On Error GoTo 0
274      Exit Sub

cmdCreateFP_Click_Error:

275       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCreateFP_Click of Form frmFakturPajak")
End Sub

Private Sub cmdCtkFaktur_Click()
      Dim rsFaktur As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset
      Dim rsSearch2 As ADODB.Recordset
      Dim NoInvoice As String
      Dim NoInvoiceACC As String
      Dim NoInvoiceKet As String
      Dim nCountFP As Double
      Dim boleh As Boolean
      Dim pesan As String
      Dim Ket As String

1        On Error GoTo cmdCtkFaktur_Click_Error

2         On Error GoTo vbError
3         boleh = True
4         pesan = ""
5         If frmPrinter.Visible Then
6             boleh = False
7         End If
8         If Trim(Text1(0)) = "" Then
9             boleh = False
10            pesan = pesan + "Alamat. "
11        End If
12        If Trim(Text1(1)) = "" Then
13            boleh = False
14            pesan = pesan + "Nama. "
15        End If
16        If Trim(Nul(tglFaktur)) = "" Then
17            boleh = False
18            pesan = pesan + "Tanggal Faktur "
19        Else
20            If Right(year(tglFaktur), 2) <> thnInv Then
21                If MsgBox("Tahun cetak Faktur berbeda dengan Tahun invoice" & vbCr & "Mengikuti tahun cetak Faktur ?", vbYesNo, "Konfirmasi") = vbYes Then
22                    thnInv = Right(year(tglFaktur), 2)
23                End If
24            End If
25        End If
26        If Trim(cboType) = "" Then
27            boleh = False
28            pesan = pesan + "Kode Transaksi. "
29        End If
              
30        If boleh Then
          
31        If OnNewFP = 0 Then
32            Set rsFaktur = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..DEALER WITH(NOLOCK)")
33        Else
34            Set rsFaktur = objO.FetchDataACC("SELECT left(NoAwal,3) as KODEDEALER FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH(NOLOCK) WHERE TAHUN = '" & Format(isnow, "YYYY") & "' and NoSekarang <> NoAkhir Order By Nomor ASC")
35        End If
36        kdDealer = Trim(rsFaktur("KODEDEALER"))
                  
37            Call saveLogMenu(gSA, ComptName, gVersi, "frmFakturPajak", "cmdCtkFaktur", 0)
              
38                erTranAll = False
39                onTR = True
40                Call objHT.OpenTransactionAll
              
cariNo:
41            If NoBaru Then
42                If OnNewFP = 0 Then
43                    Set rsSearch = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..PROFIL WITH(NOLOCK) WHERE TAHUN = '" & Right(year(tglFaktur), 2) & "'")
44                Else
45                    Set rsSearch = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..FPKonter WITH(NOLOCK) WHERE TAHUN = '" & year(tglFaktur) & "' and NoSekarang <> NoAkhir Order By Nomor ASC")
46                End If
47                If Not rsSearch.EOF Then
48                    If OnNewFP = 0 Then
49                        nCountFP = CDbl(rsSearch("NOFAKTURPAJAK"))
50                    Else
51                        nCountFP = CDbl(Right(rsSearch("NOSEKARANG"), 8))
52                    End If
53                    nCountFP = nCountFP + 1
54                    If OnNewFP = 0 Then
55                        objHT.UpdateData_ACC ("UPDATE FAKTUR" & Right(gDb, 2) & "..PROFIL SET NOFAKTURPAJAK = '" & Format(nCountFP, "00000000") & "' WHERE TAHUN = '" & Right(thnInv, 2) & "'")
56                    Else
57                        objHT.UpdateData_ACC ("UPDATE FAKTUR" & Right(gDb, 2) & "..FPKonter SET NOSEKARANG = '" & objO.FetchData_StringACC("SELECT KodeDealer FROM FAKTUR" & Right(gDb, 2) & "..DEALER") & "-" & Right(thnInv, 2) & "." & Format(nCountFP, "00000000") & "' WHERE TAHUN = '20" & thnInv & "' And Nomor = " & Val(rsSearch.Fields(0).value))
58                    End If
59                Else
60                    nCountFP = 1
61                    Call objHT.InsertData_ACC("INSERT INTO FAKTUR" & Right(gDb, 2) & "..PROFIL VALUES (" & Format(nCountFP, "00000000") & ",'" & Right(thnInv, 2) & "')")
62                End If
                   
63                Set rsSearch = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH(NOLOCK) WHERE NOFPS = '" & Trim(text(4)) + Format(nCountFP, "00000000") & "'")
64                If Not rsSearch.EOF Then
65                    GoTo cariNo
66                End If
                   
67            End If
              
68            Call setNoFaktur(cboType.text)
              
69            NoInvoice = ""
70            NoInvoiceACC = ""
71            NoInvoiceKet = ""
72            rsInvoice.Find "ST = true", , adSearchForward, 1
73            If Not rsInvoice.EOF Then
74                Do While Not rsInvoice.EOF
75                    If rsInvoice("ST") Then
76                        NoInvoice = NoInvoice & Mid(Trim(rsInvoice("INVO")), 3, 8) & "','"
77                        NoInvoiceACC = NoInvoiceACC & Trim(rsInvoice("INVO")) & "','"
78                        NoInvoiceKet = NoInvoiceKet & Trim(rsInvoice("INVO"))
79                    End If
80                    rsInvoice.MoveNext
81                Loop
82            End If
              
83            rsInvoice.MoveFirst
84            rsInvoice.Find "ST = true", , adSearchForward, 1
              
              Dim isKTP As Boolean
85            isKTP = False
86            If nTagih > 1 Then ' Or cboFaktur = "Faktur Pajak Gabungan"
87                Ket = "(" & nTagih & " rincian invoice terlampir)."
88            Else
                  Dim NoInvD As String
89                NoInvD = ""
90                Set rsSearch = objO.FetchData("SELECT NOFAKTUR FROM T_INVOICE WITH(NOLOCK) WHERE NOINV IN ('" & Trim(NoInvoice) & "')")
91                If Not rsSearch.EOF Then
92                    Do While Not rsSearch.EOF
93                        NoInvD = NoInvD & IIf(Trim(NoInvD) <> "", ",", "") & Trim(rsSearch("NOFAKTUR"))
94                        rsSearch.MoveNext
95                    Loop
96                End If
97                isKTP = True
98                Ket = "No Invoice : " & Trim(NoInvoiceKet) & IIf(Trim(NoInvD) <> "", " , " & Trim(NoInvD), "") & " , No. Polisi : " & Trim(rsInvoice("NOPOL")) '" , CUSTOMER : " & Trim(Text(2)) & ", DPP : " & setKoma(Round(CDbl(totTagih)))
99            End If
              
100           If NoBaru Then
101               text(4) = Trim(text(4)) + Format(nCountFP, "#00000000")
102               Set rsSearch = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN WITH(NOLOCK) WHERE NOINVOICE IN ('" & Trim(NoInvoiceACC) & "')")
103               If rsSearch.EOF Then
104                   Call objHT.InsertData_ACC("INSERT INTO FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR (NoFPS,TglFPS,NPWP,NPPKP,Keterangan,HargaJual,PPN,PPnBM,KodeUser,TglEntry,NamaFP,AlamatFP1,AlamatFP2,Modul,kodelgn,NoKTP) VALUES ('" & Trim(text(4)) & "','" & Format(tglFaktur, "MM/DD/YYYY") & "','" & Trim(text(2)) & "','" & Trim(text(3)) & "','" & Trim(Ket) & "'," & Round(CDbl(totTagih) * (100 / 110), 0) & "," & Round(CDbl(totTagih) * (10 / 110), 0) & ",0,'" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Text1(1)) & "','" & Trim(Text1(0)) & "','" & Trim(Text1(5)) & "','SVC','" & Trim(rsInvoice("PLG")) & "','')")
105               Else
106                   Call objHT.InsertData_ACC("INSERT INTO FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR (NoFPS,TglFPS,NPWP,NPPKP,Keterangan,HargaJual,PPN,PPnBM,KodeUser,TglEntry,NamaFP,AlamatFP1,AlamatFP2,Modul,kodelgn,NoKTP) SELECT '" & Trim(text(4)) & "','" & Format(tglFaktur, "MM/DD/YYYY") & "','" & Trim(text(2)) & "','" & Trim(text(3)) & "','" & Trim(Ket) & IIf(isKTP, "No KTP " & NulSID(rsSearch!noKTP), "") & "',SUM(ISNULL(DPP,0)) AS DPP,SUM(ISNULL(PPN,0)) AS PPN,0,'" & Trim(gSA) & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(Text1(1)) & "','" & Trim(Text1(0)) & "','" & Trim(Text1(5)) & "','SVC','" & Trim(rsInvoice("PLG")) & "','" & NulSID(rsSearch!noKTP) & "' FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN WITH(NOLOCK) WHERE NOINVOICE IN ('" & Trim(NoInvoiceACC) & "')")
107               End If
                   
                  
108               rsInvoice.Find "ST = true", , adSearchForward, 1
109               If Not rsInvoice.EOF Then
110                   Do While Not rsInvoice.EOF
111                       If rsInvoice("ST") Then
112                           Call objHT.InsertData_ACC("INSERT INTO FAKTUR" & Right(gDb, 2) & "..INVOICE VALUES ('" & Trim(text(4)) & "','" & Trim(rsInvoice("INVO")) & "','" & Format(rsInvoice("TGLINVO"), "MM/DD/YYYY") & "')")
113                           Call objHT.InsertData_ACC("UPDATE FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN SET NOFPS = '" & Trim(text(4)) & "', TGLENTRY = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', KODEENTRY = '" & Trim(gSA) & "' WHERE NOINVOICE = '" & Trim(rsInvoice("INVO")) & "'")
114                           Call objHT.InsertData_ACC("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET NOFP = '" & Trim(text(4)) & "' WHERE NOINV = '" & Mid(Trim(rsInvoice("INVO")), 3, Len(Trim(rsInvoice("INVO"))) - 2) & "'")
115                       End If
116                       rsInvoice.MoveNext
117                   Loop
118               End If
                  
119               rsInvoice.MoveFirst
120               rsInvoice.Find "ST = true", , adSearchForward, 1
121               If Not rsInvoice.EOF Then
122                   Do While Not rsInvoice.EOF
123                       If rsInvoice("ST") Then
124                           If Right(rsInvoice("INVO"), 2) = "-B" Then
125                               Call objHT.UpdateData_BengkelOnly("UPDATE T_INVOICE SET NOFP2 = '" & Trim(text(4)) & "' WHERE NOINV = '" & Mid(Trim(rsInvoice("INVO")), 3, 8) & "'")
126                           Else
127                               Call objHT.UpdateData_BengkelOnly("UPDATE T_INVOICE SET NOFP = '" & Trim(text(4)) & "' WHERE NOINV = '" & Mid(Trim(rsInvoice("INVO")), 3, Len(Trim(rsInvoice("INVO"))) - 2) & "'")
128                           End If
129                       End If
130                       rsInvoice.MoveNext
131                   Loop
132               End If
133           Else
134               rsInvoice.Find "ST = true", , adSearchForward, 1
135               If Not rsInvoice.EOF Then
136                   text(4) = Left(text(4), 3) & Mid(Trim(rsInvoice("NFP")), 4, Len(Trim(rsInvoice("NFP"))) - 3)
137               End If
138               Call objO.UpdateDataACC("UPDATE FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR SET NOFPS = '" & Trim(text(4)) & "', TGLFPS = '" & Format(tglFaktur, "MM/DD/YYYY") & "', NPWP = '" & Trim(text(2)) & "', NPPKP = '" & Trim(text(3)) & "', KODEUSER = '" & Trim(gSA) & "', TGLENTRY = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', NAMAFP = '" & Trim(Text1(1)) & "', ALAMATFP1 = '" & Trim(Text1(0)) & "', ALAMATFP2 = '" & Trim(Text1(5)) & "' WHERE NOFPS = '" & Trim(rsInvoice("NFP")) & "'")
139               Call objO.UpdateDataACC("UPDATE FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN SET NOFPS = '" & Trim(text(4)) & "', NPWP = '" & Trim(text(2)) & "', NPPKP = '" & Trim(text(3)) & "',  KODEENTRY = '" & Trim(gSA) & "', TGLENTRY = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', NAMAFP = '" & Trim(Text1(1)) & "', ALAMAT1 = '" & Trim(Text1(0)) & "', ALAMAT2 = '" & Trim(Text1(5)) & "' WHERE NOFPS = '" & Trim(rsInvoice("NFP")) & "'")
140               Call objO.UpdateDataACC("UPDATE FAKTUR" & Right(gDb, 2) & "..INVOICE SET NOFPS = '" & Trim(text(4)) & "' WHERE NOFPS = '" & Trim(rsInvoice("NFP")) & "'")
141               Call objO.UpdateDataACC("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET NOFP = '" & Trim(text(4)) & "' WHERE NOFP = '" & Trim(rsInvoice("NFP")) & "'")
                  'Call objHT.UpdateData_ACC("UPDATE FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR SET NOFPS = '" & Trim(Text(4)) & "', TGLFPS = '" & Format(tglFaktur, "MM/DD/YYYY") & "', NPWP = '" & Trim(Text(2)) & "', NPPKP = '" & Trim(Text(3)) & "', KODEUSER = '" & Trim(gSA) & "', TGLENTRY = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', NAMAFP = '" & Trim(Text1(1)) & "', ALAMATFP1 = '" & Trim(Text1(0)) & "', ALAMATFP2 = '" & Trim(Text1(5)) & "' WHERE NOFPS = '" & Trim(rsInvoice("NFP")) & "'")
                  'Call objHT.UpdateData_ACC("UPDATE FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN SET NOFPS = '" & Trim(Text(4)) & "', NPWP = '" & Trim(Text(2)) & "', NPPKP = '" & Trim(Text(3)) & "',  KODEENTRY = '" & Trim(gSA) & "', TGLENTRY = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', NAMAFP = '" & Trim(Text1(1)) & "', ALAMAT1 = '" & Trim(Text1(0)) & "', ALAMAT2 = '" & Trim(Text1(5)) & "' WHERE NOFPS = '" & Trim(rsInvoice("NFP")) & "'")
                  'Call objHT.UpdateData_ACC("UPDATE FAKTUR" & Right(gDb, 2) & "..INVOICE SET NOFPS = '" & Trim(Text(4)) & "' WHERE NOFPS = '" & Trim(rsInvoice("NFP")) & "'")
                  'Call objHT.UpdateData_ACC("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET NOFP = '" & Trim(Text(4)) & "' WHERE NOFP = '" & Trim(rsInvoice("NFP")) & "'")
                  
142               rsInvoice.MoveFirst
143               rsInvoice.Find "ST = true", , adSearchForward, 1
144               If Not rsInvoice.EOF Then
145                   Set rsSearch2 = objO.FetchData("SELECT NOINV FROM T_INVOICE WITH(NOLOCK) WHERE NOFP2 = '" & Trim(rsInvoice("NFP")) & "'")
146                   Set rsSearch = objO.FetchData("SELECT NOINV FROM T_INVOICE WITH(NOLOCK) WHERE NOFP = '" & Trim(rsInvoice("NFP")) & "'")
147                   If Not rsSearch.EOF Then
148                       Do While Not rsSearch.EOF
                              'Call objHT.UpdateData_BengkelOnly("UPDATE T_INVOICE SET NOFP = '" & Trim(Text(4)) & "' WHERE NOINV = '" & Trim(rsSearch("NOINV")) & "'")
149                           Call objO.UpdateData("UPDATE T_INVOICE SET NOFP = '" & Trim(text(4)) & "' WHERE NOINV = '" & Trim(rsSearch("NOINV")) & "'")
150                           rsSearch.MoveNext
151                       Loop
152                   End If
                       
153               End If
                  
                  ' UNTUK INVOICE -B
154               rsInvoice.MoveFirst
155               rsInvoice.Find "ST = true", , adSearchForward, 1
156               If Not rsInvoice.EOF Then
157                   If Not rsSearch2.EOF Then
158                       Do While Not rsSearch2.EOF
                              'Call objHT.UpdateData_BengkelOnly("UPDATE T_INVOICE SET NOFP2 = '" & Trim(Text(4)) & "' WHERE NOINV = '" & Trim(rsSearch2("NOINV")) & "'")
159                           Call objO.UpdateData("UPDATE T_INVOICE SET NOFP2 = '" & Trim(text(4)) & "' WHERE NOINV = '" & Trim(rsSearch2("NOINV")) & "'")
160                           rsSearch2.MoveNext
161                       Loop
162                   End If
163               End If
164           End If
              
              'Call objHT.UpdateData_BengkelOnly("UPDATE [" & NisDb & "]..SIHDR SET NOFPS = '" & Trim(Text(4)) & "' WHERE NOFAKTUR IN (SELECT NOFAKTUR FROM " & gDb & "..T_INVOICE WHERE NOINV IN (SELECT NOINV FROM T_INVOICE WHERE NOFP = '" & Trim(Text(4)) & "'))")
              'Call objHT.UpdateData_BengkelOnly("UPDATE [" & NisDb & "]..SIHDR SET NOFPS = '" & Trim(Text(4)) & "' WHERE NOFAKTUR IN (SELECT NOFAKTUR FROM " & gDb & "..T_INVOICE WHERE NOINV IN (SELECT NOINV FROM T_INVOICE WHERE NOFP2 = '" & Trim(Text(4)) & "'))")
              'Call objHT.UpdateData_BengkelOnly("UPDATE [" & NisDb & "]..GLTRN SET NOFPS = '" & Trim(Text(4)) & "' WHERE NOBUKTI IN ('" & Trim(NoInvoiceACC) & "') AND NOFPS = ''")

165           Call objO.UpdateData("UPDATE [" & NisDb & "]..SIHDR SET NOFPS = '" & Trim(text(4)) & "' WHERE NOFAKTUR IN (SELECT NOFAKTUR FROM " & gDb & "..T_INVOICE WITH(NOLOCK) WHERE NOINV IN (SELECT NOINV FROM T_INVOICE WITH(NOLOCK) WHERE NOFP = '" & Trim(text(4)) & "'))")
166           Call objO.UpdateData("UPDATE [" & NisDb & "]..SIHDR SET NOFPS = '" & Trim(text(4)) & "' WHERE NOFAKTUR IN (SELECT NOFAKTUR FROM " & gDb & "..T_INVOICE WITH(NOLOCK) WHERE NOINV IN (SELECT NOINV FROM T_INVOICE WITH(NOLOCK) WHERE NOFP2 = '" & Trim(text(4)) & "'))")
167           Call objO.UpdateData("UPDATE [" & NisDb & "]..GLTRN SET NOFPS = '" & Trim(text(4)) & "' WHERE NOBUKTI IN ('" & Trim(NoInvoiceACC) & "') AND NOFPS = ''")

                  
168           If erTranAll = False Then
169               Call objHT.CloseTransactionAll
170               onTR = False
171               Call saveLogMenu(gSA, ComptName, gVersi, "frmFakturPajak", "cmdCtkFaktur", 1)
172               text(4) = Trim(text(4))
173               MsgBox "No Faktur Pajak : " & Trim(text(4))
                  'Call cetakFaktur(Trim(Text(4)), 1)
174               If NoBaru Then
175                   If cekAdaPart(Trim(text(4))) Then
176                       chHarga = Unchecked
177                   End If
178               End If
179               noFakturCetak = Trim(text(4))
180               frmPrinter.Visible = True
181               cmdLanjut.SetFocus
182               Exit Sub
183           Else
184               Call objHT.CloseTransactionAll_Err
185               onTR = False
186               Call saveLogMenu(gSA, ComptName, gVersi, "frmFakturPajak", "cmdCtkFaktur", 2)
187               MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
188               Exit Sub
189           End If
190       Else
191           MsgBox "Lengkapi Data. " & vbCr & pesan
192       End If
193       Exit Sub
          
vbError:
194       Call objHT.CloseTransactionAll_Err
195       onTR = False
196       Call saveLogMenu(gSA, ComptName, gVersi, "frmFakturPajak", "cmdCtkFaktur", 2)
197       MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
198       Exit Sub

199      On Error GoTo 0
200      Exit Sub

cmdCtkFaktur_Click_Error:

201       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCtkFaktur_Click of Form frmFakturPajak")
End Sub

Private Sub cetakFaktur(noFPcetak As String, idx As Integer, idxPrinter As Integer)
      Dim sHlp2 As Object
      Dim rsFaktur As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset

1        On Error GoTo cetakFaktur_Error

2             Select Case idxPrinter
              Case 0
                  ' LQ 2180
3                 Set rpt = New ReportFakturPajak4LQ2180
4             Case 1
                  ' LQ 300
5                 Set rpt = New ReportFakturPajak4LQ300
6             Case 2
                  ' LX
7                 Set rpt = New ReportFakturPajak4LX300
8             Case 3
                  ' LQ 2190
9                 Set rpt = New ReportFakturPajak4LQ2190
10            End Select
              
              ' 1 -> NEW FAKTUR
              ' 2 -> CETAK FAKTUR
              
11            If idx = 2 Then
12                chInPPn = Unchecked
13                chSilang = Checked
14            End If
              
              'sSQL = "SELECT '" & IIf(chInPPn = Checked, 1, 0) & "' AS STATUS, * FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WHERE NOFPS = '" & Trim(noFPcetak) & "'"
15            sSQL = "SELECT '" & IIf(chInPPn = Checked, 1, 0) & "' AS STATUS, NoFPS,TglFPS,NPWP,NPPKP,Keterangan,(HargaJual+PPN) as Hargajual,PPN,PPnBM,KodeUser,TglEntry,NamaFP,AlamatFP1,AlamatFP2,Modul, " & _
                     "round(case when '" & IIf(chInPPn = Checked, 1, 0) & "' = 0 then hargajual else hargajuaL+PPN end,0) as dpp, " & _
                     "round(case when '" & IIf(chInPPn = Checked, 1, 0) & "' = 0 then (case when '" & IIf(chInPPn = Checked, 1, 0) & "' = 0 then hargajual else hargajual+PPN end) else (case when '" & IIf(chInPPn = Checked, 1, 0) & "' = 0 then hargajual  else hargajual+PPN end) end,0) as dpp2 " & _
                     "FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH(NOLOCK) WHERE NOFPS = '" & Trim(noFPcetak) & "'"
16            Set rsRepo = objO.FetchDataACC(sSQL)
17            rpt.database.SetDataSource rsRepo
              
18            If idx = 2 Then
19                Set rsFaktur = objO.FetchDataACC(sSQL)
20                If Not rsFaktur.EOF Then
21                    tglFaktur = Format(rsFaktur("TGLFPS"), "DD/MM/YYYY")
22                End If
23                If rsFaktur.State <> adStateClosed Then rsFaktur.Close
24            End If
              
25            rpt.txtAL.SetText (rsRepo.Fields(12).value)
26            rpt.TxtNamaPT.SetText (namaPT)
27            rpt.TxtAlamatPT.SetText (alamatPT)
28            rpt.txtNPWPPT.SetText (npwpPT)
29            rpt.txtKotaTgl.SetText (Trim(gKotaFaktur) & ", " & TglDDMMMMYYYY(Format(tglFaktur, "DD/MM/YYYY"), 1))
30            rpt.TxtBag.SetText IIf(chJabatan = Unchecked, "", Trim(text(8)))
31            rpt.TxtNamaTtd.SetText Trim(text(6))
32            If chSilang = Unchecked Then
33                rpt.Text10.SetText ""
34            End If
35            If chHarga = Unchecked Then
36                rpt.Text1.SetText ""
37            End If
38            If chGanti = Unchecked Then
39                rpt.Text2.SetText ""
40            End If
              
41            With CRViewer3
42                .ReportSource = rpt
43                .ViewReport
44                .Zoom 100
45                .Top = 0
46                .Left = 0
47                .Width = GridProgram.Width + 100
48                .Height = cmdCreateFP.Top + cmdCreateFP.Height
                  
49                cmdTutup3.Top = .Top + 430
50                cmdTutup3.Left = .Left
51                cmdTutup3.Visible = True
                  
52                cmdSettingPrinter.Top = cmdTutup3.Top
53                cmdSettingPrinter.Left = cmdTutup3.Left + cmdTutup3.Width + 50
54                cmdSettingPrinter.Visible = True
55                .Visible = True
56            End With
57            cmdTutup3.Visible = True
58            Screen.MousePointer = vbDefault
              
              Dim rsHlp As ADODB.Recordset
59            Set rsHlp = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN WITH(NOLOCK) WHERE NOFPS = '" & Trim(noFPcetak) & "'")
60            If Not rsHlp.EOF Then
61                sSQL = "SELECT * FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN WITH(NOLOCK) WHERE NOFPS = '" & Trim(noFPcetak) & "'"
62                Set rsSearch = objO.FetchDataACC(sSQL)
63            Else
64                sSQL = "SELECT * FROM T_INVOICE WITH(NOLOCK) WHERE NOFP = '" & Trim(noFPcetak) & "' OR NOFP2 = '" & Trim(noFPcetak) & "'"
65                Set rsSearch = objO.FetchData(sSQL)
66            End If
              'Set rsSearch = objO.FetchData(sSQL)
67            If Not rsSearch.EOF Then
68                If rsSearch.RecordCount > 1 Then
69                    Call cetakFakturLampiran(noFPcetak)
70                End If
71            End If

72       On Error GoTo 0
73       Exit Sub

cetakFaktur_Error:

74        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakFaktur of Form frmFakturPajak")
               
End Sub

Private Sub cetakFakturLampiran(noFPcetak As String)
      Dim sHlp2 As Object
      Dim rsHlp As ADODB.Recordset
1        On Error GoTo cetakFakturLampiran_Error

2     Set rpt2 = New ReportFakturPajakLamp

3             On Error Resume Next
              
4             Set rsHlp = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN WITH(NOLOCK) WHERE NOFPS = '" & Trim(noFPcetak) & "'")
5             If Not rsHlp.EOF Then
6                 sSQL = "SELECT TGLINVOICE,  (SELECT NAMAFP FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH(NOLOCK) WHERE NOFPS = A.NOFPS) as NAMAFP, NOINVOICE, KET1, DPP,PPN, NOFPS, KET2 AS NOD FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN a WITH(NOLOCK) WHERE NOFPS = '" & Trim(noFPcetak) & "'"
7                 Set sHlp2 = objO.FetchDataACC(sSQL)
8             Else
9                 Call setCetakLampiran(1, "CETAKLAMP", Trim(noFPcetak))
10                sSQL = "SELECT INVDATE AS TGLINVOICE, (SELECT NAMAFP FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH(NOLOCK) WHERE NOFPS = A.NOFP) AS NAMAFP, '" & gKodeFaktur & "'+NOINV AS NOINVOICE, NO_POLISI AS KET1, ROUND(TOTAL*100/110,0) AS DPP, ROUND(TOTAL*10/110,0) AS PPN, NOFP AS NOFPS, NOFAKTUR AS NOD FROM [CETAKLAMP] a WITH(NOLOCK) WHERE NOFP = '" & Trim(noFPcetak) & "'"
11                Set sHlp2 = objO.FetchDataACC(sSQL)
12            End If
              'Set sHlp2 = objO.FetchData(sSQL)
13            rpt2.database.SetDataSource sHlp2
              
14            With CRViewer4
15                .ReportSource = rpt2
16                .ViewReport
17                .Zoom 100
18                .Top = 0
19                .Left = 0
20                .Width = GridProgram.Width + 100
21                .Height = cmdCreateFP.Top + cmdCreateFP.Height
                  
22                cmdTutup4.Top = .Top + 430
23                cmdTutup4.Left = .Left
24                cmdTutup4.Visible = True
                  
25                cmdSettingPrinter2.Top = cmdTutup4.Top
26                cmdSettingPrinter2.Left = cmdTutup4.Left + cmdTutup4.Width + 50
27                cmdSettingPrinter2.Visible = True
28                .Visible = True
29            End With
30            cmdTutup4.Visible = True
31            Screen.MousePointer = vbDefault

32       On Error GoTo 0
33       Exit Sub

cetakFakturLampiran_Error:

34        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakFakturLampiran of Form frmFakturPajak")
End Sub

Private Sub cmdLanjut_Click()
    Call cetakFaktur(Trim(noFakturCetak), IIf(frmFaktur.Visible, 1, 2), IIf(prtrn(0).value, 0, IIf(prtrn(1).value, 1, IIf(prtrn(2).value, 2, 3))))
    frmPrinter.Visible = False
    If frmFaktur.Visible = True Then
        frmFaktur.Visible = False
    End If
End Sub

Private Sub cmdRefresh_Click()
    Call refreshList(Trim(cboNoFP))
End Sub

Private Sub cmdReturCetak_Click()
          Dim rsRetur As ADODB.Recordset
1        On Error GoTo cmdReturCetak_Click_Error

2         Set rsRetur = objO.FetchData("SELECT NOINV, CUSTOMER, CUST_ADRES, KOTA, LABOR AS JASA, PART_G+PART_NG+OIL_G+OIL_NG AS PART, SBLT AS SUBLET, TOTAL, B.NOFPS AS NOFAKTUR, C.HARGAJUAL, C.PPN, INVDATE as Tanggal " & _
                                       "FROM " & Log.LogBengkel & "..T_INVOICE A WITH(NOLOCK) LEFT JOIN OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogAccServ & ";Uid=" & Log.LogAccUser & ";PWD=" & Log.LogAccPass & "',[FAKTUR" & Right(gDb, 2) & "].dbo.INVOICE) B ON SUBSTRING(B.NOINVOICE,3,10) = A.NOINV LEFT JOIN OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogAccServ & ";Uid=" & Log.LogAccUser & ";PWD=" & Log.LogAccPass & "',[FAKTUR" & Right(gDb, 2) & "].dbo.SIFPSTANDAR) C ON B.NOFPS = C.NOFPS " & _
                                       "WHERE FD_RETUR BETWEEN '" & Format(DtDOAwal, "MM/DD/YYYY") & "' AND '" & Format(DtDOAkhir, "MM/DD/YYYY") & "'")
                                       'by FD_RETUR atau INVDATE
          'OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',[" & Log.LogHODbSPK & "]
          
3         If rsRetur.RecordCount > 0 Then
4             Frame1.Visible = True
5             GridR.DataSource = rsRetur
6         Else
7             MsgBox "Tidak ada Data Retur !"
8         End If

9        On Error GoTo 0
10       Exit Sub

cmdReturCetak_Click_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdReturCetak_Click of Form frmFakturPajak")
End Sub

Private Sub cmdSettingPrinter_Click()
    rpt.PrinterSetup Me.hWnd
    CRViewer3.Refresh
End Sub

Private Sub cmdSettingPrinter2_Click()
    rpt2.PrinterSetup Me.hWnd
    CRViewer4.Refresh
End Sub

Private Sub cmdTtd_Click()
    frmtandaTangan.Show
End Sub

Private Sub cmdTutup3_Click()
    CRViewer3.Visible = False
    cmdTutup3.Visible = False
    cmdSettingPrinter.Visible = False
    Call tampilGrid(IIf(cboFaktur = "Faktur Pajak", 1, 2), cboNoFP)
End Sub

Private Sub cmdTutup4_Click()
    CRViewer4.Visible = False
    cmdTutup4.Visible = False
    cmdSettingPrinter2.Visible = False
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

Private Sub BackUpMyBase(onVal As String)
    
    On Error GoTo errHandler
    GridProgram.ExportToFile onVal, True
    MsgBox "Data sudah di Export ke File Excel." & vbCr & "File : " & onVal
    Exit Sub
    
errHandler:
    Exit Sub
    err.Raise err.Number, err.Source, err.Description
End Sub


Private Sub BackUpMyBaseR(onVal As String)
    
    On Error GoTo errHandler
    GridR.ExportToFile onVal, True
    MsgBox "Data sudah di Export ke File Excel." & vbCr & "File : " & onVal
    Exit Sub
    
errHandler:
    Exit Sub
    err.Raise err.Number, err.Source, err.Description
End Sub

Public Sub NoFaktur(idxx As Integer)
Dim rsSearch As New ADODB.Recordset

     
    Select Case idxx
        Case 1
            sSQL = "SELECT DISTINCT NOFPS AS NO FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH(NOLOCK) WHERE MODUL = 'SVC' ORDER BY NOFPS"
        Case 2
            sSQL = "SELECT DISTINCT NOFPS AS NO FROM FAKTUR" & Right(gDb, 2) & "..FPGABUNGAN WITH(NOLOCK) WHERE MODUL = 'SVC' ORDER BY NOFPS"
        Case 3
            sSQL = "SELECT DISTINCT NOMB AS NO FROM T_INVOICE WITH(NOLOCK) ORDER BY NOMB"
    End Select
    Set rsSearch = objO.FetchDataACC(sSQL)
    
    If Not rsSearch.EOF Then
        With cboNoFP
            .Clear
            Do While Not rsSearch.EOF
                If Trim(rsSearch("NO")) <> "" Then
                    .AddItem Trim(rsSearch("NO"))
                End If
                rsSearch.MoveNext
            Loop
        End With
    End If
End Sub


Private Sub Command1_Click()
    Call TblKeluar_Click
End Sub

Private Sub Command2_Click()
    Frame1.Visible = False
End Sub

Private Sub Command3_Click()
    Dim f As New FormExplorer
    f.Show (vbModal)
    Dim MyPath As String
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub

    Call BackUpMyBaseR(MyPath & f.FileName & ".XLS")

End Sub

Private Sub Form_Load()
Dim rsSearch As ADODB.Recordset
    Frame1.Left = 0
    DtDOAwal = Format(isnow, "DD/MM/YYYY")
    DtDOAkhir = Format(isnow, "DD/MM/YYYY")
    With cboFaktur
        .Clear
        .AddItem ("Faktur Pajak")
        .AddItem ("Faktur Pajak Gabungan")
    End With
    If rsInvoice.State <> adStateClosed Then rsInvoice.Close
    Set rsInvoice = Nothing
    
     
    Call setNoFaktur(cboType)
    With cboTtd
        .Clear
        .AddItem ("KACAB")
        .AddItem ("ADH")
        .AddItem ("KABENG")
    End With
    
    Set rsSearch = objO.FetchData("SELECT * FROM SETTANDA WITH(NOLOCK) ")
    If Not rsSearch.EOF Then
        cboTtd = Trim(rsSearch("OTORITAS"))
        chJabatan = IIf(Trim(rsSearch("DISPLAY")) = 0, Unchecked, Checked)
        Select Case CDbl(rsSearch("TPPRINT"))
            Case 0
                prtrn(0).value = True
            Case 1
                prtrn(1).value = True
            Case 2
                prtrn(2).value = True
            Case 3
                prtrn(3).value = True
        End Select
    Else
        cboTtd = "KABENG"
        chJabatan = Checked
    End If
     
End Sub

Private Sub GridProgram_AfterColUpdate(ByVal ColIndex As Integer)

    With GridProgram
        Select Case ColIndex
            Case 0
                If .Columns(ColIndex) Then
                    nTagih = nTagih + 1
                    totTagih = totTagih + (CDbl(.Columns(9)))
                Else
                    nTagih = nTagih - 1
                    totTagih = totTagih - (CDbl(.Columns(9)))
                End If
        End Select
        nTagih = Format(CDbl(nTagih), "##,0")
        totTagih = Format(CDbl(totTagih), "##,0")
    End With
End Sub

Private Sub GridProgram_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    
    Dim rs As New ADODB.Recordset
    Set rs = rsInvoice.Clone
    If Bookmark > 1 Then rs.Bookmark = Bookmark
    If Nul(rs.Fields(12).value) = "1" Then RowStyle = "invoiceRetur"
    
End Sub

Private Sub GridProgram_FilterChange()
On Error GoTo errHandler
    Set cols = GridProgram.Columns
    Dim C As Integer
    
    C = GridProgram.col
    GridProgram.HoldFields
    rsInvoice.filter = getFilter()
    If rsInvoice.EOF Then
        For Each col In GridProgram.Columns
            col.FilterText = ""
        Next col
        rsInvoice.filter = ""
    End If
    GridProgram.col = C
    GridProgram.EditActive = True
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
Private Sub TblKeluar_Click()

    If Trim(vNamaTabel) <> "" Then
        sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
             & " drop table [tempdb]..[" & vNamaTabel & "]"
        Call objO.UpdateData(sSQL)
    End If
    Unload Me
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
            text(Index) = gMainKeyCari
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
                gMainKeyCari = text(Index)
                gSearchOption = "NPWP"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 2
            If Trim(text(Index)) <> "" Then
                Set rsSearch = objO.FetchDataACC("SELECT * FROM FAKTUR" & Right(gDb, 2) & "..SIFPSTANDAR WITH(NOLOCK) WHERE NPWP = '" & Trim(text(Index)) & "'")
                If Not rsSearch.EOF Then
                    If CDbl(Replace(Replace(Trim(text(Index)), "-", ""), ".", "")) <> 0 And Trim(text(Index)) <> "" Then
                        Text1(1) = Trim(rsSearch("NAMAFP"))
                        Text1(0) = Trim(rsSearch("ALAMATFP1"))
                        text(3) = Trim(rsSearch("NPPKP"))
                    End If
                End If
            End If
    End Select
     
End Sub

Private Function statusDealer(codeStatus As Integer)
    
    Select Case codeStatus
        Case 1
            statusDealer = "AFFNONPPN"
        Case 2
            statusDealer = "AFFPPN"
        Case 3
            statusDealer = "EXT"
    End Select
End Function

Private Function cekDuo(hrg As Double, nINV As String, Optional stat As Integer)
    
    If Right(Trim(nINV), 2) = "-B" Then
        cekDuo = 0
        If stat = 1 Then
            cekDuo = CDbl(hrg)
        End If
    Else
        cekDuo = CDbl(hrg)
        If stat = 2 Then
            cekDuo = 0
        End If
    End If
End Function

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub tglFaktur_Change()
    Call setNoFaktur(cboType.text)
End Sub
