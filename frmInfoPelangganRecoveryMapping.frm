VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmInfoPelangganRecoveryMapping 
   BackColor       =   &H00FFFFFF&
   Caption         =   ".:: Data Recovery Mapping ::."
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdRefreshLap 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Refresh"
      Height          =   735
      Left            =   9930
      Picture         =   "frmInfoPelangganRecoveryMapping.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7590
      Width           =   735
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
      Left            =   11010
      Picture         =   "frmInfoPelangganRecoveryMapping.frx":0664
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7590
      Width           =   735
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
      Left            =   9210
      Picture         =   "frmInfoPelangganRecoveryMapping.frx":0DF4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7590
      Width           =   735
   End
   Begin VB.CommandButton cmdCall 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Call"
      Height          =   735
      Left            =   7890
      Picture         =   "frmInfoPelangganRecoveryMapping.frx":1270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7635
      Width           =   975
   End
   Begin MSComDlg.CommonDialog ComDialogXls 
      Left            =   7020
      Top             =   7860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid GridPelanggan 
      Height          =   6615
      Left            =   330
      TabIndex        =   4
      Top             =   630
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   11668
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NO ORDER"
      Columns(0).DataField=   "NO_WO"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NO RANGKA"
      Columns(1).DataField=   "NORANGKA"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "PESAN LAST REMINDER"
      Columns(2).DataField=   "PESANLASTREMINDER"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "TGL REMINDER"
      Columns(3).DataField=   "TGLREMINDER"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "KM NEXT"
      Columns(4).DataField=   "KMNEXT"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "MODEL KENDARAAN"
      Columns(5).DataField=   "MDLKENDARAAN"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "TGL NEXT SERVICE"
      Columns(6).DataField=   "TGLNEXTSERVICE"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "NO ENGINE"
      Columns(7).DataField=   "NOENGINE"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "NAMA"
      Columns(8).DataField=   "NAMA"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "ALAMAT"
      Columns(9).DataField=   "ALAMAT"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "KOTA"
      Columns(10).DataField=   "KOTA"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "KODEPOS"
      Columns(11).DataField=   "KODEPOS"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "NO POLISI"
      Columns(12).DataField=   "NOPOL"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "TELP"
      Columns(13).DataField=   "TELP"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "GSM"
      Columns(14).DataField=   "GSM"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "CDMA"
      Columns(15).DataField=   "CDMA"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   16
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=3281"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(21)=   "Column(4).Visible=0"
      Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(23)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(27)=   "Column(6).Width=3281"
      Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=3175"
      Splits(0)._ColumnProps(30)=   "Column(6).Visible=0"
      Splits(0)._ColumnProps(31)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(32)=   "Column(7).Width=3281"
      Splits(0)._ColumnProps(33)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(7)._WidthInPix=3175"
      Splits(0)._ColumnProps(35)=   "Column(7).Visible=0"
      Splits(0)._ColumnProps(36)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(37)=   "Column(8).Width=3281"
      Splits(0)._ColumnProps(38)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(40)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(41)=   "Column(9).Width=3281"
      Splits(0)._ColumnProps(42)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(9)._WidthInPix=3175"
      Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(45)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(48)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(49)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(50)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(52)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(53)=   "Column(12).Width=3281"
      Splits(0)._ColumnProps(54)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(12)._WidthInPix=3175"
      Splits(0)._ColumnProps(56)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(57)=   "Column(13).Width=3281"
      Splits(0)._ColumnProps(58)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(13)._WidthInPix=3175"
      Splits(0)._ColumnProps(60)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(61)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(62)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(64)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(65)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(15).Order=16"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=98,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=94,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=91,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=92,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=93,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=90,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=87,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=88,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=89,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=86,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=78,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=74,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=54,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=50,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=47,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=48,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=49,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=46,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=43,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=44,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=45,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=32,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=29,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=30,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=31,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=28,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=25,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=26,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=27,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=58,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=55,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=56,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=57,.parent=17"
      _StyleDefs(100) =   "Named:id=33:Normal"
      _StyleDefs(101) =   ":id=33,.parent=0"
      _StyleDefs(102) =   "Named:id=34:Heading"
      _StyleDefs(103) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(104) =   ":id=34,.wraptext=-1"
      _StyleDefs(105) =   "Named:id=35:Footing"
      _StyleDefs(106) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(107) =   "Named:id=36:Selected"
      _StyleDefs(108) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(109) =   "Named:id=37:Caption"
      _StyleDefs(110) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(111) =   "Named:id=38:HighlightRow"
      _StyleDefs(112) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(113) =   "Named:id=39:EvenRow"
      _StyleDefs(114) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(115) =   "Named:id=40:OddRow"
      _StyleDefs(116) =   ":id=40,.parent=33"
      _StyleDefs(117) =   "Named:id=41:RecordSelector"
      _StyleDefs(118) =   ":id=41,.parent=34"
      _StyleDefs(119) =   "Named:id=42:FilterBar"
      _StyleDefs(120) =   ":id=42,.parent=33"
   End
   Begin VB.Label lblJml 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11130
      TabIndex        =   7
      Top             =   270
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   10890
      Shape           =   4  'Rounded Rectangle
      Top             =   7470
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LIST RECOVERY REMINDER"
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
      Left            =   1410
      TabIndex        =   6
      Top             =   270
      Width           =   9735
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      X1              =   570
      X2              =   4290
      Y1              =   510
      Y2              =   510
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      X1              =   1770
      X2              =   4290
      Y1              =   390
      Y2              =   390
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      X1              =   2610
      X2              =   4290
      Y1              =   270
      Y2              =   270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   330
      X2              =   11970
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   9090
      Shape           =   4  'Rounded Rectangle
      Top             =   7470
      Width           =   1695
   End
   Begin VB.Label Label70 
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Data  :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   9690
      TabIndex        =   5
      Top             =   270
      Width           =   1455
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   7770
      Shape           =   4  'Rounded Rectangle
      Top             =   7500
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   330
      Shape           =   4  'Rounded Rectangle
      Top             =   180
      Width           =   11715
   End
End
Attribute VB_Name = "frmInfoPelangganRecoveryMapping"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private OnLoad As Boolean

Private objPel As New ClsDataPelanggan
Private rsPel As ADODB.Recordset
Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String
Private myChassis As String

 

Private Sub cmdCall_Click()
    If MsgBox("Apakah ingin memproses reminder untuk No Rangka " & GridPelanggan.Columns(1).value & "?", vbYesNo) = vbYes Then
        Dim fCL As New frmInfoPelangganRecoveryMappingCall
        With fCL
           '.noWO = Trim(GridPelanggan.Columns(0).value)
           .IsSvc1000 = False
            
            '.txttglNextSvc.text = Trim(GridPelanggan.Columns(6).value)
            '.txtKmNextSvc.text = Trim(GridPelanggan.Columns(4).value)
            .txtNoChassis.text = Trim(GridPelanggan.Columns(1).value)
            .txtNoEngine.text = Trim(GridPelanggan.Columns(7).value)
            .txtNama.text = Trim(GridPelanggan.Columns(8).value)
            .txtAlamat.text = Trim(GridPelanggan.Columns(9).value)
            .txtKota.text = Trim(GridPelanggan.Columns(10).value)
            .txtPos.text = Trim(GridPelanggan.Columns(11).value)
            .txtNoPol.text = Trim(GridPelanggan.Columns(12).value)
            .txtTelp.text = Trim(GridPelanggan.Columns(13).value)
            .txtGSM.text = Trim(GridPelanggan.Columns(14).value)
            .txtCDMA.text = Trim(GridPelanggan.Columns(15).value)
           
           '.Left = Me.Left
           '.Top = Me.Top
           .Show vbModal
           'Unload Me
           Call FillGridPelanggan
        End With
    End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub
 

Private Sub CmdRefreshLap_Click()
   Call FillGridPelanggan
End Sub
 

Private Sub CmdXls_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsPel.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Pelanggan Kosong !"
5             Beep
6             Er.Show (vbModal)
7             Exit Sub
8         End If
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.Show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".XLS"
16         GridPelanggan.ExportToFile onVal, True
           
17                Er.HeadCap = "Export to excel !"
18                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
19                Beep
20                Er.Show (vbModal)
21         Exit Sub
Err_Export:
22          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
23          MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdXls of Form FormRecoveryReminderMRA")
Keluar:
End Sub

 

Private Sub Form_Load()
   Call FillGridPelanggan
End Sub


Private Sub FillGridPelanggan()
   On Error GoTo FillGridPelanggan_Error
      sSQL = "select NO_WO,NORANGKA,PESANLASTREMINDER,CONVERT(VARCHAR(20),TGLREMINDER,103) TGLREMINDER,KMNEXT,MDLKENDARAAN,"
      sSQL = sSQL & " CONVERT(VARCHAR(20),TGLNEXTSERVICE,103) TGLNEXTSERVICE,NOENGINE,NAMA,ALAMAT,KOTA,KODEPOS,NOPOL,TELP,GSM,CDMA"
      sSQL = sSQL & " FROM recoverymapping WITH (NOLOCK) WHERE ISVIEW = '0'"
      Set rsPel = objO.FetchData(sSQL)
      
      'Set rsPel = objPel.FetchDataPelanggan
      If rsPel.RecordCount = 0 Then Exit Sub
      Set Me.GridPelanggan.DataSource = rsPel
      
      lblJml.Caption = rsPel.RecordCount
 
   On Error GoTo 0
   Exit Sub

FillGridPelanggan_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure FillGridPelanggan of Form FormRecoveryReminderMRA")
End Sub

Private Sub GridPelanggan_FilterChange()
On Error GoTo err
   Set cols = GridPelanggan.Columns
   Dim C As Integer
   
   C = GridPelanggan.col
   GridPelanggan.HoldFields
   rsPel.filter = getFilter()
   If rsPel.EOF = True Then
      For Each col In GridPelanggan.Columns
         col.FilterText = ""
      Next col
      rsPel.filter = ""
   End If
   GridPelanggan.col = C
   GridPelanggan.EditActive = True
   lblJml.Caption = rsPel.RecordCount
   Exit Sub
err:
   Resume Next
End Sub

Private Function getFilter() As String
   On Error GoTo err

On Error GoTo err
Dim tmp As String
Dim n As Integer

   For Each col In cols
      If Trim(col.FilterText) <> "" Then
         n = n + 1
         If n > 1 Then
            tmp = tmp & " AND "
         End If
         Select Case Trim(col.DataField)
            Case "JMLHISTORY", "LASTSVC", "KMAKHIR"
                tmp = tmp & col.DataField & " >= '" & CDbl(col.FilterText) & "'"
            Case Else
                tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
         End Select
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub GridPelanggan_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 23 Then rsPel.Sort = GridPelanggan.Columns(ColIndex).DataField & " " & Urut
4        GridPelanggan.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        lblJml.Caption = rsPel.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form FormRecoveryReminderMRA")
End Sub

Private Sub GridPelanggan_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    myChassis = GridPelanggan.Columns(10).value
'    chGSM.Caption = GridPelanggan.Columns(6).value
'    chCDMA.Caption = GridPelanggan.Columns(7).value
End Sub


