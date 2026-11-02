VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmInfoRetensiUnitRetail 
   Caption         =   "Form1"
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   13200
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmInfoRetensiUnitRetail.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   10
      Top             =   120
      Width           =   450
   End
   Begin VB.CommandButton cmdSearch2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Searc&h"
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
      Left            =   4440
      Picture         =   "frmInfoRetensiUnitRetail.frx":8DF3
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6960
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
      Left            =   10320
      Picture         =   "frmInfoRetensiUnitRetail.frx":9442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6960
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
      Left            =   11280
      Picture         =   "frmInfoRetensiUnitRetail.frx":98BE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6960
      Width           =   735
   End
   Begin TrueOleDBGrid70.TDBGrid grid 
      Height          =   6135
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   10821
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No Rangka"
      Columns(0).DataField=   "norangka"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nama Pemesan"
      Columns(1).DataField=   "nama_pemesan"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Kode Lgn"
      Columns(2).DataField=   "kode_lgn"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Model"
      Columns(3).DataField=   "model"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Model"
      Columns(4).DataField=   "namamodel"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "No Polisi"
      Columns(5).DataField=   "Nopolisi"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Thn Mobil"
      Columns(6).DataField=   "thmobil"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "DDate"
      Columns(7).DataField=   "Ddate"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Nama Sales"
      Columns(8).DataField=   "namasales"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Nama Supervisor"
      Columns(9).DataField=   "Nama_Supervisor"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Tanggal"
      Columns(10).DataField=   "Tanggal"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Alamat"
      Columns(11).DataField=   "alamat"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Kelurahan"
      Columns(12).DataField=   "kelurahan"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Kecamatan"
      Columns(13).DataField=   "Kecamatan"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "Kota"
      Columns(14).DataField=   "Kota"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "Propinsi"
      Columns(15).DataField=   "propinsi"
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).Caption=   "Kodepos"
      Columns(16).DataField=   "kodepos"
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).Caption=   "Telp"
      Columns(17).DataField=   "Telp"
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).Caption=   "HP1"
      Columns(18).DataField=   "HP1"
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).Caption=   "HP2"
      Columns(19).DataField=   "HP2"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).Caption=   "Usename"
      Columns(20).DataField=   "Usename"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).Caption=   "UseGSM"
      Columns(21).DataField=   "UseGSM"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).Caption=   "Last SB"
      Columns(22).DataField=   "LastSB"
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).Caption=   "Job Last SB"
      Columns(23).DataField=   "joblastsb"
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).Caption=   "Dealer Last SB"
      Columns(24).DataField=   "DealerLastSB"
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).Caption=   "Last GR"
      Columns(25).DataField=   "LastGR"
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).Caption=   "Job Last GR"
      Columns(26).DataField=   "JobLastGR"
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).Caption=   "Dealer Last GR"
      Columns(27).DataField=   "DealerLastGR"
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).Caption=   "Last SVC"
      Columns(28).DataField=   "LastSVC"
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).Caption=   "Dealer Last SVC"
      Columns(29).DataField=   "DealerLastSVC"
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   30
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=30"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2831"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2725"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=3969"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3863"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=3545"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3440"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=3281"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=1773"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1667"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=1773"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=1667"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=1773"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=1667"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=5292"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=5186"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=10583"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=10478"
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
      Splits(0)._ColumnProps(73)=   "Column(18).Width=3281"
      Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=3175"
      Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(77)=   "Column(19).Width=3281"
      Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=3175"
      Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(81)=   "Column(20).Width=3281"
      Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=3175"
      Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(85)=   "Column(21).Width=3281"
      Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=3175"
      Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(89)=   "Column(22).Width=3281"
      Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=3175"
      Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(93)=   "Column(23).Width=3281"
      Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=3175"
      Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(97)=   "Column(24).Width=3281"
      Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=3175"
      Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(101)=   "Column(25).Width=3281"
      Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=3175"
      Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(105)=   "Column(26).Width=3281"
      Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=3175"
      Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(109)=   "Column(27).Width=3281"
      Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=3175"
      Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(113)=   "Column(28).Width=3281"
      Splits(0)._ColumnProps(114)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(28)._WidthInPix=3175"
      Splits(0)._ColumnProps(116)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(117)=   "Column(29).Width=3281"
      Splits(0)._ColumnProps(118)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(29)._WidthInPix=3175"
      Splits(0)._ColumnProps(120)=   "Column(29).Order=30"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=118,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=115,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=116,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=117,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=114,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=111,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=112,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=113,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=110,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=107,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=108,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=109,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=32,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=62,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=66,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=70,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=67,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=68,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=69,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=74,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=78,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=75,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=76,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=77,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=82,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=79,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=80,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=81,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=86,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=83,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=84,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=85,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=90,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=87,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=88,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=89,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=94,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=91,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=92,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=93,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=98,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=95,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=96,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=97,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=102,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=99,.parent=14"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=100,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=101,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=106,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=103,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=104,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=105,.parent=17"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=122,.parent=13"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=119,.parent=14"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=120,.parent=15"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=121,.parent=17"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=126,.parent=13"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=123,.parent=14"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=124,.parent=15"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=125,.parent=17"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=130,.parent=13"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=127,.parent=14"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=128,.parent=15"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=129,.parent=17"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=134,.parent=13"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=131,.parent=14"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=132,.parent=15"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=133,.parent=17"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=138,.parent=13"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=135,.parent=14"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=136,.parent=15"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=137,.parent=17"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=142,.parent=13"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=139,.parent=14"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=140,.parent=15"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=141,.parent=17"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=146,.parent=13"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=143,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=144,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=145,.parent=17"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=150,.parent=13"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=147,.parent=14"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=148,.parent=15"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=149,.parent=17"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=154,.parent=13"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=151,.parent=14"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=152,.parent=15"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=153,.parent=17"
      _StyleDefs(156) =   "Named:id=33:Normal"
      _StyleDefs(157) =   ":id=33,.parent=0"
      _StyleDefs(158) =   "Named:id=34:Heading"
      _StyleDefs(159) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(160) =   ":id=34,.wraptext=-1"
      _StyleDefs(161) =   "Named:id=35:Footing"
      _StyleDefs(162) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(163) =   "Named:id=36:Selected"
      _StyleDefs(164) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(165) =   "Named:id=37:Caption"
      _StyleDefs(166) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(167) =   "Named:id=38:HighlightRow"
      _StyleDefs(168) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(169) =   "Named:id=39:EvenRow"
      _StyleDefs(170) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(171) =   "Named:id=40:OddRow"
      _StyleDefs(172) =   ":id=40,.parent=33"
      _StyleDefs(173) =   "Named:id=41:RecordSelector"
      _StyleDefs(174) =   ":id=41,.parent=34"
      _StyleDefs(175) =   "Named:id=42:FilterBar"
      _StyleDefs(176) =   ":id=42,.parent=33"
   End
   Begin TDBDate6Ctl.TDBDate Dt1 
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   7320
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      Calendar        =   "frmInfoRetensiUnitRetail.frx":A04E
      Caption         =   "frmInfoRetensiUnitRetail.frx":A17A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInfoRetensiUnitRetail.frx":A1E6
      Keys            =   "frmInfoRetensiUnitRetail.frx":A204
      Spin            =   "frmInfoRetensiUnitRetail.frx":A262
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
   Begin TDBDate6Ctl.TDBDate Dt2 
      Height          =   315
      Left            =   2520
      TabIndex        =   5
      Top             =   7320
      Width           =   1575
      _Version        =   65536
      _ExtentX        =   2778
      _ExtentY        =   556
      Calendar        =   "frmInfoRetensiUnitRetail.frx":A28A
      Caption         =   "frmInfoRetensiUnitRetail.frx":A3B6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmInfoRetensiUnitRetail.frx":A422
      Keys            =   "frmInfoRetensiUnitRetail.frx":A440
      Spin            =   "frmInfoRetensiUnitRetail.frx":A49E
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
   Begin VB.Label JUMLAH 
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
      Left            =   11160
      TabIndex        =   9
      Top             =   120
      Width           =   975
   End
   Begin VB.Shape Shape15 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   4320
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   975
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   4200
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line3 
      BorderWidth     =   2
      X1              =   2280
      X2              =   2400
      Y1              =   7440
      Y2              =   7440
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   10200
      Shape           =   4  'Rounded Rectangle
      Top             =   6840
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   4420
      X2              =   1080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DATA PENJUALAN"
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
      Left            =   1560
      TabIndex        =   8
      Top             =   120
      Width           =   9375
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
   Begin VB.Line Line8 
      X1              =   600
      X2              =   4080
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Periode Retail :"
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
      Left            =   600
      TabIndex        =   7
      Top             =   6855
      Width           =   3495
   End
   Begin VB.Line Line9 
      BorderColor     =   &H000000FF&
      X1              =   480
      X2              =   12120
      Y1              =   6720
      Y2              =   6720
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
      Left            =   9720
      TabIndex        =   6
      Top             =   120
      Width           =   1455
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
Attribute VB_Name = "frmInfoRetensiUnitRetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmInfoRetensiUnitRetail
' DateTime  : 26/11/2020 15:46
' Author    : Rudi Kristanto
' Purpose   :
'---------------------------------------------------------------------------------------

Option Explicit

Private OnLoad As Boolean

Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String
Private rsCR As ADODB.Recordset
Private objCR As New ClsDataPelanggan
Private kriteria As String

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSearch2_Click()
    grid.DataSource = Nothing
    JUMLAH.Caption = 0
    If Me.Dt1.value = Null Or Dt2.value = Null Then Exit Sub
    sSQL = "select K.NoRangka, K.Nama_Pemesan, k.Kode_Lgn, k.model, k.[DESC] namamodel,   k.NoPolisi, k.Thmobil, k.Ddate, k.NamaSales, k.Nama_Supervisor,"
    sSQL = sSQL & " k.Tanggal, k.Alamat, k.Kelurahan, k.Kecamatan, k.Kota, k.Propinsi, k.KodePos, k.Telp, k.HP1, k.HP2, k.useName, k.UseGsm, k.lastSB,"
    sSQL = sSQL & " 'Service Berkala '+t.fs_keterangan JobLastSB, k.DealerlastSB, k.lastGR, k.joblastGR, k.DealerlastGR"
    sSQL = sSQL & " , case when datediff(d,isnull(k.lastSB,'01/01/1900'),isnull(k.lastGR,'01/01/1900')) < 0 then 'Service Berkala '+t.FS_KETERANGAN else k.joblastGR end  LastSvc,"
    sSQL = sSQL & " case when datediff(d,isnull(k.lastSB,'01/01/1900'),isnull(k.lastGR,'01/01/1900')) < 0 then k.dealerlastsb else k.DealerlastGR end  dealerLastSvc"
    sSQL = sSQL & " from ("
    sSQL = sSQL & "     select a.NoRangka, b.Nama_Pemesan, b.Kode_Lgn, a.model, c.[DESC], d.NoPolisi, d.Thmobil, d.Ddate, s.NamaSales, p.Nama_Supervisor, a.Tanggal,"
    sSQL = sSQL & "     n.Alamat, n.Kelurahan, n.Kecamatan, n.Kota, n.Propinsi, n.KodePos, n.Telp, n.HP1, n.HP2, n2.namacust1+' '+n2.namacust2+' '+n2.namacust3 useName, n2.hp1 UseGsm,"
    sSQL = sSQL & "    (select top 1 invdate from [" & dbReplikasiHO & "]..t_invoicenew with (nolock) where no_chassis = a.norangka and fs_rc <> '' order by invdate desc) lastSB,"
    sSQL = sSQL & "    (select top 1 FS_Rc from [" & dbReplikasiHO & "]..t_invoicenew with (nolock) where no_chassis = a.norangka and fs_rc <> '' order by invdate desc) JoblastSB,"
    sSQL = sSQL & "    (select top 1 fs_nm_dealer from [" & dbReplikasiHO & "]..t_invoicenew with (nolock) where no_chassis = a.norangka and fs_rc <> '' order by invdate desc) DealerlastSB,"
    sSQL = sSQL & "    (select top 1 invdate from [" & dbReplikasiHO & "]..t_invoicenew with (nolock) where no_chassis = a.norangka and fs_rc = '' order by invdate desc) lastGR,"
    sSQL = sSQL & "    (select top 1 rep1 from [" & dbReplikasiHO & "]..t_invoicenew with (nolock) where no_chassis = a.norangka and fs_rc = '' order by invdate desc) joblastGR,"
    sSQL = sSQL & "    (select top 1 fs_nm_dealer from [" & dbReplikasiHO & "]..t_invoicenew with (nolock) where no_chassis = a.norangka and fs_rc = '' order by invdate desc) DealerlastGR"
    sSQL = sSQL & "    from [" & dbReplikasiHO & "]..Summary a"
    sSQL = sSQL & "    left join SPK" & Right(gDb, 2) & "..spk b on a.NoRangka = b.NoRangka"
    sSQL = sSQL & "    left join SPK" & Right(gDb, 2) & "..ArSales S on b.Kode_Sales = s.KodeSales"
    sSQL = sSQL & "    left join SPK" & Right(gDb, 2) & "..ArSupervisor P on s.Kode_Supervisor = p.Kode_Supervisor"
    sSQL = sSQL & "    left join [" & dbReplikasiHO & "]..MODEL c on a.Model = c.MDL"
    sSQL = sSQL & "    left join SingleIDOri..UnitCustomer D on a.NoRangka = d.NoRangka"
    sSQL = sSQL & "    left join SingleIDOri..NewCustomer N on d.IDCustSales = n.IDCust"
    sSQL = sSQL & "    left join SingleIDOri..NewCustomer N2 on d.IDCustSL = n2.IDCust"
    sSQL = sSQL & "    where KodeDealer = '" & Log.KodeCabang & "' and TipeTr = 'NK' and Status = 1"
    sSQL = sSQL & "    and Tanggal between '" & Format(Dt1.value, "yyyy-mm-dd") & "' and '" & Format(Dt2.value, "yyyy-mm-dd") & "' and isnull(a.NoRangka,'') <> ''"
    sSQL = sSQL & " ) K"
    sSQL = sSQL & " left join BENGKEL" & Right(gDb, 2) & "..T_REGULAR T on k.joblastSB = t.FS_RC"
    sSQL = sSQL & " order by k.norangka"
    
    Set rsCR = objO.FetchData(sSQL)
    JUMLAH.Caption = rsCR.RecordCount
    grid.DataSource = rsCR
End Sub

Private Sub CmdXls_Click()
1         On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsCR.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Invalid CR Kosong !"
5             Beep
6             Er.show (vbModal)
7             Exit Sub
8         End If
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".XLS"
          
          ''20200513 REQ PAK KOKOK MINTA DITAMBAHKAN LOG UNTUK USER YG MENDOWNLOAD
          ''-----------------------------------------------------------------------------
16        sSQL = "insert into logDownloadExcelMRA(kodeuser,namamenu,tglExport,tglFilter1,tglFilter2,kriteria,totalrecord,destinasi,kompname)"
17        sSQL = sSQL & " values ('" & gSA & "','frmInfoRetensiUnitRetail',getdate(),'" & Format(Dt1, "yyyy-mm-dd") & "','" & Format(Dt2, "yyyy-mm-dd") & "','" & kriteria & "','" & rsCR.RecordCount & "','" & onVal & "','" & ComptName & "') "
18        Call objO.InsertData(sSQL)
          
19         grid.Columns(11).Visible = False
20         grid.Columns(12).Visible = False
21         grid.Columns(13).Visible = False
22         grid.Columns(14).Visible = False
23         grid.Columns(15).Visible = False
24         grid.Columns(16).Visible = False
25         grid.ExportToFile onVal, True
26         grid.Columns(11).Visible = True
27         grid.Columns(12).Visible = True
28         grid.Columns(13).Visible = True
29         grid.Columns(14).Visible = True
30         grid.Columns(15).Visible = True
31         grid.Columns(16).Visible = True
32                Er.HeadCap = "Export to excel !"
33                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
34                Beep
35                Er.show (vbModal)
36         Exit Sub
Err_Export:
37          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
38          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub grid_FilterChange()
    On Error GoTo err
   Set cols = grid.Columns
   Dim C As Integer
   
   C = grid.col
   grid.HoldFields
   rsCR.filter = getFilter()
   If rsCR.EOF = True Then
      For Each col In grid.Columns
         col.FilterText = ""
      Next col
      rsCR.filter = ""
   End If
   grid.col = C
   grid.EditActive = True
   JUMLAH.Caption = rsCR.RecordCount
   Exit Sub
err:
   For Each col In cols
      col.FilterText = ""
   Next col
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
         tmp = tmp & col.DataField & " LIKE '" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub grid_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 14 Then rsCR.Sort = grid.Columns(ColIndex).DataField & " " & Urut
4        grid.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        JUMLAH.Caption = rsCR.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form FormDataPelanggan")

End Sub
