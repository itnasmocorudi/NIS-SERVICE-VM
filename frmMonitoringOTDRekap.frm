VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmMonitoringOTDRekap 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Monitoring Rekap  Unit OTD"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11910
   LinkTopic       =   "frmMonitoringOTDRekap"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEditCustomer 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Edit Customer"
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
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6600
      Width           =   1410
   End
   Begin VB.CommandButton cmdRefresh 
      BackColor       =   &H0000FF00&
      Caption         =   "Refresh"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   693108737
      CurrentDate     =   42571
   End
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   5318
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   953
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   15790320
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Button=1"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(0).ButtonText=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   15790320
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
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
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.namedParent=39"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.namedParent=40"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBGrid GridFind 
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   4260
      _LayoutType     =   4
      _RowHeight      =   13
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Pilih"
      Columns(0).DataField=   "pilih"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "KodeLgn"
      Columns(1).DataField=   "idcust"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "No KTP"
      Columns(2).DataField=   "NoKTP"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nama  Customer"
      Columns(3).DataField=   "nama"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Alamat"
      Columns(4).DataField=   "alamat"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "No HP"
      Columns(5).DataField=   "hp"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "WO Terakhir"
      Columns(6).DataField=   "tglWO"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "KodeLgn Lama"
      Columns(7).DataField=   "Kdlgn"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
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
      MultiSelect     =   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=862,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=859,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=860,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=861,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=870,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=867,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=868,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=869,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=878,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=875,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=876,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=877,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=886,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=883,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=884,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=885,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=894,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=891,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=892,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=893,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H80FF&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
      _StyleDefs(89)  =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(90)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(91)  =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(92)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(93)  =   "Named:id=87:SudahDitagih"
      _StyleDefs(94)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(95)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(96)  =   ":id=87,.fontname=Tahoma"
      _StyleDefs(97)  =   "Named:id=514:Kuning"
      _StyleDefs(98)  =   ":id=514,.parent=87,.bgcolor=&HFFFF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
      _StyleDefs(99)  =   ":id=514,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(100) =   ":id=514,.fontname=Tahoma"
      _StyleDefs(101) =   "Named:id=515:Biru"
      _StyleDefs(102) =   ":id=515,.parent=514,.bgcolor=&HFF8080&"
      _StyleDefs(103) =   "Named:id=516:Merah"
      _StyleDefs(104) =   ":id=516,.parent=515,.bgcolor=&H88A0F7&,.fgcolor=&H0&"
      _StyleDefs(105) =   "Named:id=517:Hijau"
      _StyleDefs(106) =   ":id=517,.parent=516,.bgcolor=&HFF00&"
      _StyleDefs(107) =   "Named:id=518:Putih"
      _StyleDefs(108) =   ":id=518,.parent=517,.bgcolor=&HFFFFFF&"
   End
   Begin VB.Label Label56 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Data Customer Belum Lengkap"
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
      Height          =   285
      Left            =   4545
      TabIndex        =   5
      Top             =   3720
      Width           =   3570
   End
   Begin VB.Line Line28 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   4080
      Y1              =   3795
      Y2              =   3795
   End
   Begin VB.Line Line29 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   4080
      Y1              =   3915
      Y2              =   3915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tanggal OTD :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   3720
      Width           =   11655
   End
End
Attribute VB_Name = "frmMonitoringOTDRekap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As New ADODB.Recordset

Dim rsKTP  As ADODB.Recordset
Dim rsTemp2 As New ADODB.Recordset
Dim rsTemp3 As New ADODB.Recordset

Dim rsGrid As New ADODB.Recordset
Dim nCount As Integer
Dim cols As TrueOleDBGrid70.Columns
Dim tglAwal As String
Dim tglAkhir As String
Public monitoringSA As String


Private Sub headColumns()
'[DATE], CUSTOMER, CUST_ADRES, TELPFU, JAMFU, KETFU, A.MDL + ' | ' + A.FMDL, A.COLOR, A.NO_CHASSIS, A.NO_WO, A.NO_POLISI, KM

   On Error GoTo headColumns_Error

    With grid
        .Columns(0).Caption = "SA"
        .Columns(0).DataField = "sa"
        .Columns(0).Width = 800
        
        .Columns(1).Caption = "WORK ORDER"
        .Columns(1).DataField = "UNITWO"
        .Columns(1).Width = 1700
        
        .Columns(2).Caption = "UNIT INVOICE"
        .Columns(2).DataField = "UNITINVOICE"
        .Columns(2).Width = 1200
        
        .Columns(3).Caption = "NOTIFIKASI"
        .Columns(3).DataField = "notifikasi"
        .Columns(3).Width = 1100

        .Columns(4).Caption = "OTD"
        .Columns(4).DataField = "OTD"
        .Columns(4).Width = 2500

        .Columns(5).Caption = "PROMISE"
        .Columns(5).DataField = "PROMISE"
        .Columns(5).Width = 1200
        .Columns(5).Visible = False

        .Columns(6).Caption = "ACTION"
        .Columns(6).DataField = "action"
        .Columns(6).Width = 1500
        .Columns(6).Button = True
        .Columns(6).ButtonText = True
        .Columns(6).ButtonAlways = True
        
    End With

   On Error GoTo 0
   Exit Sub

headColumns_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumns of Form frm3DFU")
End Sub

Private Sub refreshGrid()
1        On Error GoTo refreshGrid_Error

2         grid.ClearFields
3         If rsGrid.State Then rsGrid.Close
      '   sSQL = "select T.FRONT , COUNT(NOINV) UnitInvoice,"
      '   sSQL = sSQL & " (select COUNT(*) total"
      '   sSQL = sSQL & " from [Queue" & Right(gDb, 2) & "]..FinishService a"
      '   sSQL = sSQL & " left join T_INVOICE b on b.INVDATE = a.TglFinish and NoPol = b.NO_POLISI"
      '   sSQL = sSQL & " left join JOBDTL j on j.NO_WO = b.NO_WO and j.NO_CHASSIS = b.NO_CHASSIS"
      '   sSQL = sSQL & " where b.FRONT = T.FRONT and TglFinish = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' and TglPanggil > '01/01/1900') panggilan,"
      '   sSQL = sSQL & " (select COUNT(*) total"
      '   sSQL = sSQL & " from [Queue" & Right(gDb, 2) & "]..FinishService a"
      '   sSQL = sSQL & " left join T_INVOICE b on b.INVDATE = a.TglFinish and NoPol = b.NO_POLISI"
      '   sSQL = sSQL & " left join JOBDTL j on j.NO_WO = b.NO_WO and j.NO_CHASSIS = b.NO_CHASSIS"
      '   sSQL = sSQL & " where b.FRONT = T.Front and TglFinish = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' and TglPanggil > '01/01/1900' and DATEDIFF(minute,TglPanggil,DATESERAH + TIMESERAH) >= 0)"
      '   sSQL = sSQL & " otd, (select SUM(rubah) from JOBSerah where No_Wo = T.NO_WO and userAdd = T.FRONT group by userAdd ) promise,"
      '   sSQL = sSQL & " 'Detail SA ' + FRONT as button"
      '   sSQL = sSQL & " from T_INVOICE T WITH (NOLOCK)"
      '   sSQL = sSQL & " INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NO_WO AND L.NO_CHASSIS = T.NO_CHASSIS"
      '   sSQL = sSQL & " WHERE L.DATESERAH = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' AND FS_LOG_DEL = ''"
      '   sSQL = sSQL & " GROUP BY T.FRONT,T.NO_WO"
      '' tambah jumlah wo
         '' 2016-08-03 || data yang diambil otd untuk hari terpilih
4        sSQL = "select T.FRONT , COUNT(t.NO_WO) WO,'Detail SA ' + FRONT as button"
5        sSQL = sSQL & " from JOBDTA T WITH (NOLOCK)"
6        sSQL = sSQL & " INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NO_WO AND L.NO_CHASSIS = T.NO_CHASSIS"
7        sSQL = sSQL & " WHERE L.DATESERAH = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' AND FS_LOG_DEL = ''"
8        sSQL = sSQL & " GROUP BY T.FRONT"


9        Set rsTemp = objO.FetchData(sSQL)

10        If Not rsTemp.EOF Then
11          Call createRsGrid
12          Do While Not rsTemp.EOF
13             With rsGrid
14                .AddNew
15                !sa = Trim(rsTemp!FRONT)
16                !UnitWO = Trim(rsTemp!wo)
17                !UnitInvoice = objO.FetchData_Double("SELECT COUNT(NOINV) FROM T_INVOICE WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND NOINV NOT LIKE '%-B' AND NO_WO IN (select t.no_wo from JOBDTA T WITH (NOLOCK) INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NO_WO AND L.NO_CHASSIS = T.NO_CHASSIS WHERE L.DATESERAH = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' AND FS_LOG_DEL = '' AND t.front = '" & Trim(rsTemp!FRONT) & "')") ''("SELECT count(noinv) FROM t_invoice WITH (NOLOCK) WHERE invdate = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' and fs_log_del = '' AND noinv not like '%-B'")
18                !notifikasi = objO.FetchData_Double("select COUNT(*) total from [Queue" & Right(gDb, 2) & "]..FinishService a left join T_INVOICE b on b.INVDATE = a.TglFinish and NoPol = b.NO_POLISI left join JOBDTL j on j.NO_WO = b.NO_WO and j.NO_CHASSIS = b.NO_CHASSIS where b.FRONT = '" & Trim(rsTemp!FRONT) & "' and TglFinish = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' and TglPanggil > '01/01/1900'")
19                !OTD = objO.FetchData_Double("select COUNT(*) total from [Queue" & Right(gDb, 2) & "]..FinishService a left join T_INVOICE b on b.INVDATE = a.TglFinish and NoPol = b.NO_POLISI left join JOBDTL j on j.NO_WO = b.NO_WO and j.NO_CHASSIS = b.NO_CHASSIS where b.FRONT = '" & Trim(rsTemp!FRONT) & "' and TglFinish = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' and TglPanggil > '01/01/1900' and DATEDIFF(minute,TglPanggil,DATESERAH + TIMESERAH) >= 0")
20                '!promise = objO.FetchData_Double("select SUM(rubah) from JOBSerah where No_Wo IN (select T.NO_WO from T_INVOICE T WITH (NOLOCK) INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NO_WO AND L.NO_CHASSIS = T.NO_CHASSIS WHERE L.DATESERAH = '" & Format(DTPicker1.value, "mm/dd/yyyy") & "' AND FS_LOG_DEL = '') and userAdd = '" & Trim(rsTemp!FRONT) & "' group by userAdd ")
21                !Action = Trim(rsTemp!Button)
22                .Update
23             End With
24             rsTemp.MoveNext
25          Loop
26            grid.DataSource = rsGrid
27            grid.FetchRowStyle = True
28            Call headColumns
           
29        End If
       

30       On Error GoTo 0
31       Exit Sub

refreshGrid_Error:

32        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form frm3DFU")

End Sub

Private Sub createRsGrid()

1        On Error GoTo createRsGrid_Error

2         If rsGrid.State Then rsGrid.Close
3         Set rsGrid = Nothing
4         With rsGrid
5             .Fields.Append "SA", adVarChar, 25
6             .Fields.Append "UnitInvoice", adVarChar, 25
7             .Fields.Append "UnitWO", adVarChar, 25
8             .Fields.Append "Notifikasi", adVarChar, 25
9             .Fields.Append "OTD", adVarChar, 25
10            .Fields.Append "Promise", adVarChar, 25
11            .Fields.Append "Action", adVarChar, 25
12            .Open
13        End With

14       On Error GoTo 0
15       Exit Sub

createRsGrid_Error:

16       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createRsGrid of Form frmMonitoringOTDRekap")
         
End Sub

Private Sub cmdEditCustomer_Click()
    If Not rsKTP.EOF Then
        rsKTP.MoveFirst
        rsKTP.filter = "pilih=" & True
        If Not rsKTP.EOF Then
            If Log.monitoringScoring = 0 Then
                With frmSIDMasterPelangganNonMerge
                    .txt(1) = NulSID(rsKTP!idcust)
                    .isEdit (True)
                    .show vbModal
                End With
            Else
                With frmSIDMasterPelanggan
                    .txt(1) = NulSID(rsKTP!idcust)
                    .isEdit (True)
                    .show vbModal
                End With
            End If
        Else
            MsgBox "Silahkan pilih data yang akan diubah.", vbInformation
        End If
    End If
    Call refreshCustomerNotLengkap
End Sub

Private Sub cmdRefresh_Click()
   If monitoringSA = "" Then
      monitoringSA = gSA
   End If
   Call refreshGrid
End Sub

Private Sub Form_Load()
   DTPicker1.value = Format(isnow, "dd/mm/yyyy")
   If monitoringSA = "" Then
      monitoringSA = gSA
   End If
   Call refreshGrid
   
   If Log.monitoringScoring = 1 Then
    Call refreshCustomerNotLengkap
   End If
End Sub

Private Sub createHeaderKTP()
    Set rsKTP = New ADODB.Recordset
    If rsKTP.State Then rsKTP.Close
    With rsKTP
        .Fields.Append "Pilih", adBoolean
        .Fields.Append "IDCust", adVarChar, 20
        .Fields.Append "NoKTP", adVarChar, 50
        .Fields.Append "Nama", adVarChar, 210
        .Fields.Append "Alamat", adVarChar, 350
        .Fields.Append "HP", adVarChar, 30
        .Fields.Append "tglWO", adDate
        .Fields.Append "KdLgn", adVarChar, 20
        .Open
        
    End With
End Sub

Private Sub refreshCustomerNotLengkap()
    sSQL = "select a.IDCust,a.TypeCustomer,a.NoKTP,a.NamaCust1 + ' ' + a.NamaCust2 + ' ' + a.NamaCust3 as nama,a.Alamat + ' ' + a.Kelurahan + ' ' +  a.Kecamatan + ' ' +  a.Kota + ' ' +  a.Propinsi as alamat,a.HP1,b.KodeLgn"
    sSQL = sSQL & " ,ISNULL((select Top 1 [DATE] from jobdta where fs_kd_plg = b.KodeLgn order by [DATE] desc),'1900-01-01') as tglwo "
    sSQL = sSQL & " from [singleidori]..newcustomer a with (nolock)"
    sSQL = sSQL & " inner join [SingleIDOri]..Bridging b with (nolock) on a.IDCust = b.IDCust and b.KodeDealer = '" & Right(gDb, 2) & "'"
    sSQL = sSQL & " left join jobdta c on b.kodelgn = c.fs_kd_plg and c.fs_log_del = '' and c.front = '" & gSA & "'"
    sSQL = sSQL & " where "
    sSQL = sSQL & " (isnull(replace(a.NoKTP,'00000000000000',''),'') = '' or (isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '' or isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '1900-01-01' or isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '3000-01-01') or isnull(replace(a.NoNPWP,'00.000.000.0-000.000',''),'') = ''"
    sSQL = sSQL & " or ISNULL(a.Agama,'') = '' or ISNULL(a.Pekerjaan,'') = '' or ISNULL(a.hobi,'') = '' or ISNULL(a.Olahraga,'') = '' or ISNULL(a.Propinsi,'') = ''"
    sSQL = sSQL & " or ISNULL(a.kota,'') = '' or ISNULL(a.Kecamatan,'') = '' or ISNULL(a.Kelurahan,'') = '' or ISNULL(a.Alamat,'') = '' or ISNULL(a.email,'') = ''"
    sSQL = sSQL & " or ISNULL(a.telp+a.HP1+a.hp2,'') = '')"
    sSQL = sSQL & " order by tglwo desc"
    Set rsTemp = objO.FetchData(sSQL)
    Call createHeaderKTP
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            With rsKTP
                .AddNew
                !Pilih = False
                !idcust = NulSID(rsTemp!idcust)
                !noKTP = NulSID(rsTemp!noKTP)
                !Nama = NulSID(rsTemp!Nama)
                !Alamat = NulSID(rsTemp!Alamat)
                !HP = NulSID(rsTemp!hp1)
                !tglWO = NulSID(rsTemp!tglWO)
                !kdLgn = NulSID(rsTemp!kodeLgn)
                .Update
            End With
            rsTemp.MoveNext
        Loop
    End If
    Set GridFind.DataSource = rsKTP
    
End Sub

Private Sub Grid_ButtonClick(ByVal ColIndex As Integer)
   frmMonitoringOTD.monitoringSA = grid.Columns(0)
   frmMonitoringOTD.tanggalMonitoring = DTPicker1.value
   frmMonitoringOTD.show vbModal
End Sub

Private Sub Grid_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
   Dim rs As New ADODB.Recordset
    Set rs = rsGrid.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    'RowStyle = "putih"
'    If rs("Panggilan") > 0 Then
'         ''jika sudah dipanggil -: background hijau
'         If Trim(CDbl(rs!SELISIHOTD)) <= 30 Then
'            RowStyle = "HijauKuning"
'         Else
'            RowStyle = "HijauMerah"
'         End If
'    Else
'         ''jika belum dipanggil -: background putih
'         If Trim(CDbl(rs!SELISIHOTD)) <= 30 Then
'            RowStyle = "PutihKuning"
'         Else
'            RowStyle = "PutihMerah"
'         End If
'    End If
End Sub

Private Sub GridFind_FilterChange()
    On Error GoTo err
   Set cols = GridFind.Columns
   Dim C As Integer
   
   C = GridFind.col
   GridFind.HoldFields
   rsKTP.filter = getFilter()
   If rsKTP.EOF = True Then
      For Each col In GridFind.Columns
         col.FilterText = ""
      Next col
      rsKTP.filter = ""
   End If
   GridFind.col = C
   GridFind.EditActive = True
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
         Select Case col.DataField
            Case "Tgl_BPK", "TglCr1", "TglCr2", "TGL_1000", "Tanggal_Transfer"
                If Len(col.FilterText) = 8 Then
                    tmp = tmp & col.DataField & " = '" & col.FilterText & "'"
                End If
            Case Else
                tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
        End Select
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function
