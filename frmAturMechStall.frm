VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmAturMechStall 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "hari"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10455
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   3360
      Top             =   3600
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
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown2 
      Bindings        =   "frmAturMechStall.frx":0000
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2355
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode Stall"
      Columns(0).DataField=   "KodeStall"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nama Stall"
      Columns(1).DataField=   "NamaStall"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
      DataField       =   "SHIFT"
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=180,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   5640
      Top             =   3600
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
   Begin TrueOleDBGrid70.TDBGrid listMekNJ 
      Height          =   4095
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7223
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1).DropDown=   "TDBDropDown2"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDown1"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "TDBDropDown1"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4).DropDown=   "TDBDropDown1"
      Columns(4).DropDown.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   4
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   4
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   4
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   4
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   4
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   11
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=11"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(10)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(34)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(38)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   0
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
      _StyleDefs(80)  =   "Named:id=33:Normal"
      _StyleDefs(81)  =   ":id=33,.parent=0"
      _StyleDefs(82)  =   "Named:id=34:Heading"
      _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(84)  =   ":id=34,.wraptext=-1"
      _StyleDefs(85)  =   "Named:id=35:Footing"
      _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(87)  =   "Named:id=36:Selected"
      _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=37:Caption"
      _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(91)  =   "Named:id=38:HighlightRow"
      _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=39:EvenRow"
      _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(95)  =   "Named:id=40:OddRow"
      _StyleDefs(96)  =   ":id=40,.parent=33"
      _StyleDefs(97)  =   "Named:id=41:RecordSelector"
      _StyleDefs(98)  =   ":id=41,.parent=34"
      _StyleDefs(99)  =   "Named:id=42:FilterBar"
      _StyleDefs(100) =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
      Bindings        =   "frmAturMechStall.frx":0015
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2355
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode Mek"
      Columns(0).DataField=   "MECH"
      Columns(0).DataWidth=   500
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nama Mek"
      Columns(1).DataField=   "Name"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Level Mek"
      Columns(2).DataField=   "levMech"
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
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1323"
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
      DataField       =   "SHIFT"
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=180,.bold=0,.fontsize=825,.italic=0"
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
   Begin TrueOleDBGrid70.TDBGrid ListMech 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   7223
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "MEKANIK 1"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "MEKANIK 1"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDown1"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "TDBDropDown1"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "STALL"
      Columns(4).DataField=   ""
      Columns(4).DropDown=   "TDBDropDown1"
      Columns(4).DropDown.vt=   8
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
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
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      MultiSelect     =   0
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
   Begin ucXPButton.XPButton cmdSimpan 
      Height          =   495
      Left            =   7800
      TabIndex        =   5
      Top             =   4800
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
      MICON           =   "frmAturMechStall.frx":002A
      PICN            =   "frmAturMechStall.frx":0046
      PICH            =   "frmAturMechStall.frx":0559
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
      Left            =   9120
      TabIndex        =   6
      Top             =   4920
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
      MICON           =   "frmAturMechStall.frx":0AFB
      PICN            =   "frmAturMechStall.frx":0B17
      PICH            =   "frmAturMechStall.frx":0FEB
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdSimpanNJ 
      Height          =   495
      Left            =   7800
      TabIndex        =   3
      Top             =   4920
      Visible         =   0   'False
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
      MICON           =   "frmAturMechStall.frx":1536
      PICN            =   "frmAturMechStall.frx":1552
      PICH            =   "frmAturMechStall.frx":1A65
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown3 
      Bindings        =   "frmAturMechStall.frx":2007
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   2520
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2355
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Kode Stall"
      Columns(0).DataField=   "KodeStall"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nama Stall"
      Columns(1).DataField=   "NamaStall"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
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
      DataField       =   "SHIFT"
      IntegralHeight  =   0   'False
      FetchRowStyle   =   0   'False
      AlternatingRowStyle=   0   'False
      DataMember      =   ""
      ColumnFooters   =   0   'False
      FootLines       =   1
      DeadAreaBackColor=   14215660
      ValueTranslate  =   0   'False
      RowDividerColor =   15790320
      RowSubDividerColor=   15790320
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=180,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   5520
      Top             =   3120
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
      BorderColor     =   &H00FFFFFF&
      X1              =   10200
      X2              =   9600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   10200
      X2              =   9360
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   10200
      X2              =   9240
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   600
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   960
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   10320
      X2              =   0
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   10320
      X2              =   0
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Stall Mekanik"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   7815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   10335
   End
End
Attribute VB_Name = "frmAturMechStall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim rsSublet As New ADODB.Recordset
Dim sMulai As String
Dim sSelesai As String
Dim batasMulai As String
Dim batasSelesai As String
Dim nHari As String
Dim jmlStall As Integer
Dim rsLevelStall As ADODB.Recordset
Dim isGridValidate As Boolean

Private Function createSublet() As Boolean
1         On Error GoTo vb_error
          
2         If rsSublet.State Then rsSublet.Close
3         Set rsSublet = Nothing
4         With rsSublet
5             .Fields.Append "STALL", adVarChar, 25
6             .Fields.Append "TIPE", adVarChar, 25
7             .Fields.Append "MEKANIK1", adVarChar, 5
8             .Fields.Append "MEKANIK2", adVarChar, 5
9             .Fields.Append "MEKANIK3", adVarChar, 5
10            .Open
11        End With
12        If rsSublet.State Then createSublet = True
13        Exit Function
          
vb_error:
14        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function


Private Function createRSNJ() As Boolean
    On Error GoTo vb_error
    
    If rsSublet.State Then rsSublet.Close
    Set rsSublet = Nothing
    With rsSublet
        .Fields.Append "STALL", adVarChar, 25
        .Fields.Append "kodestall", adVarChar, 25
        .Fields.Append "MEKANIK1", adVarChar, 5
        .Fields.Append "MEKANIK2", adVarChar, 5
        .Fields.Append "MEKANIK3", adVarChar, 5
        .Fields.Append "levelStall", adVarChar, 10
        .Fields.Append "NamaStall", adVarChar, 50
        .Fields.Append "Lift", adBoolean
        .Fields.Append "Aktif", adBoolean
        .Fields.Append "OnlineStall", adBoolean
        .Fields.Append "AvailableGR", adBoolean
        .Fields.Append "AvailablePM", adBoolean
        .Open
    End With
    If rsSublet.State Then createRSNJ = True
    Exit Function
    
vb_error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub refreshGridNJ()
    Set rsTemp = objO.FetchData("SELECT A.mech, name, isnull(masuk,0) as masuk, ISNULL(HARI,'') AS HARI, ISNULL(TANGGAL,'') AS TANGGAL, " _
                             & "ISNULL(SHIFT,0) AS SHIFT, ISNULL(JAMMULAI,'') AS JAMMULAI, ISNULL(JAMSELESAI,'') AS JAMSELESAI, ISNULL(MENITKERJA,0) AS MENITKERJA " _
                             & "FROM MEKANIK A WITH(NOLOCK) LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' ORDER BY a.MECH")
                                 
    sSQL = "SELECT a.*,"
    sSQL = sSQL & " case when b.is_online=1 then '1' else '0' end as [OnlineStall],"
    sSQL = sSQL & " case when b.avail_gr=1 then '1' else '0' end as [AvailableGR],"
    sSQL = sSQL & " case when b.avail_pm=1 then '1' else '0' end as [AvailablePM] FROM SETTINGSTALL a WITH (NOLOCK)"
    sSQL = sSQL & " left join ReplBengkel99..p2_stall_config_data b on a.noStall=b.outlet_stall_ID and b.outlet_ID = '" & Log.kdDealerDidx & "'"
    Set rsMek = objO.FetchData(sSQL)
    jmlStall = objO.FetchData_Integer("SELECT FN_JML_STALL FROM T_PROFIL WITH(NOLOCK)")
  
    Dim rsTipe As ADODB.Recordset
    Dim tipeStall As String
    
    If createRSNJ Then
        For i = 1 To jmlStall
            
            rsMek.filter = "nostall = 'Stall " & i & "'"
            If Not rsMek.EOF Then
                With rsSublet
                    .AddNew
                    !stall = "Stall " & i
                    !LevelStall = Trim(rsMek("levelstall"))
                    !mekanik1 = rsMek("tek1")
                    !mekanik2 = rsMek("tek2")
                    !MEKANIK3 = rsMek("tek3")
                    !kodestall = Trim(rsMek("kodestall"))
                    !NAMASTALL = Trim(rsMek("namastall"))
                    !lift = IIf(CInt(rsMek("lift")) = 1, True, False)
                    !aktif = IIf(CInt(rsMek("ACTIVATION")) = 1, True, False)
                    !OnlineStall = IIf(CInt(rsMek("OnlineStall")) = 1, True, False)
                    !AvailableGR = IIf(CInt(rsMek("AvailableGR")) = 1, True, False)
                    !AvailablePM = IIf(CInt(rsMek("AvailablePM")) = 1, True, False)
                    .Update
                    .MoveNext
                End With
            Else
                With rsSublet
                    .AddNew
                    !stall = "Stall " & i
                    !LevelStall = ""
                    !mekanik1 = ""
                    !mekanik2 = ""
                    !MEKANIK3 = ""
                    !kodestall = ""
                    !NAMASTALL = ""
                    !lift = False
                    !aktif = False
                    !OnlineStall = False
                    !AvailableGR = False
                    !AvailablePM = False
                    .Update
                    .MoveNext
                End With
            End If
            rsMek.filter = adFilterNone
        Next
    End If
    listMekNJ.DataSource = rsSublet
    Call headColumnsNJ
End Sub

Private Sub refreshGrid()
    ' LIHAT LAGI
    Set rsTemp = objO.FetchData("SELECT A.mech, name, isnull(masuk,0) as masuk, ISNULL(HARI,'') AS HARI, ISNULL(TANGGAL,'') AS TANGGAL, " _
                             & "ISNULL(SHIFT,0) AS SHIFT, ISNULL(JAMMULAI,'') AS JAMMULAI, ISNULL(JAMSELESAI,'') AS JAMSELESAI, ISNULL(MENITKERJA,0) AS MENITKERJA " _
                             & "FROM MEKANIK A WITH(NOLOCK) LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' ORDER BY a.MECH")
                             
    jmlStall = objO.FetchData_Integer("SELECT FN_JML_STALL FROM T_PROFIL WITH(NOLOCK)")
    Dim rsMek As ADODB.Recordset
    Dim rsTipe As ADODB.Recordset
    Dim tipeStall As String
    
    If createSublet Then
        For i = 1 To jmlStall
            tipeStall = ""
            Set rsTipe = objO.FetchData("SELECT TIPE FROM LIFT WITH (NOLOCK) WHERE STALL = 'Stall " & i & "' AND TIPE <> 'L'")
            Set rsMek = objO.FetchData("SELECT * FROM SETTINGSTALL WITH (NOLOCK) WHERE NOSTALL = 'Stall " & i & "'")
                If Not rsTipe.EOF Then
                    rsTipe.MoveFirst
                    Do While Not rsTipe.EOF
                        If rsTipe("TIPE") = "L" Then
                            tipeStall = tipeStall & "L "
                        ElseIf rsTipe("TIPE") = "N" Then
                            tipeStall = tipeStall & "NP "
                        ElseIf rsTipe("TIPE") = "E" Then
                            tipeStall = tipeStall & "EM "
                        ElseIf rsTipe("TIPE") = "X" Then
                            tipeStall = tipeStall & "SC "
                        End If
                    rsTipe.MoveNext
                    Loop
                    
                End If
            If Not rsMek.EOF Then
                With rsSublet
                    .AddNew
                    !stall = "Stall " & i
                    !tipe = tipeStall
                    !mekanik1 = rsMek("tek1")
                    !mekanik2 = rsMek("tek2")
                    !MEKANIK3 = rsMek("tek3")
                    .Update
                    .MoveNext
                End With
            Else
                With rsSublet
                    .AddNew
                    !stall = "Stall " & i
                    !tipe = tipeStall
                    !mekanik1 = ""
                    !mekanik2 = ""
                    !MEKANIK3 = ""
                    .Update
                    .MoveNext
                End With
            End If
            
        Next
    End If
    ListMech.DataSource = rsSublet
    Call headColumns
End Sub

Private Sub headColumns()
    
    With ListMech
        .Columns(0).Caption = "STALL"
        .Columns(0).DataField = "STALL"
        .Columns(0).Width = 600
        .Columns(1).Caption = "TIPE STALL"
        .Columns(1).DataField = "tipe"
        .Columns(1).Width = 1000
        .Columns(2).Caption = "MEKANIK 1"
        .Columns(2).DataField = "MEKANIK1"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "MEKANIK 2"
        .Columns(3).DataField = "MEKANIK2"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "MEKANIK 3"
        .Columns(4).DataField = "MEKANIK3"
        .Columns(4).Width = 1000
    End With
End Sub

Private Sub headColumnsNJ()
    
    With listMekNJ
        .Columns(0).Caption = "STALL"
        .Columns(0).DataField = "STALL"
        .Columns(0).Width = 600
        .Columns(1).Caption = "TIPE STALL"
        .Columns(1).DataField = "kodestall"
        .Columns(1).Width = 1000
        .Columns(2).Caption = "MEKANIK 1"
        .Columns(2).DataField = "MEKANIK1"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "MEKANIK 2"
        .Columns(3).DataField = "MEKANIK2"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "MEKANIK 3"
        .Columns(4).DataField = "MEKANIK3"
        .Columns(4).Width = 1000
        .Columns(5).Caption = "NAMA STALL"
        .Columns(5).DataField = "namaStall"
        .Columns(5).Width = 2000
        .Columns(6).Caption = "LIFT"
        .Columns(6).DataField = "lift"
        .Columns(6).Width = 1000
        .Columns(7).Caption = "AKTIF"
        .Columns(7).DataField = "aktif"
        .Columns(7).Width = 1000
        .Columns(8).Caption = "Online Stall"
        .Columns(8).DataField = "OnlineStall"
        .Columns(8).Width = 1000
        .Columns(9).Caption = "Available GR"
        .Columns(9).DataField = "AvailableGR"
        .Columns(9).Width = 1000
        .Columns(10).Caption = "Available PM"
        .Columns(10).DataField = "AvailablePM"
        .Columns(10).Width = 1000
    End With
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub
Private Sub cmdSimpan_Click()
          Dim rsHelpAbsen As ADODB.Recordset
          Dim jamMulai As String
          Dim jamSelesai As String
          Dim rsLogJPCB As ADODB.Recordset
1         Set rsTekLama = objO.FetchData("SELECT NOSTALL,TEK1 AS TEK1LAMA,TEK2 AS TEK2LAMA,TEK3 AS TEK3LAMA FROM SETTINGSTALL WITH (NOLOCK) ")
          'On Error GoTo vbError
2         Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMechStall", "simpanSettingStallMekanik", 0)
                                
3         erTranAll = False
4         onTR = True
5         Call objHT.OpenTransactionBengkel
6         Set rsLogJPCB = objHT.FetchDataBengkel("SELECT * FROM PROFILBENGKEL..LOGJPCB")
          ' SIMPAN
7         If Not rsSublet.State = adStateClosed Then
8             rsSublet.MoveFirst
9             If Not rsSublet.EOF Then
10                Do While Not rsSublet.EOF
11                    Set rsTemp = objO.FetchData("SELECT * FROM SETTINGSTALL WITH(NOLOCK) WHERE NOSTALL = '" & rsSublet("STALL") & "'")
12                    If Not rsTemp.EOF Then
13                        Call objO.UpdateData("UPDATE SETTINGSTALL SET TEK1 = '" & rsSublet("MEKANIK1") & "',TEK2 = '" & rsSublet("MEKANIK2") & "',TEK3 = '" & rsSublet("MEKANIK3") & "' WHERE nostall = '" & rsSublet("stall") & "'")
14                    Else
15                        Call objO.InsertData("INSERT INTO SETTINGSTALL  (NOSTALL, TEK1, TEK2, TEK3) VALUES ('" & rsSublet("STALL") & "', '" & rsSublet("MEKANIK1") & "', '" & rsSublet("MEKANIK2") & "', '" & rsSublet("MEKANIK3") & "')")
16                    End If
                      
17                    If Log.helpJpcb = 1 Then
18                    Do While Not rsLogJPCB.EOF
                      
19                        rsTekLama.filter = "Nostall = '" & Trim(rsSublet("stall")) & "'"
                         'lakukan pengisian data help jpcb berdasarkan mech
20                        If Not rsTekLama.EOF Then
21                            If rsSublet("MEKANIK1") <> Trim(rsTekLama("TEK1LAMA")) And Trim(rsTekLama("TEK1LAMA")) <> "" Then
22                                jamMulai = "00:00"
23                                jamSelesai = "00:00"
24                                Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsTekLama("TEK1LAMA")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
25                            End If
26                            If rsSublet("MEKANIK1") <> "" Then
27                                Set rsHelpAbsen = objHT.FetchDataBengkel("SELECT * FROM T_MEKANIK WHERE TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MASUK = '1' AND MECH = '" & rsSublet("MEKANIK1") & "'")
28                                If rsHelpAbsen.EOF Then
29                                    jamMulai = "00:00"
30                                    jamSelesai = "00:00"
31                                Else
32                                    rsHelpAbsen.MoveFirst
33                                    jamMulai = Trim(rsHelpAbsen("jammulai"))
34                                    jamSelesai = Trim(rsHelpAbsen("jamselesai"))
35                                End If
36                                Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsSublet("MEKANIK1")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
37                            End If
                              
                              '------------- MEKANIK 2 --------------
38                            If rsSublet("MEKANIK2") <> Trim(rsTekLama("TEK2LAMA")) And Trim(rsTekLama("TEK2LAMA")) <> "" Then
39                                jamMulai = "00:00"
40                                jamSelesai = "00:00"
41                                Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsTekLama("TEK2LAMA")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
42                            End If
43                            If rsSublet("MEKANIK2") <> "" Then
44                                Set rsHelpAbsen = objHT.FetchDataBengkel("SELECT * FROM T_MEKANIK WHERE TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MASUK = '1' AND MECH = '" & rsSublet("MEKANIK2") & "'")
45                                If rsHelpAbsen.EOF Then
46                                    jamMulai = "00:00"
47                                    jamSelesai = "00:00"
48                                Else
49                                    rsHelpAbsen.MoveFirst
50                                    jamMulai = Trim(rsHelpAbsen("jammulai"))
51                                    jamSelesai = Trim(rsHelpAbsen("jamselesai"))
52                                End If
53                                Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsSublet("MEKANIK2")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
54                            End If
                              '------------- MEKANIK 3 --------------
55                            If rsSublet("MEKANIK3") <> Trim(rsTekLama("TEK3LAMA")) And Trim(rsTekLama("TEK3LAMA")) <> "" Then
56                                jamMulai = "00:00"
57                                jamSelesai = "00:00"
58                                Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsTekLama("TEK3LAMA")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
59                            End If
60                            If rsSublet("MEKANIK3") <> "" Then
61                                Set rsHelpAbsen = objHT.FetchDataBengkel("SELECT * FROM T_MEKANIK WHERE TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MASUK = '1' AND MECH = '" & rsSublet("MEKANIK3") & "'")
62                                If rsHelpAbsen.EOF Then
63                                    jamMulai = "00:00"
64                                    jamSelesai = "00:00"
65                                Else
66                                    rsHelpAbsen.MoveFirst
67                                    jamMulai = Trim(rsHelpAbsen("jammulai"))
68                                    jamSelesai = Trim(rsHelpAbsen("jamselesai"))
69                                End If
70                                Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsSublet("MEKANIK3")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
71                            End If
72                        End If
73                        rsTekLama.filter = ""
                                              
                          
74                       rsLogJPCB.MoveNext
75                    Loop
76                    If rsLogJPCB.RecordCount > 0 Then
77                        rsLogJPCB.MoveFirst
78                    End If
79                    End If
                      
80                    rsSublet.MoveNext
81                Loop
82            End If
83        End If
          

84        If erTranAll = False Then
85            Call objHT.CloseTransactionBengkel
86            onTR = False
87            Unload Me
88            Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMechStall", "simpanSettingStallMech", 1)
89        Else
90            Call objHT.CloseTransactionBengkel_err
91            onTR = False
92            Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMechStall", "simpanSettingStallMech", 2)
93            MsgBox "Data tidak berhasil disimpan." & vbCr & "Ulangi penyimpanan...", vbInformation, "Gagal Simpan"
94        End If
                  
95        Exit Sub
          
vbError:
96        Call objHT.CloseTransactionBengkel_err
97        onTR = False
98        Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMechStall", "simpanSettingStallMech", 2)
99        MsgBox "Data tidak berhasil disimpan." & vbCr & "Ulangi penyimpanan...", vbInformation, "Gagal Simpan"
100       Exit Sub
End Sub

Private Sub cmdSimpanNJ_Click()

1        On Error GoTo cmdSimpanNJ_Click_Error
          Dim rsHelpAbsen As ADODB.Recordset
          Dim jamMulai As String
          Dim jamSelesai As String
          Dim rsLogJPCB As ADODB.Recordset
          Dim level As String
2         Set rsTekLama = objO.FetchData("SELECT NOSTALL,TEK1 AS TEK1LAMA,TEK2 AS TEK2LAMA,TEK3 AS TEK3LAMA FROM SETTINGSTALL WITH (NOLOCK) ")
          'On Error GoTo vbError
3         Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMechStall", "simpanSettingStallMekanik", 0)
                                
4         erTranAll = False
5         onTR = True
6         Call objHT.OpenTransactionBengkel
7         Set rsLogJPCB = objHT.FetchDataBengkel("SELECT * FROM PROFILBENGKEL..LOGJPCB")
          ' SIMPAN
8         If Not rsSublet.State = adStateClosed Then
9             rsSublet.MoveFirst
10            If Not rsSublet.EOF Then
11                Do While Not rsSublet.EOF
12                    level = ""
13                    If Not rsLevelStall.EOF Then
14                        If Trim(rsSublet("kodestall")) <> "" Then
15                            rsLevelStall.filter = "KODESTALL = '" & Trim(rsSublet("KODESTALL")) & "'"
16                            If Not rsLevelStall.EOF Then
17                                level = rsLevelStall("LEVEL")
18                            End If
19                            rsLevelStall.filter = adFilterNone
20                        End If
21                    End If
22                    Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM SETTINGSTALL WITH(NOLOCK) WHERE NOSTALL = '" & rsSublet("STALL") & "'")
23                    If Not rsTemp.EOF Then
24                        Call objHT.UpdateData_Bengkel("UPDATE SETTINGSTALL SET TEK1 = '" & rsSublet("MEKANIK1") & "',TEK2 = '" & rsSublet("MEKANIK2") & "',TEK3 = '" & rsSublet("MEKANIK3") & "',LEVELSTALL = '" & level & "',KODESTALL = '" & Trim(rsSublet("kodestall")) & "',NAMASTALL = '" & Trim(rsSublet("namastall")) & "', LIFT = '" & IIf(rsSublet("lift"), 1, 0) & "',ACTIVATION = '" & IIf(rsSublet("aktif"), 1, 0) & "' WHERE nostall = '" & rsSublet("stall") & "'")
25                    Else
26                        Call objHT.UpdateData_Bengkel("INSERT INTO SETTINGSTALL VALUES ('" & rsSublet("STALL") & "', '" & rsSublet("MEKANIK1") & "', '" & rsSublet("MEKANIK2") & "', '" & rsSublet("MEKANIK3") & "','" & Trim(level) & "','" & Trim(rsSublet("kodestall")) & "','" & Trim(rsSublet("namastall")) & "','" & IIf(rsSublet("lift"), 1, 0) & "','" & IIf(rsSublet("aktif"), 1, 0) & "')")
27                    End If
                      
28                    Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM ReplBengkel99..p2_stall_config WITH(NOLOCK) WHERE outlet_ID = '" & Log.kdDealerDidx & "'")
29                    If Not rsTemp.EOF Then
30                        Call objHT.UpdateData_Bengkel("Update ReplBengkel99..p2_stall_config set event_ID=NEWID(),timestamp=GETDATE(), tgl_entry=GETDATE() where outlet_ID= '" & Log.kdDealerDidx & "'")
31                    Else
32                        Call objHT.UpdateData_Bengkel("INSERT INTO ReplBengkel99..p2_stall_config (outlet_ID,outlet_name,event_ID,timestamp,tgl_entry) VALUES ('" & Log.kdDealerDidx & "','" & gNmCbg & "',NEWID(),GETDATE(),GETDATE())")
33                    End If
                      
34                    Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM ReplBengkel99..p2_stall_config_data WITH(NOLOCK) WHERE outlet_stall_ID = '" & rsSublet("STALL") & "' and outlet_ID = '" & Log.kdDealerDidx & "'")
35                    If Not rsTemp.EOF Then
36                        Call objHT.UpdateData_Bengkel("UPDATE ReplBengkel99..p2_stall_config_data SET stall_status = '" & IIf(rsSublet("aktif"), 1, 0) & "', avail_gr = '" & IIf(rsSublet("AvailableGR"), 1, 0) & "',avail_pm = '" & IIf(rsSublet("AvailablePM"), 1, 0) & "',is_online = '" & IIf(rsSublet("OnlineStall"), 1, 0) & "' WHERE outlet_stall_ID = '" & rsSublet("STALL") & "' and outlet_ID = '" & Log.kdDealerDidx & "'")
37                    Else
38                        Call objHT.UpdateData_Bengkel("INSERT INTO ReplBengkel99..p2_stall_config_data (stall_status,[outlet_ID],[outlet_stall_ID],[stall_name],[stall_description],[stall_status],[avail_gr],[avail_pm],[is_online]) VALUES ('" & IIf(rsSublet("aktif"), 1, 0) & "','" & Log.kdDealerDidx & "', '" & rsSublet("STALL") & "', '" & rsSublet("STALL") & "', '','1','" & IIf(rsSublet("AvailableGR"), 1, 0) & "','" & IIf(rsSublet("AvailablePM"), 1, 0) & "','" & IIf(rsSublet("OnlineStall"), 1, 0) & "')")
39                    End If
40                    If Log.helpJpcb = 1 Then
41                        Do While Not rsLogJPCB.EOF
                          
42                            rsTekLama.filter = "Nostall = '" & Trim(rsSublet("stall")) & "'"
                             'lakukan pengisian data help jpcb berdasarkan mech
43                            If Not rsTekLama.EOF Then
44                                If rsSublet("MEKANIK1") <> Trim(rsTekLama("TEK1LAMA")) And Trim(rsTekLama("TEK1LAMA")) <> "" Then
45                                    jamMulai = "00:00"
46                                    jamSelesai = "00:00"
47                                    Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsTekLama("TEK1LAMA")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
48                                End If
49                                If rsSublet("MEKANIK1") <> "" Then
50                                    Set rsHelpAbsen = objHT.FetchDataBengkel("SELECT * FROM T_MEKANIK WHERE TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MASUK = '1' AND MECH = '" & rsSublet("MEKANIK1") & "'")
51                                    If rsHelpAbsen.EOF Then
52                                        jamMulai = "00:00"
53                                        jamSelesai = "00:00"
54                                    Else
55                                        rsHelpAbsen.MoveFirst
56                                        jamMulai = Trim(rsHelpAbsen("jammulai"))
57                                        jamSelesai = Trim(rsHelpAbsen("jamselesai"))
58                                    End If
59                                    Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsSublet("MEKANIK1")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
60                                End If
                                  
                                  '------------- MEKANIK 2 --------------
61                                If rsSublet("MEKANIK2") <> Trim(rsTekLama("TEK2LAMA")) And Trim(rsTekLama("TEK2LAMA")) <> "" Then
62                                    jamMulai = "00:00"
63                                    jamSelesai = "00:00"
64                                    Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsTekLama("TEK2LAMA")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
65                                End If
66                                If rsSublet("MEKANIK2") <> "" Then
67                                    Set rsHelpAbsen = objHT.FetchDataBengkel("SELECT * FROM T_MEKANIK WHERE TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MASUK = '1' AND MECH = '" & rsSublet("MEKANIK2") & "'")
68                                    If rsHelpAbsen.EOF Then
69                                        jamMulai = "00:00"
70                                        jamSelesai = "00:00"
71                                    Else
72                                        rsHelpAbsen.MoveFirst
73                                        jamMulai = Trim(rsHelpAbsen("jammulai"))
74                                        jamSelesai = Trim(rsHelpAbsen("jamselesai"))
75                                    End If
76                                    Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsSublet("MEKANIK2")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
77                                End If
                                  '------------- MEKANIK 3 --------------
78                                If rsSublet("MEKANIK3") <> Trim(rsTekLama("TEK3LAMA")) And Trim(rsTekLama("TEK3LAMA")) <> "" Then
79                                    jamMulai = "00:00"
80                                    jamSelesai = "00:00"
81                                    Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsTekLama("TEK3LAMA")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
82                                End If
83                                If rsSublet("MEKANIK3") <> "" Then
84                                    Set rsHelpAbsen = objHT.FetchDataBengkel("SELECT * FROM T_MEKANIK WHERE TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MASUK = '1' AND MECH = '" & rsSublet("MEKANIK3") & "'")
85                                    If rsHelpAbsen.EOF Then
86                                        jamMulai = "00:00"
87                                        jamSelesai = "00:00"
88                                    Else
89                                        rsHelpAbsen.MoveFirst
90                                        jamMulai = Trim(rsHelpAbsen("jammulai"))
91                                        jamSelesai = Trim(rsHelpAbsen("jamselesai"))
92                                    End If
93                                    Call objHT.UpdateData_Bengkel("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & UCase(Trim(rsSublet("STALL"))) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & jamMulai & "','" & jamSelesai & "',0,'" & Trim(rsSublet("MEKANIK3")) & "','','SETTING STALL','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
94                                End If
95                            End If
96                            rsTekLama.filter = ""
                                                  
                              
97                           rsLogJPCB.MoveNext
98                        Loop
99                        If rsLogJPCB.RecordCount > 0 Then
100                           rsLogJPCB.MoveFirst
101                       End If
102                   End If
103                   rsSublet.MoveNext
104               Loop
105               sSQL = "INSERT INTO ReplBengkel99..kirimDataDidx"
106               sSQL = sSQL & " (tipe,nobookingapps,nobooking,noSPK,kodedealer,status,proses,event_no,tglentry,asal_cabang)"
107               sSQL = sSQL & " VALUES ('CONFIG','','" & Log.kdDealerDidx & "','','" & Right(gDb, 2) & "','STALLCONFIG ',0,'',GETDATE(),'" & Right(gDb, 2) & "')"
108               Call objHT.UpdateData_Bengkel(sSQL)
109           End If
110       End If
          

111       If erTranAll = False Then
112           Call objHT.CloseTransactionBengkel
113           onTR = False
114           Unload Me
115           Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMechStall", "simpanSettingStallMech", 1)
116       Else
117           Call objHT.CloseTransactionBengkel_err
118           onTR = False
119           Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMechStall", "simpanSettingStallMech", 2)
120           MsgBox "Data tidak berhasil disimpan." & vbCr & "Ulangi penyimpanan...", vbInformation, "Gagal Simpan"
121       End If
                  
122       Exit Sub
          
123      On Error GoTo 0
124      Exit Sub

cmdSimpanNJ_Click_Error:
125       Call objHT.CloseTransactionBengkel_err
126       onTR = False
127       Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMechStall", "simpanSettingStallMech", 2)

128       MsgBox "Data tidak berhasil disimpan." & vbCr & "Ulangi penyimpanan..." & vbCrLf & "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpanNJ_Click of Form frmAturMechStall")
End Sub

Private Sub Form_Activate()
    Me.Move 1680, 4200, Me.Width, Me.Height
    With Adodc1
        .ConnectionString = adoCon
        .CommandType = adCmdText
        .RecordSource = "SELECT MECH,NAME,LEVMECH FROM mekanik WITH (NOLOCK) ORDER BY levmech,mech"
        .Refresh
    End With
    TDBDropDown1.DataSource = Adodc1
    With Adodc2
        .ConnectionString = adoCon
        .CommandType = adCmdText
        .RecordSource = "select KodeStall,NamaStall FROM LevelStall order by level"
        .Refresh
    End With
    TDBDropDown2.DataSource = Adodc2
    With Adodc3
        .ConnectionString = adoCon
        .CommandType = adCmdText
        .RecordSource = "select KodeStall,NamaStall FROM LevelStall order by level"
        .Refresh
    End With
    TDBDropDown3.DataSource = Adodc3
End Sub

Private Sub Form_Load()
    If Log.newJpcb = 1 Then
        cmdSimpan.Visible = False
        cmdSimpanNJ.Visible = True
        ListMech.Visible = False
        listMekNJ.Visible = True
        Call refreshGridNJ
    Else
        cmdSimpanNJ.Visible = False
        cmdSimpan.Visible = True
        ListMech.Visible = True
        listMekNJ.Visible = False
        Call refreshGrid
    End If
    TDBDropDown1.Width = 3800
    TDBDropDown2.Width = 2500
    Call jamHrIni
    Set rsLevelStall = objO.FetchData("select * FROM LevelStall order by level")
End Sub

Private Sub jamHrIni()
Dim rsJam As ADODB.Recordset

    Set rsJam = objO.FetchData("SELECT * FROM T_STALL WITH(NOLOCK) WHERE HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "'")
    If Not rsJam.EOF Then
        sMulai = Trim(rsJam("WAKTUMULAI"))
        sSelesai = Trim(rsJam("WAKTUSELESAI"))
        nHari = Trim(rsJam("HARI"))
    End If
End Sub

Private Sub ListMech_AfterColEdit(ByVal ColIndex As Integer)
Dim rsSearch As ADODB.Recordset
    
End Sub

Private Sub ListMech_AfterColUpdate(ByVal ColIndex As Integer)
    Select Case ColIndex
    Case 2
        If ListMech.Columns(2) <> "" Then
            rsSublet.filter = "MEKANIK1 = '" & ListMech.Columns(2) & "' or MEKANIK2 = '" & ListMech.Columns(2) & "' or MEKANIK3 = '" & ListMech.Columns(2) & "'"
            If Not rsSublet.EOF Then
                    If rsSublet.RecordCount > 1 Then
                        rsSublet.filter = adFilterNone
                        MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                        Call refreshGrid
                    End If
                    rsSublet.filter = adFilterNone
                Exit Sub
            End If
        End If
    Case 3
        If ListMech.Columns(3) <> "" Then
            rsSublet.filter = "MEKANIK1 = '" & ListMech.Columns(3) & "' or MEKANIK2 = '" & ListMech.Columns(3) & "' or MEKANIK3 = '" & ListMech.Columns(3) & "'"
            If Not rsSublet.EOF Then
                    If rsSublet.RecordCount > 1 Then
                        rsSublet.filter = adFilterNone
                        MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                        Call refreshGrid
                    End If
                    rsSublet.filter = adFilterNone
                Exit Sub
            End If
        End If
    Case 4
        If ListMech.Columns(4) <> "" Then
            rsSublet.filter = "MEKANIK1 = '" & ListMech.Columns(4) & "' or MEKANIK2 = '" & ListMech.Columns(4) & "' or MEKANIK3 = '" & ListMech.Columns(4) & "'"
            If Not rsSublet.EOF Then
                    If rsSublet.RecordCount > 1 Then
                        rsSublet.filter = adFilterNone
                        MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                        Call refreshGrid
                    End If
                    rsSublet.filter = adFilterNone
                Exit Sub
            End If
        End If
    End Select
            
End Sub

Public Sub panggilWaktu(idx As Integer)
Dim rsHlp As ADODB.Recordset

    Set rsHlp = objO.FetchData("SELECT * FROM MECHSHIFT WITH(NOLOCK) WHERE SHIFT = '" & CDbl(idx) & "'")
    If Not rsHlp.EOF Then
        batasMulai = Trim(rsHlp("JAMMULAI"))
        batasSelesai = Trim(rsHlp("JAMSELESAI"))
    End If
End Sub

Private Sub listMekNJ_AfterColUpdate(ByVal ColIndex As Integer)
    If isGridValidate = False Then
        Call colUpdate(ColIndex)
    End If
End Sub

Private Sub colUpdate(ColIndex As Integer)
     Select Case ColIndex
        Case 1
            If Trim(listMekNJ.Columns(ColIndex)) <> "" Then
                rsSublet!NAMASTALL = Trim(listMekNJ.Columns(ColIndex)) & " " & Trim(Right(listMekNJ.Columns(0), 2)) 'Trim(listMekNJ.Columns(0))
            End If
        Case 2
            If listMekNJ.Columns(2) <> "" Then
                rsSublet.filter = "MEKANIK1 = '" & listMekNJ.Columns(2) & "' or MEKANIK2 = '" & listMekNJ.Columns(2) & "' or MEKANIK3 = '" & listMekNJ.Columns(2) & "'"
                If Not rsSublet.EOF Then
                        If rsSublet.RecordCount > 1 Then
                            rsSublet.filter = adFilterNone
                            MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                            Call refreshGridNJ
                        End If
                        rsSublet.filter = adFilterNone
                    Exit Sub
                End If
            End If
        Case 3
            If listMekNJ.Columns(3) <> "" Then
                rsSublet.filter = "MEKANIK1 = '" & listMekNJ.Columns(3) & "' or MEKANIK2 = '" & listMekNJ.Columns(3) & "' or MEKANIK3 = '" & listMekNJ.Columns(3) & "'"
                If Not rsSublet.EOF Then
                        If rsSublet.RecordCount > 1 Then
                            rsSublet.filter = adFilterNone
                            MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                            Call refreshGridNJ
                        End If
                        rsSublet.filter = adFilterNone
                    Exit Sub
                End If
            End If
        Case 4
            If listMekNJ.Columns(4) <> "" Then
                rsSublet.filter = "MEKANIK1 = '" & listMekNJ.Columns(4) & "' or MEKANIK2 = '" & listMekNJ.Columns(4) & "' or MEKANIK3 = '" & listMekNJ.Columns(4) & "'"
                If Not rsSublet.EOF Then
                        If rsSublet.RecordCount > 1 Then
                            rsSublet.filter = adFilterNone
                            MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                            Call refreshGridNJ
                        End If
                        rsSublet.filter = adFilterNone
                    Exit Sub
                End If
            End If
    End Select
End Sub


Private Sub listMekNJ_LostFocus()
    'Sendkeys ("+{TAB}")
End Sub

Private Sub listMekNJ_Validate(Cancel As Boolean)
    isGridValidate = True
    If listMekNJ.Columns(2) <> "" Then
        rsSublet.filter = "MEKANIK1 = '" & listMekNJ.Columns(2) & "' or MEKANIK2 = '" & listMekNJ.Columns(2) & "' or MEKANIK3 = '" & listMekNJ.Columns(2) & "'"
        If Not rsSublet.EOF Then
                If rsSublet.RecordCount > 1 Then
                    rsSublet.filter = adFilterNone
                    MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                    Call refreshGridNJ
                End If
                rsSublet.filter = adFilterNone
            'Exit Sub
        End If
    End If
    If listMekNJ.Columns(3) <> "" Then
        rsSublet.filter = "MEKANIK1 = '" & listMekNJ.Columns(3) & "' or MEKANIK2 = '" & listMekNJ.Columns(3) & "' or MEKANIK3 = '" & listMekNJ.Columns(3) & "'"
        If Not rsSublet.EOF Then
                If rsSublet.RecordCount > 1 Then
                    rsSublet.filter = adFilterNone
                    MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                    Call refreshGridNJ
                End If
                rsSublet.filter = adFilterNone
            'Exit Sub
        End If
    End If
    If listMekNJ.Columns(4) <> "" Then
        rsSublet.filter = "MEKANIK1 = '" & listMekNJ.Columns(4) & "' or MEKANIK2 = '" & listMekNJ.Columns(4) & "' or MEKANIK3 = '" & listMekNJ.Columns(4) & "'"
        If Not rsSublet.EOF Then
                If rsSublet.RecordCount > 1 Then
                    rsSublet.filter = adFilterNone
                    MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
                    Call refreshGridNJ
                End If
                rsSublet.filter = adFilterNone
            'Exit Sub
        End If
    End If
    isGridValidate = False
End Sub

Private Sub TDBDropDown1_Click()
'    rsSublet.filter = "MEKANIK1 = '" & TDBDropDown1 & "' or MEKANIK2 = '" & TDBDropDown1 & "' or MEKANIK3 = '" & TDBDropDown1 & "'"
'    If Not rsSublet.EOF Then
'        MsgBox "Tidak dapat memilih Mekanik " & TDBDropDown1 & " sudah dipilih di Stall ", vbCritical, "Kesalahan input"
'        rsSublet.filter = ""
'        Exit Sub
'    Else
'        rsSublet.filter = ""
'    End If
End Sub


