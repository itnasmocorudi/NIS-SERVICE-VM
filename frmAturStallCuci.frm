VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmAturStallCuci 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "hari"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Begin AIFCmp1.asxPowerButton btnHapusCuci 
      Height          =   495
      Left            =   7080
      TabIndex        =   10
      Top             =   1200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   873
      BackColor       =   255
      Caption         =   "Hapus Cuci"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin VB.CheckBox ckPengeringan 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stall Pengering"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Width           =   1695
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   600
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      Calculator      =   "frmAturStallCuci.frx":0000
      Caption         =   "frmAturStallCuci.frx":0020
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAturStallCuci.frx":008C
      Keys            =   "frmAturStallCuci.frx":00AA
      Spin            =   "frmAturStallCuci.frx":00F4
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   100
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1515126789
      MinValueVT      =   1163264005
   End
   Begin TrueOleDBGrid70.TDBGrid ListMech 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "MEKANIK 1"
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "TDBDropDown1"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1).DropDown=   "TDBDropDown1"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "STALL"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "TDBDropDown1"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
   Begin ucXPButton.XPButton cmdSimpan 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   5880
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
      MICON           =   "frmAturStallCuci.frx":011C
      PICN            =   "frmAturStallCuci.frx":0138
      PICH            =   "frmAturStallCuci.frx":064B
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
      Left            =   7560
      TabIndex        =   5
      Top             =   5880
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
      MICON           =   "frmAturStallCuci.frx":0BED
      PICN            =   "frmAturStallCuci.frx":0C09
      PICH            =   "frmAturStallCuci.frx":10DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBNumber6Ctl.TDBNumber TDBNumber2 
      Height          =   375
      Left            =   7200
      TabIndex        =   6
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
      _Version        =   65536
      _ExtentX        =   1720
      _ExtentY        =   661
      Calculator      =   "frmAturStallCuci.frx":1628
      Caption         =   "frmAturStallCuci.frx":1648
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAturStallCuci.frx":16B4
      Keys            =   "frmAturStallCuci.frx":16D2
      Spin            =   "frmAturStallCuci.frx":171C
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   100
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   1245189
      Value           =   0
      MaxValueVT      =   1515126789
      MinValueVT      =   1163264005
   End
   Begin TrueOleDBGrid70.TDBGrid listPengering 
      Height          =   2055
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Visible         =   0   'False
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3625
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "MEKANIK 1"
      Columns(0).DataField=   ""
      Columns(0).DropDown=   "TDBDropDown1"
      Columns(0).DropDown.vt=   8
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1).DropDown=   "TDBDropDown1"
      Columns(1).DropDown.vt=   8
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "STALL"
      Columns(3).DataField=   ""
      Columns(3).DropDown=   "TDBDropDown1"
      Columns(3).DropDown.vt=   8
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   4
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
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
   Begin AIFCmp1.asxPowerButton btnHapusPengering 
      Height          =   495
      Left            =   7080
      TabIndex        =   11
      Top             =   3840
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BackColor       =   255
      Caption         =   "Hapus Pengering"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jumlah Stall Pengeringan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Jumlah Stall Cuci"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   600
      Width           =   2175
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   8760
      X2              =   8160
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   8760
      X2              =   7920
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   8760
      X2              =   7800
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
      X1              =   8760
      X2              =   120
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   8760
      X2              =   120
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Setting Stall Cuci"
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
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmAturStallCuci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim rsTempSimpan As ADODB.Recordset

Dim rsStallCuci As ADODB.Recordset
Dim rsTemStallCuci As New ADODB.Recordset

Dim rsStallPengering As ADODB.Recordset
Dim rsTemStallPengering As New ADODB.Recordset

Dim sMulai As String
Dim sSelesai As String
Dim batasMulai As String
Dim batasSelesai As String
Dim nHari As String
Dim jmlStall As Integer

Private jumlahStallCuci As Integer
Private jumlahStallPengering As Integer

Private Sub refreshGridPengering()
          ' LIHAT LAGI
          
1         If createStallPengering Then
2             If rsStallPengering.RecordCount > 0 Then
                  Dim j As Integer
3                 j = 1
4                 rsStallPengering.MoveFirst
5                 Do While Not rsStallPengering.EOF
6                     With rsTemStallPengering
7                         .AddNew
8                         !noStall = Trim(rsStallPengering("nostall"))
9                         !NAMASTALL = Trim(rsStallPengering("NamaStall"))
10                        !LTCuci = Trim(rsStallPengering("LTCuci"))
11                        !namaMekanik = Trim(rsStallPengering!petugas)
12                        .Update
13                        .MoveNext
14                    End With
15                    rsStallPengering.MoveNext
16                    j = j + 1
17                Loop
18                If jumlahStallPengering > rsStallPengering.RecordCount Then
19                    For i = j To jumlahStallPengering
20                        With rsTemStallPengering
21                            .AddNew
22                            !noStall = "Pengering " & i
23                            !NAMASTALL = ""
24                            !LTCuci = "00:00"
25                            !namaMekanik = ""
26                            .Update
27                            .MoveNext
28                        End With
29                    Next
30                End If
31            Else
32                For i = 1 To jumlahStallPengering
33                    With rsTemStallPengering
34                        .AddNew
35                        !noStall = "Pengering " & i
36                        !NAMASTALL = ""
37                        !LTCuci = "00:00"
38                        !namaMekanik = ""
39                        .Update
40                        .MoveNext
41                    End With
42                Next
43            End If
              
44        End If
45        listPengering.DataSource = rsTemStallPengering
46        Call headColumnPengering
End Sub

Private Sub refreshGrid()
    ' LIHAT LAGI
    
    If createStallCuci Then
        If rsStallCuci.RecordCount > 0 Then
            Dim j As Integer
            j = 1
            rsStallCuci.MoveFirst
            Do While Not rsStallCuci.EOF
                With rsTemStallCuci
                    .AddNew
                    !noStall = Trim(rsStallCuci("nostall"))
                    !NAMASTALL = Trim(rsStallCuci("NamaStall"))
                    !LTCuci = Trim(rsStallCuci("LTCuci"))
                    !namaMekanik = Trim(rsStallCuci!petugas)
                    .Update
                    .MoveNext
                End With
                rsStallCuci.MoveNext
                j = j + 1
            Loop
            If jumlahStallCuci > rsStallCuci.RecordCount Then
                For i = j To jumlahStallCuci
                    With rsTemStallCuci
                        .AddNew
                        !noStall = "Cuci " & i
                        !NAMASTALL = ""
                        !LTCuci = "00:00"
                        !namaMekanik = ""
                        .Update
                        .MoveNext
                    End With
                Next
            End If
        Else
            For i = 1 To jumlahStallCuci
                With rsTemStallCuci
                    .AddNew
                    !noStall = "Cuci " & i
                    !NAMASTALL = ""
                    !LTCuci = "00:00"
                    !namaMekanik = ""
                    .Update
                    .MoveNext
                End With
            Next
        End If
        
    End If
    ListMech.DataSource = rsTemStallCuci
    Call headColumns
End Sub

Private Sub headColumnPengering()
    With listPengering
        .Columns(0).Caption = "No Stall"
        .Columns(0).DataField = "NoStall"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "Nama Stall"
        .Columns(1).DataField = "NamaStall"
        .Columns(1).Width = 4000
        .Columns(2).Caption = "LTCuci"
        .Columns(2).DataField = "LTCuci"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "Nama Mekanik"
        .Columns(3).DataField = "NamaMekanik"
        .Columns(3).Width = 4000
    End With
End Sub

Private Sub headColumns()
    
    With ListMech
        .Columns(0).Caption = "No Stall"
        .Columns(0).DataField = "NoStall"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "Nama Stall"
        .Columns(1).DataField = "NamaStall"
        .Columns(1).Width = 4000
        .Columns(2).Caption = "LTCuci"
        .Columns(2).DataField = "LTCuci"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "Nama Mekanik"
        .Columns(3).DataField = "NamaMekanik"
        .Columns(3).Width = 4000
    End With
    
    
End Sub

Private Sub btnHapusCuci_Click()
    If MsgBox("Apakah yakin menghapus data " & ListMech.Columns(0).value, vbYesNo) = vbYes Then
        objO.InsertData ("UPDATE QTY_STALLWASHING SET ISDEL = '1',LOGUSERDEL = '" & gSA & "',TGLUSERDEL = GETDATE() WHERE NOSTALL = '" & ListMech.Columns(0).value & "'")
        Call Form_Load
    End If
End Sub

Private Sub btnHapusPengering_Click()
    If MsgBox("Apakah yakin menghapus data " & listPengering.Columns(0).value, vbYesNo) = vbYes Then
        objO.InsertData ("UPDATE QTY_STALLWASHING SET ISDEL = '1',LOGUSERDEL = '" & gSA & "',TGLUSERDEL = GETDATE() WHERE NOSTALL = '" & listPengering.Columns(0).value & "'")
        Call Form_Load
    End If
End Sub

Private Sub ckPengeringan_Click()
    If ckPengeringan.value = 1 Then
    Call hidePengering(True)
    Else
    Call hidePengering(False)
    End If
End Sub

Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Function createStallPengering() As Boolean
1         On Error GoTo vb_error
          
2         If rsTemStallPengering.State Then rsTemStallPengering.Close
3         Set rsTemStallPengering = Nothing
4         With rsTemStallPengering
5             .Fields.Append "NoStall", adVarChar, 25
6             .Fields.Append "NamaStall", adVarChar, 50
7             .Fields.Append "LTCuci", adVarChar, 10
8             .Fields.Append "NamaMekanik", adVarChar, 20
9             .Open
10        End With
11        If rsTemStallPengering.State Then createStallPengering = True
12        Exit Function
          
vb_error:
13        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createStallCuci() As Boolean
1         On Error GoTo vb_error
          
2         If rsTemStallCuci.State Then rsTemStallCuci.Close
3         Set rsTemStallCuci = Nothing
4         With rsTemStallCuci
5             .Fields.Append "NoStall", adVarChar, 25
6             .Fields.Append "NamaStall", adVarChar, 50
7             .Fields.Append "LTCuci", adVarChar, 10
8             .Fields.Append "NamaMekanik", adVarChar, 20
9             .Open
10        End With
11        If rsTemStallCuci.State Then createStallCuci = True
12        Exit Function
          
vb_error:
13        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function


Private Sub cmdSimpan_Click()
    
    On Error GoTo vbError
    Call saveLogMenu(gSA, ComptName, gVersi, "frmAturStallCuci", "simpanSettingCuciStall", 0)
                          
    erTranAll = False
    onTR = True
    Call objHT.OpenTransactionBengkel
            
    ' SIMPAN
    If Not rsTemStallCuci.State = adStateClosed Then
        rsTemStallCuci.MoveFirst
        If Not rsTemStallCuci.EOF Then
            Do While Not rsTemStallCuci.EOF
                
                Call objO.DeleteData("DELETE FROM qty_stallwashing WHERE nostall = '" & rsTemStallCuci("nostall") & "'")
                Call objO.InsertData("INSERT INTO qty_stallwashing VALUES (getdate(),'" & rsTemStallCuci("nostall") & "','" & rsTemStallCuci("namastall") & "','Cuci','" & gSA & "','" & rsTemStallCuci("ltcuci") & "','','','0','0','" & rsTemStallCuci("NamaMekanik") & "')")
                Call objO.UpdateData("UPDATE QTY_STALLWASHING SET AKTIF = '1' WHERE nostall = '" & rsTemStallCuci("nostall") & "'")
                rsTemStallCuci.MoveNext
            Loop
        End If
    End If
    
    If Not rsTemStallPengering.State = adStateClosed Then
        If rsTemStallPengering.RecordCount > 0 Then
            rsTemStallPengering.MoveFirst
            If Not rsTemStallPengering.EOF Then
                Do While Not rsTemStallPengering.EOF
                    Call objO.DeleteData("DELETE FROM qty_stallwashing WHERE nostall = '" & rsTemStallPengering("nostall") & "'")
                    Call objO.InsertData("INSERT INTO qty_stallwashing VALUES (getdate(),'" & rsTemStallPengering("nostall") & "','" & rsTemStallPengering("namastall") & "','Pengering','" & gSA & "','" & rsTemStallPengering("ltcuci") & "','','','0','1','" & rsTemStallPengering("NamaMekanik") & "')")
                    If ckPengeringan.value = 0 Then
                        Call objO.UpdateData("UPDATE QTY_STALLWASHING SET AKTIF = '0' WHERE nostall = '" & rsTemStallPengering("nostall") & "'")
                    End If
                    rsTemStallPengering.MoveNext
                Loop
            End If
        End If
        
    End If
    
    If erTranAll = False Then
        Call objHT.CloseTransactionBengkel
        onTR = False
        Unload Me
        Call saveLogMenu(gSA, ComptName, gVersi, "frmAturStallCuci", "simpanStallCuci", 1)
    Else
        Call objHT.CloseTransactionBengkel_err
        onTR = False
        Call saveLogMenu(gSA, ComptName, gVersi, "frmAturStallCuci", "simpanStallCuci", 2)
        MsgBox "Data tidak berhasil disimpan." & vbCr & "Ulangi penyimpanan...", vbInformation, "Gagal Simpan"
    End If
    Exit Sub
vbError:
    Call objHT.CloseTransactionBengkel_err
    onTR = False
    Call saveLogMenu(gSA, ComptName, gVersi, "frmAturStallCuci", "simpanStallCuci", 2)
    MsgBox "Data tidak berhasil disimpan." & vbCr & "Ulangi penyimpanan...", vbInformation, "Gagal Simpan"
    Exit Sub
End Sub

Private Sub Form_Load()
    Call GetDataStall
    Call GetDataPengering
    TDBNumber1.value = jumlahStallCuci
    TDBNumber2.value = jumlahStallPengering
    refreshGrid
    refreshGridPengering
End Sub

Private Sub GetDataPengering()
    Set rsStallPengering = objO.FetchData("SELECT * FROM QTY_StallWashing WHERE TipeStall <> 'Cuci' AND ISDEL = '0' and aktif = '1'")
    jumlahStallPengering = rsStallPengering.RecordCount
    If jumlahStallPengering > 0 Then
        ckPengeringan.value = 1
    End If
End Sub

Private Sub GetDataStall()
    Set rsStallCuci = objO.FetchData("SELECT * FROM QTY_StallWashing WHERE TipeStall = 'Cuci' AND ISDEL = '0' AND AKTIF = '1'")
    jumlahStallCuci = rsStallCuci.RecordCount
End Sub

Private Sub hidePengering(Opt As Boolean)
    Label2.Visible = Opt
    TDBNumber2.Visible = Opt
    listPengering.Visible = Opt
    btnHapusPengering.Visible = Opt
End Sub


Private Sub ListMech_AfterColUpdate(ByVal ColIndex As Integer)
    If ListMech.Columns(1).value <> "" Then
        If cekJamPengerjaan(ListMech.Columns(2)) = True Then
            'MsgBox ListMech.Columns(2)
        Else
            MsgBox "Estimasi pengerjaan salah, Format MM:DD", vbCritical
        End If
    End If
End Sub

Private Sub ListMech_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If ListMech.Columns(1).value <> "" Then
'        If cekJamPengerjaan(ListMech.Columns(2)) = True Then
'            MsgBox ListMech.Columns(2)
'        Else
'            MsgBox "Estimasi pengerjaan salah, Format MM:DD", vbCritical
'
'        End If
'    End If
        
End Sub

Private Function cekJamPengerjaan(total As String) As Boolean
    Dim detik As Integer
    Dim menit As Integer
    
    On Error GoTo vb_error
    
    menit = CInt(Left(Trim(total), InStr(Trim(total), ":") - 1))
    detik = Mid(Trim(total), InStr(Trim(total), ":") + 1, Len(Trim(total)) - InStr(Trim(total), ":"))
    
    detik = detik + (menit * 60)
    
    If detik > 0 Then
        cekJamPengerjaan = True
    Else
        cekJamPengerjaan = False
    End If
    Exit Function
vb_error:
    cekJamPengerjaan = False
End Function

Private Sub TDBNumber1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        jumlahStallCuci = TDBNumber1.value
        refreshGrid
    End If
End Sub

Private Sub TDBNumber2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        jumlahStallPengering = TDBNumber2.value
        refreshGridPengering
    End If
End Sub
