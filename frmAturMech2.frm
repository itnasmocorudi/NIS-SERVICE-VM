VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmAturMech2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "hari"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6570
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H000000FF&
      Caption         =   "&Batal"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton cmdSimpan 
      BackColor       =   &H0000FF00&
      Caption         =   "&Simpan"
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
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdLihat 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Lihat Data"
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4920
      Width           =   1335
   End
   Begin TDBDate6Ctl.TDBDate tanggal 
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   4920
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   661
      Calendar        =   "frmAturMech2.frx":0000
      Caption         =   "frmAturMech2.frx":012C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmAturMech2.frx":0198
      Keys            =   "frmAturMech2.frx":01B6
      Spin            =   "frmAturMech2.frx":0214
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
      Text            =   "04/02/2015"
      ValidateMode    =   0
      ValueVT         =   1028128775
      Value           =   42039
      CenturyMode     =   0
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10575
      Left            =   0
      Picture         =   "frmAturMech2.frx":023C
      ScaleHeight     =   10575
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   -4440
      Width           =   495
   End
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
   Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
      Bindings        =   "frmAturMech2.frx":1B12
      Height          =   1455
      Left            =   960
      TabIndex        =   0
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2566
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "SHIFT"
      Columns(0).DataField=   "SHIFT"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "JamMulai"
      Columns(1).DataField=   "JamMulai"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Shift"
      Columns(2).DataField=   "JamSelesai"
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
      Splits(0)._ColumnProps(10)=   "Column(1).Visible=0"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1323"
      Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
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
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7223
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "KODE"
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NAMA"
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   129
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "SHIFT"
      Columns(2).DataField=   ""
      Columns(2).DropDown=   "TDBDropDown1"
      Columns(2).DropDown.vt=   8
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3).NumberFormat=   "Short Time"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4).NumberFormat=   "Short Time"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   4
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "masuk"
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=1402"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=1323"
      Splits(0)._ColumnProps(14)=   "Column(2).Button=1"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(2).AutoDropDown=1"
      Splits(0)._ColumnProps(17)=   "Column(2).AutoCompletion=1"
      Splits(0)._ColumnProps(18)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
      _StyleDefs(60)  =   "Named:id=33:Normal"
      _StyleDefs(61)  =   ":id=33,.parent=0"
      _StyleDefs(62)  =   "Named:id=34:Heading"
      _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(64)  =   ":id=34,.wraptext=-1"
      _StyleDefs(65)  =   "Named:id=35:Footing"
      _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(67)  =   "Named:id=36:Selected"
      _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=37:Caption"
      _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(71)  =   "Named:id=38:HighlightRow"
      _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(73)  =   "Named:id=39:EvenRow"
      _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(75)  =   "Named:id=40:OddRow"
      _StyleDefs(76)  =   ":id=40,.parent=33"
      _StyleDefs(77)  =   "Named:id=41:RecordSelector"
      _StyleDefs(78)  =   ":id=41,.parent=34"
      _StyleDefs(79)  =   "Named:id=42:FilterBar"
      _StyleDefs(80)  =   ":id=42,.parent=33"
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   5760
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   5520
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   6360
      X2              =   5400
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   1080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   1320
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   600
      X2              =   1440
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   6480
      X2              =   0
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   6480
      X2              =   480
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hari"
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
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmAturMech2"
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
Dim pilihTanggal As Date

Private Function createSublet() As Boolean
1        On Error GoTo createSublet_Error

2         On Error GoTo vb_error
          
3         If rsSublet.State Then rsSublet.Close
4         With rsSublet
5             .Fields.Append "KODEMECH", adVarChar, 25
6             .Fields.Append "NAMAMECH", adVarChar, 30
7             .Fields.Append "MASUK", adBoolean
8             .Fields.Append "HARI", adVarChar, 10
9             .Fields.Append "TGL", adVarChar, 10
10            .Fields.Append "SHIFT", adVarChar, 1
11            .Fields.Append "JAMMULAI", adVarChar, 5
12            .Fields.Append "JAMSELESAI", adVarChar, 5
13            .Fields.Append "MENIT", adInteger
14            .Open
15        End With
16        If rsSublet.State Then createSublet = True
17        Exit Function
          
vb_error:
18        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical

19       On Error GoTo 0
20       Exit Function

createSublet_Error:

21        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createSublet of Form frmAturMech2")
End Function

Private Sub refreshGrid()
          ' LIHAT LAGI
1        On Error GoTo refreshGrid_Error

2         Set rsTemp = objO.FetchData("SELECT A.mech, name, isnull(masuk,0) as masuk, ISNULL(HARI,'') AS HARI, ISNULL(TANGGAL,'') AS TANGGAL, " _
                                   & "ISNULL(SHIFT,0) AS SHIFT, ISNULL(JAMMULAI,'') AS JAMMULAI, ISNULL(JAMSELESAI,'') AS JAMSELESAI, ISNULL(MENITKERJA,0) AS MENITKERJA " _
                                   & "FROM MEKANIK A WITH(NOLOCK) LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH AND TANGGAL = '" & Format(pilihTanggal, "DD/MM/YYYY") & "' AND HARI = '" & IIf(Weekday(pilihTanggal) = 1, 6, Weekday(pilihTanggal) - 2) & WeekdayName(IIf(Weekday(pilihTanggal) = 1, 8, Weekday(pilihTanggal)) - 1) & "' ORDER BY a.MECH")
                                   
3         If Not rsTemp.EOF Then
4             If createSublet Then
5                 Do While Not rsTemp.EOF
6                     With rsSublet
7                         .AddNew
8                         !KODEMECH = rsTemp("MECH")
9                         !NAMAMECH = rsTemp("NAME")
10                        !masuk = rsTemp("MASUK")
11                        !hari = rsTemp("HARI")
12                        !Tgl = rsTemp("TANGGAL")
13                        !Shift = rsTemp("SHIFT")
14                        !jamMulai = rsTemp("JAMMULAI")
15                        !jamSelesai = rsTemp("JAMSELESAI")
16                        !menit = rsTemp("MENITKERJA")
17                        .Update
18                        .MoveNext
19                    End With
20                    rsTemp.MoveNext
21                Loop
22            End If
23        End If
24        ListMech.DataSource = rsSublet
25        Call headColumns

26       On Error GoTo 0
27       Exit Sub

refreshGrid_Error:

28        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form frmAturMech2")
End Sub

Private Sub headColumns()
          
1        On Error GoTo headColumns_Error

2         With ListMech
3             .Columns(0).Caption = "KODE"
4             .Columns(0).DataField = "KODEMECH"
5             .Columns(0).Width = 600
6             .Columns(1).Caption = "NAMA"
7             .Columns(1).DataField = "NAMAMECH"
8             .Columns(1).Width = 2200
9             .Columns(2).Caption = "SHIFT"
10            .Columns(2).DataField = "SHIFT"
11            .Columns(2).Width = 600
12            .Columns(3).Caption = "MULAI"
13            .Columns(3).DataField = "JAMMULAI"
14            .Columns(3).Width = 800
15            .Columns(4).Caption = "SELESAI"
16            .Columns(4).DataField = "JAMSELESAI"
17            .Columns(4).Width = 800
18            .Columns(5).Caption = "ADA"
19            .Columns(5).DataField = "MASUK"
20            .Columns(5).Width = 400
21            .Columns(5).Alignment = dbgCenter
22        End With

23       On Error GoTo 0
24       Exit Sub

headColumns_Error:

25        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumns of Form frmAturMech2")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdLihat_Click()
1        On Error GoTo cmdLihat_Click_Error

2         pilihTanggal = Tanggal 'Format(Tanggal, "MM/DD/YYYY")
3         If DateDiff("d", Format(isnow, "DD/MM/YYYY"), Format(pilihTanggal, "DD/MM/YYYY")) >= 0 Then
4         hariNow = WeekdayName(IIf(Weekday(pilihTanggal) = 1, 7, Weekday(pilihTanggal) - 1)) & ", " & Format(pilihTanggal, "DD - MMMM - YYYY")
5         Me.Caption = "Absensi Mekanik  ->  " & hariNow
6         TDBDropDown1.Width = 800
7         Call jamHrIni
8         Call refreshGrid
9         Else
10            MsgBox "Tidak menginput data absen kemarin ", vbInformation
11            Unload Me
12        End If

13       On Error GoTo 0
14       Exit Sub

cmdLihat_Click_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdLihat_Click of Form frmAturMech2")
          
End Sub

Private Sub cmdSimpan_Click()
    Dim rsLogJPCB As ADODB.Recordset
    Dim rsHlp As ADODB.Recordset
    Dim noStall As String
    'On Error GoTo vbError
   On Error GoTo cmdSimpan_Click_Error

    Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMech2", "simpanAbsent", 0)
                          
    erTranAll = False
    onTR = True
    Call objHT.OpenTransactionBengkel
    '2015-03-11 RUDI PENAMBAHAN HELPJPCB
    Set rsLogJPCB = objO.FetchData("SELECT * FROM PROFILBENGKEL..LOGJPCB")
    
    ' SIMPAN
    If Not rsSublet.State = adStateClosed Then
        rsSublet.MoveFirst
        If Not rsSublet.EOF Then
            Do While Not rsSublet.EOF
                If Trim(rsSublet("JAMMULAI")) <> "" And Trim(rsSublet("JAMSELESAI")) <> "" Then
                    rsSublet("MENIT") = CDbl(hitMenit(rsSublet("JAMMULAI"), rsSublet("JAMSELESAI")))
                    rsSublet("HARI") = nHari
                    rsSublet("TGL") = Format(pilihTanggal, "DD/MM/YYYY")
                End If
                
                ''2016-05-17 || penambahan nostall dari setting stall
                noStall = ""
                Set rsTemp = objO.FetchData("SELECT NOSTALL FROM settingstall WITH (NOLOCK) WHERE tek1 = '" & rsSublet("KODEMECH") & "' OR tek2 = '" & rsSublet("KODEMECH") & "' OR tek3 = '" & rsSublet("KODEMECH") & "' AND activation = 1 ")
                If Not rsTemp.EOF Then
                  rsTemp.MoveFirst
                  noStall = Trim(rsTemp!noStall)
                  
                End If
                
                Set rsTemp = objO.FetchData("SELECT * FROM T_MEKANIK WITH(NOLOCK) WHERE MECH = '" & rsSublet("KODEMECH") & "' AND HARI = '" & rsSublet("HARI") & "' AND TANGGAL = '" & rsSublet("TGL") & "'")
                If Not rsTemp.EOF Then
                    If rsSublet("HARI") = "" Then
                        If rsTemp("TERPAKAI") > 0 Then
                            Call objO.UpdateData("UPDATE T_MEKANIK SET MENITKERJA = 0, MASUK = 0 WHERE MECH = '" & rsSublet("KODEMECH") & "' AND TANGGAL = '" & Format(pilihTanggal, "DD/MM/YYYY") & "'")
                        Else
                            Call objO.DeleteData("DELETE T_MEKANIK WHERE MECH = '" & rsSublet("KODEMECH") & "' AND TANGGAL = '" & Format(pilihTanggal, "DD/MM/YYYY") & "'")
                        End If
                    Else
                        Call objO.UpdateData("UPDATE T_MEKANIK SET TANGGAL = '" & rsSublet("TGL") & "', SHIFT = '" & IIf(rsSublet("SHIFT") = "", 0, rsSublet("SHIFT")) & "', jammulai =  '" & rsSublet("JAMMULAI") & "', jamselesai = '" & rsSublet("JAMSELESAI") & "', masuk = '" & IIf(rsSublet("MASUK"), 1, 0) & "', menitkerja = '" & rsSublet("MENIT") & "' WHERE MECH = '" & rsSublet("KODEMECH") & "' AND HARI = '" & rsSublet("HARI") & "' AND TANGGAL = '" & rsSublet("TGL") & "'")
                    End If
                Else
                    If rsSublet("HARI") <> "" Then
                        Call objO.InsertData("INSERT INTO T_MEKANIK VALUES ('" & rsSublet("KODEMECH") & "', '" & rsSublet("HARI") & "', '" & rsSublet("TGL") & "', '" & IIf(rsSublet("SHIFT") = "", 0, rsSublet("SHIFT")) & "', '" & rsSublet("JAMMULAI") & "', '" & rsSublet("JAMSELESAI") & "', '" & rsSublet("MENIT") & "', 0, '" & IIf(rsSublet("MASUK"), 1, 0) & "',0)")
                    End If
                End If
                
               'Dim rsHlp As ADODB.Recordset
                ' ADJUSTMET JAM MEKANIK UTK PERPANJANGAN WAKTU
                Set rsHlp = objO.FetchData("SELECT * FROM STALL A WITH(NOLOCK) INNER JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH WHERE A.MECH = '" & Trim(rsSublet("KODEMECH")) & "' AND A.TANGGAL = '" & Format(DateAdd("d", -1, pilihTanggal), "MM/DD/YYYY") & "' AND B.TANGGAL = '" & Format(DateAdd("d", -1, pilihTanggal), "DD/MM/YYYY") & "' AND FLAG = 'P'")
                If Not rsHlp.EOF Then
                    Call objO.UpdateData("UPDATE T_MEKANIK SET ALOKASI = ALOKASI + (SELECT ALOKASI - TERPAKAI FROM T_MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(rsSublet("KODEMECH")) & "' AND TANGGAL = '" & Format(DateAdd("d", -1, pilihTanggal), "DD/MM/YYYY") & "') WHERE MECH = '" & Trim(rsSublet("KODEMECH")) & "' AND TANGGAL = '" & Format(pilihTanggal, "DD/MM/YYYY") & "'")
                End If
                
                'LAKUKAN PROSES PENGISIAN DATA HELP JPCB
                'Set rsHlp = objO.FetchData("SELECT * FROM HELPJPCB WHERE TANGGAL = '" & Format(rsSublet("TGL"), "MM/DD/YYYY") & "' AND MECH = '" & Trim(rsSublet("KODEMECH")) & "'")
                If Log.helpJpcb = 1 Then
                    If rsSublet("masuk") And noStall <> "" Then
                        'lakukan pengecekan sudah ada data mekanik dan tanggal absen di help jpcb
                        If rsLogJPCB.RecordCount > 0 Then
                            rsLogJPCB.MoveFirst
                            Do While Not rsLogJPCB.EOF
                                Call objO.InsertData("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[USERID]) VALUES ('-','T','A','" & Trim(noStall) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & rsSublet("JAMMULAI") & "','" & rsSublet("JAMSELESAI") & "',0,'" & Trim(rsSublet("KODEMECH")) & "','','ABSEN','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
                                rsLogJPCB.MoveNext
                            Loop
                        End If
                    ElseIf rsSublet("masuk") = False And noStall <> "" Then
                    
                        If rsLogJPCB.RecordCount > 0 Then
                            rsLogJPCB.MoveFirst
                            Do While Not rsLogJPCB.EOF
                                '2015-03-12 UPDATE DIGANTI INSERT
                                'lakukan pengisian data help jpcb berdasarkan mech
                                Call objO.InsertData("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('-','T','A','" & Trim(noStall) & "','" & Format(isnow, "MM/DD/YYYY") & "','00:00','00:00',0,'" & Trim(rsSublet("KODEMECH")) & "','','ABSEN','00:00','CREATE','" & rsLogJPCB("KOMP") & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
                                rsLogJPCB.MoveNext
                            Loop
                        End If
                    End If
                End If
                rsSublet.MoveNext
            Loop
        End If
    End If
    

    If erTranAll = False Then
        Call objHT.CloseTransactionBengkel
        onTR = False
        Unload Me
        Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMech2", "simpanAbsent", 1)
    Else
        Call objHT.CloseTransactionBengkel_err
        onTR = False
        Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMech2", "simpanAbsent", 2)
        MsgBox "Data tidak berhasil disimpan." & vbCr & "Ulangi penyimpanan...", vbInformation, "Gagal Simpan"
    End If
            
    Exit Sub
    
vbError:
    Call objHT.CloseTransactionBengkel_err
    onTR = False
    Call saveLogMenu(gSA, ComptName, gVersi, "frmAturMech2", "simpanAbsent", 2)
    MsgBox "Data tidak berhasil disimpan." & vbCr & "Ulangi penyimpanan...", vbInformation, "Gagal Simpan"
    Exit Sub

   On Error GoTo 0
   Exit Sub

cmdSimpan_Click_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmAturMech2")
End Sub

Private Sub Form_Activate()
1        On Error GoTo Form_Activate_Error

2         Me.Move 1680, 4200, Me.Width, Me.Height
3         With Adodc1
4             .ConnectionString = adoCon
5             .CommandType = adCmdText
6             .RecordSource = "SELECT * FROM MECHSHIFT WITH(NOLOCK) ORDER BY SHIFT"
7             .Refresh
8         End With

9        On Error GoTo 0
10       Exit Sub

Form_Activate_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Activate of Form frmAturMech2")
End Sub

Private Sub Form_Load()
1        On Error GoTo Form_Load_Error

2         Tanggal = Format(isnow, "DD/MM/YYYY")
3         pilihTanggal = Tanggal 'Format(Tanggal, "MM/DD/YYYY")
          
4         hariNow = WeekdayName(IIf(Weekday(Tanggal) = 1, 7, Weekday(pilihTanggal) - 1)) & ", " & Format(pilihTanggal, "DD - MMMM - YYYY")
5         Me.Caption = "Absensi Mekanik  ->  " & hariNow
6         TDBDropDown1.Width = 800
7         Call jamHrIni
8         Call refreshGrid

9        On Error GoTo 0
10       Exit Sub

Form_Load_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Load of Form frmAturMech2")
End Sub

Private Sub jamHrIni()
      Dim rsJam As ADODB.Recordset

1        On Error GoTo jamHrIni_Error

2         Set rsJam = objO.FetchData("SELECT * FROM T_STALL WITH(NOLOCK) WHERE HARI = '" & IIf(Weekday(pilihTanggal) = 1, 6, Weekday(pilihTanggal) - 2) & WeekdayName(IIf(Weekday(pilihTanggal) = 1, 8, Weekday(pilihTanggal)) - 1) & "'")
3         If Not rsJam.EOF Then
4             sMulai = Trim(rsJam("WAKTUMULAI"))
5             sSelesai = Trim(rsJam("WAKTUSELESAI"))
6             nHari = Trim(rsJam("HARI"))
7         End If

8        On Error GoTo 0
9        Exit Sub

jamHrIni_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure jamHrIni of Form frmAturMech2")
End Sub

Private Sub ListMech_AfterColEdit(ByVal ColIndex As Integer)
      Dim rsSearch As ADODB.Recordset
1        On Error GoTo ListMech_AfterColEdit_Error

2         Select Case ColIndex
              Case 3, 4
3                 If Len(ListMech.Columns(ColIndex)) <> 5 Then
4                     rsSublet.Find "KODEMECH = '" & ListMech.Columns(0) & "'", , adSearchForward, 1
5                     rsSublet.Fields(ColIndex + 3) = Format(ListMech.Columns(ColIndex), "HH:NN")
6                 Else
7                     If Trim(ListMech.Columns(2)) = "0" Then
8                         ListMech.Columns(ColIndex) = ""
9                     End If
10                End If
11            Case 2
12                If Trim(ListMech.Columns(ColIndex)) = 0 Then
13                    rsSublet.Find "KODEMECH = '" & ListMech.Columns(0) & "'", , adSearchForward, 1
                       
14                    Set rsSearch = objO.FetchData("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND MECH = '" & Trim(ListMech.Columns(0)) & "'")
15                    If Not rsSearch.EOF Then
16                        MsgBox "Teknisi sudah dipakai tidak bisa dihapus !!", vbCritical
17                    Else
18                        rsSublet.Fields(ColIndex) = False
19                        rsSublet.Fields(ColIndex + 3) = 0
20                        rsSublet.Fields(ColIndex + 4) = ""
21                        rsSublet.Fields(ColIndex + 5) = ""
22                        rsSublet.Fields(ColIndex + 6) = 0
23                    End If
24                End If
25        End Select

26       On Error GoTo 0
27       Exit Sub

ListMech_AfterColEdit_Error:

28        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure ListMech_AfterColEdit of Form frmAturMech2")
End Sub

Private Sub ListMech_AfterColUpdate(ByVal ColIndex As Integer)
      Dim rsHlp As ADODB.Recordset
      Dim rsSearch As ADODB.Recordset

1        On Error GoTo ListMech_AfterColUpdate_Error

2         Select Case ColIndex
              Case 2      ' COMBO SHIFT
3                 If Trim(ListMech.Columns(ColIndex)) <> "" Then
4                     rsSublet.Find "KODEMECH = '" & ListMech.Columns(0) & "'", , adSearchForward, 1
5                     Set rsHlp = objO.FetchData("SELECT * FROM MECHSHIFT WITH(NOLOCK) WHERE SHIFT = '" & Trim(ListMech.Columns(ColIndex)) & "'")
6                     If Not rsHlp.EOF Then
7                         rsSublet("JAMMULAI") = Trim(rsHlp("JAMMULAI"))
8                         rsSublet("JAMSELESAI") = Trim(rsHlp("JAMSELESAI"))
9                         batasMulai = Trim(rsHlp("JAMMULAI"))
10                        batasSelesai = Trim(rsHlp("JAMSELESAI"))
11                        rsSublet("MASUK") = True
12                    Else
13                        MsgBox "Tidak ada shift tersebut !!"
14                        rsSublet("SHIFT") = 0
15                        rsSublet("JAMMULAI") = "00:00"
16                        rsSublet("JAMSELESAI") = "00:00"
17                        batasMulai = "00:00"
18                        batasSelesai = "00:00"
19                        rsSublet("MASUK") = False
20                    End If
21                Else
22                    ListMech.Columns(ColIndex) = 0
23                End If
24            Case 3      ' JAMMULAI
25                If Len(ListMech.Columns(ColIndex)) <= 5 Then
26                    If Trim(ListMech.Columns(2)) <> "0" Then
27                        Call panggilWaktu(CDbl(ListMech.Columns(2)))
28                        If Trim(Format(ListMech.Columns(ColIndex), "HH:NN")) >= batasMulai Then
29                            rsSublet.Find "KODEMECH = '" & ListMech.Columns(0) & "'", , adSearchForward, 1
30                            rsSublet("JAMMULAI") = Format(ListMech.Columns(ColIndex), "HH:NN")
31                            If CekJam(ListMech.Columns(ColIndex)) Then
32                                If Trim(ListMech.Columns(4)) <> "" Then
33                                    If ListMech.Columns(ColIndex) > ListMech.Columns(4) Then
34                                        MsgBox "Jam Mulai tidak boleh lebih besar dari Jam Selesai !"
35                                        rsSublet("JAMMULAI") = ""
36                                        rsSublet.Fields(6) = ""
37                                        ListMech.Columns(ColIndex) = ""
38                                    End If
39                                End If
40                                Set rsSearch = objO.FetchData("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND MECH = '" & Trim(ListMech.Columns(0)) & "' AND PAUSE = ''")
41                                If Not rsSearch.EOF Then
42                                    MsgBox "Teknisi sudah dipakai tidak bisa dihapus !!", vbCritical
43                                    ListMech.Columns(ColIndex) = batasMulai
44                                End If
                                  
45                                Set rsSearch = objO.FetchData("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND MECH = '" & Trim(ListMech.Columns(0)) & "' AND PAUSE <> '' AND START < '" & Trim(ListMech.Columns(ColIndex)) & "'")
46                                If Not rsSearch.EOF Then
47                                    MsgBox "Teknisi sudah dipakai tidak bisa dihapus !!", vbCritical
48                                    ListMech.Columns(ColIndex) = batasMulai
49                                End If
50                            Else
51                                rsSublet("JAMMULAI") = batasMulai
52                            End If
53                        Else
54                            Call ListMech_AfterColUpdate(2)
55                        End If
56                    Else
57                        MsgBox "Shift Mekanik harus dipilih"
58                        rsSublet("JAMMULAI") = ""
59                        rsSublet.Fields(6) = ""
60                        ListMech.Columns(ColIndex) = ""
61                    End If
62                Else
63                    MsgBox "Format Jam tidak sesuai !!" & vbCr & "Input Jam antara " & sMulai & " sampai Jam " & sSelesai, vbCritical, "SALAH INPUT"
64                    ListMech.Columns(ColIndex) = ""
65                End If
66            Case 4      ' JAMSELESAI
67                If Trim(ListMech.Columns(3)) <> "" Then
68                    If Len(ListMech.Columns(ColIndex)) <= 5 Then
69                        If Trim(ListMech.Columns(2)) <> "0" Then
70                            Call panggilWaktu(CDbl(ListMech.Columns(2)))
71                            If Trim(Format(ListMech.Columns(ColIndex), "HH:NN")) <= batasSelesai Then
72                                rsSublet.Find "KODEMECH = '" & ListMech.Columns(0) & "'", , adSearchForward, 1
                                  'ListMech.Columns(ColIndex) = Format(ListMech.Columns(ColIndex), "HH:NN")
73                                rsSublet("JAMSELESAI") = Format(ListMech.Columns(ColIndex), "HH:NN")
74                                If CekJam(ListMech.Columns(ColIndex)) Then
75                                    If ListMech.Columns(ColIndex) < ListMech.Columns(3) Then
76                                        MsgBox "Jam Selesai tidak boleh lebih kecil dari Jam Mulai !"
77                                        rsSublet("JAMSELESAI") = ""
78                                        rsSublet.Fields(7) = ""
79                                        ListMech.Columns(ColIndex) = ""
80                                    End If
81                                    Set rsSearch = objO.FetchData("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND MECH = '" & Trim(ListMech.Columns(0)) & "' AND PAUSE = ''")
82                                    If Not rsSearch.EOF Then
83                                        MsgBox "Teknisi sudah dipakai tidak bisa dihapus !!", vbCritical
84                                        ListMech.Columns(ColIndex) = batasSelesai
85                                    End If
                                      
86                                    Set rsSearch = objO.FetchData("SELECT * FROM TEMPPROGRESS WITH(NOLOCK) WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND MECH = '" & Trim(ListMech.Columns(0)) & "' AND PAUSE <> '' AND PAUSE > '" & Trim(ListMech.Columns(ColIndex)) & "'")
87                                    If Not rsSearch.EOF Then
88                                        MsgBox "Teknisi sudah dipakai tidak bisa dihapus !!", vbCritical
89                                        ListMech.Columns(ColIndex) = batasSelesai
90                                    End If
91                                Else
92                                    rsSublet("JAMSELESAI") = batasSelesai
93                                End If
94                            Else
95                                Call ListMech_AfterColUpdate(2)
96                            End If
97                        Else
98                            MsgBox "Shift Mekanik harus dipilih "
99                            rsSublet("JAMSELESAI") = ""
100                           rsSublet.Fields(6) = ""
101                           ListMech.Columns(ColIndex) = ""
102                       End If
103                   Else
104                       MsgBox "Format Jam tidak sesuai !!" & vbCr & "Input Jam antara " & sMulai & " sampai Jam " & sSelesai, vbCritical, "SALAH INPUT"
105                       ListMech.Columns(ColIndex) = ""
106                   End If
107               Else
108                   MsgBox "Isi Jam Mulai terlebih dahulu"
109                   ListMech.Columns(ColIndex) = ""
110               End If
111           Case 5      ' ABSENSI
112               If ListMech.Columns(ColIndex) = "0" Then
113                   rsSublet.Find "KODEMECH = '" & ListMech.Columns(0) & "'", , adSearchForward, 1
114                   rsSublet.Fields(2) = False
115                   rsSublet.Fields(5) = 0
116                   rsSublet.Fields(6) = ""
117                   rsSublet.Fields(7) = ""
118                   rsSublet.Fields(8) = 0
119               End If
120       End Select

121      On Error GoTo 0
122      Exit Sub

ListMech_AfterColUpdate_Error:

123       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure ListMech_AfterColUpdate of Form frmAturMech2")
End Sub

Public Sub panggilWaktu(idx As Integer)
      Dim rsHlp As ADODB.Recordset

1        On Error GoTo panggilWaktu_Error

2         Set rsHlp = objO.FetchData("SELECT * FROM MECHSHIFT WITH(NOLOCK) WHERE SHIFT = '" & CDbl(idx) & "'")
3         If Not rsHlp.EOF Then
4             batasMulai = Trim(rsHlp("JAMMULAI"))
5             batasSelesai = Trim(rsHlp("JAMSELESAI"))
6         End If

7        On Error GoTo 0
8        Exit Sub

panggilWaktu_Error:

9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilWaktu of Form frmAturMech2")
End Sub

