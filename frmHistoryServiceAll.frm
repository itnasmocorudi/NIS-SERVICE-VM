VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmHistoryServiceAll 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12375
   Icon            =   "frmHistoryServiceAll.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   12375
   Begin MSMask.MaskEdBox txtnoka 
      Height          =   375
      Left            =   2520
      TabIndex        =   27
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Frame frmeGantiNoPol 
      BackColor       =   &H00C0FFFF&
      Caption         =   "SARAN NEXT JOB ( CR 7 )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   3120
      TabIndex        =   23
      Top             =   2280
      Visible         =   0   'False
      Width           =   7455
      Begin ucXPButton.XPButton cmdSimpanPekerjaan 
         Height          =   495
         Left            =   6000
         TabIndex        =   26
         Top             =   3550
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
         MICON           =   "frmHistoryServiceAll.frx":000C
         PICN            =   "frmHistoryServiceAll.frx":0028
         PICH            =   "frmHistoryServiceAll.frx":682B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   5280
         Top             =   2640
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
         Bindings        =   "frmHistoryServiceAll.frx":D3D1
         Height          =   1455
         Left            =   1200
         TabIndex        =   24
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   2566
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
      Begin TrueOleDBGrid70.TDBGrid ListNextJob 
         Height          =   3015
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5318
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Saran Job Order"
         Columns(0).DataField=   "SNJ"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Mekanik"
         Columns(1).DataField=   ""
         Columns(1).DropDown=   "TDBDropDown1"
         Columns(1).DropDown.vt=   8
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   36
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
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
      Begin VB.Image btnBatal 
         Height          =   375
         Index           =   0
         Left            =   7080
         Picture         =   "frmHistoryServiceAll.frx":D3E6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   0
      Picture         =   "frmHistoryServiceAll.frx":10D01
      ScaleHeight     =   7215
      ScaleWidth      =   495
      TabIndex        =   16
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   735
      Left            =   6720
      TabIndex        =   10
      Top             =   1200
      Width           =   3255
      Begin VB.OptionButton OptH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sublet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton OptH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Part"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton OptH 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Service"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   12
         Top             =   0
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label5 
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
         Index           =   10
         Left            =   1800
         TabIndex        =   15
         Top             =   0
         Width           =   60
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   2
         X1              =   0
         X2              =   1440
         Y1              =   1140
         Y2              =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View History"
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
         Index           =   9
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1215
      End
   End
   Begin ucXPButton.XPButton cmdExit 
      Height          =   495
      Left            =   10800
      TabIndex        =   17
      Top             =   6600
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
      MICON           =   "frmHistoryServiceAll.frx":1217D
      PICN            =   "frmHistoryServiceAll.frx":12199
      PICH            =   "frmHistoryServiceAll.frx":1266D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdExport 
      Height          =   495
      Left            =   9600
      TabIndex        =   18
      Top             =   6600
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
      MICON           =   "frmHistoryServiceAll.frx":12BB8
      PICN            =   "frmHistoryServiceAll.frx":12BD4
      PICH            =   "frmHistoryServiceAll.frx":13169
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   4215
      Index           =   2
      Left            =   480
      TabIndex        =   19
      Top             =   2280
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   7435
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tanggal"
      Columns(0).DataField=   "Tanggal"
      Columns(0).NumberFormat=   "Short Date"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "No Invoice"
      Columns(1).DataField=   "NoINV"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "No WO"
      Columns(2).DataField=   "NoWO"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Total"
      Columns(3).DataField=   "Total"
      Columns(3).NumberFormat=   "General Number"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Sublet"
      Columns(4).DataField=   "Sublet"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Qty"
      Columns(5).DataField=   "Qty"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Harga"
      Columns(6).DataField=   "Harga"
      Columns(6).NumberFormat=   "General Number"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2461"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2355"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2302"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2196"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=2646"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2540"
      Splits(0)._ColumnProps(16)=   "Column(3)._ColStyle=2"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=4948"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=4842"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=1244"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1138"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=2646"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=2540"
      Splits(0)._ColumnProps(29)=   "Column(6)._ColStyle=2"
      Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13,.alignment=1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
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
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   4215
      Index           =   1
      Left            =   360
      TabIndex        =   20
      Top             =   2280
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   7435
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tanggal"
      Columns(0).DataField=   "Tanggal"
      Columns(0).NumberFormat=   "Short Date"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "No Invoice"
      Columns(1).DataField=   "NoINV"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "No WO"
      Columns(2).DataField=   "NoWO"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Kode Part"
      Columns(3).DataField=   "KodePart"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Nama Part"
      Columns(4).DataField=   "NamaPart"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Qty"
      Columns(5).DataField=   "Qty"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
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
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2990"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2884"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2540"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=3889"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3784"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=5636"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=5530"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(21)=   "Column(5).Width=1244"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1138"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=78,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=75,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=76,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=77,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
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
   Begin TrueOleDBGrid70.TDBGrid Grid 
      Height          =   4215
      Index           =   0
      Left            =   480
      TabIndex        =   21
      Top             =   2280
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   7435
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tanggal"
      Columns(0).DataField=   "Tanggal"
      Columns(0).NumberFormat=   "Short Date"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "No Invoice"
      Columns(1).DataField=   "NoINV"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Total Invoice"
      Columns(2).DataField=   "TotInv"
      Columns(2).NumberFormat=   "Currency"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "No WO"
      Columns(3).DataField=   "NoWO"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "KM"
      Columns(4).DataField=   "KM"
      Columns(4).NumberFormat=   "General Number"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "SA"
      Columns(5).DataField=   "SA"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "MECH"
      Columns(6).DataField=   "MECH"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Job Order"
      Columns(7).DataField=   "JobOrder"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "Pesan Khusus"
      Columns(8).DataField=   "Pesan"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Saran"
      Columns(9).DataField=   "Saran"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "Catatan"
      Columns(10).DataField=   "Catat"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "Permintaan"
      Columns(11).DataField=   "Minta"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).Caption=   "Keluhan"
      Columns(12).DataField=   "Keluh"
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).Caption=   "Dealer Service"
      Columns(13).DataField=   "Dealer"
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).Caption=   "CR 7"
      Columns(14).DataField=   "SARANPERBAIKAN"
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).Caption=   "pilihdealer"
      Columns(15).DataField=   "pilihdealer"
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
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=16"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2461"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2355"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2461"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2355"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=2646"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2540"
      Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(14)=   "Column(3).Width=2302"
      Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2196"
      Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(18)=   "Column(4).Width=2302"
      Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2196"
      Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(22)=   "Column(5).Width=1244"
      Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=1138"
      Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(26)=   "Column(6).Width=1244"
      Splits(0)._ColumnProps(27)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(6)._WidthInPix=1138"
      Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(30)=   "Column(7).Width=6165"
      Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=6059"
      Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(34)=   "Column(8).Width=4419"
      Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=4313"
      Splits(0)._ColumnProps(37)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(38)=   "Column(9).Width=4419"
      Splits(0)._ColumnProps(39)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(9)._WidthInPix=4313"
      Splits(0)._ColumnProps(41)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(42)=   "Column(10).Width=4419"
      Splits(0)._ColumnProps(43)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(44)=   "Column(10)._WidthInPix=4313"
      Splits(0)._ColumnProps(45)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(46)=   "Column(11).Width=4419"
      Splits(0)._ColumnProps(47)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(48)=   "Column(11)._WidthInPix=4313"
      Splits(0)._ColumnProps(49)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(50)=   "Column(12).Width=4419"
      Splits(0)._ColumnProps(51)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(12)._WidthInPix=4313"
      Splits(0)._ColumnProps(53)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(54)=   "Column(13).Width=4419"
      Splits(0)._ColumnProps(55)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(56)=   "Column(13)._WidthInPix=4313"
      Splits(0)._ColumnProps(57)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(58)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(59)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(60)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(61)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(62)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(63)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(64)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(65)=   "Column(15).Visible=0"
      Splits(0)._ColumnProps(66)=   "Column(15).Order=16"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=82,.parent=13,.alignment=1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=50,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=46,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=43,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=44,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=45,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=28,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=25,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=26,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=27,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=32,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=29,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=30,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=31,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=94,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=91,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=92,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=93,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=98,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=95,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=96,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=97,.parent=17"
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
   Begin ucXPButton.XPButton cmdAktifkanJob 
      Height          =   495
      Left            =   7800
      TabIndex        =   22
      Top             =   6600
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
      MICON           =   "frmHistoryServiceAll.frx":137A4
      PICN            =   "frmHistoryServiceAll.frx":137C0
      PICH            =   "frmHistoryServiceAll.frx":1A63A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdRef 
      Height          =   495
      Left            =   5760
      TabIndex        =   35
      Top             =   0
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
      MICON           =   "frmHistoryServiceAll.frx":219B4
      PICN            =   "frmHistoryServiceAll.frx":219D0
      PICH            =   "frmHistoryServiceAll.frx":21F4C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape3 
      BackStyle       =   1  'Opaque
      Height          =   270
      Left            =   5640
      Top             =   0
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Warna"
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
      Index           =   12
      Left            =   720
      TabIndex        =   34
      Top             =   1560
      Width           =   630
   End
   Begin VB.Label Label5 
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
      Index           =   11
      Left            =   2520
      TabIndex        =   33
      Top             =   1560
      Width           =   60
   End
   Begin VB.Label Text 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   2640
      TabIndex        =   32
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Model"
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
      Index           =   8
      Left            =   720
      TabIndex        =   31
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label5 
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
      Index           =   7
      Left            =   2520
      TabIndex        =   30
      Top             =   1320
      Width           =   60
   End
   Begin VB.Label Text 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   2640
      TabIndex        =   29
      Top             =   1320
      Width           =   1755
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Rangka"
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
      Index           =   6
      Left            =   840
      TabIndex        =   28
      Top             =   90
      Width           =   1035
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      BorderWidth     =   8
      X1              =   495
      X2              =   11955
      Y1              =   6480
      Y2              =   6480
   End
   Begin VB.Label Text 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Polisi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   2640
      TabIndex        =   9
      Top             =   1800
      Width           =   915
   End
   Begin VB.Label Text 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Rangka"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   1140
   End
   Begin VB.Label Text 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   1755
   End
   Begin VB.Label Label5 
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
      Index           =   5
      Left            =   2520
      TabIndex        =   6
      Top             =   1800
      Width           =   60
   End
   Begin VB.Label Label5 
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
      Left            =   2520
      TabIndex        =   5
      Top             =   840
      Width           =   60
   End
   Begin VB.Label Label5 
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
      Left            =   2520
      TabIndex        =   4
      Top             =   1080
      Width           =   60
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Polisi"
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
      Left            =   720
      TabIndex        =   3
      Top             =   1800
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No Rangka"
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
      TabIndex        =   2
      Top             =   840
      Width           =   1035
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nama Customer"
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
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: INFORMASI - HISTORY SERVICE ::."
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
      Index           =   1
      Left            =   4320
      TabIndex        =   0
      Top             =   540
      Width           =   3975
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000080FF&
      BorderWidth     =   8
      X1              =   525
      X2              =   12005
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Width           =   11655
   End
End
Attribute VB_Name = "frmHistoryServiceAll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsX As ADODB.Recordset
Private rsXp As ADODB.Recordset
Private rsXs As ADODB.Recordset
Private rsP As New ADODB.Recordset
Private rsS As New ADODB.Recordset

''20180928|DEV ''PENAMBAHAN DATA INSPECTION
Private rsIns As New ADODB.Recordset
Private rsInspection As ADODB.Recordset
Public isOrder As Boolean
Public noBaris As Integer
Public modelKendaraan As String
Public pilihNoWo As String
Public pilihDealer As String


Private Sub btnBatal_Click(Index As Integer)
    If MsgBox("Apakah yakin menutup Saran Next Job?", vbYesNo, "Menutup Saran Next Job") = vbYes Then
        frmeGantiNoPol.Visible = False
    End If
        
End Sub

Private Sub cmdAktifkanJob_Click()
          
1        On Error GoTo cmdAktifkanJob_Click_Error

          Dim sqlStr As String
2         frmeGantiNoPol.Visible = True
3         cmdAktifkanJob.Enabled = False
          'lakukan pengisian data grid
          
          ''20180515REV01'INSPECTION
4         sqlStr = "SELECT top 1 * FROM [" & dbReplikasiHO & "]..INSPECTION WHERE NOWO = '" & pilihNoWo & "' AND DEALER = '" & pilihDealer & "' order by tglentry desc"
5         Set rsX = objO.FetchData(sqlStr)
          
6         If Not rsX.EOF Then
7             Call CreateTableInspection
              Dim sHelp As String
              Dim i As Integer
8             For i = 1 To 7
9                 sHelp = "SNJ" & i
10                If rsX(sHelp) <> "" Then
11                    rsIns.AddNew
12                    rsIns("SNJ") = rsX(sHelp)
13                    rsIns("MEK") = ""
14                    rsIns("Status") = False
15                    rsIns.Update
16                End If
17            Next
18            ListNextJob.DataSource = rsIns
19            Call headColumns
20        End If
          
21        With Adodc1
22            .ConnectionString = adoCon
23            .CommandType = adCmdText
24            .RecordSource = "SELECT m.MECH,m.[NAME],m.LEVMECH FROM mekanik m WITH (NOLOCK) INNER JOIN t_mekanik t ON t.mech = m.mech WHERE t.tanggal = '" & Format(isnow, "DD/MM/YYYY") & "' and masuk = '1' ORDER BY m.levmech,m.mech"
25            .Refresh
26        End With
27        TDBDropDown1.DataSource = Adodc1


28       On Error GoTo 0
29       Exit Sub

cmdAktifkanJob_Click_Error:

30        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdAktifkanJob_Click of Form frmHistoryServiceAll"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Err_Export
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

    If OptH(0).value = True Then
        grid(0).ExportToFile onVal, True
    ElseIf OptH(1).value = True Then
        grid(1).ExportToFile onVal, True
    ElseIf OptH(2).value = True Then
        grid(2).ExportToFile onVal, True
    End If
    
    sSQL = "UPDATE profilbengkel..logViewHistory SET userDownload = '" & gSA & "',tgldownload = getdate() WHERE NOrangka = '" & NulSID(Text(0)) & "' and userview = '" & gSA & "' and convert(varchar(20),tglview,101) = '" & Format(isnow, "mm/dd/yyyy") & "'"
    Call objO.InsertData(sSQL)
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            

            Er.show (vbModal)
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub cmdRef_Click()
    If NulSID(txtnoka.Text) <> "" Then
        Dim rsHistory As ADODB.Recordset
        sSQL = "SELECT NO_CHASSIS,CUSTOMER,mdl + ' | ' + FMDL model,COLOR,NO_POLISI FROM CUSTDAT WHERE NO_CHASSIS = '" & NulSID(txtnoka.Text) & "'"
        Set rsHistory = objO.FetchData(sSQL)
        If Not rsHistory.EOF Then
            Call loadData(rsHistory!NO_CHASSIS, rsHistory!customer, rsHistory!NO_POLISI, rsHistory!model, rsHistory!Color)
            sSQL = "INSERT INTO PROFILBENGKEL..logViewHistory (norangka,userView,tglView,CompName,versiNis,userDownload) "
            sSQL = sSQL & " VALUES ('" & NulSID(rsHistory!NO_CHASSIS) & "','" & gSA & "','" & Format(isnow, "YYYY-MM-DD HH:MM:SS") & "','" & ComptName & "','" & gVersi & "','')"
            Call objO.InsertData(sSQL)
        Else
            MsgBox "Norangka tidak ditemukan dalam sistem bengkel " & gNmCbg, vbInformation
            txtnoka.Text = ""
            clearForm
        End If
    End If
End Sub

Private Sub cmdSimpanPekerjaan_Click()
    If Not rsIns.EOF Then
    rsIns.MoveFirst
    Do While Not rsIns.EOF
        If rsIns("status") Then
            If isOrder Then
                If isiJobOrder(Trim(rsIns("SNJ")), rsIns("MEK")) = False Then
                    Exit Sub
                End If
            Else
                Call isiJobBook(Trim(rsIns("SNJ")), rsIns("MEK"))
            End If
            noBaris = noBaris + 1
        End If
        rsIns.MoveNext
    Loop
    End If
    MsgBox "Data berhasil ditambahkan", vbInformation, "Informasi Penambahan Job Order"
    frmeGantiNoPol.Visible = False
End Sub

Private Sub loadData(noRangka As String, namacustomer As String, NoPolisi As String, model As String, warna As String)
    Text(0).Caption = NulSID(noRangka)
    Text(1).Caption = NulSID(namacustomer)
    Text(2).Caption = NulSID(NoPolisi)
    Text(3).Caption = NulSID(model)
    Text(4).Caption = NulSID(warna)
    
    grid(0).Visible = False
    grid(1).Visible = False
    grid(2).Visible = False
    
    If gTag = "ONLINE" Then  'On Line HO
        If Log.RPLHistori = 1 Then
            Call LoadDataReplikasi
        Else
            Call LoadDataHistory
        End If
'        Call LoadDataHistory
        Label5(9).Caption = "View History HO"
    Else
        Call LoadDataHistoryLocal
        Label5(9).Caption = "View History Local"
    End If
    
    
        cmdAktifkanJob.Visible = False
End Sub

Private Sub LoadDataReplikasi()
          Dim nINV As String
          Dim sqlStr As String
          Dim keteranganHide As String
          
1         If OptH(0).value = True Then       'Service
          
              '' 20220210DEV
2             If Log.hideService = 1 Then
3                 keteranganHide = "" '" AND D.KM > 10 "
4             End If
5             grid(0).Visible = True
              '' ------------------------------------------------
              '' 20180912 PENAMBAHAN JOIN DENGAN NORANGKA
              '' ------------------------------------------------
              ''20180928|DEV ''PENAMBAHAN DATA INSPECTION
              
6             sqlStr = "select distinct isNull(J.CATAT,'') as CATAT , D.DATE as Tanggal,D.NOINV,D.TOTAL as TotInv, D.NO_WO as NOWO, D.KM,D.FRONT as SA ,D.MECH ," _
                   & " REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' " _
                   & " + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder, D.REMARK1 as Pesan, D.SUGGEST as Saran, " _
                   & " D.REQ1 + '; ' + D.REQ2 + '; ' + D.REQ3 + '; ' + D.REQ4 + '; ' + D.REQ5 as Minta,D.FS_NM_DEALER as Dealer, " _
                   & " REPLACE(REPLACE(REPLACE(D.CMPL1 + '; ' + D.CMPL2 + '; ' + D.CMPL3 + '; ' + D.CMPL4 + '; ' + D.CMPL5 ,'; ;',''),'; ;',''),'; ;','') as Keluh, " _
                   & " ISNULL(REPLACE(REPLACE(REPLACE(I.SNJ1 + '; ' + I.SNJ2 + '; ' + I.SNJ3 + '; ' + I.SNJ4 + '; ' + I.SNJ5 + '; ' + I.SNJ6 + '; '  + I.SNJ7,'; ;',''),'; ;',''),'; ;',''),'') as SARANPERBAIKAN,isnull(i.dealer,'') as pilihdealer From " _
                   & " [" & dbReplikasiHO & "]..T_INVOICENEW D WITH (NOLOCK) Left Join [" & dbReplikasiHO & "]..JobDTA J WITH (NOLOCK) On D.No_WO = J.NO_WO AND D.NO_CHASSIS = J.NO_CHASSIS AND J.NOINV = Replace(Replace(D.NOINV,'-A',''),'-B','') " _
                   & " left join [" & dbReplikasiHO & "]..INSPECTION I WITH (NOLOCK) ON I.DEALER = D.FS_KD_DEALER AND I.NoWo = D.NO_WO " _
                   & " where D.NO_CHASSIS ='" & Text(0).Caption & "' AND D.FS_LOG_DEL = '' order by D.DATE DESC, D.NOINV DESC, D.KM DESC"
                                         
              '' 20181003 | REQ | MB THANIA | GANTI QUERY HISTORY SERVICE
              '' ====================================================================================
7             sqlStr = " select distinct isNull(J.CATAT,'') as CATAT , D.DATE as Tanggal,D.NOINV,D.TOTAL as TotInv, D.NO_WO as NOWO, D.KM,D.FRONT as SA ,D.MECH ,"
8             sqlStr = sqlStr & " REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; '"
9             sqlStr = sqlStr & " + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder, D.REMARK1 as Pesan, D.SUGGEST as Saran,"
10            sqlStr = sqlStr & " D.REQ1 + '; ' + D.REQ2 + '; ' + D.REQ3 + '; ' + D.REQ4 + '; ' + D.REQ5 as Minta,D.FS_NM_DEALER as Dealer,"
11            sqlStr = sqlStr & " REPLACE(REPLACE(REPLACE(D.CMPL1 + '; ' + D.CMPL2 + '; ' + D.CMPL3 + '; ' + D.CMPL4 + '; ' + D.CMPL5 ,'; ;',''),'; ;',''),'; ;','') as Keluh,"
12            sqlStr = sqlStr & " REPLACE(REPLACE(REPLACE(left(i.SNJ1, CHARINDEX('|', i.SNJ1)) + '; ' + left(i.SNJ2, CHARINDEX('|', i.SNJ2)) + '; ' + left(i.SNJ3, CHARINDEX('|', i.SNJ3)) + '; ' + left(i.SNJ4, CHARINDEX('|', i.SNJ4)) + '; ' + left(i.SNJ5, CHARINDEX('|', i.SNJ5)) + '; ' + left(i.SNJ6, CHARINDEX('|', i.SNJ6)) + '; ' + left(i.SNJ7, CHARINDEX('|', i.SNJ7)),'; ;',''),'; ;',''),'; ;','')  as SARANPERBAIKAN,isnull(i.dealer,'') as pilihdealer"
13            sqlStr = sqlStr & " from [" & dbReplikasiHO & "]..T_INVOICENEW D WITH (NOLOCK)"
14            sqlStr = sqlStr & " Left Join [" & dbReplikasiHO & "]..JobDTA J WITH (NOLOCK) On D.No_WO = J.NO_WO AND D.NO_CHASSIS = J.NO_CHASSIS and J.NOINV = Replace(Replace(D.NOINV,'-A',''),'-B','')"
15            sqlStr = sqlStr & " left join [" & dbReplikasiHO & "]..INSPECTION I WITH (NOLOCK) On i.NoWo = d.NO_WO and d.FS_KD_DEALER = i.DEALER"
16            sqlStr = sqlStr & " where D.NO_CHASSIS ='" & Text(0).Caption & "' AND D.FS_LOG_DEL = '' " & keteranganHide & " order by D.DATE DESC, D.NOINV DESC, D.KM DESC"
                   
                                 
17            Set rsX = objO.FetchData(sqlStr)
18            If rsX.EOF Then Exit Sub
19            Set grid(0).DataSource = rsX
20        ElseIf OptH(1).value = True Then   'Part
21            grid(1).Visible = True
              '' 20220210DEV
22            If Log.hideService = 1 Then
23                keteranganHide = "" '" AND A.KM > 10 "
24            End If
25            createTablePart True
26            sqlStr = "select [Date] as Tanggal, A.NoInv, A.no_wo as NOWO, KODEITEM, Keterangan, QTY " & _
                       "from [" & dbReplikasiHO & "]..historynew as A WITH (NOLOCK) inner join [" & dbReplikasiHO & "]..t_partnew as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') and B.FS_LOG_DEL = '' " & _
                       "where NO_CHASSIS = '" & Text(0).Caption & "' and B.NOINV <> '' and A.FS_LOG_DEL = '' " & keteranganHide & " order by [DATE] DESC, KODEITEM"
27            Set rsXp = objO.FetchData(sqlStr)
28            If rsXp.RecordCount > 0 Then
29                rsXp.MoveFirst
30                nINV = ""
31                Do Until rsXp.EOF
32                    If Trim(rsXp.Fields(1).value) <> nINV Then
33                        rsP.AddNew
                          ''2016-09-22 || penambahan format tanggal pada result value
34                        rsP!Tanggal = Format(rsXp.Fields(0).value, "dd/mm/yyyy")
                          ''rsP!Tanggal = rsXp.Fields(0).value
35                        rsP!noInv = rsXp.Fields(1).value
36                        rsP!noWO = rsXp.Fields(2).value
37                        rsP!kodePart = rsXp.Fields(3).value
38                        rsP!namaPart = rsXp.Fields(4).value
39                        rsP!qty = Val(rsXp.Fields(5).value)
40                    Else
41                        rsP.AddNew
42                        rsP!Tanggal = ""
43                        rsP!noInv = ""
44                        rsP!noWO = ""
45                        rsP!kodePart = rsXp.Fields(3).value
46                        rsP!namaPart = rsXp.Fields(4).value
47                        rsP!qty = Val(rsXp.Fields(5).value)
48                    End If
49                    rsP.Update
50                    If Not rsXp.EOF Then nINV = Trim(rsXp.Fields(1).value)
51                    rsXp.MoveNext
52                Loop
53            End If
54            Set grid(1).DataSource = rsP
55        ElseIf OptH(2).value = True Then   'Sublet
56            grid(2).Visible = True
              '' 20220210DEV
57            If Log.hideService = 1 Then
58                keteranganHide = "" '" AND A.KM > 10 "
59            End If
60            CreateTableSublet True
61            sqlStr = "SELECT [DATE], A.NOINV, A.NO_WO, SBLT, NAMABARANG, QTY, HARGA FROM [" & dbReplikasiHO & "]..historynew A WITH (NOLOCK) " & _
                       "INNER JOIN [" & dbReplikasiHO & "]..HIST_SUBLETNEW B WITH (NOLOCK) ON A.no_wo = B.no_wo and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') AND A.FS_LOG_DEL = '' " & _
                       "WHERE A.NO_CHASSIS = '" & Text(0).Caption & "' AND B.noinv <> '' " & keteranganHide & " order by [DATE] DESC, KODEITEM"
              
62            Set rsXs = objO.FetchData(sqlStr)
63            If rsXs.RecordCount > 0 Then
64                rsXs.MoveFirst
65                nINV = ""
66                Do Until rsXs.EOF
67                    If Trim(rsXs.Fields(1).value) <> nINV Then
68                        rsS.AddNew
                          ''2016-09-22 || penambahan format tanggal pada result value
69                        rsS!Tanggal = Format(rsXs.Fields(0).value, "dd/mm/yyyy")
      '                    rsS!Tanggal = rsXs.Fields(0).value
70                        rsS!noInv = rsXs.Fields(1).value
71                        rsS!noWO = rsXs.Fields(2).value
72                        rsS!total = rsXs.Fields(3).value
73                        rsS!SUBLET = rsXs.Fields(4).value
74                        rsS!qty = Val(rsXs.Fields(5).value)
75                        rsS!Harga = Val(rsXs.Fields(6).value)
76                    Else
77                        rsS.AddNew
78                        rsS!Tanggal = ""
79                        rsS!noInv = ""
80                        rsS!noWO = ""
81                        rsS!total = ""
82                        rsS!SUBLET = rsXs.Fields(4).value
83                        rsS!qty = Val(rsXs.Fields(5).value)
84                        rsS!Harga = Val(rsXs.Fields(6).value)
85                    End If
86                    rsS.Update
87                    If Not rsXs.EOF Then nINV = Trim(rsXs.Fields(1).value)
88                    rsXs.MoveNext
89                Loop
90            End If
91            Set grid(2).DataSource = rsS
92        End If
End Sub


Private Sub LoadDataHistory()
          Dim nINV As String
          Dim sqlStr As String
          Dim keteranganHide As String
          
1         If OptH(0).value = True Then       'Service
2             grid(0).Visible = True
              '' 20220210DEV
3             If Log.hideService = 1 Then
4                 keteranganHide = "" '" AND D.KM > 10 "
5             End If
              'sqlStr = "select D.DATE as Tanggal,D.NOINV,D.TOTAL as TotInv, D.NO_WO as NOWO, D.KM,D.FRONT as SA ,D.MECH ," _
                   & " REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder, D.REMARK1 as Pesan, D.SUGGEST as Saran," _
                   & " REPLACE(REPLACE(REPLACE(D.CMPL1 + '; ' + D.CMPL2 + '; ' + D.CMPL3 + '; ' + D.CMPL4 + '; ' + D.CMPL5,'; ;',''),'; ;',''),'; ;','') as Complain, D.Catat as Catat," _
                   & " D.REQ1 + '; ' + D.REQ2 + '; ' + D.REQ3 + '; ' + D.REQ4 + '; ' + D.REQ5 as Minta,D.FS_NM_DEALER as Dealer from  " _
                   & " T_INVOICENEW D where D.NO_CHASSIS ='" & Text(0).Caption & "' AND D.FS_LOG_DEL = '' order by D.DATE DESC, D.NOINV DESC, D.KM DESC"
                  
6             sqlStr = "select distinct isNull(J.CATAT,'') as CATAT , D.DATE as Tanggal,D.NOINV,D.TOTAL as TotInv, D.NO_WO as NOWO, D.KM,D.FRONT as SA ,D.MECH ," _
                   & " REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' " _
                   & " + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder, D.REMARK1 as Pesan, D.SUGGEST as Saran, " _
                   & " D.REQ1 + '; ' + D.REQ2 + '; ' + D.REQ3 + '; ' + D.REQ4 + '; ' + D.REQ5 as Minta,D.FS_NM_DEALER as Dealer, " _
                   & " REPLACE(REPLACE(REPLACE(D.CMPL1 + '; ' + D.CMPL2 + '; ' + D.CMPL3 + '; ' + D.CMPL4 + '; ' + D.CMPL5 ,'; ;',''),'; ;',''),'; ;','') as Keluh from " _
                   & " T_INVOICENEW D WITH (NOLOCK) Left Join JobDTA J WITH (NOLOCK) On D.No_WO = J.NO_WO where D.NO_CHASSIS ='" & Text(0).Caption & "' AND D.FS_LOG_DEL = '' " & keteranganHide & " order by D.DATE DESC, D.NOINV DESC, D.KM DESC"
                  
7             Set rsX = objO.FetchDataHO(sqlStr)
8             If rsX.EOF Then Exit Sub
9             Set grid(0).DataSource = rsX
10        ElseIf OptH(1).value = True Then   'Part
11            grid(1).Visible = True
12            createTablePart True
              '' 20220210DEV
13            If Log.hideService = 1 Then
14                keteranganHide = "" '" AND A.KM > 10 "
15            End If
16            sqlStr = "select [Date] as Tanggal, A.NoInv, A.no_wo as NOWO, KODEITEM, Keterangan, QTY " & _
                       "from historynew as A WITH (NOLOCK) inner join t_partnew as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') and B.FS_LOG_DEL = '' " & _
                       "where NO_CHASSIS = '" & Text(0).Caption & "' and B.NOINV <> '' and A.FS_LOG_DEL = '' " & keteranganHide & " order by [DATE] DESC, KODEITEM"
              
17            Set rsXp = objO.FetchDataHO(sqlStr)
18            If rsXp.RecordCount > 0 Then
19                rsXp.MoveFirst
20                nINV = ""
21                Do Until rsXp.EOF
22                    If Trim(rsXp.Fields(1).value) <> nINV Then
23                        rsP.AddNew
                          ''2016-09-22 || penambahan format tanggal pada result value
24                        rsP!Tanggal = Format(rsXp.Fields(0).value, "dd/mm/yyyy")
                          ''rsP!Tanggal = rsXp.Fields(0).value
25                        rsP!noInv = rsXp.Fields(1).value
26                        rsP!noWO = rsXp.Fields(2).value
27                        rsP!kodePart = rsXp.Fields(3).value
28                        rsP!namaPart = rsXp.Fields(4).value
29                        rsP!qty = Val(rsXp.Fields(5).value)
30                    Else
31                        rsP.AddNew
32                        rsP!Tanggal = ""
33                        rsP!noInv = ""
34                        rsP!noWO = ""
35                        rsP!kodePart = rsXp.Fields(3).value
36                        rsP!namaPart = rsXp.Fields(4).value
37                        rsP!qty = Val(rsXp.Fields(5).value)
38                    End If
39                    rsP.Update
40                    If Not rsXp.EOF Then nINV = Trim(rsXp.Fields(1).value)
41                    rsXp.MoveNext
42                Loop
43            End If
44            Set grid(1).DataSource = rsP
45        ElseIf OptH(2).value = True Then   'Sublet
46            grid(2).Visible = True
47            CreateTableSublet True
              '' 20220210DEV
48            If Log.hideService = 1 Then
49                keteranganHide = "" '" AND A.KM > 10 "
50            End If
51            sqlStr = "SELECT [DATE], A.NOINV, A.NO_WO, SBLT, NAMABARANG, QTY, HARGA FROM historynew A WITH (NOLOCK) " & _
                       "INNER JOIN HIST_SUBLETNEW B WITH (NOLOCK) ON A.no_wo = B.no_wo and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') AND A.FS_LOG_DEL = '' " & _
                       "WHERE A.NO_CHASSIS = '" & Text(0).Caption & "' AND B.noinv <> '' " & keteranganHide & " order by [DATE] DESC, KODEITEM"
              
52            Set rsXs = objO.FetchDataHO(sqlStr)
53            If rsXs.RecordCount > 0 Then
54                rsXs.MoveFirst
55                nINV = ""
56                Do Until rsXs.EOF
57                    If Trim(rsXs.Fields(1).value) <> nINV Then
58                        rsS.AddNew
                          ''2016-09-22 || penambahan format tanggal pada result value
59                        rsS!Tanggal = Format(rsXs.Fields(0).value, "dd/mm/yyyy")
      '                    rsS!Tanggal = rsXs.Fields(0).value
60                        rsS!noInv = rsXs.Fields(1).value
61                        rsS!noWO = rsXs.Fields(2).value
62                        rsS!total = rsXs.Fields(3).value
63                        rsS!SUBLET = rsXs.Fields(4).value
64                        rsS!qty = Val(rsXs.Fields(5).value)
65                        rsS!Harga = Val(rsXs.Fields(6).value)
66                    Else
67                        rsS.AddNew
68                        rsS!Tanggal = ""
69                        rsS!noInv = ""
70                        rsS!noWO = ""
71                        rsS!total = ""
72                        rsS!SUBLET = rsXs.Fields(4).value
73                        rsS!qty = Val(rsXs.Fields(5).value)
74                        rsS!Harga = Val(rsXs.Fields(6).value)
75                    End If
76                    rsS.Update
77                    If Not rsXs.EOF Then nINV = Trim(rsXs.Fields(1).value)
78                    rsXs.MoveNext
79                Loop
80            End If
81            Set grid(2).DataSource = rsS
82        End If
End Sub

Private Sub LoadDataHistoryLocal()
          Dim nINV As String
          Dim sqlStr As String
          Dim keteranganHide As String
          
1         If OptH(0).value = True Then       'Service
              '' 20220210DEV
2             If Log.hideService = 1 Then
3                 keteranganHide = "" '" AND D.KM > 10 "
4             End If
5             grid(0).Visible = True
              'sqlStr = "select D.DATE as Tanggal,D.NOINV,D.TOTAL as TotInv, D.NO_WO as NOWO, D.KM,D.FRONT as SA ,D.MECH ," _
                   & " REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder, D.REMARK1 as Pesan, D.SUGGEST as Saran," _
                   & " D.REQ1 + '; ' + D.REQ2 + '; ' + D.REQ3 + '; ' + D.REQ4 + '; ' + D.REQ5 as Minta,D.FS_NM_DEALER as Dealer from  " _
                   & " T_INVOICE D where D.NO_CHASSIS ='" & Text(0).Caption & "' AND D.FS_LOG_DEL = '' order by D.DATE DESC, D.NOINV DESC, D.KM DESC"
                   
6             sqlStr = "select IsNull(J.CATAT,'') as CATAT, D.DATE as Tanggal,D.NOINV,D.TOTAL as TotInv, D.NO_WO as NOWO, D.KM,D.FRONT as SA ,D.MECH ," _
                   & " REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' " _
                   & " + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder, D.REMARK1 as Pesan, D.SUGGEST as Saran, " _
                   & " D.REQ1 + '; ' + D.REQ2 + '; ' + D.REQ3 + '; ' + D.REQ4 + '; ' + D.REQ5 as Minta,D.FS_NM_DEALER as Dealer, " _
                   & " REPLACE(REPLACE(REPLACE(D.CMPL1 + '; ' + D.CMPL2 + '; ' + D.CMPL3 + '; ' + D.CMPL4 + '; ' + D.CMPL5 ,'; ;',''),'; ;',''),'; ;','') as Keluh from " _
                   & " T_INVOICE D WITH (NOLOCK) Left Join JobDTA J WITH (NOLOCK) On D.No_WO = J.NO_WO where D.NO_CHASSIS ='" & Text(0).Caption & "' AND D.FS_LOG_DEL = '' " & keteranganHide & " order by D.DATE DESC, D.NOINV DESC, D.KM DESC"
                   
                   
7             Set rsX = objO.FetchData(sqlStr)
8             Set grid(0).DataSource = rsX
9         ElseIf OptH(1).value = True Then   'Part
10            grid(1).Visible = True
11            createTablePart True
              '' 20220210DEV
12            If Log.hideService = 1 Then
13                keteranganHide = "" '" AND A.KM > 10 "
14            End If
15            sqlStr = "select [Date] as Tanggal, A.NoInv, A.no_wo as NOWO, KODEITEM, Keterangan, QTY " & _
                       "from history as A WITH (NOLOCK) inner join t_part as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') and B.FS_LOG_DEL = '' " & _
                       "where NO_CHASSIS = '" & Text(0).Caption & "' and B.NOINV <> '' and A.FS_LOG_DEL = '' " & keteranganHide & " order by [DATE] DESC, KODEITEM"
              
16            Set rsXp = objO.FetchData(sqlStr)
17            If rsXp.RecordCount > 0 Then
18                rsXp.MoveFirst
19                nINV = ""
20                Do Until rsXp.EOF
21                    If Trim(rsXp.Fields(1).value) <> nINV Then
22                        rsP.AddNew
                          ''2016-09-22 || penambahan format tanggal pada result value
23                        rsP!Tanggal = Format(rsXp.Fields(0).value, "dd/mm/yyyy")
      '                    rsP!Tanggal = rsXp.Fields(0).value
24                        rsP!noInv = rsXp.Fields(1).value
25                        rsP!noWO = rsXp.Fields(2).value
26                        rsP!kodePart = rsXp.Fields(3).value
27                        rsP!namaPart = rsXp.Fields(4).value
28                        rsP!qty = Val(rsXp.Fields(5).value)
29                    Else
30                        rsP.AddNew
31                        rsP!Tanggal = ""
32                        rsP!noInv = ""
33                        rsP!noWO = ""
34                        rsP!kodePart = rsXp.Fields(3).value
35                        rsP!namaPart = rsXp.Fields(4).value
36                        rsP!qty = Val(rsXp.Fields(5).value)
37                    End If
38                    rsP.Update
39                    If Not rsXp.EOF Then nINV = Trim(rsXp.Fields(1).value)
40                    rsXp.MoveNext
41                Loop
42            End If
43            Set grid(1).DataSource = rsP
44        ElseIf OptH(2).value = True Then   'Sublet
45            grid(2).Visible = True
46            CreateTableSublet True
              '' 20220210DEV
47            If Log.hideService = 1 Then
48                keteranganHide = "" '" AND A.KM > 10 "
49            End If
50            sqlStr = "SELECT [DATE], A.NOINV, A.NO_WO, SBLT, NAMABARANG, QTY, HARGA FROM history A WITH (NOLOCK) " & _
                       "INNER JOIN HIST_SUBLET B WITH (NOLOCK) ON A.no_wo = B.no_wo and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') AND A.FS_LOG_DEL = '' " & _
                       "WHERE A.NO_CHASSIS = '" & Text(0).Caption & "' AND B.noinv <> '' " & keteranganHide & " order by [DATE] DESC, KODEITEM"
              
51            Set rsXs = objO.FetchData(sqlStr)
52            If rsXs.RecordCount > 0 Then
53                rsXs.MoveFirst
54                nINV = ""
55                Do Until rsXs.EOF
56                    If Trim(rsXs.Fields(1).value) <> nINV Then
57                        rsS.AddNew
                          ''2016-09-22 || penambahan format tanggal pada result value
58                        rsS!Tanggal = Format(rsXs.Fields(0).value, "dd/mm/yyyy")
                          
      '                    rsS!Tanggal = rsXs.Fields(0).value
59                        rsS!noInv = rsXs.Fields(1).value
60                        rsS!noWO = rsXs.Fields(2).value
61                        rsS!total = rsXs.Fields(3).value
62                        rsS!SUBLET = rsXs.Fields(4).value
63                        rsS!qty = Val(rsXs.Fields(5).value)
64                        rsS!Harga = Val(rsXs.Fields(6).value)
65                    Else
66                        rsS.AddNew
67                        rsS!Tanggal = ""
68                        rsS!noInv = ""
69                        rsS!noWO = ""
70                        rsS!total = ""
71                        rsS!SUBLET = rsXs.Fields(4).value
72                        rsS!qty = Val(rsXs.Fields(5).value)
73                        rsS!Harga = Val(rsXs.Fields(6).value)
74                    End If
75                    rsS.Update
76                    If Not rsXs.EOF Then nINV = Trim(rsXs.Fields(1).value)
77                    rsXs.MoveNext
78                Loop
79            End If
80            Set grid(2).DataSource = rsS
81        End If
End Sub

Private Sub createTablePart(Optional Edit As Boolean = False)
    On Error GoTo vb_error
    If rsP.State Then rsP.Close
    With rsP
        .Fields.Append "Tanggal", adVarChar, 15
        .Fields.Append "NoInv", adVarChar, 15
        .Fields.Append "NoWO", adVarChar, 15
        .Fields.Append "KodePart", adVarChar, 50
        .Fields.Append "NamaPart", adVarChar, 50
        .Fields.Append "Qty", adDouble, 9
        .Open
    End With
    Exit Sub
vb_error:
    err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub CreateTableSublet(Optional Edit As Boolean = False)
1         On Error GoTo vb_error
2         If rsS.State Then rsS.Close
3         Set rsS = Nothing
4         With rsS
5             .Fields.Append "Tanggal", adVarChar, 15
6             .Fields.Append "NoInv", adVarChar, 15
7             .Fields.Append "NoWO", adVarChar, 15
8             .Fields.Append "Total", adVarChar, 50
9             .Fields.Append "Sublet", adVarChar, 50
10            .Fields.Append "Qty", adDouble, 9
11            .Fields.Append "Harga", adDouble, 9
12            .Open
13        End With
14        Exit Sub
vb_error:
15        err.Raise err.Number, err.Source, err.Description
End Sub
 

Private Sub Form_Load()
    clearForm
End Sub

Private Sub Grid_RowColChange(Index As Integer, LastRow As Variant, ByVal LastCol As Integer)
    If grid(0).Visible Then

        If NulSID(grid(0).Columns("SARANPERBAIKAN").value) <> "" And NulSID(grid(0).Columns("pilihDealer").value) <> "" Then
            cmdAktifkanJob.Enabled = True
            pilihNoWo = Trim(grid(0).Columns("nowo").value)
            pilihDealer = NulSID(grid(0).Columns("pilihDealer").value)
            
        Else
            cmdAktifkanJob.Enabled = False
        End If
    Else
        cmdAktifkanJob.Enabled = False
    End If
End Sub

Private Sub OptH_Click(Index As Integer)
    grid(0).Visible = False
    grid(1).Visible = False
    grid(2).Visible = False
    If gTag = "ONLINE" Then  'On Line HO
        'Call LoadDataHistory
        If Log.RPLHistori = 1 Then
            Call LoadDataReplikasi
        Else
            Call LoadDataHistory
        End If
    Else
        Call LoadDataHistoryLocal
    End If
End Sub


Private Sub CreateTableInspection()
1         On Error GoTo vb_error
2         If rsIns.State Then rsIns.Close
3         Set rsIns = Nothing
4         With rsIns
5             .Fields.Append "SNJ", adVarChar, 100
6             .Fields.Append "MEK", adVarChar, 5
7             .Fields.Append "Status", adBoolean
8             .Open
9         End With
10        Exit Sub
vb_error:
11        err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub headColumns()
    
    With ListNextJob
        .Columns(0).DataField = "SNJ"
        .Columns(0).Width = 4500
        .Columns(1).DataField = "MEK"
        .Columns(1).Width = 1000
        .Columns(2).DataField = "Status"
        .Columns(2).Width = 1000
    End With
End Sub


Private Function isiJobOrder(fs_item As String, mekanik As String) As Boolean
          Dim jobRate As Double
          Dim Text46, Text44, Text50, Text11 As String
          Dim isPrediag As Boolean
          Dim i As Integer
          Dim job As String
          
1         fs_item = Replace(NulSID(fs_item), " - ", "|")
2         If InStr(1, NulSID(fs_item), "|") > 0 Then
3             fs_item = Mid(NulSID(fs_item), 1, InStr(1, NulSID(fs_item), "|") - 1)
          
4         End If
5         jobRate = objO.FetchData_Double("select DISTINCT RATE from [" & dbReplikasiHO & "]..HO_FLATE WITH (NOLOCK) where fs_item = '" & fs_item & "' AND mdl = '" & modelKendaraan & "' and FS_LOG_DEL = ''")
6         If jobRate = 0 Then
7             jobRate = objO.FetchData_Double("select DISTINCT RATE from HO_FLATE WITH (NOLOCK) where fs_item = '" & fs_item & "' AND mdl = '" & modelKendaraan & "' and FS_LOG_DEL = ''")
8         End If
9         With frmOrder
10            Text46 = .Text(46)
11            Text44 = .Text(44)
12            Text50 = .Text(50)
13            Text11 = .Text(11)
14            job = objO.FetchData_String("select top 1 JOB from [" & dbReplikasiHO & "]..HO_FLATE WITH (NOLOCK) where fs_item = '" & fs_item & "' AND mdl = '" & modelKendaraan & "' and FS_LOG_DEL = ''")
15            If job = "" Then
16                job = objO.FetchData_String("select top 1 JOB from HO_FLATE WITH (NOLOCK) where fs_item = '" & fs_item & "' AND mdl = '" & modelKendaraan & "' and FS_LOG_DEL = ''")
17            End If
18            If job = "" Then
19                MsgBox "Jenis Pekerjaan " & fs_item & " Tidak ada untuk model kendaraan ini!", vbCritical
20                isiJobOrder = False
21                Exit Function
22            End If
23            If .sgnPreDiag.Visible Then
24                isPrediag = True
25            Else
26                isPrediag = False
27            End If
28            .Text(50) = mekanik
              
29            .Text(56) = Format(CDbl(.Text(56)) + Val(hitungLabour(CDbl(jobRate), Trim(.Text(11)), isPrediag)), "#0,0")
30                With .dGrid
31                    .TextMatrix(noBaris, 1) = Trim(fs_item)
32                    .TextMatrix(noBaris, 2) = jobRate
33                    .TextMatrix(noBaris, 3) = False
34                    .TextMatrix(noBaris, 4) = Trim(job)
35                    .TextMatrix(noBaris, 5) = Val(hitungLabour(CDbl(jobRate), Trim(Text11), isPrediag))
36                    .TextMatrix(noBaris, 6) = Trim(mekanik)
37                End With
                 
38                .baris = noBaris + 1
                 
39                .ntRate = 0
40                .Text(114) = .ntRate
41                For i = 1 To 7
42                    If Trim(.dGrid.TextMatrix(i, 2)) <> "" Then
                          
                        '  If Log.RPLHistori = 0 Then
43                            .ntRate = .ntRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(.Text(46)) & "' AND MDL = '" & Trim(.Text(10)) & "'")
                         ' Else
                        '      .ntRate = .ntRate + objO.FetchData_Double("select Rate from [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(.text(46)) & "' AND MDL = '" & Trim(.text(10)) & "'")
                         ' End If
44                    End If
45                Next
46                .Text(114) = .ntRate
                             
47                .Text(61) = Format(CDbl(.Text(56)) + CDbl(.Text(57)) + CDbl(.Text(58)) + CDbl(.Text(59)) + CDbl(.Text(60)) + CDbl(.Text(77)), "#0,#")
              
48        End With
49        isiJobOrder = True
End Function


Private Sub isiJobBook(fs_item As String, mekanik As String)
          Dim jobRate As Double
1         fs_item = Replace(NulSID(fs_item), " - ", "|")
2         If InStr(1, NulSID(fs_item), "|") > 0 Then
3             fs_item = Mid(NulSID(fs_item), 1, InStr(1, NulSID(fs_item), "|") - 1)
4         End If
          'jobRate = objO.FetchData_Double("select DISTINCT RATE from HO_FLATE WITH (NOLOCK) where fs_item = '" & fs_item & "' AND mdl = '" & modelKendaraan & "' and FS_LOG_DEL = ''")
5         jobRate = objO.FetchData_Double("select DISTINCT RATE from [" & dbReplikasiHO & "]..HO_FLATE WITH (NOLOCK) where fs_item = '" & fs_item & "' AND mdl = '" & modelKendaraan & "' and FS_LOG_DEL = ''")
6         If jobRate = 0 Then
7             jobRate = objO.FetchData_Double("select DISTINCT RATE from HO_FLATE WITH (NOLOCK) where fs_item = '" & fs_item & "' AND mdl = '" & modelKendaraan & "' and FS_LOG_DEL = ''")
8         End If
          'With frmBookingDMSNonDSS
9          With frmBooking
10            With .dGrid
11                .TextMatrix(noBaris, 1) = Trim(fs_item)
12                .TextMatrix(noBaris, 2) = jobRate
13                .TextMatrix(noBaris, 3) = False
14                .TextMatrix(noBaris, 4) = objO.FetchData_String("select top 1 JOB from HO_FLATE WITH (NOLOCK) where fs_item = '" & fs_item & "' AND mdl = '" & modelKendaraan & "' and FS_LOG_DEL = ''")
15            End With
16                .Text(39) = Format(CDbl(.Text(39)) + Val(hitungLabour(CDbl(jobRate), Trim(.Text(11)))), "#0,0")
                  
                  'If Log.RPLHistori = 0 Then
17                    .nJRate = .nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & fs_item & "' AND MDL = '" & modelKendaraan & "'")
                 ' Else
                 '     .nJRate = .nJRate + objO.FetchData_Double("select Rate from [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & fs_item & "' AND MDL = '" & modelKendaraan & "'")
                  'End If
18                .hrgLabour = Val(hitungLabour(CDbl(jobRate), Trim(.Text(11)))) / IIf(CDbl(jobRate) = 0, 1, CDbl(jobRate))
19                .Text(53) = Format(CDbl(.Text(39)) + CDbl(.Text(40)) + CDbl(.Text(52)) + CDbl(.Text(55)), "#0,#")
20                .baris = noBaris + 1
21        End With
End Sub

Private Sub txtnoka_GotFocus()
    If gMainKeyCari <> "" Then
        txtnoka = gMainKeyCari
        gMainKeyCari = ""
        
    End If
End Sub

Private Sub txtnoka_KeyDown(KeyCode As Integer, Shift As Integer)
     Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
         
        gMainKeyCari = txtnoka
        gSearchOption = "CUSTOMER2"
        frmCari.show (vbModal)
             
    End If
End Sub

Private Sub txtnoka_KeyPress(KeyAscii As Integer)
 If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub txtnoka_Validate(Cancel As Boolean)
    If NulSID(txtnoka.Text) <> "" Then
'        Dim rsHistory As ADODB.Recordset
'        sSQL = "SELECT NO_CHASSIS,CUSTOMER,mdl + ' | ' + FMDL model,COLOR,NO_POLISI FROM CUSTDAT WHERE NO_CHASSIS = '" & NulSID(txtnoka.Text) & "'"
'        Set rsHistory = objO.FetchData(sSQL)
'        If Not rsHistory.EOF Then
'            Call loadData(rsHistory!NO_CHASSIS, rsHistory!customer, rsHistory!NO_POLISI, rsHistory!model, rsHistory!Color)
'            sSQL = "INSERT INTO PROFILBENGKEL..logViewHistory (norangka,userView,tglView,CompName,versiNis,userDownload) "
'            sSQL = sSQL & " VALUES ('" & NulSID(rsHistory!NO_CHASSIS) & "','" & gSA & "','" & Format(isnow, "YYYY-MM-DD") & "','" & ComptName & "','" & gVersi & "','')"
'            Call objO.InsertData(sSQL)
'        Else
'            MsgBox "Norangka tidak ditemukan dalam sistem bengkel " & gNmCbg, vbInformation
'            txtnoka.Text = ""
'            clearForm
'        End If
    End If
End Sub

Private Sub clearForm()
    Text(0) = ""
    Text(1) = ""
    Text(2) = ""
    Text(3) = ""
    Text(4) = ""
End Sub
