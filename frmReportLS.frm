VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "arprogbar.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmReportLS 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Report LS"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   10665
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6240
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin ARProgBarCtrl.ARProgressBar ARProgressBar1 
         Height          =   375
         Left            =   3720
         TabIndex        =   5
         Top             =   2880
         Visible         =   0   'False
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   661
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Border          =   0
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   2280
         TabIndex        =   4
         Top             =   2880
         Width           =   1215
      End
      Begin MSAdodcLib.Adodc AdodcSA 
         Height          =   330
         Left            =   360
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
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
         Caption         =   "AdodcSA"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbSA 
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   120
         Width           =   2775
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   93
         Text            =   "- Pilih SA -"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TrueOleDBGrid70.TDBGrid gridPeriode 
         Height          =   2175
         Left            =   2280
         TabIndex        =   3
         Top             =   600
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3836
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1).NumberFormat=   "Standard"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2).NumberFormat=   "FormatText Event"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
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
         Splits(0)._ColumnProps(12)=   "Column(2)._ColStyle=8196"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
         _StyleDefs(73)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(74)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(75)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(76)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         _StyleDefs(77)  =   "Named:id=87:SudahDitagih"
         _StyleDefs(78)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(79)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(80)  =   ":id=87,.fontname=Tahoma"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "SA : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Caption         =   "Export Report LS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin TrueOleDBGrid70.TDBGrid gridReport 
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4471
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Dealer View"
      Columns(0).DataField=   "periode"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "SA"
      Columns(1).DataField=   "namaSA"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "CPUS"
      Columns(2).DataField=   "grcpus"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "LS"
      Columns(3).DataField=   "grLS"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Non LS"
      Columns(4).DataField=   "nonls"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "% LS"
      Columns(5).DataField=   "LS"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "% Auto LS"
      Columns(6).DataField=   "lstam"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   7
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=7"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5292"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5186"
      Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=8193"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1349"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1244"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=2117"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=2011"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=516"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2117"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2011"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2117"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2011"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=516"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(27)=   "Column(5).Width=2117"
      Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2011"
      Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=516"
      Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(32)=   "Column(6).Width=2117"
      Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2011"
      Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=516"
      Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
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
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&H80FFFF&"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13,.alignment=2,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14,.alignment=2"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14,.alignment=2"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14,.alignment=2"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14,.alignment=2"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14,.alignment=2"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14,.alignment=2"
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
      _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
      _StyleDefs(79)  =   "Named:id=40:OddRow"
      _StyleDefs(80)  =   ":id=40,.parent=33"
      _StyleDefs(81)  =   "Named:id=41:RecordSelector"
      _StyleDefs(82)  =   ":id=41,.parent=34"
      _StyleDefs(83)  =   "Named:id=42:FilterBar"
      _StyleDefs(84)  =   ":id=42,.parent=33"
      _StyleDefs(85)  =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(86)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(87)  =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(88)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(89)  =   "Named:id=87:SudahDitagih"
      _StyleDefs(90)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(91)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(92)  =   ":id=87,.fontname=Tahoma"
   End
End
Attribute VB_Name = "frmReportLS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim rsPeriode As New ADODB.Recordset
Dim rsReport As New ADODB.Recordset
Dim rsSA As ADODB.Recordset

Dim kriteria As String

Private Sub cmdExport_Click()
   Dim f As New FormExplorer
   f.show vbModal
   Dim MyPath As String
   MyPath = f.PathSaving
   If MyPath = "" Then Exit Sub
   If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
   If f.FileName = "" Then Exit Sub
   'Call export(MyPath & f.FileName & ".xls")
   gridReport.ExportToFile MyPath & f.FileName & ".xls", True
   MsgBox "Data berhasil diexport ", vbInformation
   
End Sub

Private Sub cmdRefresh_Click()
1        cmdRefresh.Enabled = False
2        If rsPeriode.RecordCount > 0 Then
3           rsPeriode.MoveFirst
4           Call createReport
5           ARProgressBar1.Visible = True
6           cmdExport.Visible = True
            
7           ARProgressBar1.Max = rsPeriode.RecordCount + 1
            Dim i As Integer
            Dim totalUE  As Double
            Dim LS As Double
8           i = 1
9           gridReport.Visible = True
10          kriteria = ""
            
            ''Set rsSA = objO.FetchData("select FS_KD_PETUGAS KODE_PETUGAS, FS_NM_PETUGAS NAMA from T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA'")
            
11          If cbSA.value <> "- Pilih SA -" Then
12             kriteria = kriteria & " AND FRONT = '" & Trim(cbSA.value) & "'"
13          End If
14          Set rsSA = objO.FetchData("select FS_KD_PETUGAS KODE_PETUGAS, FS_NM_PETUGAS NAMA from T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA' " & IIf(kriteria <> "", " AND FS_KD_PETUGAS = '" & Trim(cbSA.value) & "'", ""))
           ' Label1.Visible = True
      '      Label2.Visible = True
      '      Label4.Visible = True
15          Do While Not rsPeriode.EOF
16             If rsPeriode!Pilih = True Then
17                If Not rsSA.EOF Then
18                    rsSA.MoveFirst
19                    Do While Not rsSA.EOF
20                        With rsReport
21                           .AddNew
22                           !periode = "Minggu " & rsPeriode!periode & " - " & rsPeriode!Tgl
23                           !namaSA = NulSID(rsSA!KODE_PETUGAS)
      '                       totalUE = objO.FetchData_Double("SELECT COUNT(NO_WO) AS TOTAL FROM JOBDTA WITH (NOLOCK) WHERE ([DATE] BETWEEN '" & Format(rsPeriode!tglAwal, "mm/dd/yyyy") & "' AND '" & Format(rsPeriode!tglAkhir, "mm/dd/yyyy") & "') AND TPSVC <> 'W' AND PDS = 0 AND FRONT = '" & NulSID(rsSA!KODE_PETUGAS) & "' AND FS_LOG_DEL = '' and FS_RC = '' ")
      '                       LS = objO.FetchData_Double("select COUNT(no_wo) as TOTAL from DTLFU where (respon = 'k' OR respon = 'LS') AND SA = '" & NulSID(rsSA!KODE_PETUGAS) & "' and NO_WO IN (SELECT DISTINCT (NO_WO) AS TOTAL FROM JOBDTA WITH (NOLOCK) WHERE ([DATE] BETWEEN '" & Format(rsPeriode!tglAwal, "mm/dd/yyyy") & "' AND '" & Format(rsPeriode!tglAkhir, "mm/dd/yyyy") & "') AND TPSVC <> 'W' AND PDS = 0 AND FRONT = '" & NulSID(rsSA!KODE_PETUGAS) & "' AND FS_LOG_DEL = ''  and FS_RC = ''  )")
                             
24                           totalUE = objO.FetchData_Double("SELECT COUNT(NO_WO) AS TOTAL FROM JOBDTA WITH (NOLOCK) WHERE ([DATE] BETWEEN '" & Format(rsPeriode!tglAwal, "mm/dd/yyyy") & "' AND '" & Format(rsPeriode!tglAkhir, "mm/dd/yyyy") & "') AND TPSVC <> 'W' AND PDS = 0 AND FRONT = '" & NulSID(rsSA!KODE_PETUGAS) & "' AND FS_LOG_DEL = '' and FS_RC <> 'R001' ")
25                           LS = objO.FetchData_Double("select COUNT(no_wo) as TOTAL from DTLFU where (respon = 'k' OR respon = 'LS') AND SA = '" & NulSID(rsSA!KODE_PETUGAS) & "' and NO_WO IN (SELECT DISTINCT (NO_WO) AS TOTAL FROM JOBDTA WITH (NOLOCK) WHERE ([DATE] BETWEEN '" & Format(rsPeriode!tglAwal, "mm/dd/yyyy") & "' AND '" & Format(rsPeriode!tglAkhir, "mm/dd/yyyy") & "') AND TPSVC <> 'W' AND PDS = 0 AND FRONT = '" & NulSID(rsSA!KODE_PETUGAS) & "' AND FS_LOG_DEL = ''  and FS_RC <> 'R001'  )")

26                           !grcpus = totalUE
27                           !grLS = LS
28                           !nonls = totalUE - LS
29                           If totalUE > 0 Then
30                            !LS = Round(((LS / totalUE) * 100), 0) & " %"
31                           Else
32                            !LS = "0" & " %"
33                           End If
34                           If totalUE > 0 Then
35                            If Round(((LS / totalUE) * 100), 0) < rsPeriode!MINLS Then
                                   ''!LsTam = NulSID(rsPeriode!MAXLS) & " %"
                                   '' 20181101 | MB THANIA | DIGANTI NILAI MINIMAL YANG DIKIRIM KE TAM
36                                 !LsTam = NulSID(rsPeriode!MINLS) & " %"
37                            Else
38                                 !LsTam = Round(((LS / totalUE) * 100), 0) & " %"
39                            End If
40                           Else
                                  ''!LsTam = NulSID(rsPeriode!MAXLS) & " %"
41                                !LsTam = NulSID(rsPeriode!MINLS) & " %"
42                           End If
                             
43                           .Update
44                        End With
45                        rsSA.MoveNext
46                    Loop
47                    rsSA.MoveFirst
48                End If
49             End If
50             rsPeriode.MoveNext
51             i = i + 1
52             ARProgressBar1.value = i
53             DoEvents
54          Loop
55          gridReport.Refresh
56       End If
57       ARProgressBar1.Visible = False
58       cmdRefresh.Enabled = True
End Sub

Private Sub createReport()

1        On Error GoTo createReport_Error

2        If rsReport.State Then rsReport.Close
3        Set rsReport = Nothing
4        With rsReport
5            .Fields.Append "Periode", adVarChar, 50
6            .Fields.Append "namasa", adVarChar, 20
7            .Fields.Append "grcpus", adVarChar, 20
8            .Fields.Append "grls", adVarChar, 20
9            .Fields.Append "nonLs", adVarChar, 20
10           .Fields.Append "LS", adVarChar, 20
11           .Fields.Append "LsTam", adVarChar, 20
12           .Open
13       End With
          
14       gridReport.DataSource = rsReport

15       On Error GoTo 0
16       Exit Sub

createReport_Error:

17       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createReport of Form frmReportLS")
         
End Sub

Private Sub Command1_Click()
    Unload Me
    
End Sub

Private Sub Form_Load()
   Call refreshGrid
   
End Sub

Public Sub createHeaderPeriode()
   
   On Error GoTo createHeaderPeriode_Error
      With gridPeriode
         
         .Columns(1).Caption = "Periode"
         .Columns(1).DataField = "Periode"
         .Columns(1).Width = 800
         .Columns(1).Visible = False
         
         .Columns(2).Caption = "Tanggal Awal"
         .Columns(2).DataField = "tglAwal"
         .Columns(2).Width = 1500
         
         .Columns(3).Caption = "Tanggal Akhir"
         .Columns(3).DataField = "tglAkhir"
         .Columns(3).Width = 1500
         
         .Columns(0).Caption = "Pilih"
         .Columns(0).DataField = "Pilih"
         .Columns(0).Width = 500
         
      End With
   On Error GoTo 0
   Exit Sub

createHeaderPeriode_Error:

   MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createHeaderPeriode of Form frmReportLS")
   
End Sub

Public Sub createPeriode()
   
   On Error GoTo createPeriode_Error
   If rsPeriode.State Then rsPeriode.Close
   Set rsPeriode = Nothing
   With rsPeriode
       .Fields.Append "Pilih", adBoolean
       .Fields.Append "periode", adVarChar, 20
       .Fields.Append "tglAwal", adVarChar, 50
       .Fields.Append "tglAkhir", adVarChar, 50
       .Fields.Append "tgl", adVarChar, 100
       .Fields.Append "minLS", adInteger
       .Fields.Append "maxLS", adInteger
       .Open
   End With
    
   gridPeriode.DataSource = rsPeriode ''  rsPeriode.AddNew
   ''Call createHeaderPeriode

   On Error GoTo 0
   Exit Sub

createPeriode_Error:

   MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createPeriode of Form frmReportLS")
   
End Sub

Private Sub refreshGrid()
   With AdodcSA
        .ConnectionString = SQLCONBengkel
        .CommandType = adCmdText
        .RecordSource = "select FS_KD_PETUGAS KODE_PETUGAS, FS_NM_PETUGAS NAMA from T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA'"
        .Refresh
   End With
   
   cbSA.DataSourceList = AdodcSA
   cbSA.DataFieldList = "KODE_PETUGAS"
   cbSA.Columns(0).Width = 1200
   cbSA.Columns(1).Width = 4000
   cbSA.Columns(0).Caption = "Kode Petugas"
   cbSA.Columns(1).Caption = "Nama Petugas"
   
   Call createPeriode
   Call createHeaderPeriode
   If Log.RPLHistori = 0 Then
    Set rsTemp = objO.FetchData("SELECT TOP 10 b.minls,b.maxls, a.minggu as periode,a.tglAwal,a.tglAkhir, convert(varchar(100),a.tglAwal,105) + ' s/d ' + convert(varchar(100),a.tglAkhir,105) tgl FROM periode a WITH (NOLOCK) LEFT JOIN master_setting b on a.minggu = b.[week] ORDER BY tglAwal DESC")
   Else
    Set rsTemp = objO.FetchData("SELECT TOP 10 b.minls,b.maxls, a.minggu as periode,a.tglAwal,a.tglAkhir, convert(varchar(100),a.tglAwal,105) + ' s/d ' + convert(varchar(100),a.tglAkhir,105) tgl FROM periode a WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..master_setting b on a.minggu = b.[week] ORDER BY tglAwal DESC")
   End If
   If Not rsTemp.EOF Then
      rsTemp.MoveFirst
      Do While Not rsTemp.EOF
         With rsPeriode
            .AddNew
            !Pilih = False
            !periode = Trim(rsTemp!periode)
            !tglAwal = Format(Trim(rsTemp!tglAwal), "dd/mm/yyyy")
            !tglAkhir = Format(Trim(rsTemp!tglAkhir), "dd/mm/yyyy")
            !Tgl = Trim(rsTemp!Tgl)
            !MINLS = NulSID(rsTemp!MINLS)
            !MAXLS = NulSID(rsTemp!MAXLS)
            
            .Update
         End With
         rsTemp.MoveNext
      Loop
      rsPeriode.MoveFirst
   End If
    
  ' gridPeriode.DataSource = rsPeriode
  gridPeriode.Refresh
End Sub
