VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form FormMRABoard 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16050
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   16050
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   8835
      Left            =   645
      TabIndex        =   10
      Top             =   540
      Visible         =   0   'False
      Width           =   15405
      Begin VB.CommandButton cmdHistory 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&History"
         Height          =   735
         Left            =   12480
         Picture         =   "FormMRABoard.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5805
         Width           =   855
      End
      Begin VB.Frame Frame3 
         Caption         =   "Legend Warna"
         Height          =   750
         Left            =   285
         TabIndex        =   21
         Top             =   5610
         Width           =   4965
         Begin VB.Frame Frame4 
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            Height          =   360
            Left            =   135
            TabIndex        =   22
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label1 
            Caption         =   "Unit yang sudah difollow up MRA"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1080
            TabIndex        =   23
            Top             =   255
            Width           =   3765
         End
      End
      Begin VB.CommandButton cmdReminderClose 
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
         Left            =   14400
         Picture         =   "FormMRABoard.frx":066D
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   5805
         Width           =   735
      End
      Begin VB.CommandButton cmdCall 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Call"
         Height          =   735
         Left            =   13365
         Picture         =   "FormMRABoard.frx":0DFD
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5805
         Width           =   975
      End
      Begin TrueOleDBGrid70.TDBGrid gridReminder 
         Height          =   4875
         Left            =   270
         TabIndex        =   11
         Top             =   660
         Width           =   15030
         _ExtentX        =   26511
         _ExtentY        =   8599
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
         Columns(1).Caption=   "Type Kendaraan"
         Columns(1).DataField=   "type"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Nama"
         Columns(2).DataField=   "Nama"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Alamat"
         Columns(3).DataField=   "Alamat"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Kota"
         Columns(4).DataField=   "Kota"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Telp"
         Columns(5).DataField=   "Telp"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "GSM"
         Columns(6).DataField=   "GSM"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "CDMA"
         Columns(7).DataField=   "CDMA"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Tgl Reminder Sales"
         Columns(8).DataField=   "tglReminderSales"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Wiraniaga"
         Columns(9).DataField=   "namasales"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Tanggal BOC"
         Columns(10).DataField=   "tglBOC"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "warna"
         Columns(11).DataField=   "warna"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   12
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=12"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=4498"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4419"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=4498"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4419"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4498"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4419"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=6244"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=6165"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=2725"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
         Splits(0)._ColumnProps(48)=   "Column(11).Visible=0"
         Splits(0)._ColumnProps(49)=   "Column(11).Order=12"
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
         Caption         =   "PROSES REMINDER MONITORING BOARD"
         MultipleLines   =   0
         CellTipsWidth   =   0
         CellTipsDelay   =   10
         DeadAreaBackColor=   12577002
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
         _StyleDefs(24)  =   "Splits(0).Style:id=25,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=27,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=56,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=55,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=57,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=58,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=25"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=26"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=27"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=55"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=25"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=26"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=27"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=55"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=25"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=26"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=27"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=55"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=25"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=26"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=27"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=55"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=25"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=26"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=27"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=55"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=25"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=26"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=27"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=55"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=86,.parent=25"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=83,.parent=26"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=84,.parent=27"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=85,.parent=55"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=90,.parent=25"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=87,.parent=26"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=88,.parent=27"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=89,.parent=55"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=32,.parent=25"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=26"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=27"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=55"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=24,.parent=25"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=21,.parent=26"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=22,.parent=27"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=23,.parent=55"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=20,.parent=25"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=17,.parent=26"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=18,.parent=27"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=19,.parent=55"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=16,.parent=25"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=13,.parent=26"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=14,.parent=27"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=15,.parent=55"
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
         _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(99)  =   "Named:id=40:OddRow"
         _StyleDefs(100) =   ":id=40,.parent=33"
         _StyleDefs(101) =   "Named:id=41:RecordSelector"
         _StyleDefs(102) =   ":id=41,.parent=34"
         _StyleDefs(103) =   "Named:id=42:FilterBar"
         _StyleDefs(104) =   ":id=42,.parent=33"
         _StyleDefs(105) =   "Named:id=303:BiruPutih"
         _StyleDefs(106) =   ":id=303,.parent=39,.bgcolor=&HFF8080&,.fgcolor=&HFFFFFF&,.bold=-1,.fontsize=825"
         _StyleDefs(107) =   ":id=303,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(108) =   ":id=303,.fontname=MS Sans Serif"
         _StyleDefs(109) =   "Named:id=304:KuningHitam"
         _StyleDefs(110) =   ":id=304,.parent=303,.bgcolor=&HFFFF&,.fgcolor=&H0&"
         _StyleDefs(111) =   "Named:id=305:HijauPutih"
         _StyleDefs(112) =   ":id=305,.parent=303,.bgcolor=&H80FF80&,.fgcolor=&H0&"
         _StyleDefs(113) =   "Named:id=306:MerahPutih"
         _StyleDefs(114) =   ":id=306,.parent=305,.bgcolor=&HFF&"
         _StyleDefs(115) =   "Named:id=307:BiruHitam"
         _StyleDefs(116) =   ":id=307,.parent=303,.bgcolor=&HC08000&,.fgcolor=&H0&"
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   12330
         Shape           =   4  'Rounded Rectangle
         Top             =   5715
         Width           =   2880
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   4080
      Left            =   990
      TabIndex        =   18
      Top             =   2115
      Visible         =   0   'False
      Width           =   8610
      Begin VB.CommandButton btnTutup 
         BackColor       =   &H000000FF&
         Caption         =   "Tutup"
         Height          =   360
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   30
         Visible         =   0   'False
         Width           =   1050
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   3795
         Left            =   90
         OleObjectBlob   =   "FormMRABoard.frx":111D
         TabIndex        =   19
         Top             =   210
         Visible         =   0   'False
         Width           =   8415
      End
   End
   Begin VB.CommandButton CmdOpen 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Graph"
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
      Left            =   14100
      Picture         =   "FormMRABoard.frx":38FC
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6465
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
      Left            =   14925
      Picture         =   "FormMRABoard.frx":3EF5
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6465
      Width           =   735
   End
   Begin VB.CommandButton cmdProcess 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pr&ocess"
      Height          =   732
      Left            =   13140
      Picture         =   "FormMRABoard.frx":4685
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6480
      Width           =   732
   End
   Begin ucXPButton.XPButton thnMinus 
      Height          =   495
      Left            =   735
      TabIndex        =   1
      Top             =   585
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   ""
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormMRABoard.frx":4CF2
      PICN            =   "FormMRABoard.frx":4D0E
      PICH            =   "FormMRABoard.frx":4F19
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton blnMinus 
      Height          =   495
      Left            =   13065
      TabIndex        =   2
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   ""
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormMRABoard.frx":5124
      PICN            =   "FormMRABoard.frx":5140
      PICH            =   "FormMRABoard.frx":534B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton thnPlus 
      Height          =   495
      Left            =   3810
      TabIndex        =   3
      Top             =   585
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   ""
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormMRABoard.frx":5556
      PICN            =   "FormMRABoard.frx":5572
      PICH            =   "FormMRABoard.frx":5769
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton blnPlus 
      Height          =   495
      Left            =   15420
      TabIndex        =   4
      Top             =   660
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BTYPE           =   3
      TX              =   ""
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
      BCOL            =   16315377
      BCOLO           =   16315377
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FormMRABoard.frx":5960
      PICN            =   "FormMRABoard.frx":597C
      PICH            =   "FormMRABoard.frx":5B73
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TrueOleDBGrid70.TDBGrid gridMonitoring 
      Height          =   5085
      Left            =   720
      TabIndex        =   9
      Top             =   1215
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   8969
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
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).DataField=   ""
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).DataField=   ""
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).DataField=   ""
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).DataField=   ""
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).DataField=   ""
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   33
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   2
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=33"
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
      Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=2725"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=2646"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=2725"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=2646"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=2725"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=2646"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=2725"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=2646"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(69)=   "Column(17).Width=2725"
      Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=2646"
      Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(73)=   "Column(18).Width=2725"
      Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=2646"
      Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(77)=   "Column(19).Width=2725"
      Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=2646"
      Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(81)=   "Column(20).Width=2725"
      Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=2646"
      Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(85)=   "Column(21).Width=2725"
      Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=2646"
      Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(89)=   "Column(22).Width=2725"
      Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=2646"
      Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(93)=   "Column(23).Width=2725"
      Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=2646"
      Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(97)=   "Column(24).Width=2725"
      Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=2646"
      Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(101)=   "Column(25).Width=2725"
      Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=2646"
      Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(105)=   "Column(26).Width=2725"
      Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=2646"
      Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(109)=   "Column(27).Width=2725"
      Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=2646"
      Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(113)=   "Column(28).Width=2725"
      Splits(0)._ColumnProps(114)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(28)._WidthInPix=2646"
      Splits(0)._ColumnProps(116)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(117)=   "Column(29).Width=2725"
      Splits(0)._ColumnProps(118)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(29)._WidthInPix=2646"
      Splits(0)._ColumnProps(120)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(121)=   "Column(30).Width=2725"
      Splits(0)._ColumnProps(122)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(30)._WidthInPix=2646"
      Splits(0)._ColumnProps(124)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(125)=   "Column(31).Width=2725"
      Splits(0)._ColumnProps(126)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(31)._WidthInPix=2646"
      Splits(0)._ColumnProps(128)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(129)=   "Column(32).Width=2725"
      Splits(0)._ColumnProps(130)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(32)._WidthInPix=2646"
      Splits(0)._ColumnProps(132)=   "Column(32).Order=33"
      Splits(1)._UserFlags=   0
      Splits(1).RecordSelectorWidth=   688
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).FetchRowStyle=   -1  'True
      Splits(1).DividerColor=   14933984
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=33"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(1)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(1)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(1)._ColumnProps(8)=   "Column(1).Button=1"
      Splits(1)._ColumnProps(9)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(10)=   "Column(1).ButtonText=1"
      Splits(1)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(1)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(1)._ColumnProps(14)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(15)=   "Column(3).Width=2725"
      Splits(1)._ColumnProps(16)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
      Splits(1)._ColumnProps(18)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(19)=   "Column(4).Width=2725"
      Splits(1)._ColumnProps(20)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
      Splits(1)._ColumnProps(22)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(23)=   "Column(5).Width=2725"
      Splits(1)._ColumnProps(24)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(25)=   "Column(5)._WidthInPix=2646"
      Splits(1)._ColumnProps(26)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(27)=   "Column(6).Width=2725"
      Splits(1)._ColumnProps(28)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(29)=   "Column(6)._WidthInPix=2646"
      Splits(1)._ColumnProps(30)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(31)=   "Column(7).Width=2725"
      Splits(1)._ColumnProps(32)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(33)=   "Column(7)._WidthInPix=2646"
      Splits(1)._ColumnProps(34)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(35)=   "Column(8).Width=2725"
      Splits(1)._ColumnProps(36)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(37)=   "Column(8)._WidthInPix=2646"
      Splits(1)._ColumnProps(38)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(39)=   "Column(9).Width=2725"
      Splits(1)._ColumnProps(40)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(41)=   "Column(9)._WidthInPix=2646"
      Splits(1)._ColumnProps(42)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(43)=   "Column(10).Width=2725"
      Splits(1)._ColumnProps(44)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(45)=   "Column(10)._WidthInPix=2646"
      Splits(1)._ColumnProps(46)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(47)=   "Column(11).Width=2725"
      Splits(1)._ColumnProps(48)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(49)=   "Column(11)._WidthInPix=2646"
      Splits(1)._ColumnProps(50)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(51)=   "Column(12).Width=2725"
      Splits(1)._ColumnProps(52)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(53)=   "Column(12)._WidthInPix=2646"
      Splits(1)._ColumnProps(54)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(55)=   "Column(13).Width=2725"
      Splits(1)._ColumnProps(56)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(57)=   "Column(13)._WidthInPix=2646"
      Splits(1)._ColumnProps(58)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(59)=   "Column(14).Width=2725"
      Splits(1)._ColumnProps(60)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(61)=   "Column(14)._WidthInPix=2646"
      Splits(1)._ColumnProps(62)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(63)=   "Column(15).Width=2725"
      Splits(1)._ColumnProps(64)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(65)=   "Column(15)._WidthInPix=2646"
      Splits(1)._ColumnProps(66)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(67)=   "Column(16).Width=2725"
      Splits(1)._ColumnProps(68)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(69)=   "Column(16)._WidthInPix=2646"
      Splits(1)._ColumnProps(70)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(71)=   "Column(17).Width=2725"
      Splits(1)._ColumnProps(72)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(73)=   "Column(17)._WidthInPix=2646"
      Splits(1)._ColumnProps(74)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(75)=   "Column(18).Width=2725"
      Splits(1)._ColumnProps(76)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(77)=   "Column(18)._WidthInPix=2646"
      Splits(1)._ColumnProps(78)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(79)=   "Column(19).Width=2725"
      Splits(1)._ColumnProps(80)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(81)=   "Column(19)._WidthInPix=2646"
      Splits(1)._ColumnProps(82)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(83)=   "Column(20).Width=2725"
      Splits(1)._ColumnProps(84)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(85)=   "Column(20)._WidthInPix=2646"
      Splits(1)._ColumnProps(86)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(87)=   "Column(21).Width=2725"
      Splits(1)._ColumnProps(88)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(89)=   "Column(21)._WidthInPix=2646"
      Splits(1)._ColumnProps(90)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(91)=   "Column(22).Width=2725"
      Splits(1)._ColumnProps(92)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(93)=   "Column(22)._WidthInPix=2646"
      Splits(1)._ColumnProps(94)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(95)=   "Column(23).Width=2725"
      Splits(1)._ColumnProps(96)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(97)=   "Column(23)._WidthInPix=2646"
      Splits(1)._ColumnProps(98)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(99)=   "Column(24).Width=2725"
      Splits(1)._ColumnProps(100)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(101)=   "Column(24)._WidthInPix=2646"
      Splits(1)._ColumnProps(102)=   "Column(24).Order=25"
      Splits(1)._ColumnProps(103)=   "Column(25).Width=2725"
      Splits(1)._ColumnProps(104)=   "Column(25).DividerColor=0"
      Splits(1)._ColumnProps(105)=   "Column(25)._WidthInPix=2646"
      Splits(1)._ColumnProps(106)=   "Column(25).Order=26"
      Splits(1)._ColumnProps(107)=   "Column(26).Width=2725"
      Splits(1)._ColumnProps(108)=   "Column(26).DividerColor=0"
      Splits(1)._ColumnProps(109)=   "Column(26)._WidthInPix=2646"
      Splits(1)._ColumnProps(110)=   "Column(26).Order=27"
      Splits(1)._ColumnProps(111)=   "Column(27).Width=2725"
      Splits(1)._ColumnProps(112)=   "Column(27).DividerColor=0"
      Splits(1)._ColumnProps(113)=   "Column(27)._WidthInPix=2646"
      Splits(1)._ColumnProps(114)=   "Column(27).Order=28"
      Splits(1)._ColumnProps(115)=   "Column(28).Width=2725"
      Splits(1)._ColumnProps(116)=   "Column(28).DividerColor=0"
      Splits(1)._ColumnProps(117)=   "Column(28)._WidthInPix=2646"
      Splits(1)._ColumnProps(118)=   "Column(28).Order=29"
      Splits(1)._ColumnProps(119)=   "Column(29).Width=2725"
      Splits(1)._ColumnProps(120)=   "Column(29).DividerColor=0"
      Splits(1)._ColumnProps(121)=   "Column(29)._WidthInPix=2646"
      Splits(1)._ColumnProps(122)=   "Column(29).Order=30"
      Splits(1)._ColumnProps(123)=   "Column(30).Width=2725"
      Splits(1)._ColumnProps(124)=   "Column(30).DividerColor=0"
      Splits(1)._ColumnProps(125)=   "Column(30)._WidthInPix=2646"
      Splits(1)._ColumnProps(126)=   "Column(30).Order=31"
      Splits(1)._ColumnProps(127)=   "Column(31).Width=2725"
      Splits(1)._ColumnProps(128)=   "Column(31).DividerColor=0"
      Splits(1)._ColumnProps(129)=   "Column(31)._WidthInPix=2646"
      Splits(1)._ColumnProps(130)=   "Column(31).Order=32"
      Splits(1)._ColumnProps(131)=   "Column(32).Width=2725"
      Splits(1)._ColumnProps(132)=   "Column(32).DividerColor=0"
      Splits(1)._ColumnProps(133)=   "Column(32)._WidthInPix=2646"
      Splits(1)._ColumnProps(134)=   "Column(32).Order=33"
      Splits.Count    =   2
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
      Caption         =   "PROGRESS REMINDER MONITORING BOARD"
      MultipleLines   =   0
      CellTipsWidth   =   0
      CellTipsDelay   =   10
      DeadAreaBackColor=   12577002
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
      _StyleDefs(24)  =   "Splits(0).Style:id=29,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=48,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=30,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=31,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=32,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=44,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=43,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=45,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=46,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=47,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=49,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=50,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=130,.parent=29"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=127,.parent=30"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=128,.parent=31"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=129,.parent=43"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=122,.parent=29"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=119,.parent=30"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=120,.parent=31"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=121,.parent=43"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=114,.parent=29"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=111,.parent=30"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=112,.parent=31"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=113,.parent=43"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=106,.parent=29"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=103,.parent=30"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=104,.parent=31"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=105,.parent=43"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=98,.parent=29"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=95,.parent=30"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=96,.parent=31"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=97,.parent=43"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=90,.parent=29"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=87,.parent=30"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=88,.parent=31"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=89,.parent=43"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=82,.parent=29"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=79,.parent=30"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=80,.parent=31"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=81,.parent=43"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=74,.parent=29"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=71,.parent=30"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=72,.parent=31"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=73,.parent=43"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=29"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=30"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=31"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=43"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=28,.parent=29"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=30"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=31"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=43"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=58,.parent=29"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=55,.parent=30"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=56,.parent=31"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=57,.parent=43"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=138,.parent=29"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=135,.parent=30"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=136,.parent=31"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=137,.parent=43"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=146,.parent=29"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=143,.parent=30"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=144,.parent=31"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=145,.parent=43"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=154,.parent=29"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=151,.parent=30"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=152,.parent=31"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=153,.parent=43"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=162,.parent=29"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=159,.parent=30"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=160,.parent=31"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=161,.parent=43"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=170,.parent=29"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=167,.parent=30"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=168,.parent=31"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=169,.parent=43"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=178,.parent=29"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=175,.parent=30"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=176,.parent=31"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=177,.parent=43"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=186,.parent=29"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=183,.parent=30"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=184,.parent=31"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=185,.parent=43"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=194,.parent=29"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=191,.parent=30"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=192,.parent=31"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=193,.parent=43"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=202,.parent=29"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=199,.parent=30"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=200,.parent=31"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=201,.parent=43"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=210,.parent=29"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=207,.parent=30"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=208,.parent=31"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=209,.parent=43"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=218,.parent=29"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=215,.parent=30"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=216,.parent=31"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=217,.parent=43"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=226,.parent=29"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=223,.parent=30"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=224,.parent=31"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=225,.parent=43"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=234,.parent=29"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=231,.parent=30"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=232,.parent=31"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=233,.parent=43"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=242,.parent=29"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=239,.parent=30"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=240,.parent=31"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=241,.parent=43"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=250,.parent=29"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=247,.parent=30"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=248,.parent=31"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=249,.parent=43"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=258,.parent=29"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=255,.parent=30"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=256,.parent=31"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=257,.parent=43"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=266,.parent=29"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=263,.parent=30"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=264,.parent=31"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=265,.parent=43"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=274,.parent=29"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=271,.parent=30"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=272,.parent=31"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=273,.parent=43"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=282,.parent=29"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=279,.parent=30"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=280,.parent=31"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=281,.parent=43"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=290,.parent=29"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=287,.parent=30"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=288,.parent=31"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=289,.parent=43"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=298,.parent=29"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=295,.parent=30"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=296,.parent=31"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=297,.parent=43"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=311,.parent=29"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=308,.parent=30"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=309,.parent=31"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=310,.parent=43"
      _StyleDefs(168) =   "Splits(1).Style:id=13,.parent=1"
      _StyleDefs(169) =   "Splits(1).CaptionStyle:id=22,.parent=4"
      _StyleDefs(170) =   "Splits(1).HeadingStyle:id=14,.parent=2"
      _StyleDefs(171) =   "Splits(1).FooterStyle:id=15,.parent=3"
      _StyleDefs(172) =   "Splits(1).InactiveStyle:id=16,.parent=5"
      _StyleDefs(173) =   "Splits(1).SelectedStyle:id=18,.parent=6"
      _StyleDefs(174) =   "Splits(1).EditorStyle:id=17,.parent=7"
      _StyleDefs(175) =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(176) =   "Splits(1).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(177) =   "Splits(1).OddRowStyle:id=21,.parent=10"
      _StyleDefs(178) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(179) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(180) =   "Splits(1).Columns(0).Style:id=134,.parent=13"
      _StyleDefs(181) =   "Splits(1).Columns(0).HeadingStyle:id=131,.parent=14"
      _StyleDefs(182) =   "Splits(1).Columns(0).FooterStyle:id=132,.parent=15"
      _StyleDefs(183) =   "Splits(1).Columns(0).EditorStyle:id=133,.parent=17"
      _StyleDefs(184) =   "Splits(1).Columns(1).Style:id=126,.parent=13,.bgcolor=&H80FF&"
      _StyleDefs(185) =   "Splits(1).Columns(1).HeadingStyle:id=123,.parent=14"
      _StyleDefs(186) =   "Splits(1).Columns(1).FooterStyle:id=124,.parent=15"
      _StyleDefs(187) =   "Splits(1).Columns(1).EditorStyle:id=125,.parent=17,.bgcolor=&HFF00&"
      _StyleDefs(188) =   "Splits(1).Columns(2).Style:id=118,.parent=13"
      _StyleDefs(189) =   "Splits(1).Columns(2).HeadingStyle:id=115,.parent=14"
      _StyleDefs(190) =   "Splits(1).Columns(2).FooterStyle:id=116,.parent=15"
      _StyleDefs(191) =   "Splits(1).Columns(2).EditorStyle:id=117,.parent=17"
      _StyleDefs(192) =   "Splits(1).Columns(3).Style:id=110,.parent=13"
      _StyleDefs(193) =   "Splits(1).Columns(3).HeadingStyle:id=107,.parent=14"
      _StyleDefs(194) =   "Splits(1).Columns(3).FooterStyle:id=108,.parent=15"
      _StyleDefs(195) =   "Splits(1).Columns(3).EditorStyle:id=109,.parent=17"
      _StyleDefs(196) =   "Splits(1).Columns(4).Style:id=102,.parent=13"
      _StyleDefs(197) =   "Splits(1).Columns(4).HeadingStyle:id=99,.parent=14"
      _StyleDefs(198) =   "Splits(1).Columns(4).FooterStyle:id=100,.parent=15"
      _StyleDefs(199) =   "Splits(1).Columns(4).EditorStyle:id=101,.parent=17"
      _StyleDefs(200) =   "Splits(1).Columns(5).Style:id=94,.parent=13"
      _StyleDefs(201) =   "Splits(1).Columns(5).HeadingStyle:id=91,.parent=14"
      _StyleDefs(202) =   "Splits(1).Columns(5).FooterStyle:id=92,.parent=15"
      _StyleDefs(203) =   "Splits(1).Columns(5).EditorStyle:id=93,.parent=17"
      _StyleDefs(204) =   "Splits(1).Columns(6).Style:id=86,.parent=13"
      _StyleDefs(205) =   "Splits(1).Columns(6).HeadingStyle:id=83,.parent=14"
      _StyleDefs(206) =   "Splits(1).Columns(6).FooterStyle:id=84,.parent=15"
      _StyleDefs(207) =   "Splits(1).Columns(6).EditorStyle:id=85,.parent=17"
      _StyleDefs(208) =   "Splits(1).Columns(7).Style:id=78,.parent=13"
      _StyleDefs(209) =   "Splits(1).Columns(7).HeadingStyle:id=75,.parent=14"
      _StyleDefs(210) =   "Splits(1).Columns(7).FooterStyle:id=76,.parent=15"
      _StyleDefs(211) =   "Splits(1).Columns(7).EditorStyle:id=77,.parent=17"
      _StyleDefs(212) =   "Splits(1).Columns(8).Style:id=70,.parent=13"
      _StyleDefs(213) =   "Splits(1).Columns(8).HeadingStyle:id=67,.parent=14"
      _StyleDefs(214) =   "Splits(1).Columns(8).FooterStyle:id=68,.parent=15"
      _StyleDefs(215) =   "Splits(1).Columns(8).EditorStyle:id=69,.parent=17"
      _StyleDefs(216) =   "Splits(1).Columns(9).Style:id=54,.parent=13"
      _StyleDefs(217) =   "Splits(1).Columns(9).HeadingStyle:id=51,.parent=14"
      _StyleDefs(218) =   "Splits(1).Columns(9).FooterStyle:id=52,.parent=15"
      _StyleDefs(219) =   "Splits(1).Columns(9).EditorStyle:id=53,.parent=17"
      _StyleDefs(220) =   "Splits(1).Columns(10).Style:id=62,.parent=13"
      _StyleDefs(221) =   "Splits(1).Columns(10).HeadingStyle:id=59,.parent=14"
      _StyleDefs(222) =   "Splits(1).Columns(10).FooterStyle:id=60,.parent=15"
      _StyleDefs(223) =   "Splits(1).Columns(10).EditorStyle:id=61,.parent=17"
      _StyleDefs(224) =   "Splits(1).Columns(11).Style:id=142,.parent=13"
      _StyleDefs(225) =   "Splits(1).Columns(11).HeadingStyle:id=139,.parent=14"
      _StyleDefs(226) =   "Splits(1).Columns(11).FooterStyle:id=140,.parent=15"
      _StyleDefs(227) =   "Splits(1).Columns(11).EditorStyle:id=141,.parent=17"
      _StyleDefs(228) =   "Splits(1).Columns(12).Style:id=150,.parent=13"
      _StyleDefs(229) =   "Splits(1).Columns(12).HeadingStyle:id=147,.parent=14"
      _StyleDefs(230) =   "Splits(1).Columns(12).FooterStyle:id=148,.parent=15"
      _StyleDefs(231) =   "Splits(1).Columns(12).EditorStyle:id=149,.parent=17"
      _StyleDefs(232) =   "Splits(1).Columns(13).Style:id=158,.parent=13"
      _StyleDefs(233) =   "Splits(1).Columns(13).HeadingStyle:id=155,.parent=14"
      _StyleDefs(234) =   "Splits(1).Columns(13).FooterStyle:id=156,.parent=15"
      _StyleDefs(235) =   "Splits(1).Columns(13).EditorStyle:id=157,.parent=17"
      _StyleDefs(236) =   "Splits(1).Columns(14).Style:id=166,.parent=13"
      _StyleDefs(237) =   "Splits(1).Columns(14).HeadingStyle:id=163,.parent=14"
      _StyleDefs(238) =   "Splits(1).Columns(14).FooterStyle:id=164,.parent=15"
      _StyleDefs(239) =   "Splits(1).Columns(14).EditorStyle:id=165,.parent=17"
      _StyleDefs(240) =   "Splits(1).Columns(15).Style:id=174,.parent=13"
      _StyleDefs(241) =   "Splits(1).Columns(15).HeadingStyle:id=171,.parent=14"
      _StyleDefs(242) =   "Splits(1).Columns(15).FooterStyle:id=172,.parent=15"
      _StyleDefs(243) =   "Splits(1).Columns(15).EditorStyle:id=173,.parent=17"
      _StyleDefs(244) =   "Splits(1).Columns(16).Style:id=182,.parent=13"
      _StyleDefs(245) =   "Splits(1).Columns(16).HeadingStyle:id=179,.parent=14"
      _StyleDefs(246) =   "Splits(1).Columns(16).FooterStyle:id=180,.parent=15"
      _StyleDefs(247) =   "Splits(1).Columns(16).EditorStyle:id=181,.parent=17"
      _StyleDefs(248) =   "Splits(1).Columns(17).Style:id=190,.parent=13"
      _StyleDefs(249) =   "Splits(1).Columns(17).HeadingStyle:id=187,.parent=14"
      _StyleDefs(250) =   "Splits(1).Columns(17).FooterStyle:id=188,.parent=15"
      _StyleDefs(251) =   "Splits(1).Columns(17).EditorStyle:id=189,.parent=17"
      _StyleDefs(252) =   "Splits(1).Columns(18).Style:id=198,.parent=13"
      _StyleDefs(253) =   "Splits(1).Columns(18).HeadingStyle:id=195,.parent=14"
      _StyleDefs(254) =   "Splits(1).Columns(18).FooterStyle:id=196,.parent=15"
      _StyleDefs(255) =   "Splits(1).Columns(18).EditorStyle:id=197,.parent=17"
      _StyleDefs(256) =   "Splits(1).Columns(19).Style:id=206,.parent=13"
      _StyleDefs(257) =   "Splits(1).Columns(19).HeadingStyle:id=203,.parent=14"
      _StyleDefs(258) =   "Splits(1).Columns(19).FooterStyle:id=204,.parent=15"
      _StyleDefs(259) =   "Splits(1).Columns(19).EditorStyle:id=205,.parent=17"
      _StyleDefs(260) =   "Splits(1).Columns(20).Style:id=214,.parent=13"
      _StyleDefs(261) =   "Splits(1).Columns(20).HeadingStyle:id=211,.parent=14"
      _StyleDefs(262) =   "Splits(1).Columns(20).FooterStyle:id=212,.parent=15"
      _StyleDefs(263) =   "Splits(1).Columns(20).EditorStyle:id=213,.parent=17"
      _StyleDefs(264) =   "Splits(1).Columns(21).Style:id=222,.parent=13"
      _StyleDefs(265) =   "Splits(1).Columns(21).HeadingStyle:id=219,.parent=14"
      _StyleDefs(266) =   "Splits(1).Columns(21).FooterStyle:id=220,.parent=15"
      _StyleDefs(267) =   "Splits(1).Columns(21).EditorStyle:id=221,.parent=17"
      _StyleDefs(268) =   "Splits(1).Columns(22).Style:id=230,.parent=13"
      _StyleDefs(269) =   "Splits(1).Columns(22).HeadingStyle:id=227,.parent=14"
      _StyleDefs(270) =   "Splits(1).Columns(22).FooterStyle:id=228,.parent=15"
      _StyleDefs(271) =   "Splits(1).Columns(22).EditorStyle:id=229,.parent=17"
      _StyleDefs(272) =   "Splits(1).Columns(23).Style:id=238,.parent=13"
      _StyleDefs(273) =   "Splits(1).Columns(23).HeadingStyle:id=235,.parent=14"
      _StyleDefs(274) =   "Splits(1).Columns(23).FooterStyle:id=236,.parent=15"
      _StyleDefs(275) =   "Splits(1).Columns(23).EditorStyle:id=237,.parent=17"
      _StyleDefs(276) =   "Splits(1).Columns(24).Style:id=246,.parent=13"
      _StyleDefs(277) =   "Splits(1).Columns(24).HeadingStyle:id=243,.parent=14"
      _StyleDefs(278) =   "Splits(1).Columns(24).FooterStyle:id=244,.parent=15"
      _StyleDefs(279) =   "Splits(1).Columns(24).EditorStyle:id=245,.parent=17"
      _StyleDefs(280) =   "Splits(1).Columns(25).Style:id=254,.parent=13"
      _StyleDefs(281) =   "Splits(1).Columns(25).HeadingStyle:id=251,.parent=14"
      _StyleDefs(282) =   "Splits(1).Columns(25).FooterStyle:id=252,.parent=15"
      _StyleDefs(283) =   "Splits(1).Columns(25).EditorStyle:id=253,.parent=17"
      _StyleDefs(284) =   "Splits(1).Columns(26).Style:id=262,.parent=13"
      _StyleDefs(285) =   "Splits(1).Columns(26).HeadingStyle:id=259,.parent=14"
      _StyleDefs(286) =   "Splits(1).Columns(26).FooterStyle:id=260,.parent=15"
      _StyleDefs(287) =   "Splits(1).Columns(26).EditorStyle:id=261,.parent=17"
      _StyleDefs(288) =   "Splits(1).Columns(27).Style:id=270,.parent=13"
      _StyleDefs(289) =   "Splits(1).Columns(27).HeadingStyle:id=267,.parent=14"
      _StyleDefs(290) =   "Splits(1).Columns(27).FooterStyle:id=268,.parent=15"
      _StyleDefs(291) =   "Splits(1).Columns(27).EditorStyle:id=269,.parent=17"
      _StyleDefs(292) =   "Splits(1).Columns(28).Style:id=278,.parent=13"
      _StyleDefs(293) =   "Splits(1).Columns(28).HeadingStyle:id=275,.parent=14"
      _StyleDefs(294) =   "Splits(1).Columns(28).FooterStyle:id=276,.parent=15"
      _StyleDefs(295) =   "Splits(1).Columns(28).EditorStyle:id=277,.parent=17"
      _StyleDefs(296) =   "Splits(1).Columns(29).Style:id=286,.parent=13"
      _StyleDefs(297) =   "Splits(1).Columns(29).HeadingStyle:id=283,.parent=14"
      _StyleDefs(298) =   "Splits(1).Columns(29).FooterStyle:id=284,.parent=15"
      _StyleDefs(299) =   "Splits(1).Columns(29).EditorStyle:id=285,.parent=17"
      _StyleDefs(300) =   "Splits(1).Columns(30).Style:id=294,.parent=13"
      _StyleDefs(301) =   "Splits(1).Columns(30).HeadingStyle:id=291,.parent=14"
      _StyleDefs(302) =   "Splits(1).Columns(30).FooterStyle:id=292,.parent=15"
      _StyleDefs(303) =   "Splits(1).Columns(30).EditorStyle:id=293,.parent=17"
      _StyleDefs(304) =   "Splits(1).Columns(31).Style:id=302,.parent=13"
      _StyleDefs(305) =   "Splits(1).Columns(31).HeadingStyle:id=299,.parent=14"
      _StyleDefs(306) =   "Splits(1).Columns(31).FooterStyle:id=300,.parent=15"
      _StyleDefs(307) =   "Splits(1).Columns(31).EditorStyle:id=301,.parent=17"
      _StyleDefs(308) =   "Splits(1).Columns(32).Style:id=315,.parent=13"
      _StyleDefs(309) =   "Splits(1).Columns(32).HeadingStyle:id=312,.parent=14"
      _StyleDefs(310) =   "Splits(1).Columns(32).FooterStyle:id=313,.parent=15"
      _StyleDefs(311) =   "Splits(1).Columns(32).EditorStyle:id=314,.parent=17"
      _StyleDefs(312) =   "Named:id=33:Normal"
      _StyleDefs(313) =   ":id=33,.parent=0"
      _StyleDefs(314) =   "Named:id=34:Heading"
      _StyleDefs(315) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(316) =   ":id=34,.wraptext=-1"
      _StyleDefs(317) =   "Named:id=35:Footing"
      _StyleDefs(318) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(319) =   "Named:id=36:Selected"
      _StyleDefs(320) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(321) =   "Named:id=37:Caption"
      _StyleDefs(322) =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(323) =   "Named:id=38:HighlightRow"
      _StyleDefs(324) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(325) =   "Named:id=39:EvenRow"
      _StyleDefs(326) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(327) =   "Named:id=40:OddRow"
      _StyleDefs(328) =   ":id=40,.parent=33"
      _StyleDefs(329) =   "Named:id=41:RecordSelector"
      _StyleDefs(330) =   ":id=41,.parent=34"
      _StyleDefs(331) =   "Named:id=42:FilterBar"
      _StyleDefs(332) =   ":id=42,.parent=33"
      _StyleDefs(333) =   "Named:id=303:BiruPutih"
      _StyleDefs(334) =   ":id=303,.parent=39,.bgcolor=&HFF8080&,.fgcolor=&HFFFFFF&,.bold=-1,.fontsize=825"
      _StyleDefs(335) =   ":id=303,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(336) =   ":id=303,.fontname=MS Sans Serif"
      _StyleDefs(337) =   "Named:id=304:KuningHitam"
      _StyleDefs(338) =   ":id=304,.parent=303,.bgcolor=&HFFFF&,.fgcolor=&H0&"
      _StyleDefs(339) =   "Named:id=305:HijauPutih"
      _StyleDefs(340) =   ":id=305,.parent=303,.bgcolor=&H80FF80&"
      _StyleDefs(341) =   "Named:id=306:MerahPutih"
      _StyleDefs(342) =   ":id=306,.parent=305,.bgcolor=&HFF&"
      _StyleDefs(343) =   "Named:id=307:BiruHitam"
      _StyleDefs(344) =   ":id=307,.parent=303,.bgcolor=&HC08000&,.fgcolor=&H0&"
   End
   Begin TrueOleDBGrid70.TDBGrid gridBooking 
      Height          =   2970
      Left            =   690
      TabIndex        =   14
      Top             =   6345
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5239
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tanggal Service"
      Columns(0).DataField=   "tanggal"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Total"
      Columns(1).DataField=   "total"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2196"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2117"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1058"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=979"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
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
      Caption         =   "BOOKING SERVICE"
      MultipleLines   =   0
      CellTipsWidth   =   0
      CellTipsDelay   =   10
      DeadAreaBackColor=   12577002
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
      _StyleDefs(24)  =   "Splits(0).Style:id=25,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=27,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=56,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=55,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=57,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=58,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=25"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=26"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=27"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=55"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=25,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=26"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=27"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=55"
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
      _StyleDefs(65)  =   "Named:id=303:BiruPutih"
      _StyleDefs(66)  =   ":id=303,.parent=39,.bgcolor=&HFF8080&,.fgcolor=&HFFFFFF&,.bold=-1,.fontsize=825"
      _StyleDefs(67)  =   ":id=303,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(68)  =   ":id=303,.fontname=MS Sans Serif"
      _StyleDefs(69)  =   "Named:id=304:KuningHitam"
      _StyleDefs(70)  =   ":id=304,.parent=303,.bgcolor=&HFFFF&,.fgcolor=&H0&"
      _StyleDefs(71)  =   "Named:id=305:HijauPutih"
      _StyleDefs(72)  =   ":id=305,.parent=303,.bgcolor=&H80FF80&,.fgcolor=&H0&"
      _StyleDefs(73)  =   "Named:id=306:MerahPutih"
      _StyleDefs(74)  =   ":id=306,.parent=305,.bgcolor=&HFF&"
      _StyleDefs(75)  =   "Named:id=307:BiruHitam"
      _StyleDefs(76)  =   ":id=307,.parent=303,.bgcolor=&HC08000&,.fgcolor=&H0&"
   End
   Begin TrueOleDBGrid70.TDBGrid gridNotBOC 
      Height          =   3000
      Left            =   3855
      TabIndex        =   15
      Top             =   6330
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   5292
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Tanggal Service"
      Columns(0).DataField=   "tanggal"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Total"
      Columns(1).DataField=   "total"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2381"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2302"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1058"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=979"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
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
      Caption         =   "Not BOC"
      MultipleLines   =   0
      CellTipsWidth   =   0
      CellTipsDelay   =   10
      DeadAreaBackColor=   12577002
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
      _StyleDefs(24)  =   "Splits(0).Style:id=25,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=27,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=56,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=55,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=57,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=58,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=25"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=26"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=27"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=55"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=25,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=26"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=27"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=55"
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
      _StyleDefs(65)  =   "Named:id=303:BiruPutih"
      _StyleDefs(66)  =   ":id=303,.parent=39,.bgcolor=&HFF8080&,.fgcolor=&HFFFFFF&,.bold=-1,.fontsize=825"
      _StyleDefs(67)  =   ":id=303,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(68)  =   ":id=303,.fontname=MS Sans Serif"
      _StyleDefs(69)  =   "Named:id=304:KuningHitam"
      _StyleDefs(70)  =   ":id=304,.parent=303,.bgcolor=&HFFFF&,.fgcolor=&H0&"
      _StyleDefs(71)  =   "Named:id=305:HijauPutih"
      _StyleDefs(72)  =   ":id=305,.parent=303,.bgcolor=&H80FF80&,.fgcolor=&H0&"
      _StyleDefs(73)  =   "Named:id=306:MerahPutih"
      _StyleDefs(74)  =   ":id=306,.parent=305,.bgcolor=&HFF&"
      _StyleDefs(75)  =   "Named:id=307:BiruHitam"
      _StyleDefs(76)  =   ":id=307,.parent=303,.bgcolor=&HC08000&,.fgcolor=&H0&"
   End
   Begin TrueOleDBGrid70.TDBGrid gridRekap 
      Height          =   3000
      Left            =   7710
      TabIndex        =   16
      Top             =   6330
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   5292
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Hasil Konfirmasi"
      Columns(0).DataField=   "tanggal"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Total"
      Columns(1).DataField=   "total"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=5292"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5212"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=1058"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=979"
      Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=1"
      Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
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
      Caption         =   "HASIL KONFIRMASI"
      MultipleLines   =   0
      CellTipsWidth   =   0
      CellTipsDelay   =   10
      DeadAreaBackColor=   12577002
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
      _StyleDefs(24)  =   "Splits(0).Style:id=25,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=64,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=27,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=56,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=55,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=57,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=58,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=63,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=65,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=66,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=16,.parent=25"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=26"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=27"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=55"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=20,.parent=25,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=26"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=27"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=55"
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
      _StyleDefs(65)  =   "Named:id=303:BiruPutih"
      _StyleDefs(66)  =   ":id=303,.parent=39,.bgcolor=&HFF8080&,.fgcolor=&HFFFFFF&,.bold=-1,.fontsize=825"
      _StyleDefs(67)  =   ":id=303,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(68)  =   ":id=303,.fontname=MS Sans Serif"
      _StyleDefs(69)  =   "Named:id=304:KuningHitam"
      _StyleDefs(70)  =   ":id=304,.parent=303,.bgcolor=&HFFFF&,.fgcolor=&H0&"
      _StyleDefs(71)  =   "Named:id=305:HijauPutih"
      _StyleDefs(72)  =   ":id=305,.parent=303,.bgcolor=&H80FF80&,.fgcolor=&H0&"
      _StyleDefs(73)  =   "Named:id=306:MerahPutih"
      _StyleDefs(74)  =   ":id=306,.parent=305,.bgcolor=&HFF&"
      _StyleDefs(75)  =   "Named:id=307:BiruHitam"
      _StyleDefs(76)  =   ":id=307,.parent=303,.bgcolor=&HC08000&,.fgcolor=&H0&"
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   13095
      Shape           =   4  'Rounded Rectangle
      Top             =   6390
      Width           =   2790
   End
   Begin VB.Label nBulan 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AGUSTUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13635
      TabIndex        =   5
      Top             =   750
      Width           =   1695
   End
   Begin VB.Label nTahun 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Periode  :  2009"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1350
      TabIndex        =   6
      Top             =   720
      Width           =   2295
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   435
      Left            =   13545
      Shape           =   4  'Rounded Rectangle
      Top             =   690
      Width           =   1875
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   705
      X2              =   15885
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Shape Shape12 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   435
      Left            =   1215
      Shape           =   4  'Rounded Rectangle
      Top             =   615
      Width           =   2595
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "MRA Monitoring Reminder Board"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   13710
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   750
      X2              =   4470
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   1950
      X2              =   4470
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      X1              =   2790
      X2              =   4470
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00404040&
      Height          =   495
      Left            =   540
      Shape           =   4  'Rounded Rectangle
      Top             =   45
      Width           =   15510
   End
   Begin VB.Image Image1 
      Height          =   8655
      Left            =   15
      Picture         =   "FormMRABoard.frx":5D6A
      Top             =   15
      Width           =   495
   End
End
Attribute VB_Name = "FormMRABoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsRecord As New ADODB.Recordset

Private rsReminder As New ADODB.Recordset
Private rsReminderCall As New ADODB.Recordset
Private rsGrafik As New ADODB.Recordset
Private rsTemp As ADODB.Recordset
Private Bulan(12) As String
Private isProses As Boolean
Private Bln As Integer
Private Thn As Integer
Private targetReminder As Double
Private bookingFix As Double

Private Sub createHeaderGrid()
    Dim Tanggal As Date
    Dim totalHeader As Integer
    Dim i As Integer
   On Error GoTo createHeaderGrid_Error

    i = 0
    Tanggal = Bln & "/01/" & Thn
    
    Set rsTemp = objO.FetchData("select datediff(day, dateadd(day, 1-day('8/01/2017'), '8/01/2017'),dateadd(month, 1, dateadd(day, 1-day('8/01/2017'), '8/01/2017')))")
    If Not rsTemp.EOF Then
        totalHeader = CInt(rsTemp.Fields(0).value)
    Else
        MsgBox "Tidak dapat konek kedatabase", vbInformation
        Exit Sub
    End If
    
    With gridMonitoring
        .Columns(0).Width = 5000
        .Columns(0).Caption = "ITEM"
        .Columns(0).DataField = "item"
        .Splits(0).Columns(0).Locked = True
        .Splits(1).Columns(0).Visible = False
        i = 1
        
        For i = 1 To totalHeader
            .Splits(1).Columns(i).Width = 600
            .Splits(1).Columns(i).Caption = i
            .Splits(1).Columns(i).DataField = "tgl" & i
            .Splits(1).Columns(i).Alignment = dbgCenter
            .Splits(1).Columns(i).FetchStyle = True
            .Splits(1).Columns(i).Button = True
            '.Columns(i).ButtonAlways = True
            .Splits(1).Columns(i).ButtonText = True
            '.Splits(1).Columns(i).Locked = True
            .Splits(0).Columns(i).Visible = False
        Next
        
        .Columns(i).Width = 5000
        .Columns(i).Caption = "Total"
        .Columns(i).DataField = "total"
        .Splits(0).Columns(i).Locked = True
        .Splits(0).Columns(i).Visible = False
        .Splits(1).Columns(i).Visible = True
    End With
    
'    If totalHeader < 33 Then
'        totalHeader = 33 - totalHeader
'        With gridMonitoring
'
'            For i = 1 To totalHeader
'                .Splits(0).Columns(i).Visible = False
'                .Splits(1).Columns(i).Visible = False
'            Next
'        End With
'    End If

   On Error GoTo 0
   Exit Sub

createHeaderGrid_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createHeaderGrid of Form FormMRABoard"
    
End Sub

Private Sub createTableRecord()
    If rsRecord.State Then rsRecord.Close
    Set rsRecord = Nothing
    With rsRecord
        .Fields.Append "item", adVarChar, 100
        .Fields.Append "tgl1", adVarChar, 100
        .Fields.Append "tgl2", adVarChar, 100
        .Fields.Append "tgl3", adVarChar, 100
        .Fields.Append "tgl4", adVarChar, 100
        .Fields.Append "tgl5", adVarChar, 100
        .Fields.Append "tgl6", adVarChar, 100
        .Fields.Append "tgl7", adVarChar, 100
        .Fields.Append "tgl8", adVarChar, 100
        .Fields.Append "tgl9", adVarChar, 100
        .Fields.Append "tgl10", adVarChar, 100
        .Fields.Append "tgl11", adVarChar, 100
        .Fields.Append "tgl12", adVarChar, 100
        .Fields.Append "tgl13", adVarChar, 100
        .Fields.Append "tgl14", adVarChar, 100
        .Fields.Append "tgl15", adVarChar, 100
        .Fields.Append "tgl16", adVarChar, 100
        .Fields.Append "tgl17", adVarChar, 100
        .Fields.Append "tgl18", adVarChar, 100
        .Fields.Append "tgl19", adVarChar, 100
        .Fields.Append "tgl20", adVarChar, 100
        .Fields.Append "tgl21", adVarChar, 100
        .Fields.Append "tgl22", adVarChar, 100
        .Fields.Append "tgl23", adVarChar, 100
        .Fields.Append "tgl24", adVarChar, 100
        .Fields.Append "tgl25", adVarChar, 100
        .Fields.Append "tgl26", adVarChar, 100
        .Fields.Append "tgl27", adVarChar, 100
        .Fields.Append "tgl28", adVarChar, 100
        .Fields.Append "tgl29", adVarChar, 100
        .Fields.Append "tgl30", adVarChar, 100
        .Fields.Append "tgl31", adVarChar, 100
        .Fields.Append "total", adVarChar, 100
        .Open
    End With
    
    Call createHeaderGrid
    Set gridMonitoring.DataSource = rsRecord
    
    
End Sub

Private Sub createTableReminder()
    If rsReminderCall.State Then rsReminderCall.Close
    Set rsReminderCall = Nothing
    With rsReminderCall
        .Fields.Append "NoRangka", adVarChar, 100
        .Fields.Append "Type", adVarChar, 100
        .Fields.Append "Nama", adVarChar, 100
        .Fields.Append "Alamat", adVarChar, 200
        .Fields.Append "Kota", adVarChar, 100
        .Fields.Append "KodePos", adVarChar, 100
        .Fields.Append "Telp", adVarChar, 100
        .Fields.Append "GSM", adVarChar, 100
        .Fields.Append "CDMA", adVarChar, 100
        .Fields.Append "tglReminderSales", adDate
        .Fields.Append "namasales", adVarChar, 100
        .Fields.Append "tglBOC", adDate
        .Fields.Append "warna", adVarChar, 100
        .Open
    End With
     
End Sub

Private Sub createTableGrafik()
    If rsGrafik.State Then rsGrafik.Close
    Set rsGrafik = Nothing
    With rsGrafik
        .Fields.Append "keterangan", adVarChar, 100
        .Fields.Append "nilai", adInteger
        .Open
    End With
     
End Sub

Private Sub btnTutup_Click()
    Frame2.Visible = False
    btnTutup.Visible = False
End Sub

Private Sub cmdCall_Click()
      '    If Nul(rsReminder.Fields(0).value) = "" Then
      '        MsgBox "Tidak dapat proses reminder, data norangka kosong", vbInformation
      '        Exit Sub
      '    End If
          
1         If Nul(gridReminder.Columns(0).value) = "" Then
2             MsgBox "Tidak dapat proses reminder, data norangka kosong", vbInformation
3             Exit Sub
4         End If
          
5         sSQL = "SELECT * FROM custdat WITH (NOLOCK) WHERE no_chassis = '" & Nul(gridReminder.Columns(0).value) & "'"
6         Set rsTemp = objO.FetchData(sSQL)
7         If Not rsTemp.EOF Then
8             sSQL = "SELECT a.NoInv, a.NoWo, a.TglSvcNow as DateLastSvc,a.KmNow, "
9             sSQL = sSQL & " (CASE WHEN A.JT_Tgl > A.JT_Bln then A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then A.Jt_Bln END) as DateNextSvc,"
10            sSQL = sSQL & " a.KmNext as KmNextSvc, c.UseName, c.UseAddr, c.UseCity, c.UsePos, c.No_Polisi as NoPolisi, C.FMDL as KetModel,"
11            sSQL = sSQL & " a.NoRangka as NoChassis, c.UsePSTN as Telpon , c.UseGSM, c.UseCDMA, a.NoRangka as NoChassisSMS, c.FS_KD_PLG as KodeLGN,"
12            sSQL = sSQL & " b.FS_RC, A.SbeNext, C.COLOR as WrnModel, C.USEEMAIL AS EMAIL, C.MDL, C.MODEL, C.FS_TH_MOBIL, C.COLOR, C.FS_KD_DEALER, C.FS_NM_DEALER FROM M_ESTIMASI A WITH(NOLOCK) Inner Join"
13            sSQL = sSQL & " (SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND ISNULL(TGLSVCNOW,'') <> '') S, MAX((CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END)) AS JT_TGL"
14            sSQL = sSQL & " FROM M_ESTIMASI G WITH(NOLOCK) Where (CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END) Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S"
15            sSQL = sSQL & " INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = ''"
16            sSQL = sSQL & " INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis"
17            sSQL = sSQL & " Where a.NoRangka = '" & Nul(gridReminder.Columns(0).value) & "'"
              Dim rsRem As ADODB.Recordset
18            Set rsRem = objO.FetchData(sSQL)
              Dim fCL As New RemainderCallRetention
19            If Not rsRem.EOF Then
                  
20                Set rsTemp = objO.FetchData("SELECT * FROM [listReminderBoard] WHERE norangka = '" & Nul(gridReminder.Columns(0).value) & "' and nowo = '" & Trim(rsRem!noWO) & "'")
21                If Not rsTemp.EOF Then
22                    MsgBox "Data telah direminder pada tanggal " & rsTemp!tglentry, vbInformation
23                    Exit Sub
24                End If
                  'Dim fCL As New RemainderCall
25                With fCL
26                   .xRsRem = rsRem
27                   .xRsAll = rsRem
28                   .xTglNext = Format(rsRem!DateNextSvc, "dd/mm/yyyy") 'isTgl
29                   .xNoPol = Nul(rsRem!NoPolisi) 'isNoPol
30                   .IsSvc1000 = False
31                   .xIsMRABoard = True
                     '.Left = Me.Left
                     '.Top = Me.Top
32                   .show vbModal
33                   Call cmdReminderClose_Click
34                End With
35            Else
                  'CEK SBI
36                sSQL = "Select DISTINCT b.SPK_No as NoSPK, f.Tgl_BPK as TglDO, i.TglCR1, i.KmCR1, i.TglCR2, isNull(i.KMCR2,0) as KmCR2 , b.SNama as UseName, b.SAlamat as UseAddr, b.SKota as UseCity, b.SKodePos as UsePos, a.NoPol as NoPolisi, h.NoMesin as NoEngine, b.NoRangka as NoChassis, b.STelp as Telpon, b.SGsm as GSM, b.SCdma as CDMA, d.KodeKategory as KodeType, d.NamaKategory as NamaType, e.KodeWarna, e.NamaEnglish as NamaWarna, g.NamaSales as Wiraniaga, b.Nama_Faktur as NamaSTNK, (a.Alamat1 + a.Alamat2 + a.Alamat3 + a.Alamat4 + a.Alamat5) As AlamatSTNK, b.Kota as KotaSTNK, a.KodePos as KodePosSTNK, a.TmptLahir as TempatLahir, a.TglLahir as TanggalLahir, a.Agama, a.Pekerjaan, a.Hoby, i.Tgl1, b.NoRangka as No_Chassis, B.Kode_Lgn , h.KodeTahun, e.NamaInd  ,B.nama_pemesan 'Namaspk', b.alamat1+' '+b.alamat2 'Alamatspk', b.telp_rumah 'Telpspk', b.hp1 'GSM1spk', b.hp2 'GSM2spk' "
37                sSQL = sSQL & " From [SPK" & Right(Log.LogSPK, 2) & "]..AFI a WITH(NOLOCK)"
38                sSQL = sSQL & " INNER JOIN  [SPK" & Right(Log.LogSPK, 2) & "]..SPK b WITH(NOLOCK) ON a.SPK_No = b.SPK_No"
39                sSQL = sSQL & " INNER JOIN [SPK" & Right(Log.LogSPK, 2) & "]..INMerek c WITH(NOLOCK) ON b.Kode_Merk = c.KodeMerek"
40                sSQL = sSQL & " INNER JOIN [SPK" & Right(Log.LogSPK, 2) & "]..InKategory d WITH(NOLOCK) ON b.Kode_Type = d.KodeKategory"
41                sSQL = sSQL & " INNER JOIN [SPK" & Right(Log.LogSPK, 2) & "]..InWarnaExt e WITH(NOLOCK) ON b.Kode_Warna = e.KodeWarna"
42                sSQL = sSQL & " INNER JOIN [SPK" & Right(Log.LogSPK, 2) & "]..BPK f WITH(NOLOCK) ON a.SPK_No = f.SPK_No"
43                sSQL = sSQL & " INNER JOIN [SPK" & Right(Log.LogSPK, 2) & "]..ArSales g WITH(NOLOCK) ON b.Kode_Sales = g.KodeSales"
44                sSQL = sSQL & " INNER JOIN [SPK" & Right(Log.LogSPK, 2) & "]..INMst h WITH(NOLOCK) ON b.NoRangka = h.kodeItem"
45                sSQL = sSQL & " INNER JOIN [SPK" & Right(Log.LogSPK, 2) & "]..T_MRS i WITH(NOLOCK) ON b.SPK_No = i.NoSPK"
46                sSQL = sSQL & " Where i.Km1 = '1000' And a.batal = 0 and b.norangka = '" & Nul(gridReminder.Columns(0).value) & "'"
47                Set rsRem = objO.FetchData(sSQL)
                  'Dim fCL As New RemainderCall
48                If Not rsRem.EOF Then
49                    Set rsTemp = objO.FetchData("SELECT * FROM [listReminderBoard] WHERE norangka = '" & Nul(gridReminder.Columns(0).value) & "' ")
50                    If Not rsTemp.EOF Then
51                        MsgBox "Data telah direminder pada tanggal " & rsTemp!tglentry, vbInformation
52                        Exit Sub
53                    End If
                      
54                    With fCL
55                       .xRsRem = rsRem
56                       .xRsAll = rsRem
57                       .xTglNext = Format(NulTgl(rsRem!TglCR1), "dd/mm/yyyy")   'isTgl
58                       .xNoPol = Nul(rsRem!NoPolisi) 'isNoPol
59                       .IsSvc1000 = True
60                       .xIsMRABoard = True
                         '.Left = Me.Left
                         '.Top = Me.Top
61                       .show vbModal
62                       Call cmdReminderClose_Click
63                    End With
64                Else
65                    MsgBox "Tidak ada reminder untuk norangka " & Nul(gridReminder.Columns(0).value), vbInformation
66                    Exit Sub
67                End If
68            End If
              
69        Else
70            MsgBox "No Rangka " & Nul(gridReminder.Columns(0).value) & " tidak ditemukan didata customer bengkel.", vbInformation
71            Exit Sub
72        End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHistory_Click()
    Dim ft As New FormTransparan
    If Nul(gridReminder.Columns(0).value) = "" Then
        MsgBox "Tidak dapat proses lihat history, data norangka kosong", vbInformation
        Exit Sub
    End If
    
    If MsgBox("Apakah ingin melihat History Service HO?", vbYesNo, "Ambil Data Dari HO") = vbYes Then
        If Log.LogHOConn = 0 And Log.RPLHistori = 0 Then
            '' 2015-09-08 | Konek HO ''
            If MsgBox("Gagal Koneksi ke Server HO !" & vbCrLf & "Apakah ingin menyambungkan koneksi kembali?", vbYesNo) = vbYes Then
                If Get_ConnectionProfil_Bengkel99 Then
                    Log.LogHOConn = 1
                Else
                    Exit Sub
                End If
            Else
                Exit Sub
            End If
    '        MsgBox "Gagal Koneksi ke Server HO !", vbInformation + vbOKOnly, "History On Line HO"
    '        Exit Sub
        End If
        'jika mengambil data dari ho
       
        ft.xChassis = Nul(gridReminder.Columns(0).value)
        ft.xIsForm = "HIST_HO"
        ft.show (vbModal)
    Else
        'jika mengambil data local
        'Dim ft As New FormTransparan
        ft.xChassis = Nul(gridReminder.Columns(0).value)
        ft.xIsForm = "HIST"
        ft.show (vbModal)
    End If
End Sub

Private Sub cmdOpen_Click()
    If isProses = True Then
        Call createTableGrafik
        
        With rsGrafik
            .AddNew
            !keterangan = "TARGET TOTAL REMINDER"
            !nilai = targetReminder
            .Update
            .AddNew
            !keterangan = "BOOKING"
            !nilai = bookingFix
            .Update
        End With
        
        With MSChart1
            Set .DataSource = rsGrafik
              '.ChartData
             .ShowLegend = True
             '.SeriesColumn = True
             '.Stacking = True
        End With
        MSChart1.Visible = True
        Frame2.Visible = True
        btnTutup.Visible = True
    Else
        MsgBox "Silahkan proses data terlebih dahulu", vbInformation
    End If
End Sub

Private Sub cmdProcess_Click()
    Call createTableRecord
    Call refreshDatagrid
    isProses = True
End Sub

Private Sub refreshDatagrid()
          Dim totalHeader As Integer
          Dim i As Integer
          Dim tanggalKolom As String
          Dim rsItem As ADODB.Recordset
          Dim rsBooking As ADODB.Recordset
          Dim rsNotBOC As ADODB.Recordset
          Dim rsRekap As ADODB.Recordset
1         Set rsItem = objO.FetchData("select pesan from M_VoiceOfCustomer with (nolock)")
2         targetReminder = 0
3         bookingFix = 0
4         On Error GoTo exo
              
5             Set rsBooking = objO.FetchData("SELECT b.date as tanggal,COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "' GROUP BY b.date ORDER BY b.date")
              'Set rsNotBOC = objO.FetchDataSPK("SELECT convert(varchar(20),tgl_entry_retensi,103) as tanggal,count(kodepr) as total FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and isnull(boc,'') <> 'BOC' group by tgl_entry_retensi order by tgl_entry_retensi")
              
              '' 20200408DEV
6             If Log.sentralisasiStars = 1 Then
7                 Set rsNotBOC = objO.FetchData("SELECT convert(varchar(20),tgl_entry_retensi,103) as tanggal,count(kodepr) as total FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and isnull(boc,'') <> 'BOC' and kodecabang = '" & Log.kodeCabang & "' group by convert(varchar(20),tgl_entry_retensi,103) order by convert(varchar(20),tgl_entry_retensi,103)")
8             Else
9                 Set rsNotBOC = objO.FetchDataSPK("SELECT convert(varchar(20),tgl_entry_retensi,103) as tanggal,count(kodepr) as total FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and isnull(boc,'') <> 'BOC' group by convert(varchar(20),tgl_entry_retensi,103) order by convert(varchar(20),tgl_entry_retensi,103)")
10            End If
              'Set rsRekap = objO.FetchDataSPK("SELECT  RetensiAlasan as tanggal,count(kodepr) as total FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' group by RetensiAlasan ")
              
11            sSQL = "select a.pesan as tanggal,isnull(b.total,0) as total"
12            sSQL = sSQL & " from M_VoiceOfCustomer a"
13            sSQL = sSQL & " left join ("
14            sSQL = sSQL & " select c.pesan, COUNT(b.pesan_telp) as total"
15            sSQL = sSQL & " from listReminderBoard a"
16            sSQL = sSQL & " inner join M_ListReminderCall b on a.NoRangka = b.NoRangka and a.NoWo = b.LastNoWO"
17            sSQL = sSQL & " inner join M_VoiceOfCustomer c on c.pesan = b.Pesan_Telp WHERE year(a.tglentry) = '" & Thn & "' and month(a.tglentry) = '" & Bln & "'"
18            sSQL = sSQL & " group by c.pesan"
19            sSQL = sSQL & " ) b on a.pesan = b.pesan"
              
20            Set rsRekap = objO.FetchData(sSQL)
              
21            Set gridBooking.DataSource = rsBooking
22            Set gridNotBOC.DataSource = rsNotBOC
23            Set GridRekap.DataSource = rsRekap
              
24            GridRekap.Refresh
25            gridNotBOC.Refresh
26            gridBooking.Refresh
27            Set rsTemp = objO.FetchData("select datediff(day, dateadd(day, 1-day('8/01/2017'), '8/01/2017'),dateadd(month, 1, dateadd(day, 1-day('8/01/2017'), '8/01/2017')))")
28            If Not rsTemp.EOF Then
29                totalHeader = CInt(rsTemp.Fields(0).value)
30            End If
31            rsItem.MoveFirst
              
32            With rsRecord
33                .AddNew
34                !Item = UCase("O/S Reminder Sales")
                  '20180320 REV ' PENAMBAHAN KETERANGAN 1000KM UNTUK AKTIVITAS SALES
                  '' 20200408DEV
35                If Log.sentralisasiStars = 1 Then
36                    !Tgl1 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '1' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
37                    !Tgl2 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '2' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
38                    !Tgl3 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '3' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
39                    !Tgl4 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '4' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
40                    !Tgl5 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '5' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
41                    !Tgl6 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '6' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
42                    !Tgl7 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '7' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
43                    !Tgl8 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '8' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
44                    !Tgl9 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '9' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
45                    !Tgl10 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '10' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
46                    !Tgl11 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '11' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
47                    !Tgl12 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '12' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
48                    !Tgl13 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '13' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
49                    !Tgl14 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '14' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
50                    !Tgl15 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '15' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
51                    !Tgl16 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '16' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
52                    !Tgl17 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '17' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
53                    !Tgl18 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '18' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
54                    !Tgl19 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '19' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
55                    !Tgl20 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '20' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
56                    !Tgl21 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '21' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
57                    !Tgl22 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '22' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
58                    !Tgl23 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '23' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
59                    !Tgl24 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '24' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
60                    !Tgl25 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '25' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
61                    !Tgl26 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '26' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
62                    !Tgl27 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '27' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
63                    !Tgl28 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '28' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
64                    !Tgl29 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '29' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
65                    !Tgl30 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '30' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
66                    !Tgl31 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '31' and Keterangan like '%reminder service 1000 km%' and kodecabang = '" & Log.kodeCabang & "'")
67                Else
68                    !Tgl1 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '1' and Keterangan like '%reminder service 1000 km%'")
69                    !Tgl2 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '2' and Keterangan like '%reminder service 1000 km%'")
70                    !Tgl3 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '3' and Keterangan like '%reminder service 1000 km%'")
71                    !Tgl4 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '4' and Keterangan like '%reminder service 1000 km%'")
72                    !Tgl5 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '5' and Keterangan like '%reminder service 1000 km%'")
73                    !Tgl6 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '6' and Keterangan like '%reminder service 1000 km%'")
74                    !Tgl7 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '7' and Keterangan like '%reminder service 1000 km%'")
75                    !Tgl8 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '8' and Keterangan like '%reminder service 1000 km%'")
76                    !Tgl9 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '9' and Keterangan like '%reminder service 1000 km%'")
77                    !Tgl10 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '10' and Keterangan like '%reminder service 1000 km%'")
78                    !Tgl11 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '11' and Keterangan like '%reminder service 1000 km%'")
79                    !Tgl12 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '12' and Keterangan like '%reminder service 1000 km%'")
80                    !Tgl13 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '13' and Keterangan like '%reminder service 1000 km%'")
81                    !Tgl14 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '14' and Keterangan like '%reminder service 1000 km%'")
82                    !Tgl15 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '15' and Keterangan like '%reminder service 1000 km%'")
83                    !Tgl16 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '16' and Keterangan like '%reminder service 1000 km%'")
84                    !Tgl17 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '17' and Keterangan like '%reminder service 1000 km%'")
85                    !Tgl18 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '18' and Keterangan like '%reminder service 1000 km%'")
86                    !Tgl19 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '19' and Keterangan like '%reminder service 1000 km%'")
87                    !Tgl20 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '20' and Keterangan like '%reminder service 1000 km%'")
88                    !Tgl21 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '21' and Keterangan like '%reminder service 1000 km%'")
89                    !Tgl22 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '22' and Keterangan like '%reminder service 1000 km%'")
90                    !Tgl23 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '23' and Keterangan like '%reminder service 1000 km%'")
91                    !Tgl24 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '24' and Keterangan like '%reminder service 1000 km%'")
92                    !Tgl25 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '25' and Keterangan like '%reminder service 1000 km%'")
93                    !Tgl26 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '26' and Keterangan like '%reminder service 1000 km%'")
94                    !Tgl27 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '27' and Keterangan like '%reminder service 1000 km%'")
95                    !Tgl28 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '28' and Keterangan like '%reminder service 1000 km%'")
96                    !Tgl29 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '29' and Keterangan like '%reminder service 1000 km%'")
97                    !Tgl30 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '30' and Keterangan like '%reminder service 1000 km%'")
98                    !Tgl31 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '31' and Keterangan like '%reminder service 1000 km%'")
99                End If
100               !total = CInt(!Tgl1) + CInt(!Tgl2) + CInt(!Tgl3) + CInt(!Tgl4) + CInt(!Tgl5) + CInt(!Tgl6) + CInt(!Tgl7) + CInt(!Tgl8) + CInt(!Tgl9) + CInt(!Tgl10) + CInt(!Tgl11) + CInt(!Tgl12) + CInt(!Tgl13) + CInt(!Tgl14) + CInt(!Tgl15) + CInt(!Tgl16) + CInt(!Tgl17) + CInt(!Tgl18) + CInt(!Tgl19) + CInt(!Tgl20) + CInt(!Tgl21) + CInt(!Tgl22) + CInt(!Tgl23) + CInt(!Tgl24) + CInt(!Tgl25) + CInt(!Tgl26) + CInt(!Tgl27) + CInt(!Tgl28) + CInt(!Tgl29) + CInt(!Tgl30) + CInt(!Tgl31)
101               .Update
                  
102               .AddNew
103               !Item = UCase("O/S Booking Sales")
                  '' 20200408DEV
104               If Log.sentralisasiStars = 1 Then
105                   !Tgl1 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '1' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
106                   !Tgl2 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '2' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
107                   !Tgl3 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '3' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
108                   !Tgl4 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '4' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
109                   !Tgl5 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '5' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
110                   !Tgl6 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '6' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
111                   !Tgl7 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '7' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
112                   !Tgl8 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '8' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
113                   !Tgl9 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '9' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
114                   !Tgl10 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '10' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
115                   !Tgl11 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '11' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
116                   !Tgl12 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '12' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
117                   !Tgl13 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '13' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
118                   !Tgl14 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '14' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
119                   !Tgl15 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '15' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
120                   !Tgl16 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '16' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
121                   !Tgl17 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '17' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
122                   !Tgl18 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '18' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
123                   !Tgl19 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '19' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
124                   !Tgl20 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '20' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
125                   !Tgl21 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '21' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
126                   !Tgl22 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '22' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
127                   !Tgl23 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '23' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
128                   !Tgl24 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '24' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
129                   !Tgl25 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '25' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
130                   !Tgl26 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '26' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
131                   !Tgl27 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '27' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
132                   !Tgl28 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '28' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
133                   !Tgl29 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '29' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
134                   !Tgl30 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '30' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
135                   !Tgl31 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '31' and boc = 'BOC' and kodecabang = '" & Log.kodeCabang & "'")
136               Else
137                   !Tgl1 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '1' and boc = 'BOC'")
138                   !Tgl2 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '2' and boc = 'BOC'")
139                   !Tgl3 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '3' and boc = 'BOC'")
140                   !Tgl4 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '4' and boc = 'BOC'")
141                   !Tgl5 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '5' and boc = 'BOC'")
142                   !Tgl6 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '6' and boc = 'BOC'")
143                   !Tgl7 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '7' and boc = 'BOC'")
144                   !Tgl8 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '8' and boc = 'BOC'")
145                   !Tgl9 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '9' and boc = 'BOC'")
146                   !Tgl10 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '10' and boc = 'BOC'")
147                   !Tgl11 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '11' and boc = 'BOC'")
148                   !Tgl12 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '12' and boc = 'BOC'")
149                   !Tgl13 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '13' and boc = 'BOC'")
150                   !Tgl14 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '14' and boc = 'BOC'")
151                   !Tgl15 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '15' and boc = 'BOC'")
152                   !Tgl16 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '16' and boc = 'BOC'")
153                   !Tgl17 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '17' and boc = 'BOC'")
154                   !Tgl18 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '18' and boc = 'BOC'")
155                   !Tgl19 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '19' and boc = 'BOC'")
156                   !Tgl20 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '20' and boc = 'BOC'")
157                   !Tgl21 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '21' and boc = 'BOC'")
158                   !Tgl22 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '22' and boc = 'BOC'")
159                   !Tgl23 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '23' and boc = 'BOC'")
160                   !Tgl24 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '24' and boc = 'BOC'")
161                   !Tgl25 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '25' and boc = 'BOC'")
162                   !Tgl26 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '26' and boc = 'BOC'")
163                   !Tgl27 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '27' and boc = 'BOC'")
164                   !Tgl28 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '28' and boc = 'BOC'")
165                   !Tgl29 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '29' and boc = 'BOC'")
166                   !Tgl30 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '30' and boc = 'BOC'")
167                   !Tgl31 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '31' and boc = 'BOC'")
168               End If
169               !total = CInt(!Tgl1) + CInt(!Tgl2) + CInt(!Tgl3) + CInt(!Tgl4) + CInt(!Tgl5) + CInt(!Tgl6) + CInt(!Tgl7) + CInt(!Tgl8) + CInt(!Tgl9) + CInt(!Tgl10) + CInt(!Tgl11) + CInt(!Tgl12) + CInt(!Tgl13) + CInt(!Tgl14) + CInt(!Tgl15) + CInt(!Tgl16) + CInt(!Tgl17) + CInt(!Tgl18) + CInt(!Tgl19) + CInt(!Tgl20) + CInt(!Tgl21) + CInt(!Tgl22) + CInt(!Tgl23) + CInt(!Tgl24) + CInt(!Tgl25) + CInt(!Tgl26) + CInt(!Tgl27) + CInt(!Tgl28) + CInt(!Tgl29) + CInt(!Tgl30) + CInt(!Tgl31)
170               .Update
                  
171               .AddNew
172               !Item = UCase("Booking Sales Fix")
                  
173               !Tgl1 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '1'")
174               !Tgl2 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '2'")
175               !Tgl3 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '3'")
176               !Tgl4 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '4'")
177               !Tgl5 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '5'")
178               !Tgl6 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '6'")
179               !Tgl7 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '7'")
180               !Tgl8 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '8'")
181               !Tgl9 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '9'")
182               !Tgl10 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '10'")
183               !Tgl11 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '11'")
184               !Tgl12 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '12'")
185               !Tgl13 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '13'")
186               !Tgl14 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '14'")
187               !Tgl15 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '15'")
188               !Tgl16 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '16'")
189               !Tgl17 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '17'")
190               !Tgl18 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '18'")
191               !Tgl19 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '19'")
192               !Tgl20 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '20'")
193               !Tgl21 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '21'")
194               !Tgl22 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '22'")
195               !Tgl23 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '23'")
196               !Tgl24 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '24'")
197               !Tgl25 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '25'")
198               !Tgl26 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '26'")
199               !Tgl27 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '27'")
200               !Tgl28 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '28'")
201               !Tgl29 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '29'")
202               !Tgl30 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '30'")
203               !Tgl31 = objO.FetchData_Double("SELECT COUNT(a.nobooking) total FROM listReminderBoard A INNER JOIN BOOK b on a.NoBooking = b.NOFAK and b.FS_LOG_DEL = '' where year(b.DATE) = '" & Thn & "' and MONTH(b.date) = '" & Bln & "'  and DAY(b.date) = '31'")
204               !total = CInt(!Tgl1) + CInt(!Tgl2) + CInt(!Tgl3) + CInt(!Tgl4) + CInt(!Tgl5) + CInt(!Tgl6) + CInt(!Tgl7) + CInt(!Tgl8) + CInt(!Tgl9) + CInt(!Tgl10) + CInt(!Tgl11) + CInt(!Tgl12) + CInt(!Tgl13) + CInt(!Tgl14) + CInt(!Tgl15) + CInt(!Tgl16) + CInt(!Tgl17) + CInt(!Tgl18) + CInt(!Tgl19) + CInt(!Tgl20) + CInt(!Tgl21) + CInt(!Tgl22) + CInt(!Tgl23) + CInt(!Tgl24) + CInt(!Tgl25) + CInt(!Tgl26) + CInt(!Tgl27) + CInt(!Tgl28) + CInt(!Tgl29) + CInt(!Tgl30) + CInt(!Tgl31)
205               bookingFix = !total
206               .Update
207           End With
208           Do While Not rsItem.EOF
209               If UCase(Trim(rsItem!pesan)) <> UCase("CUSTOMER DEAL & BOOKING ON CALL") Then
210                   With rsRecord
211                       .AddNew
212                       !Item = UCase(Trim(rsItem!pesan))
213                       For i = 1 To totalHeader
214                           If 1 = i Then
                                  '' 20200408DEV
215                               If Log.sentralisasiStars = 1 Then
216                                   !Tgl1 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '1' and retensialasan= '" & Trim(rsItem!pesan) & "'")
217                               Else
218                                   !Tgl1 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '1' and retensialasan= '" & Trim(rsItem!pesan) & "'")
219                               End If
220                           ElseIf 2 = i Then
                                  '' 20200408DEV
221                               If Log.sentralisasiStars = 1 Then
222                                   !Tgl2 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '2' and retensialasan= '" & Trim(rsItem!pesan) & "'")
223                               Else
224                                   !Tgl2 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '2' and retensialasan= '" & Trim(rsItem!pesan) & "'")
225                               End If
226                           ElseIf 3 = i Then
                                  '' 20200408DEV
227                               If Log.sentralisasiStars = 1 Then
228                                   !Tgl3 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '3' and retensialasan= '" & Trim(rsItem!pesan) & "'")
229                               Else
230                                   !Tgl3 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '3' and retensialasan= '" & Trim(rsItem!pesan) & "'")
231                               End If
232                           ElseIf 4 = i Then
                                  '' 20200408DEV
233                               If Log.sentralisasiStars = 1 Then
234                                   !Tgl4 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '4' and retensialasan= '" & Trim(rsItem!pesan) & "'")
235                               Else
236                                   !Tgl4 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '4' and retensialasan= '" & Trim(rsItem!pesan) & "'")
237                               End If
238                           ElseIf 5 = i Then
                                  '' 20200408DEV
239                               If Log.sentralisasiStars = 1 Then
240                                   !Tgl5 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '5' and retensialasan= '" & Trim(rsItem!pesan) & "'")
241                               Else
242                                   !Tgl5 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '5' and retensialasan= '" & Trim(rsItem!pesan) & "'")
243                               End If
244                           ElseIf 6 = i Then
                                  '' 20200408DEV
245                               If Log.sentralisasiStars = 1 Then
246                                   !Tgl6 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '6' and retensialasan= '" & Trim(rsItem!pesan) & "'")
247                               Else
248                                   !Tgl6 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '6' and retensialasan= '" & Trim(rsItem!pesan) & "'")
249                               End If
250                           ElseIf 7 = i Then
                                  '' 20200408DEV
251                               If Log.sentralisasiStars = 1 Then
252                                   !Tgl7 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '7' and retensialasan= '" & Trim(rsItem!pesan) & "'")
253                               Else
254                                   !Tgl7 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '7' and retensialasan= '" & Trim(rsItem!pesan) & "'")
255                               End If
256                           ElseIf 8 = i Then
                                  '' 20200408DEV
257                               If Log.sentralisasiStars = 1 Then
258                                   !Tgl8 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '8' and retensialasan= '" & Trim(rsItem!pesan) & "'")
259                               Else
260                                   !Tgl8 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '8' and retensialasan= '" & Trim(rsItem!pesan) & "'")
261                               End If
262                           ElseIf 9 = i Then
                                  '' 20200408DEV
263                               If Log.sentralisasiStars = 1 Then
264                                   !Tgl9 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '9' and retensialasan= '" & Trim(rsItem!pesan) & "'")
265                               Else
266                                   !Tgl9 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '9' and retensialasan= '" & Trim(rsItem!pesan) & "'")
267                               End If
268                           ElseIf 10 = i Then
                                  '' 20200408DEV
269                               If Log.sentralisasiStars = 1 Then
270                                   !Tgl10 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '10' and retensialasan= '" & Trim(rsItem!pesan) & "'")
271                               Else
272                                   !Tgl10 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '10' and retensialasan= '" & Trim(rsItem!pesan) & "'")
273                               End If
274                           ElseIf 11 = i Then
                                  '' 20200408DEV
275                               If Log.sentralisasiStars = 1 Then
276                                   !Tgl11 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '11' and retensialasan= '" & Trim(rsItem!pesan) & "'")
277                               Else
278                                   !Tgl11 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '11' and retensialasan= '" & Trim(rsItem!pesan) & "'")
279                               End If
280                           ElseIf 12 = i Then
                                  '' 20200408DEV
281                               If Log.sentralisasiStars = 1 Then
282                                   !Tgl12 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '12' and retensialasan= '" & Trim(rsItem!pesan) & "'")
283                               Else
284                                   !Tgl12 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '12' and retensialasan= '" & Trim(rsItem!pesan) & "'")
285                               End If
286                           ElseIf 13 = i Then
                                  '' 20200408DEV
287                               If Log.sentralisasiStars = 1 Then
288                                   !Tgl13 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '13' and retensialasan= '" & Trim(rsItem!pesan) & "'")
289                               Else
290                                   !Tgl13 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '13' and retensialasan= '" & Trim(rsItem!pesan) & "'")
291                               End If
                                  
292                           ElseIf 14 = i Then
                                  '' 20200408DEV
293                               If Log.sentralisasiStars = 1 Then
294                                   !Tgl14 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '14' and retensialasan= '" & Trim(rsItem!pesan) & "'")
295                               Else
296                                   !Tgl14 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '14' and retensialasan= '" & Trim(rsItem!pesan) & "'")
297                               End If
298                           ElseIf 15 = i Then
                                  '' 20200408DEV
299                               If Log.sentralisasiStars = 1 Then
300                                   !Tgl15 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '15' and retensialasan= '" & Trim(rsItem!pesan) & "'")
301                               Else
302                                   !Tgl15 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '15' and retensialasan= '" & Trim(rsItem!pesan) & "'")
303                               End If
304                           ElseIf 16 = i Then
                                  '' 20200408DEV
305                               If Log.sentralisasiStars = 1 Then
306                                   !Tgl16 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '16' and retensialasan= '" & Trim(rsItem!pesan) & "'")
307                               Else
308                                   !Tgl16 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '16' and retensialasan= '" & Trim(rsItem!pesan) & "'")
309                               End If
310                           ElseIf 17 = i Then
                                  '' 20200408DEV
311                               If Log.sentralisasiStars = 1 Then
312                                   !Tgl17 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '17' and retensialasan= '" & Trim(rsItem!pesan) & "'")
313                               Else
314                                   !Tgl17 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '17' and retensialasan= '" & Trim(rsItem!pesan) & "'")
315                               End If
316                           ElseIf 18 = i Then
                                  '' 20200408DEV
317                               If Log.sentralisasiStars = 1 Then
318                                   !Tgl18 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '18' and retensialasan= '" & Trim(rsItem!pesan) & "'")
319                               Else
320                                   !Tgl18 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '18' and retensialasan= '" & Trim(rsItem!pesan) & "'")
321                               End If
322                           ElseIf 19 = i Then
                                  '' 20200408DEV
323                               If Log.sentralisasiStars = 1 Then
324                                   !Tgl19 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '19' and retensialasan= '" & Trim(rsItem!pesan) & "'")
325                               Else
326                                   !Tgl19 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '19' and retensialasan= '" & Trim(rsItem!pesan) & "'")
327                               End If
328                           ElseIf 20 = i Then
                                  '' 20200408DEV
329                               If Log.sentralisasiStars = 1 Then
330                                   !Tgl20 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '20' and retensialasan= '" & Trim(rsItem!pesan) & "'")
331                               Else
332                                   !Tgl20 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '20' and retensialasan= '" & Trim(rsItem!pesan) & "'")
333                               End If
334                           ElseIf 21 = i Then
                                  '' 20200408DEV
335                               If Log.sentralisasiStars = 1 Then
336                                   !Tgl21 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '21' and retensialasan= '" & Trim(rsItem!pesan) & "'")
337                               Else
338                                   !Tgl21 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '21' and retensialasan= '" & Trim(rsItem!pesan) & "'")
339                               End If
340                           ElseIf 22 = i Then
                                  '' 20200408DEV
341                               If Log.sentralisasiStars = 1 Then
342                                   !Tgl22 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '22' and retensialasan= '" & Trim(rsItem!pesan) & "'")
343                               Else
344                                   !Tgl22 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '22' and retensialasan= '" & Trim(rsItem!pesan) & "'")
345                               End If
346                           ElseIf 23 = i Then
                                  '' 20200408DEV
347                               If Log.sentralisasiStars = 1 Then
348                                   !Tgl23 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '23' and retensialasan= '" & Trim(rsItem!pesan) & "'")
349                               Else
350                                   !Tgl23 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '23' and retensialasan= '" & Trim(rsItem!pesan) & "'")
351                               End If
352                           ElseIf 24 = i Then
                                  '' 20200408DEV
353                               If Log.sentralisasiStars = 1 Then
354                                   !Tgl24 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '24' and retensialasan= '" & Trim(rsItem!pesan) & "'")
355                               Else
356                                   !Tgl24 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '24' and retensialasan= '" & Trim(rsItem!pesan) & "'")
357                               End If
358                           ElseIf 25 = i Then
                                  '' 20200408DEV
359                               If Log.sentralisasiStars = 1 Then
360                                   !Tgl25 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '25' and retensialasan= '" & Trim(rsItem!pesan) & "'")
361                               Else
362                                   !Tgl25 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '25' and retensialasan= '" & Trim(rsItem!pesan) & "'")
363                               End If
364                           ElseIf 26 = i Then
                                  '' 20200408DEV
365                               If Log.sentralisasiStars = 1 Then
366                                   !Tgl26 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '26' and retensialasan= '" & Trim(rsItem!pesan) & "'")
367                               Else
368                                   !Tgl26 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '26' and retensialasan= '" & Trim(rsItem!pesan) & "'")
369                               End If
370                           ElseIf 27 = i Then
                                  '' 20200408DEV
371                               If Log.sentralisasiStars = 1 Then
372                                   !Tgl27 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '27' and retensialasan= '" & Trim(rsItem!pesan) & "'")
373                               Else
374                                   !Tgl27 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '27' and retensialasan= '" & Trim(rsItem!pesan) & "'")
375                               End If
376                           ElseIf 28 = i Then
                                  '' 20200408DEV
377                               If Log.sentralisasiStars = 1 Then
378                                   !Tgl28 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '28' and retensialasan= '" & Trim(rsItem!pesan) & "'")
379                               Else
380                                   !Tgl28 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '28' and retensialasan= '" & Trim(rsItem!pesan) & "'")
381                               End If
382                           ElseIf 29 = i Then
                                  '' 20200408DEV
383                               If Log.sentralisasiStars = 1 Then
384                                   !Tgl29 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '29' and retensialasan= '" & Trim(rsItem!pesan) & "'")
385                               Else
386                                   !Tgl29 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '29' and retensialasan= '" & Trim(rsItem!pesan) & "'")
387                               End If
388                           ElseIf 30 = i Then
                                  '' 20200408DEV
389                               If Log.sentralisasiStars = 1 Then
390                                   !Tgl30 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '30' and retensialasan= '" & Trim(rsItem!pesan) & "'")
391                               Else
392                                   !Tgl30 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '30' and retensialasan= '" & Trim(rsItem!pesan) & "'")
393                               End If
394                           ElseIf 31 = i Then
                                  '' 20200408DEV
395                               If Log.sentralisasiStars = 1 Then
396                                   !Tgl31 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [" & dbReplikasiHO & "]..ProspectFollow WITH (NOLOCK) WHERE kodecabang = '" & Log.kodeCabang & "' AND year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '31' and retensialasan= '" & Trim(rsItem!pesan) & "'")
397                               Else
398                                   !Tgl31 = objO.FetchDataSPK_Double("SELECT count(kodepr) FROM [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow WITH (NOLOCK) WHERE year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '31' and retensialasan= '" & Trim(rsItem!pesan) & "'")
399                               End If
400                           End If
401                       Next
402                       !total = CInt(!Tgl1) + CInt(!Tgl2) + CInt(!Tgl3) + CInt(!Tgl4) + CInt(!Tgl5) + CInt(!Tgl6) + CInt(!Tgl7) + CInt(!Tgl8) + CInt(!Tgl9) + CInt(!Tgl10) + CInt(!Tgl11) + CInt(!Tgl12) + CInt(!Tgl13) + CInt(!Tgl14) + CInt(!Tgl15) + CInt(!Tgl16) + CInt(!Tgl17) + CInt(!Tgl18) + CInt(!Tgl19) + CInt(!Tgl20) + CInt(!Tgl21) + CInt(!Tgl22) + CInt(!Tgl23) + CInt(!Tgl24) + CInt(!Tgl25) + CInt(!Tgl26) + CInt(!Tgl27) + CInt(!Tgl28) + CInt(!Tgl29) + CInt(!Tgl30) + CInt(!Tgl31)
403                       targetReminder = targetReminder + !total
404                       .Update
405                   End With
406               End If
                  
407               rsItem.MoveNext
408           Loop
          
409       rsRecord.MoveFirst
          'Set gridMonitoring.DataSource = Nothing
410           Set gridMonitoring.DataSource = rsRecord
      '        Set partGrid.DataSource = rsRecord
      '        Call createHeaderGrid
              'gridMonitoring.Refresh
411       Exit Sub
exo:
412       MsgBox err.desc
End Sub

Private Sub cmdReminderClose_Click()
    Frame1.Visible = False
End Sub

Private Sub Form_Load()
    Call FillAllCombo
    Bln = Month(isnow) 'Month(Now)
    Thn = year(isnow)  'Year(Now)
    nBulan.Caption = Bulan(Bln)
    nTahun.Caption = "Periode  :  " & Thn
    isProses = False
    Call createTableRecord
End Sub

Private Sub FillAllCombo()
   Bulan(0) = "TOTAL"
   Bulan(1) = "JANUARI"
   Bulan(2) = "FEBRUARI"
   Bulan(3) = "MARET"
   Bulan(4) = "APRIL"
   Bulan(5) = "MEI"
   Bulan(6) = "JUNI"
   Bulan(7) = "JULI"
   Bulan(8) = "AGUSTUS"
   Bulan(9) = "SEPTEMBER"
   Bulan(10) = "OKTOBER"
   Bulan(11) = "NOVEMBER"
   Bulan(12) = "DESEMBER"
End Sub
  

Private Sub gridMonitoring_BeforeRowColChange(Cancel As Integer)
    'MsgBox gridMonitoring.CurrentCellVisible
End Sub

Private Sub gridMonitoring_ButtonClick(ByVal ColIndex As Integer)
    Call doClick(ColIndex)
End Sub
 

Private Sub doClick(intCol As Integer)
      'MsgBox "Col " & rsRecord.Fields(intCol).value
1         If rsRecord.Fields(intCol).value > 0 Then
2             Frame1.Visible = True
              
              Dim keterangan As String
              
3             If UCase(rsRecord.Fields(0).value) = UCase("O/S Reminder Sales") Then
4                 keterangan = ""
5             ElseIf UCase(rsRecord.Fields(0).value) = UCase("O/S Booking Sales") Then
6                 keterangan = " and c.boc = 'BOC'"
7             Else
8                 keterangan = " and c.retensialasan= '" & Trim(UCase(rsRecord.Fields(0).value)) & "'"
9             End If
              
              '20180320 ' REV ' PENAMBAHAN KETERANGAN
              '' 20200408DEV
10            If Log.sentralisasiStars = 1 Then
11                sSQL = "select a.NoRangka,a.Type,b.Nama,b.Alamat,b.Kota,b.KodePos,b.Telp,b.GSM,b.CDMA,'0' as warna,CONVERT(varchar(20),c.tglfollow,103) tglFollowSales ,c.keterangan,d.NamaSales"
12                sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
13                sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on a.KodePr = c.KodePr and a.KodeProspect = c.KodeProspect and a.KodePlan = c.KodePlan and c.kodecabang = '" & Log.kodeCabang & "'"
14                sSQL = sSQL & " left join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.kodeCabang & "'"
15                sSQL = sSQL & " LEFT JOIN [SPK" & Right(Log.LogSPK, 2) & "]..ArSales d on d.KodeSales = b.SalesPr"
16            Else
17                sSQL = "select a.NoRangka,a.Type,b.Nama,b.Alamat,b.Kota,b.KodePos,b.Telp,b.GSM,b.CDMA,'0' as warna,CONVERT(varchar(20),c.tglfollow,103) tglFollowSales ,c.keterangan,d.NamaSales"
18                sSQL = sSQL & " from [STARS" & Right(Log.LogSPK, 2) & "]..PlanTrn a"
19                sSQL = sSQL & " inner join [STARS" & Right(Log.LogSPK, 2) & "]..ProspectFollow c on a.KodePr = c.KodePr and a.KodeProspect = c.KodeProspect and a.KodePlan = c.KodePlan"
20                sSQL = sSQL & " left join [STARS" & Right(Log.LogSPK, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
21                sSQL = sSQL & " LEFT JOIN [SPK" & Right(Log.LogSPK, 2) & "]..ArSales d on d.KodeSales = b.SalesPr"
22            End If
23            If UCase(rsRecord.Fields(0).value) = UCase("O/S Booking Sales") Then
24                sSQL = sSQL & " where " & IIf(Log.sentralisasiStars = 1, "a.kodecabang = '" & Log.kodeCabang & "' and ", "") & " a.TypePlan = '3' and year(tgl_boc) = '" & Thn & "' and MONTH(tgl_boc) = '" & Bln & "' and DAY(tgl_boc) = '" & intCol & "'"
25            Else
26                sSQL = sSQL & " where " & IIf(Log.sentralisasiStars = 1, "a.kodecabang = '" & Log.kodeCabang & "' and ", "") & " a.TypePlan = '3' and year(tgl_entry_retensi) = '" & Thn & "' and MONTH(tgl_entry_retensi) = '" & Bln & "' and DAY(tgl_entry_retensi) = '" & intCol & "'"
27            End If
              
28            sSQL = sSQL & keterangan
29            Set rsReminder = objO.FetchDataSPK(sSQL)
30            If UCase(rsRecord.Fields(0).value) = UCase("Booking Sales Fix") Then
31                sSQL = "SELECT a.NoRangka,b.FMDL as type,b.CUSTOMER as nama,b.CUST_ADRES as alamat,b.KOTA ,b.TELPON as telp,b.GSM,b.CDMA,'01/01/1900' as tglFollowSales,'' as NamaSales,'0' as warna"
32                sSQL = sSQL & " FROM listReminderBoard a"
33                sSQL = sSQL & " INNER JOIN BOOK b on a.NoBooking = b.NOFAK and a.NoPolisi = b.NO_POLISI"
34                sSQL = sSQL & " WHERE year(b.date) = '" & Thn & "' and month(b.date) = '" & Bln & "' and day(b.date) = '" & intCol & "'"
35                Set rsReminder = objO.FetchData(sSQL)
36                gridReminder.Columns(8).Visible = False
37                gridReminder.Columns(9).Visible = False
38                gridReminder.Columns(10).Visible = False
39            Else
40                gridReminder.Columns(10).Visible = False
41            End If
              
              Dim noWO As String
42            Call createTableReminder
43            If Not rsReminder.EOF Then
44                rsReminder.MoveFirst
                  '' mencari unit sudah masuk reminder
45                Do While Not rsReminder.EOF
46                    With rsReminderCall
                          '20180320 ' REV ' UNTUK KM 1K NOWO DIAMBIL DARI T MRS DENGAN NOSPK SBG WO
47                        If UCase(rsReminder!keterangan) = UCase("Reminder service 1000 Km") Then
48                            sSQL = "SELECT nospk as NoWo FROM " & Log.LogSPK & "..T_MRS WITH(NOLOCK) WHERE NoRangka = '" & Nul(rsReminder.Fields(0).value) & "'"
49                        Else
50                            sSQL = "SELECT a.NoWo  "
51                            sSQL = sSQL & " FROM M_ESTIMASI A WITH(NOLOCK) Inner Join"
52                            sSQL = sSQL & " (SELECT G.NORANGKA, (SELECT MAX(SBENEXT) FROM m_estimasi WITH(NOLOCK) WHERE NORANGKA=G.NORANGKA AND ISNULL(TGLSVCNOW,'') <> '') S, MAX((CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END)) AS JT_TGL"
53                            sSQL = sSQL & " FROM M_ESTIMASI G WITH(NOLOCK) Where (CASE WHEN G.JT_Tgl > G.JT_Bln then G.JT_Bln WHEN G.JT_Tgl < G.JT_Bln then G.JT_Tgl WHEN G.JT_Tgl = G.JT_Bln then G.JT_Tgl END) Is Not Null GROUP BY G.NORANGKA) AS Z ON A.NORANGKA = Z.NORANGKA AND A.SBENEXT = Z.S"
54                            sSQL = sSQL & " INNER JOIN History b WITH(NOLOCK) ON a.NoWo = b.No_Wo AND A.NORANGKA = B.NO_CHASSIS AND B.FS_LOG_DEL = ''"
55                            sSQL = sSQL & " INNER JOIN Custdat c WITH(NOLOCK) On a.NoRangka = c.No_Chassis"
56                            sSQL = sSQL & " Where a.NoRangka = '" & Nul(rsReminder.Fields(0).value) & "'"
57                        End If
                          
                          
58                        noWO = ""
59                        Set rsTemp = objO.FetchData(sSQL)
60                        If Not rsTemp.EOF Then
61                            noWO = Nul(rsTemp!noWO)
62                        End If
63                        .AddNew
      '                    .Fields.Append "NoRangka", adVarChar, 100
      '                    .Fields.Append "Type", adVarChar, 100
      '                    .Fields.Append "Nama", adVarChar, 100
      '                    .Fields.Append "Alamat", adVarChar, 200
      '                    .Fields.Append "Kota", adVarChar, 100
      '                    .Fields.Append "KodePos", adVarChar, 100
      '                    .Fields.Append "Telp", adVarChar, 100
      '                    .Fields.Append "GSM", adVarChar, 100
      '                    .Fields.Append "CDMA", adVarChar, 100
      '                    .Fields.Append "warna", adVarChar, 100
64                        !noRangka = Nul(rsReminder!noRangka)
65                        !Type = Nul(rsReminder!Type)
66                        !Nama = Nul(rsReminder!Nama)
67                        !Alamat = Nul(rsReminder!Alamat)
68                        !Kota = Nul(rsReminder!Kota)
69                        !telp = Nul(rsReminder!telp)
70                        !GSM = Nul(rsReminder!GSM)
71                        !CDMA = Nul(rsReminder!CDMA)
72                        !tglReminderSales = TglDDMMMMYYYY(rsReminder!tglFollowSales)
73                        !namasales = Nul(rsReminder!namasales)
74                        Set rsTemp = objO.FetchData("SELECT * FROM [listReminderBoard] WHERE norangka = '" & Nul(rsReminder.Fields(0).value) & "' and nowo = '" & Trim(noWO) & "'")
75                        If Not rsTemp.EOF Then
76                            !warna = "1"
77                        Else
78                            !warna = "0"
79                        End If
                          
80                        .Update
81                    End With
                      
82                    rsReminder.MoveNext
83                Loop
84                rsReminder.MoveFirst
85            End If
86            Set gridReminder.DataSource = rsReminderCall
87            gridReminder.Refresh
88        End If
End Sub

Private Sub gridMonitoring_FetchCellStyle(ByVal Condition As Integer, ByVal Split As Integer, Bookmark As Variant, ByVal col As Integer, ByVal CellStyle As TrueOleDBGrid70.StyleDisp)
'    Select Case col
'    Case 4
'        Dim Col1 As Long, Col2 As Long
'        Col1 = IIf(gridMonitoring.Columns(4).CellText(Bookmark) = "", 0, gridMonitoring.Columns(4).CellText(Bookmark))
'        'Col2 = CLng(gridMonitoring.Columns(2).CellText(Bookmark))
'        If Col1 > 0 Then
'            CellStyle.BackColor = vbBlue
'            CellStyle = "BiruHitam"
'        End If
''    Case 7
''        Dim S As String
''        S = TDBGrid1.Columns(7).CellText(Bookmark)
''        If S = Text1.text Then CellStyle.Font.Bold = True
''    Case Else
''        Debug.Print "FetchCellStyle not handled: " & Col
'End Select


End Sub

Private Sub gridMonitoring_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsRecord.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    RowStyle = "KuningHitam"
    If rs("item") <> "" Then
        If UCase(rs("item")) = UCase("nomor tidak valid") Or UCase(rs("item")) = UCase("Nomor tidak aktif") Then
            RowStyle = "MerahPutih"
        ElseIf UCase(rs("item")) = UCase("O/S Reminder Sales") Then
            RowStyle = "BiruPutih"
        ElseIf UCase(rs("item")) = UCase("O/S Booking Sales") Or UCase(rs("item")) = UCase("Booking Sales Fix") Then
            RowStyle = "HijauPutih"
'        ElseIf UCase(rs("WARNA")) = "HIJAU" Then
'            RowStyle = "HIJAU"
'        Else
'            RowStyle = "PUTIH"
        End If
    End If
End Sub

Private Sub gridReminder_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsReminderCall.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    RowStyle = "Normal"
    If rs("warna") = "1" Then
        RowStyle = "MerahPutih"
    End If
End Sub



Private Sub thnPlus_Click()
    Thn = Thn + 1
    nTahun.Caption = "Periode  :  " & Thn
    'Call cmdProcess_Click
End Sub

Private Sub thnMinus_Click()
    Thn = Thn - 1
    If Thn < 2009 Then
        Thn = 2009
    Else
        nTahun.Caption = "Periode  :  " & Thn
        'Call cmdProcess_Click
    End If
End Sub

Private Sub blnMinus_Click()
    Bln = Bln - 1
    If Bln = 0 Then Bln = 12
    nBulan.Caption = Bulan(Bln)
    'Call cmdProcess_Click
End Sub

Private Sub blnPlus_Click()
    Bln = Bln + 1
    If Bln = 13 Then Bln = 1
    nBulan.Caption = Bulan(Bln)
    'Call cmdProcess_Click
End Sub

