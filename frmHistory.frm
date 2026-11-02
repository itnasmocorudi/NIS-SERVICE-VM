VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmHistory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Service History"
   ClientHeight    =   6900
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   11580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   11580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Export ke &Excel"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6480
      Width           =   2175
   End
   Begin VB.CommandButton cmdExit 
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
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6060
      Left            =   450
      TabIndex        =   3
      Top             =   345
      Width           =   11115
      Begin VB.OptionButton opHist 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sublet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6360
         TabIndex        =   11
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton opHist 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Part"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton opHist 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Service"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   6360
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin TrueDBGrid70.TDBGrid TData1 
         Height          =   4560
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   8043
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tanggal"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No Invoice"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "No. WO"
         Columns(3).DataField=   ""
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "KM"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "TSA"
         Columns(5).DataField=   ""
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "MCH"
         Columns(6).DataField=   ""
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Pekerjaan"
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
         Columns(12).Caption=   "Permintaan"
         Columns(12).DataField=   ""
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Keluhan"
         Columns(13).DataField=   ""
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Dealer Service"
         Columns(14).DataField=   ""
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   15
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=15"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1931"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1852"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1852"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1773"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=1244"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=1164"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=1270"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=1191"
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
         Splits(0)._ColumnProps(57)=   "Column(14).Width=5292"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=5212"
         Splits(0)._ColumnProps(60)=   "Column(14).WrapText=1"
         Splits(0)._ColumnProps(61)=   "Column(14).Order=15"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bgcolor=&H0&,.fgcolor=&HFFFFFF&"
         _StyleDefs(9)   =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(10)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=25,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=68,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=26,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=27,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=28,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=60,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=59,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=61,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=62,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=67,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=69,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=70,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=54,.parent=25"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=26"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=27"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=59"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=20,.parent=25"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=17,.parent=26"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=18,.parent=27"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=19,.parent=59"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=25"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=26"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=27"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=59"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=16,.parent=25"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=13,.parent=26"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=14,.parent=27"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=15,.parent=59"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=58,.parent=25"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=26"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=27"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=59"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=66,.parent=25"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=26"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=27"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=59"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=74,.parent=25"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=26"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=27"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=59"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=24,.parent=25"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=21,.parent=26"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=22,.parent=27"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=23,.parent=59"
         _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=90,.parent=25"
         _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=87,.parent=26"
         _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=88,.parent=27"
         _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=89,.parent=59"
         _StyleDefs(68)  =   "Splits(0).Columns(9).Style:id=46,.parent=25"
         _StyleDefs(69)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=26"
         _StyleDefs(70)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=27"
         _StyleDefs(71)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=59"
         _StyleDefs(72)  =   "Splits(0).Columns(10).Style:id=94,.parent=25"
         _StyleDefs(73)  =   "Splits(0).Columns(10).HeadingStyle:id=91,.parent=26"
         _StyleDefs(74)  =   "Splits(0).Columns(10).FooterStyle:id=92,.parent=27"
         _StyleDefs(75)  =   "Splits(0).Columns(10).EditorStyle:id=93,.parent=59"
         _StyleDefs(76)  =   "Splits(0).Columns(11).Style:id=86,.parent=25"
         _StyleDefs(77)  =   "Splits(0).Columns(11).HeadingStyle:id=83,.parent=26"
         _StyleDefs(78)  =   "Splits(0).Columns(11).FooterStyle:id=84,.parent=27"
         _StyleDefs(79)  =   "Splits(0).Columns(11).EditorStyle:id=85,.parent=59"
         _StyleDefs(80)  =   "Splits(0).Columns(12).Style:id=78,.parent=25"
         _StyleDefs(81)  =   "Splits(0).Columns(12).HeadingStyle:id=75,.parent=26"
         _StyleDefs(82)  =   "Splits(0).Columns(12).FooterStyle:id=76,.parent=27"
         _StyleDefs(83)  =   "Splits(0).Columns(12).EditorStyle:id=77,.parent=59"
         _StyleDefs(84)  =   "Splits(0).Columns(13).Style:id=32,.parent=25"
         _StyleDefs(85)  =   "Splits(0).Columns(13).HeadingStyle:id=29,.parent=26"
         _StyleDefs(86)  =   "Splits(0).Columns(13).FooterStyle:id=30,.parent=27"
         _StyleDefs(87)  =   "Splits(0).Columns(13).EditorStyle:id=31,.parent=59"
         _StyleDefs(88)  =   "Splits(0).Columns(14).Style:id=82,.parent=25,.wraptext=-1"
         _StyleDefs(89)  =   "Splits(0).Columns(14).HeadingStyle:id=79,.parent=26"
         _StyleDefs(90)  =   "Splits(0).Columns(14).FooterStyle:id=80,.parent=27"
         _StyleDefs(91)  =   "Splits(0).Columns(14).EditorStyle:id=81,.parent=59"
         _StyleDefs(92)  =   "Named:id=33:Normal"
         _StyleDefs(93)  =   ":id=33,.parent=0"
         _StyleDefs(94)  =   "Named:id=34:Heading"
         _StyleDefs(95)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(96)  =   ":id=34,.wraptext=-1"
         _StyleDefs(97)  =   "Named:id=35:Footing"
         _StyleDefs(98)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(99)  =   "Named:id=36:Selected"
         _StyleDefs(100) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(101) =   "Named:id=37:Caption"
         _StyleDefs(102) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(103) =   "Named:id=38:HighlightRow"
         _StyleDefs(104) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(105) =   "Named:id=39:EvenRow"
         _StyleDefs(106) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(107) =   "Named:id=40:OddRow"
         _StyleDefs(108) =   ":id=40,.parent=33"
         _StyleDefs(109) =   "Named:id=41:RecordSelector"
         _StyleDefs(110) =   ":id=41,.parent=34"
         _StyleDefs(111) =   "Named:id=42:FilterBar"
         _StyleDefs(112) =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   2055
         TabIndex        =   1
         Top             =   540
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   60
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   2055
         TabIndex        =   0
         Top             =   240
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   2055
         TabIndex        =   7
         Top             =   840
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1290
         TabIndex        =   8
         Top             =   885
         Width           =   585
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1095
         TabIndex        =   5
         Top             =   285
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1185
         TabIndex        =   4
         Top             =   585
         Width           =   690
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   30
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdHistPrice 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   480
      TabIndex        =   6
      Top             =   6480
      Visible         =   0   'False
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&History Price"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   6855
      Left            =   0
      Picture         =   "frmHistory.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim History As New XArrayDB
Dim baris As Integer

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub HeaderCol(idx As Integer)
    Select Case idx
        Case 0
            ' SERVICE
            TData1.Columns(0).Caption = "Tanggal"
            TData1.Columns(0).Width = 1000
            TData1.Columns(1).Caption = "No Invoice"
            TData1.Columns(1).Width = 1000
            TData1.Columns(2).Caption = "Total Invoice"
            TData1.Columns(2).Width = 1800
            'TData1.Columns(2).NumberFormat = "#,##0.00"
            TData1.Columns(3).Caption = "No WO"
            TData1.Columns(3).Width = 1000
            TData1.Columns(4).Caption = "KM"
            TData1.Columns(4).Width = 1000
            'TData1.Columns(4).NumberFormat = "#,##0"
            TData1.Columns(5).Caption = "TSA"
            TData1.Columns(5).Width = 600
            TData1.Columns(6).Caption = "MCH"
            TData1.Columns(6).Width = 600
            TData1.Columns(7).Caption = "Job Order"
            TData1.Columns(7).Width = 4000
            TData1.Columns(8).Caption = "Disc"
            TData1.Columns(8).Width = 600
            TData1.Columns(9).Caption = "Pesan Khusus"
            TData1.Columns(9).Width = 2000
            TData1.Columns(10).Caption = "Saran"
            TData1.Columns(10).Width = 2000
            TData1.Columns(11).Caption = "Catatan"
            TData1.Columns(11).Width = 2000
            TData1.Columns(12).Caption = "Permintaan"
            TData1.Columns(12).Width = 2000
            TData1.Columns(13).Caption = "Keluhan"
            TData1.Columns(13).Width = 2000
            TData1.Columns(14).Caption = "Dealer Service"
            TData1.Columns(14).Width = 3000
        Case 1
            ' PART
            TData1.Columns(0).Caption = "Tanggal"
            TData1.Columns(0).Width = 1000
            TData1.Columns(1).Caption = "No Invoice"
            TData1.Columns(1).Width = 1000
            TData1.Columns(2).Caption = "No WO"
            TData1.Columns(2).Width = 1000
            TData1.Columns(3).Caption = "Total Part"
            TData1.Columns(3).Width = 1000
            'TData1.Columns(3).NumberFormat = "#,##0.00"
            TData1.Columns(4).Caption = "Kode Part"
            TData1.Columns(4).Width = 1500
            TData1.Columns(5).Caption = "Penggantian Part"
            TData1.Columns(5).Width = 3000
            TData1.Columns(6).Caption = "QTY"
            TData1.Columns(6).Width = 600
            'TData1.Columns(6).NumberFormat = "#,##0.00"
            TData1.Columns(7).Caption = "Harga"
            TData1.Columns(7).Width = 1000
            'TData1.Columns(7).NumberFormat = "#,##0.00"
            
        Case 2
            ' SUBLET
            TData1.Columns(0).Caption = "Tanggal"
            TData1.Columns(0).Width = 1000
            TData1.Columns(1).Caption = "No Invoice"
            TData1.Columns(1).Width = 1000
            TData1.Columns(2).Caption = "No WO"
            TData1.Columns(2).Width = 1000
            TData1.Columns(3).Caption = "Total Sublet"
            TData1.Columns(3).Width = 1000
            TData1.Columns(4).Caption = "Kode Sublet"
            TData1.Columns(4).Width = 1500
            TData1.Columns(5).Caption = "Order Sublet"
            TData1.Columns(5).Width = 3000
            TData1.Columns(6).Caption = "QTY"
            TData1.Columns(6).Width = 600
            'TData1.Columns(6).NumberFormat = "#,##0.00"
            TData1.Columns(7).Caption = "Harga"
            TData1.Columns(7).Width = 1000
            'TData1.Columns(7).NumberFormat = "#,##0.00"
    End Select
End Sub

Private Sub BeriItem(noChas As String, nKol As Integer)
Dim rsHlp As ADODB.Recordset
Dim rsSearch As ADODB.Recordset

    baris = 0
    Kolom = nKol
    Set rsSearch = objO.FetchData("select REP1, REP2, REP3, REP4, REP5, REP6, REP7 from HISTORY WITH (NOLOCK) where NO_CHASSIS = '" & noChas & "' AND FS_LOG_DEL = ''")
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If Trim(rsSearch("REP1")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP2")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP3")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP4")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP5")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP6")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP7")) <> "" Then
                baris = baris + 1
            End If
            rsSearch.MoveNext
        Loop
    End If
    
    baris = baris + CDbl(objO.FetchData_Count("SELECT NO_CHASSIS FROM JOBADD2 WITH (NOLOCK) WHERE NO_CHASSIS = '" & noChas & "'"))
    baris = baris + CDbl(objO.FetchData_Count("SELECT NO_CHASSIS FROM JOBSUBLET WITH (NOLOCK) WHERE NO_CHASSIS = '" & noChas & "'"))

    History.Clear
    History.ReDim 0, baris + 20, 0, Kolom
    i = 0
    sSQL = "select A.DATE,A.KM,A.FRONT,A.MECH AS MEKANIK1,A.MECH AS MEKANIK2,A.MECH AS MEKANIK3,A.MECH AS MEKANIK4,A.MECH AS MEKANIK5,A.MECH AS MEKANIK6,A.MECH AS MEKANIK7, " _
         & " A.CUSTOMER , A.REP1, A.REP2, A.REP3, A.REP4, A.REP5, A.REP6, A.REP7, D.DISC1, D.DISC2, D.DISC3, D.DISC4, D.DISC5, D.DISC6, D.DISC7, A.NO_WO, A.NOINV, A.REMARK1, A.SUGGEST, A.NO_CHASSIS, A.TOTAL, A.FS_NM_DEALER, " _
         & " B.REQ1, B.REQ2, B.REQ3, B.REQ4, B.REQ5, C.CMPL1, C.CMPL2, C.CMPL3, C.CMPL4, C.CMPL5, catat from HISTORY A WITH (NOLOCK) " _
         & " LEFT JOIN JOBDTA B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS AND A.NOINV = B.NOINV " _
         & " LEFT JOIN JOBDTL C WITH (NOLOCK) ON A.NO_WO = C.NO_WO LEFT JOIN T_INVOICE D WITH (NOLOCK) ON A.NO_WO = D.NO_WO AND A.NO_CHASSIS = D.NO_CHASSIS AND A.NOINV = D.NOINV " _
         & " where A.NO_CHASSIS ='" & noChas & "' AND A.FS_LOG_DEL = '' order by A.DATE DESC, A.NOINV DESC, A.KM DESC"
    Set rsSearch = objO.FetchData(sSQL)
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            History(i, 0) = Trim(rsSearch("DATE"))
            History(i, 1) = Trim(rsSearch("NOINV"))
            History(i, 2) = Format(CDbl(rsSearch("TOTAL")), "#,##0.00")
            History(i, 3) = Trim(rsSearch("NO_WO"))
            History(i, 4) = Format(Trim(rsSearch("KM")), "#,##0")
            History(i, 5) = Trim(rsSearch("FRONT"))
            History(i, 6) = Trim(rsSearch("MEKANIK1"))
            History(i, 7) = Trim(rsSearch("REP1"))
            History(i, 8) = Trim(rsSearch("DISC1")) & "%"
            History(i, 9) = Trim(rsSearch("REMARK1"))
            History(i, 10) = Trim(rsSearch("SUGGEST"))
            History(i, 11) = Trim(rsSearch("CATAT"))
            History(i, 12) = Trim(rsSearch("REQ1")) & "; " & Trim(rsSearch("REQ2")) & "; " & Trim(rsSearch("REQ3")) & "; " & Trim(rsSearch("REQ4")) & "; " & Trim(rsSearch("REQ5"))
            History(i, 13) = Trim(rsSearch("CMPL1")) & "; " & Trim(rsSearch("CMPL2")) & "; " & Trim(rsSearch("CMPL3")) & "; " & Trim(rsSearch("CMPL4")) & "; " & Trim(rsSearch("CMPL5"))
            History(i, 14) = Trim(rsSearch("FS_NM_DEALER"))
            
            If Trim(rsSearch("REP2")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK2"))
                History(i, 7) = Trim(rsSearch("REP2"))
                History(i, 8) = Trim(rsSearch("DISC2")) & "%"
            End If
            If Trim(rsSearch("REP3")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK3"))
                History(i, 7) = Trim(rsSearch("REP3"))
                History(i, 8) = Trim(rsSearch("DISC3")) & "%"
            End If
            If Trim(rsSearch("REP4")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK4"))
                History(i, 7) = Trim(rsSearch("REP4"))
                History(i, 8) = Trim(rsSearch("DISC4")) & "%"
            End If
            If Trim(rsSearch("REP5")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK5"))
                History(i, 7) = Trim(rsSearch("REP5"))
                History(i, 8) = Trim(rsSearch("DISC5")) & "%"
            End If
            If Trim(rsSearch("REP6")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK6"))
                History(i, 7) = Trim(rsSearch("REP6"))
                History(i, 8) = Trim(rsSearch("DISC6")) & "%"
            End If
            If Trim(rsSearch("REP7")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK7"))
                History(i, 7) = Trim(rsSearch("REP7"))
                History(i, 8) = Trim(rsSearch("DISC7")) & "%"
            End If
            Set rsHlp = objO.FetchData("SELECT MEKANIK,REP,DISC FROM JOBADD2 WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsSearch("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsSearch("NO_WO")) & "'")
            If Not rsHlp.EOF Then
                rsHlp.MoveFirst
                Do While Not rsHlp.EOF
                    i = i + 1
                    History(i, 6) = Trim(rsHlp("MEKANIK"))
                    History(i, 7) = Trim(rsHlp("REP"))
                    History(i, 8) = Trim(rsHlp("DISC")) & "%"
                    rsHlp.MoveNext
                Loop
            End If
            Set rsHlp = objO.FetchData("SELECT NAMABARANG FROM JOBSUBLET WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsSearch("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsSearch("NO_WO")) & "'")
            If Not rsHlp.EOF Then
                rsHlp.MoveFirst
                Do While Not rsHlp.EOF
                    i = i + 1
                    History(i, 6) = "SUB"
                    History(i, 7) = Trim(rsHlp("NAMABARANG"))
                    History(i, 8) = ""
                    rsHlp.MoveNext
                Loop
            End If
            rsSearch.MoveNext
            i = i + 1
        Loop
    End If
    TData1.Array = History
    TData1.ReBind
End Sub

Private Sub BeriItemPart(noChas As String, nKol As Integer)
    Dim OldInv As String
    Dim rsSearch As ADODB.Recordset
    
    baris = 0
    Kolom = nKol
    sSQL = "select count(B.NoInv) as JML" _
         & " from history as A WITH (NOLOCK) inner join t_part as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and B.FS_LOG_DEL = ''" _
         & " where NO_CHASSIS = '" & noChas & "' and B.NOINV <> '' and A.FS_LOG_DEL = ''"
    baris = objO.FetchData_Integer(sSQL)
    History.Clear
    History.ReDim 0, baris, 0, Kolom
    i = 0
    OldInv = "LAMA"
    sSQL = "select [Date], A.NoInv, A.no_wo, PART_G+PART_NG+OIL_G+OIL_NG AS  PART, FRONT, KODEITEM, Keterangan, QTY, HARGA, OLI" _
         & " from history as A WITH (NOLOCK) inner join t_part as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and B.FS_LOG_DEL = ''" _
         & " where NO_CHASSIS = '" & noChas & "' and B.NOINV <> '' and A.FS_LOG_DEL = '' order by [DATE] DESC, KODEITEM"
    Set rsSearch = objO.FetchData(sSQL)
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If OldInv <> Trim(rsSearch("NOINV")) Then
                History(i, 0) = Trim(rsSearch("DATE"))
                History(i, 1) = Trim(rsSearch("NOINV"))
                History(i, 2) = Trim(rsSearch("NO_WO"))
                History(i, 3) = Format(Trim(rsSearch("PART")), "#,##0.00")
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("Keterangan"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            Else
                History(i, 0) = ""
                History(i, 1) = ""
                History(i, 2) = ""
                History(i, 3) = ""
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("Keterangan"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            End If
            OldInv = Trim(rsSearch("NOINV"))
            rsSearch.MoveNext
            i = i + 1
        Loop
    End If
    TData1.Array = History
    TData1.ReBind
End Sub

Private Sub BeriItemSublet(noChas As String, nKol As Integer)
    Dim OldInv As String
    Dim rsSearch As ADODB.Recordset
    
    baris = 0
    Kolom = nKol
    baris = objO.FetchData_Integer("SELECT COUNT(*) AS JML FROM history A WITH (NOLOCK) INNER JOIN hist_sublet B WITH (NOLOCK) " _
            & "ON A.no_wo = B.no_wo AND A.FS_LOG_DEL = '' WHERE A.NO_CHASSIS = '" & noChas & "' AND B.noinv <> ''")
    
    History.Clear
    History.ReDim 0, baris, 0, Kolom
    i = 0
    
    OldInv = "LAMA"
    sSQL = "SELECT [DATE], A.NOINV, A.NO_WO, SBLT, KODEITEM, NAMABARANG, QTY, HARGA FROM history A WITH (NOLOCK) " _
         & "INNER JOIN hist_sublet B WITH (NOLOCK) ON A.no_wo = B.no_wo AND A.FS_LOG_DEL = '' WHERE A.NO_CHASSIS = '" & noChas & "' " _
         & "AND B.noinv <> '' order by [DATE] DESC, KODEITEM"
    Set rsSearch = objO.FetchData(sSQL)
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If OldInv <> Trim(rsSearch("NOINV")) Then
                History(i, 0) = Trim(rsSearch("DATE"))
                History(i, 1) = Trim(rsSearch("NOINV"))
                History(i, 2) = Trim(rsSearch("NO_WO"))
                History(i, 3) = Format(Trim(rsSearch("SBLT")), "#,##0.00")
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("NAMABARANG"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            Else
                History(i, 0) = ""
                History(i, 1) = ""
                History(i, 2) = ""
                History(i, 3) = ""
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("NAMABARANG"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            End If
            
            OldInv = Trim(rsSearch("NOINV"))
            rsSearch.MoveNext
            i = i + 1
        Loop
    End If
    TData1.Array = History
    TData1.ReBind
End Sub

Private Sub cmdExit_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
End Sub

Private Sub cmdExport_Click()
    On Error GoTo Err_Export
    Dim Er As New ErrInfo
    Dim MyPath As String
    Dim onVal As String
    MyPath = ""
    Dim f As New FormExplorer
    f.Show (vbModal)
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    onVal = MyPath & f.FileName & ".XLS"

    
     TData1.ExportToFile onVal, True
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            
            Beep
            Er.Show (vbModal)
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub cmdHistPrice_Click()
    If TData1.row > 0 Then
        frmHistoryPrice.text(0) = History(TData1.row, 1)
        frmHistoryPrice.Show vbModal
    End If
End Sub

Private Sub Form_Load()
    'Call hapusForm
    text(0) = dRangka
    TData1.EvenRowStyle.BackColor = &HC0FFC0
    TData1.OddRowStyle.BackColor = &HFFC0C0
    TData1.AlternatingRowStyle = True
    opHist(0).value = True
    If gTag = "BOOKING" Then
        Call BeriItemHO(Trim(frmBooking.text(41)), 14)
        'Call ambilHistoryHO(Trim(frmBooking.Text(41)))
    ElseIf gTag = "ORDER" Then
        Call BeriItemHO(Trim(frmOrder.text(15)), 14)
        'Call ambilHistoryHO(Trim(frmOrder.Text(15)))
    ElseIf gTag = "ORDER_local" Then
        Call BeriItemHO_Local(Trim(frmOrder.text(15)), 14)
        'Call ambilHistoryHO(Trim(frmOrder.Text(15)))
    ElseIf gTag = "BTLORDER" Then
        Call BeriItemHO(Trim(frmBtlOrd.text(15)), 14)
        'Call ambilHistoryHO(Trim(frmBtlOrd.Text(15)))
    ElseIf gTag = "BTLBOOKING" Then
        Call BeriItemHO(Trim(frmBtlBooking.text(15)), 14)
        'Call ambilHistoryHO(Trim(frmBtlBooking.Text(15)))
    ElseIf gTag = "STALL" Then
        Call BeriItemHO(Trim(frmStallProgress.text(0)), 14)
        'Call ambilHistoryHO(Trim(frmStallProgress.Text(0)))
    'ElseIf gTag = "DISSTALL" Then
        'Call BeriItemHO(Trim(frmStallDisplay.Text(0)), 14)
        'Call ambilHistoryHO(Trim(frmStallDisplay.Text(0)))
    ElseIf gTag = "3dfu" Then
        Call BeriItemHO(Trim(frm3DFU.text(4)), 14)
        'Call ambilHistoryHO(Trim(frm3DFU.Text(4)))
    End If
End Sub

Private Sub ambilHistoryHO(nChasis As String)
Dim rsHlp As ADODB.Recordset
Dim rsHlp2 As ADODB.Recordset
    
    On Error GoTo vbError
    If Log.LogHOConn = 1 Then
    If ConHO = True Then
        Set rsHlp = objO.FetchDataHO_NoError("SELECT * FROM HISTORYnew WITH (NOLOCK) WHERE NO_CHASSIS = '" & nChasis & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC")
        If Not rsHlp.EOF Then
            Do While Not rsHlp.EOF
                 
                Set rsHlp2 = objO.FetchData("SELECT * FROM HISTORY WITH (NOLOCK) WHERE NO_CHASSIS = '" & nChasis & "' AND [DATE] = '" & Format(rsHlp("DATE"), "MM/DD/YYYY") & "'")
                If rsHlp2.EOF Then
                    Call objO.InsertData("INSERT INTO HISTORY (CUSTOMER,CUST_ADRES,KOTA,NO_POLISI,NO_ENGINE,NO_CHASSIS,FRONT,MECH,KM,[DATE],PART_G, PART_NG,OIL_G,OIL_NG,MTRL,LABOR,BODY,SBLT,MET,TOTAL,JOB, " _
                                 & "REP1,REP2,REP3,REP4,REP5,REP6,REP7,MEKANIK1,MEKANIK2,MEKANIK3,MEKANIK4,MEKANIK5,MEKANIK6,MEKANIK7,MEKANIK8,KLPKERJA1,KLPKERJA2,KLPKERJA3,KLPKERJA4, " _
                                 & "KLPKERJA5,KLPKERJA6,KLPKERJA7,KLPKERJA8,NO_WO,NOINV,NOFAK,STATUS,PAID,DISC,DISCT,REMARK1,FN_RETURN,FS_RC,FS_KD_DEALER,FS_NM_DEALER,TIPEWO,KODETIPE,SBE, " _
                                 & "FS_KD_PLG,FS_KD_WIL,THS,MOU,PDS,FS_LOG_DEL,FD_TGL_DEL,ALASAN_DEL,INVDATE) VALUES ('" & rsHlp("CUSTOMER") & "', '" & rsHlp("CUST_ADRES") & "', '" & rsHlp("KOTA") & "', " _
                                 & "'" & rsHlp("NO_POLISI") & "', '" & rsHlp("NO_ENGINE") & "', '" & rsHlp("NO_CHASSIS") & "', '" & rsHlp("FRONT") & "', '" & rsHlp("MECH") & "', " _
                                 & "'" & rsHlp("KM") & "', '" & rsHlp("DATE") & "', '" & rsHlp("PART_G") & "', '" & rsHlp("PART_NG") & "', '" & rsHlp("OIL_G") & "','" & rsHlp("OIL_NG") & "', '" & rsHlp("MTRL") & "', '" & rsHlp("LABOR") & "', " _
                                 & "'" & rsHlp("BODY") & "', '" & rsHlp("SBLT") & "', '" & rsHlp("MET") & "', '" & rsHlp("TOTAL") & "', '" & rsHlp("JOB") & "','" & rsHlp("REP1") & "', " _
                                 & "'" & rsHlp("REP2") & "', '" & rsHlp("REP3") & "', '" & rsHlp("REP4") & "', '" & rsHlp("REP5") & "', '" & rsHlp("REP6") & "', '" & rsHlp("REP7") & "', " _
                                 & "'" & rsHlp("MEKANIK1") & "', '" & rsHlp("MEKANIK2") & "', '" & rsHlp("MEKANIK3") & "', '" & rsHlp("MEKANIK4") & "', '" & rsHlp("MEKANIK5") & "', '" & rsHlp("MEKANIK6") & "', " _
                                 & "'" & rsHlp("MEKANIK7") & "', '" & rsHlp("MEKANIK8") & "', '" & rsHlp("KLPKERJA1") & "', '" & rsHlp("KLPKERJA2") & "', '" & rsHlp("KLPKERJA3") & "', '" & rsHlp("KLPKERJA4") & "', " _
                                 & "'" & rsHlp("KLPKERJA5") & "', '" & rsHlp("KLPKERJA6") & "', '" & rsHlp("KLPKERJA7") & "', '" & rsHlp("KLPKERJA8") & "', '" & rsHlp("NO_WO") & "', '" & rsHlp("NOINV") & "', '" & rsHlp("NOFAK") & "', " _
                                 & "'" & rsHlp("STATUS") & "', '" & rsHlp("PAID") & "', '" & rsHlp("DISC") & "', '" & rsHlp("DISCT") & "', '" & rsHlp("REMARK1") & "', '" & rsHlp("FN_RETURN") & "', '" & rsHlp("FS_RC") & "', " _
                                 & "'" & rsHlp("FS_KD_DEALER") & "', '" & rsHlp("FS_NM_DEALER") & "', '" & rsHlp("TIPEWO") & "', '" & rsHlp("KODETIPE") & "', '" & rsHlp("SBE") & "', '" & rsHlp("FS_KD_PLG") & "', " _
                                 & "'" & rsHlp("FS_KD_WIL") & "', '" & rsHlp("THS") & "', '" & rsHlp("MOU") & "', '" & rsHlp("PDS") & "', '" & rsHlp("FS_LOG_DEL") & "', '" & rsHlp("FD_TGL_DEL") & "', '" & rsHlp("ALASANDEL") & "','" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "') ")
                                
                    If Trim(rsHlp("MEMOPART")) <> "" Then
                        Call inputPart(Trim(rsHlp("NO_WO")), Trim(rsHlp("NOINV")), Trim(rsHlp("MEMOPART")))
                    End If
                    
                    If Trim(rsHlp("MEMOSUBLET")) <> "" Then
                        Call inputSUBLET(Trim(rsHlp("NO_WO")), Trim(rsHlp("NOINV")), Trim(rsHlp("MEMOSUBLET")))
                    End If
                End If
                rsHlp2.Close
                rsHlp.MoveNext
            Loop
        End If
        rsHlp.Close
    End If
    End If
    Exit Sub
vbError:
    MsgBox "History Gagal di transfer...."
    Exit Sub
End Sub

Private Sub inputSUBLET(nWO As String, nINV As String, mmSUBLET As String)
Dim tHlp As String
Dim tHlp2 As String
Dim tHlp3 As String
Dim tmSUBLET As String
Dim lNow As Integer
Dim lNow2 As Integer
Dim i As Integer
Dim nw, kdItem, nmItem, noinv, nChasis, kdSat, nNota, nKet As String
Dim qty, Harga As Double
    
    lNow = Len(mmSUBLET)
    tmSUBLET = mmSUBLET
    Do While lNow > 0
        ' PISAH TANDA |
        tHlp = Left(tmSUBLET, InStr(tmSUBLET, "|"))
        lNow = lNow - Len(tHlp)
        
        lNow2 = Len(tHlp)
        i = 10
        tHlp3 = tHlp
        Do While lNow2 > 0
            If i > 1 Then
                tHlp2 = Left(Trim(tHlp3), InStr(Trim(tHlp3), "^"))
            Else
                tHlp2 = tHlp3
            End If
            lNow2 = lNow2 - Len(tHlp2)
            If i = 10 Then
                nw = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 9 Then
                noinv = Trim(Replace(tHlp2, "^", ""))
            ElseIf i = 8 Then
                nChasis = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 7 Then
                kdItem = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 6 Then
                nmItem = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 5 Then
                kdSat = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 4 Then
                qty = CDbl(Replace(tHlp2, "^", ""))
            ElseIf i = 3 Then
                Harga = CDbl(Replace(tHlp2, "^", ""))
            ElseIf i = 2 Then
                nNota = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 1 Then
                nKet = Replace(Trim(tHlp2), "|", "")
            End If
            i = i - 1
            tHlp3 = Mid(Trim(tHlp3), Len(tHlp2) + 1, Len(tHlp3) - Len(tHlp2))
        Loop
        Call objO.InsertData("INSERT INTO HIST_SUBLET VALUES ('" & nw & "','" & noinv & "','" & nChasis & "','" & kdItem & "','" & nmItem & "','" & kdSat & "'," & qty & ", " & Harga & ",'" & nNota & "', '" & nKet & "')")
        tmSUBLET = Mid(Trim(tmSUBLET), Len(tHlp) + 1, Len(tmSUBLET) - Len(tHlp))
    Loop
End Sub

Private Sub inputPart(nWO As String, nINV As String, mmPART As String)
Dim tHlp As String
Dim tHlp2 As String
Dim tHlp3 As String
Dim tmPART As String
Dim lNow As Integer
Dim lNow2 As Integer
Dim i As Integer
Dim nw, kdItem, nmItem, noinv As String
Dim qty, Harga As Double
Dim oli, disc As Integer
    
    lNow = Len(mmPART)
    tmPART = mmPART
    Do While lNow > 0
        ' PISAH TANDA |
        tHlp = Left(tmPART, InStr(tmPART, "|"))
        lNow = lNow - Len(tHlp)
        
        lNow2 = Len(tHlp)
        i = 8
        tHlp3 = tHlp
'        tHlp2 = Left(tHlp3, InStr(tHlp3, " "))
        Do While lNow2 > 0
            If i > 1 Then
                tHlp2 = Left(Trim(tHlp3), InStr(Trim(tHlp3), "^"))
            Else
                tHlp2 = tHlp3
            End If
            lNow2 = lNow2 - Len(tHlp2)
            If i = 8 Then
                nw = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 7 Then
                kdItem = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 6 Then
                nmItem = Replace(Trim(tHlp2), "^", "")
            ElseIf i = 5 Then
                qty = CDbl(Replace(tHlp2, "^", ""))
            ElseIf i = 4 Then
                Harga = CDbl(Replace(tHlp2, "^", ""))
            ElseIf i = 3 Then
                oli = IIf(Replace(tHlp2, "^", ""), 1, 0)
            ElseIf i = 2 Then
                noinv = Trim(Replace(tHlp2, "^", ""))
            ElseIf i = 1 Then
                disc = CDbl(Replace(tHlp2, "|", ""))
            End If
            i = i - 1
            tHlp3 = Mid(Trim(tHlp3), Len(tHlp2) + 1, Len(tHlp3) - Len(tHlp2))
        Loop
        Call objO.InsertData("INSERT INTO T_PART VALUES ('" & nw & "','" & kdItem & "', '" & nmItem & "'," & qty & ", " & Harga & ",'" & oli & "', '" & noinv & "'," & disc & ",'','')")
        tmPART = Mid(Trim(tmPART), Len(tHlp) + 1, Len(tmPART) - Len(tHlp))
    Loop
End Sub


Private Sub opHist_Click(Index As Integer)
    If gTag = "ORDER" Then
        Select Case Index
            Case 0
                ' SERVICE
                Call HeaderCol(0)
                Call BeriItemHO(text(0), 14)
            Case 1
                ' PART
                Call HeaderCol(1)
                Call BeriItemPartHO(text(0), 7)
            Case 2
                ' SUBLET
                Call HeaderCol(2)
                Call BeriItemSubletHO(text(0), 7)
        End Select
    Else
        Select Case Index
            Case 0
                ' SERVICE
                Call HeaderCol(0)
                Call BeriItemHO_Local(text(0), 14)
            Case 1
                ' PART
                Call HeaderCol(1)
                Call BeriItemPartHO_Local(text(0), 7)
            Case 2
                ' SUBLET
                Call HeaderCol(2)
                Call BeriItemSubletHO_Local(text(0), 7)
        End Select
    End If
End Sub

Private Sub text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
End Sub

'''================================ &&& UPDATE TGL 27 JULI 2012 - BACA HISTORY KE HO  &&& ==============================================
Private Sub BeriItemHO(noChas As String, nKol As Integer)
Dim rsHlp As ADODB.Recordset
Dim rsSearch As ADODB.Recordset
    Dim nHO1 As String
    nHO1 = "YES"
    baris = 0
    Kolom = nKol
    If Log.LogHOConn = 1 Then
        Set rsSearch = objO.FetchDataHO_NoError("select REP1, REP2, REP3, REP4, REP5, REP6, REP7 from HISTORYNEW WITH (NOLOCK) where NO_CHASSIS = '" & noChas & "' AND FS_LOG_DEL = ''")
        If rsSearch.RecordCount = 0 Then
            nHO1 = "NO"
            Set rsSearch = objO.FetchData("select REP1, REP2, REP3, REP4, REP5, REP6, REP7 from HISTORY WITH (NOLOCK) where NO_CHASSIS = '" & noChas & "' AND FS_LOG_DEL = ''")
        End If
    Else
        nHO1 = "NO"
        Set rsSearch = objO.FetchData("select REP1, REP2, REP3, REP4, REP5, REP6, REP7 from HISTORY WITH (NOLOCK) where NO_CHASSIS = '" & noChas & "' AND FS_LOG_DEL = ''")
    End If
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If Trim(rsSearch("REP1")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP2")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP3")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP4")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP5")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP6")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP7")) <> "" Then
                baris = baris + 1
            End If
            rsSearch.MoveNext
        Loop
    End If
    If nHO1 = "YES" Then
        baris = baris + CDbl(objO.FetchDataHO_Count_NoError("SELECT NO_CHASSIS FROM HIST_JOBADDNEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & noChas & "'"))
        baris = baris + CDbl(objO.FetchDataHO_Count_NoError("SELECT NO_CHASSIS FROM HIST_SUBLETNEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & noChas & "'"))
    Else
        baris = baris + CDbl(objO.FetchData_Count("SELECT NO_CHASSIS FROM HIST_JOBADD WITH (NOLOCK) WHERE NO_CHASSIS = '" & noChas & "'"))
        baris = baris + CDbl(objO.FetchData_Count("SELECT NO_CHASSIS FROM HIST_SUBLET WITH (NOLOCK) WHERE NO_CHASSIS = '" & noChas & "'"))
    End If
    i = 0
    sSQL = "select A.DATE,A.KM,A.FRONT,A.MECH AS MEKANIK1,A.MECH AS MEKANIK2,A.MECH AS MEKANIK3,A.MECH AS MEKANIK4,A.MECH AS MEKANIK5,A.MECH AS MEKANIK6,A.MECH AS MEKANIK7, " _
         & " A.CUSTOMER , A.REP1, A.REP2, A.REP3, A.REP4, A.REP5, A.REP6, A.REP7, D.DISC1, D.DISC2, D.DISC3, D.DISC4, D.DISC5, D.DISC6, D.DISC7, A.NO_WO, A.NOINV, A.REMARK1, A.SUGGEST, A.NO_CHASSIS, A.TOTAL, A.FS_NM_DEALER, " _
         & " B.REQ1, B.REQ2, B.REQ3, B.REQ4, B.REQ5, '', '', '', '', '', catat from HISTORYNEW A WITH (NOLOCK) " _
         & " LEFT JOIN JOBDTA B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS AND A.NOINV = B.NOINV " _
         & " LEFT JOIN T_INVOICENEW D WITH (NOLOCK) ON A.NO_WO = D.NO_WO AND A.NO_CHASSIS = D.NO_CHASSIS AND A.NOINV = D.NOINV " _
         & " where A.NO_CHASSIS ='" & noChas & "' AND A.FS_LOG_DEL = '' order by A.DATE DESC, A.NOINV DESC, A.KM DESC"
    
    If nHO1 = "YES" Then
getDataHO:
        Set rsSearch = objO.FetchDataHO_NoError(sSQL)
        If nHis = False Then
            If MsgBox("Gagal Mengambil Data History HO, Coba Lagi ?", vbQuestion + vbYesNo, "History HO") = vbYes Then
                GoTo getDataHO
            End If
        End If
    Else
        sSQL = Replace(sSQL, "NEW", "")
        Set rsSearch = objO.FetchData(sSQL)
        nHis = True
    End If
    If nHis = True Then
        If rsSearch.RecordCount <> 0 Then baris = baris + rsSearch.RecordCount
        
        History.Clear
        History.ReDim 0, baris + 5, 0, Kolom
    
        If Not rsSearch.BOF Then
            Do While Not rsSearch.EOF
                History(i, 0) = Trim(rsSearch("DATE"))
                History(i, 1) = Trim(rsSearch("NOINV"))
                History(i, 2) = Format(CDbl(rsSearch("TOTAL")), "#,##0.00")
                History(i, 3) = Trim(rsSearch("NO_WO"))
                History(i, 4) = Format(Trim(rsSearch("KM")), "#,##0")
                History(i, 5) = Trim(rsSearch("FRONT"))
                History(i, 6) = Trim(rsSearch("MEKANIK1"))
                History(i, 7) = Trim(rsSearch("REP1"))
                History(i, 8) = Trim(rsSearch("DISC1")) & "%"
                History(i, 9) = Trim(rsSearch("REMARK1"))
                History(i, 10) = Trim(rsSearch("SUGGEST"))
                History(i, 11) = Trim(rsSearch("CATAT"))
                History(i, 12) = Trim(rsSearch("REQ1")) & "; " & Trim(rsSearch("REQ2")) & "; " & Trim(rsSearch("REQ3")) & "; " & Trim(rsSearch("REQ4")) & "; " & Trim(rsSearch("REQ5"))
                History(i, 13) = "" 'Trim(rsSearch("CMPL1")) & "; " & Trim(rsSearch("CMPL2")) & "; " & Trim(rsSearch("CMPL3")) & "; " & Trim(rsSearch("CMPL4")) & "; " & Trim(rsSearch("CMPL5"))
                History(i, 14) = Trim(rsSearch("FS_NM_DEALER"))
                
                If Trim(rsSearch("REP2")) <> "" Then
                    i = i + 1
                    History(i, 6) = Trim(rsSearch("MEKANIK2"))
                    History(i, 7) = Trim(rsSearch("REP2"))
                    History(i, 8) = Trim(rsSearch("DISC2")) & "%"
                End If
                If Trim(rsSearch("REP3")) <> "" Then
                    i = i + 1
                    History(i, 6) = Trim(rsSearch("MEKANIK3"))
                    History(i, 7) = Trim(rsSearch("REP3"))
                    History(i, 8) = Trim(rsSearch("DISC3")) & "%"
                End If
                If Trim(rsSearch("REP4")) <> "" Then
                    i = i + 1
                    History(i, 6) = Trim(rsSearch("MEKANIK4"))
                    History(i, 7) = Trim(rsSearch("REP4"))
                    History(i, 8) = Trim(rsSearch("DISC4")) & "%"
                End If
                If Trim(rsSearch("REP5")) <> "" Then
                    i = i + 1
                    History(i, 6) = Trim(rsSearch("MEKANIK5"))
                    History(i, 7) = Trim(rsSearch("REP5"))
                    History(i, 8) = Trim(rsSearch("DISC5")) & "%"
                End If
                If Trim(rsSearch("REP6")) <> "" Then
                    i = i + 1
                    History(i, 6) = Trim(rsSearch("MEKANIK6"))
                    History(i, 7) = Trim(rsSearch("REP6"))
                    History(i, 8) = Trim(rsSearch("DISC6")) & "%"
                End If
                If Trim(rsSearch("REP7")) <> "" Then
                    i = i + 1
                    History(i, 6) = Trim(rsSearch("MEKANIK7"))
                    History(i, 7) = Trim(rsSearch("REP7"))
                    History(i, 8) = Trim(rsSearch("DISC7")) & "%"
                End If
                Set rsHlp = objO.FetchDataHO_NoError("SELECT MEKANIK,REP,DISC FROM HIST_JOBADDNEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsSearch("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsSearch("NO_WO")) & "'")
                If Not rsHlp.EOF Then
                    rsHlp.MoveFirst
                    Do While Not rsHlp.EOF
                        i = i + 1
                        History(i, 6) = Trim(rsHlp("MEKANIK"))
                        History(i, 7) = Trim(rsHlp("REP"))
                        History(i, 8) = Trim(rsHlp("DISC")) & "%"
                        rsHlp.MoveNext
                    Loop
                End If
                Set rsHlp = objO.FetchDataHO_NoError("SELECT NAMABARANG FROM HIST_SUBLETNEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsSearch("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsSearch("NO_WO")) & "'")
                If Not rsHlp.EOF Then
                    rsHlp.MoveFirst
                    Do While Not rsHlp.EOF
                        i = i + 1
                        History(i, 6) = "SUB"
                        History(i, 7) = Trim(rsHlp("NAMABARANG"))
                        History(i, 8) = ""
                        rsHlp.MoveNext
                    Loop
                End If
                rsSearch.MoveNext
                i = i + 1
            Loop
        End If
    TData1.Array = History
    TData1.ReBind
    End If
End Sub

Private Sub BeriItemPartHO(noChas As String, nKol As Integer)
    Dim OldInv As String
    Dim rsSearch As ADODB.Recordset
    
    baris = 0
    Kolom = nKol
    sSQL = "select count(B.NoInv) as JML" _
         & " from historynew as A WITH (NOLOCK) inner join t_partnew as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') and B.FS_LOG_DEL = ''" _
         & " where NO_CHASSIS = '" & noChas & "' and B.NOINV <> '' and A.FS_LOG_DEL = ''"
    If Log.LogHOConn = 1 Then
        baris = objO.FetchDataHO_Count_NoError(sSQL)
        If baris = 0 Then
            baris = objO.FetchData_Count(Replace(UCase(sSQL), "NEW", ""))
        End If
    Else
        baris = objO.FetchData_Count(Replace(UCase(sSQL), "NEW", ""))
    End If
    
    i = 0
    OldInv = "LAMA"
    sSQL = "select [Date], A.NoInv, A.no_wo, PART_G+PART_NG+OIL_G+OIL_NG AS  PART, FRONT, KODEITEM, Keterangan, QTY, HARGA, OLI" _
         & " from historynew as A WITH (NOLOCK) inner join t_partnew as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') and B.FS_LOG_DEL = ''" _
         & " where NO_CHASSIS = '" & noChas & "' and B.NOINV <> '' and A.FS_LOG_DEL = '' order by [DATE] DESC, KODEITEM"
    
    If Log.LogHOConn = 1 Then
        Set rsSearch = objO.FetchDataHO_NoError(sSQL)
        If rsSearch.RecordCount = 0 Then
            Set rsSearch = objO.FetchData(Replace(UCase(sSQL), "NEW", ""))
        End If
    Else
        Set rsSearch = objO.FetchData(Replace(UCase(sSQL), "NEW", ""))
    End If
    If rsSearch.RecordCount <> 0 Then baris = baris + rsSearch.RecordCount
    
    History.Clear
    History.ReDim 0, baris + 2, 0, Kolom
    
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If OldInv <> Trim(rsSearch("NOINV")) Then
                History(i, 0) = Trim(rsSearch("DATE"))
                History(i, 1) = Trim(rsSearch("NOINV"))
                History(i, 2) = Trim(rsSearch("NO_WO"))
                History(i, 3) = Format(Trim(rsSearch("PART")), "#,##0.00")
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("Keterangan"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            Else
                History(i, 0) = ""
                History(i, 1) = ""
                History(i, 2) = ""
                History(i, 3) = ""
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("Keterangan"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            End If
            OldInv = Trim(rsSearch("NOINV"))
            rsSearch.MoveNext
            i = i + 1
        Loop
    End If
    TData1.Array = History
    TData1.ReBind
End Sub

Private Sub BeriItemSubletHO(noChas As String, nKol As Integer)
    Dim OldInv As String
    Dim rsSearch As ADODB.Recordset
    
    baris = 0
    Kolom = nKol
    sSQL = "SELECT COUNT(*) AS JML FROM historynew A WITH (NOLOCK) INNER JOIN HIST_SUBLETNEW B WITH (NOLOCK) " _
            & "ON A.no_wo = B.no_wo and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') AND A.FS_LOG_DEL = '' WHERE A.NO_CHASSIS = '" & noChas & "' AND B.noinv <> ''"
            
    If Log.LogHOConn = 1 Then
        baris = objO.FetchDataHO_Count_NoError(sSQL)
        If baris = 0 Then
            baris = objO.FetchData_Count(Replace(UCase(sSQL), "NEW", ""))
        End If
    Else
        baris = objO.FetchData_Count(Replace(UCase(sSQL), "NEW", ""))
    End If

    i = 0
    
    OldInv = "LAMA"
    sSQL = "SELECT [DATE], A.NOINV, A.NO_WO, SBLT, KODEITEM, NAMABARANG, QTY, HARGA FROM historynew A WITH (NOLOCK) " _
         & "INNER JOIN HIST_SUBLETNEW B WITH (NOLOCK) ON A.no_wo = B.no_wo and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') AND A.FS_LOG_DEL = '' WHERE A.NO_CHASSIS = '" & noChas & "' " _
         & "AND B.noinv <> '' order by [DATE] DESC, KODEITEM"
    If Log.LogHOConn = 1 Then
        Set rsSearch = objO.FetchDataHO_NoError(sSQL)
        If rsSearch.RecordCount = 0 Then
            Set rsSearch = objO.FetchData(Replace(UCase(sSQL), "NEW", ""))
        End If
    Else
        Set rsSearch = objO.FetchData(Replace(UCase(sSQL), "NEW", ""))
    End If
    If rsSearch.RecordCount <> 0 Then baris = baris + rsSearch.RecordCount
    
    History.Clear
    History.ReDim 0, baris + 2, 0, Kolom
    
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If OldInv <> Trim(rsSearch("NOINV")) Then
                History(i, 0) = Trim(rsSearch("DATE"))
                History(i, 1) = Trim(rsSearch("NOINV"))
                History(i, 2) = Trim(rsSearch("NO_WO"))
                History(i, 3) = Format(Trim(rsSearch("SBLT")), "#,##0.00")
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("NAMABARANG"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            Else
                History(i, 0) = ""
                History(i, 1) = ""
                History(i, 2) = ""
                History(i, 3) = ""
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("NAMABARANG"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            End If
            
            OldInv = Trim(rsSearch("NOINV"))
            rsSearch.MoveNext
            i = i + 1
        Loop
    End If
    TData1.Array = History
    TData1.ReBind
End Sub



'''================================ &&& UPDATE TGL 04 AGUSTUS 2012 - BACA HISTORY KE LOCAL &&& ==============================================
Private Sub BeriItemHO_Local(noChas As String, nKol As Integer)
Dim rsHlp As ADODB.Recordset
Dim rsSearch As ADODB.Recordset
    baris = 0
    Kolom = nKol
    Set rsSearch = objO.FetchData("select REP1, REP2, REP3, REP4, REP5, REP6, REP7 from HISTORY WITH (NOLOCK) where NO_CHASSIS = '" & noChas & "' AND FS_LOG_DEL = ''")
        
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If Trim(rsSearch("REP1")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP2")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP3")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP4")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP5")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP6")) <> "" Then
                baris = baris + 1
            End If
            If Trim(rsSearch("REP7")) <> "" Then
                baris = baris + 1
            End If
            rsSearch.MoveNext
        Loop
    End If
    
        baris = baris + CDbl(objO.FetchData_Count("SELECT NO_CHASSIS FROM HIST_JOBADD WITH (NOLOCK) WHERE NO_CHASSIS = '" & noChas & "'"))
        baris = baris + CDbl(objO.FetchData_Count("SELECT NO_CHASSIS FROM HIST_SUBLET WITH (NOLOCK) WHERE NO_CHASSIS = '" & noChas & "'"))
    
    i = 0
'''    sSQL = "select A.DATE,A.KM,A.FRONT,A.MECH AS MEKANIK1,A.MECH AS MEKANIK2,A.MECH AS MEKANIK3,A.MECH AS MEKANIK4,A.MECH AS MEKANIK5,A.MECH AS MEKANIK6,A.MECH AS MEKANIK7, " _
'''         & " A.CUSTOMER , A.REP1, A.REP2, A.REP3, A.REP4, A.REP5, A.REP6, A.REP7, D.DISC1, D.DISC2, D.DISC3, D.DISC4, D.DISC5, D.DISC6, D.DISC7, A.NO_WO, A.NOINV, A.REMARK1, A.SUGGEST, A.NO_CHASSIS, A.TOTAL, A.FS_NM_DEALER, " _
'''         & " B.REQ1, B.REQ2, B.REQ3, B.REQ4, B.REQ5, '', '', '', '', '', catat from HISTORYNEW A " _
'''         & " LEFT JOIN JOBDTA B ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS AND A.NOINV = B.NOINV " _
'''         & " LEFT JOIN T_INVOICENEW D ON A.NO_WO = D.NO_WO AND A.NO_CHASSIS = D.NO_CHASSIS AND A.NOINV = D.NOINV " _
'''         & " where A.NO_CHASSIS ='" & noChas & "' AND A.FS_LOG_DEL = '' order by A.DATE DESC, A.NOINV DESC, A.KM DESC"
    
    sSQL = "select A.DATE,A.KM,A.FRONT,A.MECH AS MEKANIK1,A.MECH AS MEKANIK2,A.MECH AS MEKANIK3,A.MECH AS MEKANIK4,A.MECH AS MEKANIK5,A.MECH AS MEKANIK6,A.MECH AS MEKANIK7, " _
         & " A.CUSTOMER , A.REP1, A.REP2, A.REP3, A.REP4, A.REP5, A.REP6, A.REP7, D.DISC1, D.DISC2, D.DISC3, D.DISC4, D.DISC5, D.DISC6, D.DISC7, A.NO_WO, A.NOINV, A.REMARK1, A.SUGGEST, A.NO_CHASSIS, A.TOTAL, A.FS_NM_DEALER, " _
         & " B.REQ1, B.REQ2, B.REQ3, B.REQ4, B.REQ5, C.CMPL1, C.CMPL2, C.CMPL3, C.CMPL4, C.CMPL5, catat from HISTORY A WITH (NOLOCK) " _
         & " LEFT JOIN JOBDTA B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS AND A.NOINV = B.NOINV " _
         & " LEFT JOIN JOBDTL C WITH (NOLOCK) ON A.NO_WO = C.NO_WO LEFT JOIN T_INVOICE D WITH (NOLOCK) ON A.NO_WO = D.NO_WO AND A.NO_CHASSIS = D.NO_CHASSIS AND A.NOINV = D.NOINV " _
         & " where A.NO_CHASSIS ='" & noChas & "' AND A.FS_LOG_DEL = '' order by A.DATE DESC, A.NOINV DESC, A.KM DESC"
    
    
        sSQL = Replace(sSQL, "NEW", "")
        Set rsSearch = objO.FetchData(sSQL)
    
    
    If rsSearch.RecordCount <> 0 Then baris = baris + rsSearch.RecordCount
    
    History.Clear
    History.ReDim 0, baris + 5, 0, Kolom

    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            History(i, 0) = Trim(rsSearch("DATE"))
            History(i, 1) = Trim(rsSearch("NOINV"))
            History(i, 2) = Format(CDbl(rsSearch("TOTAL")), "#,##0.00")
            History(i, 3) = Trim(rsSearch("NO_WO"))
            History(i, 4) = Format(Trim(rsSearch("KM")), "#,##0")
            History(i, 5) = Trim(rsSearch("FRONT"))
            History(i, 6) = Trim(rsSearch("MEKANIK1"))
            History(i, 7) = Trim(rsSearch("REP1"))
            History(i, 8) = Trim(rsSearch("DISC1")) & "%"
            History(i, 9) = Trim(rsSearch("REMARK1"))
            History(i, 10) = Trim(rsSearch("SUGGEST"))
            History(i, 11) = Trim(rsSearch("CATAT"))
            History(i, 12) = Trim(rsSearch("REQ1")) & "; " & Trim(rsSearch("REQ2")) & "; " & Trim(rsSearch("REQ3")) & "; " & Trim(rsSearch("REQ4")) & "; " & Trim(rsSearch("REQ5"))
            History(i, 13) = Trim(rsSearch("CMPL1")) & "; " & Trim(rsSearch("CMPL2")) & "; " & Trim(rsSearch("CMPL3")) & "; " & Trim(rsSearch("CMPL4")) & "; " & Trim(rsSearch("CMPL5"))
            History(i, 14) = Trim(rsSearch("FS_NM_DEALER"))
            
            If Trim(rsSearch("REP2")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK2"))
                History(i, 7) = Trim(rsSearch("REP2"))
                History(i, 8) = Trim(rsSearch("DISC2")) & "%"
            End If
            If Trim(rsSearch("REP3")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK3"))
                History(i, 7) = Trim(rsSearch("REP3"))
                History(i, 8) = Trim(rsSearch("DISC3")) & "%"
            End If
            If Trim(rsSearch("REP4")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK4"))
                History(i, 7) = Trim(rsSearch("REP4"))
                History(i, 8) = Trim(rsSearch("DISC4")) & "%"
            End If
            If Trim(rsSearch("REP5")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK5"))
                History(i, 7) = Trim(rsSearch("REP5"))
                History(i, 8) = Trim(rsSearch("DISC5")) & "%"
            End If
            If Trim(rsSearch("REP6")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK6"))
                History(i, 7) = Trim(rsSearch("REP6"))
                History(i, 8) = Trim(rsSearch("DISC6")) & "%"
            End If
            If Trim(rsSearch("REP7")) <> "" Then
                i = i + 1
                History(i, 6) = Trim(rsSearch("MEKANIK7"))
                History(i, 7) = Trim(rsSearch("REP7"))
                History(i, 8) = Trim(rsSearch("DISC7")) & "%"
            End If
            If Log.LogHOConn = 1 Then
            Set rsHlp = objO.FetchDataHO_NoError("SELECT MEKANIK,REP,DISC FROM HIST_JOBADDNEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsSearch("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsSearch("NO_WO")) & "'")
                If Not rsHlp.EOF Then
                    rsHlp.MoveFirst
                    Do While Not rsHlp.EOF
                        i = i + 1
                        History(i, 6) = Trim(rsHlp("MEKANIK"))
                        History(i, 7) = Trim(rsHlp("REP"))
                        History(i, 8) = Trim(rsHlp("DISC")) & "%"
                        rsHlp.MoveNext
                    Loop
                End If
                Set rsHlp = objO.FetchDataHO_NoError("SELECT NAMABARANG FROM HIST_SUBLETNEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(rsSearch("NO_CHASSIS")) & "' AND NO_WO = '" & Trim(rsSearch("NO_WO")) & "'")
                If Not rsHlp.EOF Then
                    rsHlp.MoveFirst
                    Do While Not rsHlp.EOF
                        i = i + 1
                        History(i, 6) = "SUB"
                        History(i, 7) = Trim(rsHlp("NAMABARANG"))
                        History(i, 8) = ""
                        rsHlp.MoveNext
                    Loop
                End If
            End If
            rsSearch.MoveNext
            i = i + 1
        Loop
    End If
    TData1.Array = History
    TData1.ReBind
End Sub

Private Sub BeriItemPartHO_Local(noChas As String, nKol As Integer)
    Dim OldInv As String
    Dim rsSearch As ADODB.Recordset
    
    baris = 0
    Kolom = nKol
    sSQL = "select count(B.NoInv) as JML" _
         & " from historynew as A WITH (NOLOCK) inner join t_partnew as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') and B.FS_LOG_DEL = ''" _
         & " where NO_CHASSIS = '" & noChas & "' and B.NOINV <> '' and A.FS_LOG_DEL = ''"
    
        baris = objO.FetchData_Count(Replace(UCase(sSQL), "NEW", ""))
    
    
    i = 0
    OldInv = "LAMA"
    sSQL = "select [Date], A.NoInv, A.no_wo, PART_G+PART_NG+OIL_G+OIL_NG AS  PART, FRONT, KODEITEM, Keterangan, QTY, HARGA, OLI" _
         & " from historynew as A WITH (NOLOCK) inner join t_partnew as B WITH (NOLOCK) on A.NO_WO = B.NO_WO and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') and B.FS_LOG_DEL = ''" _
         & " where NO_CHASSIS = '" & noChas & "' and B.NOINV <> '' and A.FS_LOG_DEL = '' order by [DATE] DESC, KODEITEM"
    
    
        Set rsSearch = objO.FetchData(Replace(UCase(sSQL), "NEW", ""))
    
    If rsSearch.RecordCount <> 0 Then baris = baris + rsSearch.RecordCount
    
    History.Clear
    History.ReDim 0, baris + 2, 0, Kolom
    
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If OldInv <> Trim(rsSearch("NOINV")) Then
                History(i, 0) = Trim(rsSearch("DATE"))
                History(i, 1) = Trim(rsSearch("NOINV"))
                History(i, 2) = Trim(rsSearch("NO_WO"))
                History(i, 3) = Format(Trim(rsSearch("PART")), "#,##0.00")
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("Keterangan"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            Else
                History(i, 0) = ""
                History(i, 1) = ""
                History(i, 2) = ""
                History(i, 3) = ""
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("Keterangan"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            End If
            OldInv = Trim(rsSearch("NOINV"))
            rsSearch.MoveNext
            i = i + 1
        Loop
    End If
    TData1.Array = History
    TData1.ReBind
End Sub

Private Sub BeriItemSubletHO_Local(noChas As String, nKol As Integer)
    Dim OldInv As String
    Dim rsSearch As ADODB.Recordset
    
    baris = 0
    Kolom = nKol
    sSQL = "SELECT COUNT(*) AS JML FROM historynew A WITH (NOLOCK) INNER JOIN HIST_SUBLETNEW B WITH (NOLOCK) " _
            & "ON A.no_wo = B.no_wo and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') AND A.FS_LOG_DEL = '' WHERE A.NO_CHASSIS = '" & noChas & "' AND B.noinv <> ''"
            
    
        baris = objO.FetchData_Count(Replace(UCase(sSQL), "NEW", ""))
    

    i = 0
    
    OldInv = "LAMA"
    sSQL = "SELECT [DATE], A.NOINV, A.NO_WO, SBLT, KODEITEM, NAMABARANG, QTY, HARGA FROM historynew A WITH (NOLOCK) " _
         & "INNER JOIN HIST_SUBLETNEW B WITH (NOLOCK) ON A.no_wo = B.no_wo and A.NOINV = Replace(Replace(B.NOINV,'-A',''),'-B','') AND A.FS_LOG_DEL = '' WHERE A.NO_CHASSIS = '" & noChas & "' " _
         & "AND B.noinv <> '' order by [DATE] DESC, KODEITEM"
    
        Set rsSearch = objO.FetchData(Replace(UCase(sSQL), "NEW", ""))
    
    If rsSearch.RecordCount <> 0 Then baris = baris + rsSearch.RecordCount
    
    History.Clear
    History.ReDim 0, baris + 2, 0, Kolom
    
    If Not rsSearch.BOF Then
        Do While Not rsSearch.EOF
            If OldInv <> Trim(rsSearch("NOINV")) Then
                History(i, 0) = Trim(rsSearch("DATE"))
                History(i, 1) = Trim(rsSearch("NOINV"))
                History(i, 2) = Trim(rsSearch("NO_WO"))
                History(i, 3) = Format(Trim(rsSearch("SBLT")), "#,##0.00")
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("NAMABARANG"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            Else
                History(i, 0) = ""
                History(i, 1) = ""
                History(i, 2) = ""
                History(i, 3) = ""
                History(i, 4) = Trim(rsSearch("KODEITEM"))
                History(i, 5) = Trim(rsSearch("NAMABARANG"))
                History(i, 6) = Trim(rsSearch("QTY"))
                History(i, 7) = Format(Trim(rsSearch("HARGA")), "#,##0.00")
            End If
            
            OldInv = Trim(rsSearch("NOINV"))
            rsSearch.MoveNext
            i = i + 1
        Loop
    End If
    TData1.Array = History
    TData1.ReBind
End Sub


