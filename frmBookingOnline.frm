VERSION 5.00
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmBookingOnline 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "hari"
   ClientHeight    =   8340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14310
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   3960
      Top             =   600
   End
   Begin VB.Frame frameBatalBooking 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2040
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox txtAlasan 
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   5295
      End
      Begin AIFCmp1.asxPowerButton btnSimpanBatal 
         Height          =   495
         Left            =   4080
         TabIndex        =   12
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BackColor       =   33023
         Caption         =   "SIMPAN"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextColor       =   16777215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "PEMBATALAN BOOKING ONLINE"
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
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   120
         Width           =   5055
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF0000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   5535
      End
   End
   Begin AIFCmp1.asxPowerButton tblImport 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "DATA ONLINE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton tblSimpan 
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Top             =   4800
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   873
      BackColor       =   49152
      Caption         =   "BUAT BOOKING"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton tblBatal 
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   4800
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BackColor       =   33023
      Caption         =   "BATAL"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxPowerButton asxPowerButton1 
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   4800
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BackColor       =   255
      Caption         =   "BATAL BOOKING"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin TrueOleDBGrid70.TDBGrid listImport 
      Height          =   3495
      Left            =   0
      TabIndex        =   13
      Top             =   1200
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   6165
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
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
      Columns(11)._VlistStyle=   4
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
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
      Splits(0)._ColumnProps(33)=   "Column(8).Width=3281"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=54,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=51,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=52,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=53,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=46,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=28,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=25,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=26,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=27,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=91,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=88,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=89,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=90,.parent=17"
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
      _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
      _StyleDefs(99)  =   "Named:id=40:OddRow"
      _StyleDefs(100) =   ":id=40,.parent=33"
      _StyleDefs(101) =   "Named:id=41:RecordSelector"
      _StyleDefs(102) =   ":id=41,.parent=34"
      _StyleDefs(103) =   "Named:id=42:FilterBar"
      _StyleDefs(104) =   ":id=42,.parent=33"
      _StyleDefs(105) =   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(106) =   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(107) =   "Named:id=132:TigaDigitKosong"
      _StyleDefs(108) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(109) =   "Named:id=87:SudahDitagih"
      _StyleDefs(110) =   ":id=87,.parent=132,.bgcolor=&HC0C0C0&,.fgcolor=&HFF&,.locked=-1,.bold=-1"
      _StyleDefs(111) =   ":id=87,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(112) =   ":id=87,.fontname=Tahoma"
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Yang Dipilih :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   6120
      TabIndex        =   8
      Top             =   720
      Width           =   3060
   End
   Begin VB.Label nTagih 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   9360
      TabIndex        =   7
      Top             =   720
      Width           =   180
   End
   Begin VB.Label lblOT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "DATA BOOKING ONLINE BELUM TRANSFER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1365
      Width           =   9735
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   9840
      X2              =   9240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   9840
      X2              =   9000
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   9840
      X2              =   8880
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
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "DATA BOOKING ONLINE"
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
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   7695
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
      Width           =   9735
   End
End
Attribute VB_Name = "frmBookingOnline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsImport As ADODB.Recordset
Dim rsTempImport As New ADODB.Recordset
Dim nTotal As Integer

Private Sub createImport()
    On Error GoTo vb_error
    
    If rsTempImport.State Then rsTempImport.Close
    With rsTempImport
        .Fields.Append "No", adInteger
        .Fields.Append "Import", adBoolean
        .Fields.Append "Id", adInteger
        .Fields.Append "NoPolisi", adVarChar, 25
        .Fields.Append "NamaCustomer", adVarChar, 50
        .Fields.Append "JenisKendaraan", adVarChar, 50
        .Fields.Append "TanggalBooking", adVarChar, 25
        .Fields.Append "JamBooking", adVarChar, 25
        .Fields.Append "NoTel", adVarChar, 25
        .Fields.Append "Email", adVarChar, 25
        .Fields.Append "Kota", adVarChar, 25
        .Fields.Append "ProsesBooking", adVarChar, 25
        .Open
    End With
    'If rsTempImport.State Then createImport = True
    Set listImport.DataSource = rsTempImport
    Call headColumns
    Exit Sub
    
vb_error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Sub


Private Sub headColumns()
    With listImport
        .Columns(0).Caption = "No"
        .Columns(0).DataField = "No"
        .Columns(0).Width = 600
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Caption = "Proses Booking"
        .Columns(1).DataField = "Import"
        .Columns(1).Width = 500
      '  .Splits(0).Columns(1).Locked = True
        
        .Columns(3).Caption = "No Polisi"
        .Columns(3).DataField = "NoPolisi"
        .Columns(3).Width = 1000
        .Splits(0).Columns(3).Locked = True
        
        .Columns(2).Caption = "ID Booking"
        .Columns(2).DataField = "ID"
        .Columns(2).Width = 1000
        .Splits(0).Columns(2).Locked = True
        
        .Columns(4).Caption = "Nama Customer"
        .Columns(4).DataField = "NamaCustomer"
        .Columns(4).Width = 2000
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).DataField = "NoTel"
        .Columns(5).Width = 1000
        .Columns(5).Caption = "No Telepon"
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Caption = "Jenis Kendaraan"
        .Columns(6).DataField = "JenisKendaraan"
        .Columns(6).Width = 2000
        .Splits(0).Columns(6).Locked = True
        
        .Columns(7).Caption = "Tanggal Booking"
        .Columns(7).DataField = "TanggalBooking"
        .Columns(7).Width = 1000
        .Splits(0).Columns(7).Locked = True
        
        .Columns(8).Caption = "Jam Booking"
        .Columns(8).DataField = "JamBooking"
        .Columns(8).Width = 1000
        .Splits(0).Columns(8).Locked = True
        
        .Columns(9).Caption = "Status"
        .Columns(9).DataField = "ProsesBooking"
        .Columns(9).Width = 2000
        .Splits(0).Columns(9).Locked = True
        
        .Columns(10).Caption = "No Booking"
        .Columns(10).DataField = "NoBooking"
        .Columns(10).Width = 1500
        .Splits(0).Columns(10).Locked = True
        
        .Columns(11).Visible = False
        .Columns(11).DataField = "ProsesBooking"
        .Columns(11).Width = 1500
        .Splits(0).Columns(11).Locked = True
    End With
End Sub

Private Sub refreshGrid()
    Dim i As Integer
    Set rsImport = objO.FetchData("select fProses,replace(notel,'`','') as Notel,Email,Kota,convert(varchar,tanggal_booking,103) tanggalbooking,convert(varchar,jam_booking,108) jambooking,id,nopol as NoPolisi,nama_pelanggan as NamaCustomer,jenis_mobil as JenisKendaraan from bookingonline WITH (NOLOCK) where dealer = '" & Log.KodeDealer & "'")
    If rsImport.RecordCount > 0 Then
            Call createImport
            Dim hasilProsesBooking As String
            i = 1
            nTotal = 0
            rsImport.MoveFirst
            Do While Not rsImport.EOF
                hasilProsesBooking = ""
                If rsImport("fProses") = 0 Then
                    hasilProsesBooking = "Belum Proses"
                    nTotal = nTotal + 1
                ElseIf rsImport("fProses") = 1 Then
                    hasilProsesBooking = "Booking"
                ElseIf rsImport("fProses") = 2 Then
                    hasilProsesBooking = "Batal Booking"
                End If
                
                With rsTempImport
                    .AddNew
                    !No = i
                    !Import = True
                    !id = rsImport("ID")
                    !NoPolisi = Trim(rsImport("NoPolisi"))
                    !NamaCustomer = Trim(rsImport("NamaCustomer"))
                    !JenisKendaraan = Trim(rsImport("JenisKendaraan"))
                    !TanggalBooking = Trim(rsImport("TanggalBooking"))
                    !JamBooking = Trim(rsImport("JamBooking"))
                    !Notel = Trim(rsImport("Notel"))
                    !Email = Trim(rsImport("Email"))
                    !Kota = Trim(rsImport("Kota"))
                    !ProsesBooking = hasilProsesBooking
                    .Update
                    .MoveNext
                End With
                i = i + 1
                rsImport.MoveNext
            Loop
    End If
    listImport.DataSource = rsTempImport
    listImport.FetchRowStyle = True
    
    Call headColumns
    rsTempImport.MoveFirst
    
    tblSimpan.Enabled = True
    nTagih = nTotal
End Sub


Private Sub cmdKeluar_Click()
    Unload Me
End Sub

Private Sub asxPowerButton1_Click()
    If nTotal > 0 Then
        If MsgBox("Apakan yakin membatalkan " & nTotal & " Pesan Booking?", vbYesNo) = vbYes Then
            txtAlasan.text = ""
            frameBatalBooking.Visible = True
        End If
    End If
End Sub

Private Sub btnSimpanBatal_Click()
    If Not rsTempImport.EOF Then
        rsTempImport.MoveFirst
        Do While Not rsTempImport.EOF
            If rsTempImport("import") = True Then
                Call objO.InsertData("UPDATE BOOKINGONLINE SET FPROSES = 2,HASILFU = '" & Trim(txtAlasan.text) & "' WHERE ID = '" & rsTempImport("Id") & "'")
            End If
            rsTempImport.MoveNext
        Loop
        MsgBox "Data Berhasil Diubah", vbInformation
        frameBatalBooking.Visible = False
        Call refreshGrid
    End If
End Sub

Private Sub Form_Load()
    nTotal = 0
    refreshGrid
End Sub

Private Sub ListImport_AfterColUpdate(ByVal ColIndex As Integer)
    With listImport
        Select Case ColIndex
            Case 1
                If .Columns(ColIndex) Then
                    nTotal = nTotal + 1
                Else
                    nTotal = nTotal - 1
                End If
                
        End Select
        nTagih = nTotal
    End With
End Sub

Private Sub ListImport_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsTempImport
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    If Trim(rs.Fields(11).value) <> "Belum Proses" Then
        RowStyle = "SudahDitagih"
    End If
End Sub

Private Sub tblBatal_Click()
    Unload Me
End Sub

Private Sub tblImport_Click()
    Dim f As New frmBookingImportDataOnline '.Show vbModal
    f.Show vbModal
End Sub

Private Sub tblSimpan_Click()
    Dim noPol, noTelpBookingOnline, namaCustomerBookingOnline, kotaBookingOnline, tanggalBookingOnline, jamBookingOnline As String
    Dim IDBookingOnline As Integer
    
    If Not rsTempImport.EOF Then
        Dim rsCloneBooking As New ADODB.Recordset
        Set rsCloneBooking = rsTempImport
        
        rsCloneBooking.MoveFirst
        noPol = ""
        Do While Not rsCloneBooking.EOF
            If rsCloneBooking("import") And rsCloneBooking("ProsesBooking") = "Belum Proses" Then
                noPol = rsCloneBooking("NoPolisi")
                IDBookingOnline = CInt(rsCloneBooking("ID"))
                noTelpBookingOnline = Trim(rsCloneBooking("Notel"))
                namaCustomerBookingOnline = Trim(rsCloneBooking("NamaCustomer"))
                kotaBookingOnline = Trim(rsCloneBooking("Kota"))
                tanggalBookingOnline = Trim(rsCloneBooking("TanggalBooking"))
                jamBookingOnline = Trim(rsCloneBooking("JamBooking"))
                
            End If
            rsCloneBooking.MoveNext
        Loop
     '   rsTempImport.MoveFirst
        If noPol <> "" Then
            If MsgBox("Apakah yakkin membuat Booking dari No Polisi = " & noPol & " ?", vbYesNo) = vbYes Then
                If objO.FetchData_Count("SELECT ID FROM BOOKINGONLINE WHERE ID = '" & IDBookingOnline & "' AND FPROSES = 0") = 0 Then
                    MsgBox "Unit dengan No Polisi = " & noPol & " telah dilakukan booking / batal booking", vbInformation
                Else
                    If gFiskal = Format(Month(gTglTampung), "00") Then
'                        If Log.DmsNonDss = "1" Then
'                            '-----------------------------------
'                            '.:: SET AS CHILD ::.
'                            '-----------------------------------
'                            TOChild Layar.hWnd, frmBookingDMSNonDSS.hWnd
'                            '-----------------------------------
'                            '-----------------------------------
'                            isBookByMrs = False
'                            frmBookingDMSNonDSS.passingISBookingOnline = True
'                            frmBookingDMSNonDSS.passingNoPolBookingOnline = noPol
'
'                            frmBookingDMSNonDSS.passingIDBookingOnline = IDBookingOnline
'
'                            frmBookingDMSNonDSS.passingNoTelpBookingOnline = noTelpBookingOnline
'                            frmBookingDMSNonDSS.passingNamaCustomerBookingOnline = namaCustomerBookingOnline
'                            frmBookingDMSNonDSS.passingKotaBookingOnline = kotaBookingOnline
'                            frmBookingDMSNonDSS.passingTanggalBookingOnline = tanggalBookingOnline
'                            frmBookingDMSNonDSS.passingJamBookingOnline = jamBookingOnline
'                            frmBookingDMSNonDSS.Height = frmBookingDMSNonDSS.asxPowerBanner1.Height
'                            frmBookingDMSNonDSS.Show
'                        Else
                            '-----------------------------------
                            '.:: SET AS CHILD ::.
                            '-----------------------------------
                            TOChild Layar.hWnd, frmBooking.hWnd
                            '-----------------------------------
                            '-----------------------------------
                            isBookByMrs = False
                            frmBooking.passingISBookingOnline = True
                            frmBooking.passingNoPolBookingOnline = noPol
    
                            frmBooking.passingIDBookingOnline = IDBookingOnline
    
                            frmBooking.passingNoTelpBookingOnline = noTelpBookingOnline
                            frmBooking.passingNamaCustomerBookingOnline = namaCustomerBookingOnline
                            frmBooking.passingKotaBookingOnline = kotaBookingOnline
                            frmBooking.passingTanggalBookingOnline = tanggalBookingOnline
                            frmBooking.passingJamBookingOnline = jamBookingOnline
                            frmBooking.Height = frmBooking.asxPowerBanner1.Height
                            frmBooking.Show
'                        End If
                        refreshGrid
                    Else
                        MsgBox "Salah Fiskal, Hubungi Administrasi untuk melakukan perubahan Fiskal !"
                    End If
                End If
            End If
        Else
            MsgBox "Tidak ada No Polisi yang dapat dibuat Booking", vbInformation
            nTotal = 0
            refreshGrid
        End If
        
    End If
    
End Sub

Private Sub Timer1_Timer()
    'Call refreshGrid
End Sub
