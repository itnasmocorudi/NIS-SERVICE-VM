VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FormSMS 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   0
   ClientTop       =   4005
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FormSMS.frx":0000
   ScaleHeight     =   8850
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExitForm 
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
      Left            =   11160
      Picture         =   "FormSMS.frx":EC75
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   6360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame fr 
      BackColor       =   &H00BFE8EA&
      Caption         =   " "
      Height          =   8175
      Index           =   0
      Left            =   360
      TabIndex        =   15
      Top             =   480
      Width           =   12135
      Begin VB.CommandButton CmdXlxInbox 
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
         Left            =   8880
         Picture         =   "FormSMS.frx":F405
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   6840
         Width           =   735
      End
      Begin VB.CommandButton CmdSearch 
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
         Left            =   5640
         Picture         =   "FormSMS.frx":F881
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   6840
         Width           =   735
      End
      Begin VB.CommandButton CmdRefreshLap 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   735
         Left            =   9600
         Picture         =   "FormSMS.frx":FED0
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   6840
         Width           =   735
      End
      Begin VB.CommandButton CmdRead 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Read"
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
         Left            =   8160
         Picture         =   "FormSMS.frx":10534
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   6840
         Width           =   735
      End
      Begin VB.TextBox txtIn3 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   80
         TabIndex        =   34
         Top             =   5310
         Width           =   9975
      End
      Begin VB.TextBox txtIn2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   80
         TabIndex        =   33
         Top             =   5000
         Width           =   9975
      End
      Begin VB.TextBox txtStatus 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   30
         Top             =   6045
         Width           =   2415
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   27
         Top             =   5685
         Width           =   2415
      End
      Begin VB.TextBox txtWaktu 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   24
         Top             =   6045
         Width           =   2415
      End
      Begin VB.TextBox txtPengirim 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   21
         Top             =   5685
         Width           =   2415
      End
      Begin VB.TextBox txtIn1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   80
         TabIndex        =   18
         Top             =   4680
         Width           =   9975
      End
      Begin VB.CommandButton cmdExitIn 
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
         Left            =   10680
         Picture         =   "FormSMS.frx":10B2D
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   6840
         Width           =   735
      End
      Begin TrueOleDBGrid70.TDBGrid GridInbox 
         Height          =   3975
         Left            =   0
         TabIndex        =   16
         Top             =   480
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7011
         _LayoutType     =   4
         _RowHeight      =   28
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No."
         Columns(0).DataField=   "Id"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Pesan"
         Columns(1).DataField=   "pesan"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Pengirim"
         Columns(2).DataField=   "sender"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Waktu"
         Columns(3).DataField=   "waktu"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "User"
         Columns(4).DataField=   "milik"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Status"
         Columns(5).DataField=   "status"
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
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=7064"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=6959"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3969"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3863"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=3519"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3413"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1323"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1217"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=7064"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=6959"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
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
         _StyleDefs(81)  =   "Named:id=29:notSent"
         _StyleDefs(82)  =   ":id=29,.parent=33,.fgcolor=&HFF0000&"
      End
      Begin VB.TextBox txtNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   11160
         TabIndex        =   40
         Top             =   480
         Width           =   375
      End
      Begin TDBDate6Ctl.TDBDate Dt1 
         Height          =   315
         Left            =   1560
         TabIndex        =   75
         Top             =   6960
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         Calendar        =   "FormSMS.frx":112BD
         Caption         =   "FormSMS.frx":113E9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FormSMS.frx":11455
         Keys            =   "FormSMS.frx":11473
         Spin            =   "FormSMS.frx":114D1
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
         Left            =   3720
         TabIndex        =   77
         Top             =   6960
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         Calendar        =   "FormSMS.frx":114F9
         Caption         =   "FormSMS.frx":11625
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FormSMS.frx":11691
         Keys            =   "FormSMS.frx":116AF
         Spin            =   "FormSMS.frx":1170D
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
      Begin VB.Shape Shape14 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   5520
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   975
      End
      Begin VB.Line Line26 
         X1              =   3360
         X2              =   3600
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   78
         Top             =   6960
         Width           =   75
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Tgl Inbox"
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
         Left            =   240
         TabIndex        =   76
         Top             =   6960
         Width           =   975
      End
      Begin VB.Line Line25 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   5400
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   8040
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   2415
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         X1              =   720
         X2              =   4440
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   1440
         X2              =   4440
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         X1              =   2160
         X2              =   4440
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "INBOX"
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
         Left            =   4680
         TabIndex        =   38
         Top             =   120
         Width           =   6615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   32
         Top             =   6045
         Width           =   75
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   31
         Top             =   6045
         Width           =   660
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   29
         Top             =   5685
         Width           =   75
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4560
         TabIndex        =   28
         Top             =   5685
         Width           =   435
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   26
         Top             =   6045
         Width           =   75
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   6045
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   5685
         Width           =   75
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pengirim"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   5685
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   20
         Top             =   4695
         Width           =   75
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   4680
         Width           =   555
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   10560
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   975
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0FF&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11595
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00BFE8EA&
      Height          =   8175
      Index           =   3
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   11775
      Begin VB.TextBox txtIsi 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   2400
         MaxLength       =   999
         MultiLine       =   -1  'True
         TabIndex        =   109
         Top             =   5040
         Width           =   9075
      End
      Begin VB.CommandButton CmdXlsTemp 
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
         Left            =   9600
         Picture         =   "FormSMS.frx":11735
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   7200
         Width           =   735
      End
      Begin VB.TextBox txtType 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3000
         MaxLength       =   40
         TabIndex        =   9
         Top             =   4600
         Width           =   4575
      End
      Begin VB.TextBox txtId 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2400
         MaxLength       =   30
         TabIndex        =   8
         Top             =   4600
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add"
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
         Left            =   5640
         Picture         =   "FormSMS.frx":11BB1
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   7200
         Width           =   735
      End
      Begin VB.CommandButton cmdEdit 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit"
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
         Left            =   6360
         Picture         =   "FormSMS.frx":121CA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   7200
         Width           =   735
      End
      Begin VB.CommandButton cmdCreate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Crea&te"
         Height          =   735
         Left            =   7080
         Picture         =   "FormSMS.frx":1285C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   7200
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save"
         Height          =   735
         Left            =   7080
         Picture         =   "FormSMS.frx":12EA9
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   7200
         Width           =   735
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Cancel"
         Height          =   735
         Left            =   7800
         Picture         =   "FormSMS.frx":13575
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   7200
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
         Left            =   10680
         Picture         =   "FormSMS.frx":13BD4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   7200
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Delete"
         Height          =   735
         Left            =   8520
         Picture         =   "FormSMS.frx":14364
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   7200
         Width           =   735
      End
      Begin TrueOleDBGrid70.TDBGrid GridTemplate 
         Height          =   3975
         Left            =   0
         TabIndex        =   10
         Top             =   480
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7011
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No."
         Columns(0).DataField=   "Id"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Type SMS"
         Columns(1).DataField=   "TypeSMS"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Pesan"
         Columns(2).DataField=   "Pesan1"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=5292"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=5186"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=13467"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=13361"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
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
      Begin MSComDlg.CommonDialog ComDialogXls 
         Left            =   9720
         Top             =   6960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label49 
         BackStyle       =   0  'Transparent
         Caption         =   "#MD"
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
         Left            =   3120
         TabIndex        =   116
         Top             =   6720
         Width           =   615
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   ": Model Kendaraan"
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
         Index           =   2
         Left            =   3720
         TabIndex        =   115
         Top             =   6720
         Width           =   2415
      End
      Begin VB.Label Label48 
         BackStyle       =   0  'Transparent
         Caption         =   ": Total Bayar"
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
         Left            =   3720
         TabIndex        =   114
         Top             =   7440
         Width           =   1455
      End
      Begin VB.Label Label47 
         BackStyle       =   0  'Transparent
         Caption         =   "#RP"
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
         Left            =   3120
         TabIndex        =   113
         Top             =   7440
         Width           =   735
      End
      Begin VB.Label Label46 
         BackStyle       =   0  'Transparent
         Caption         =   "#KM"
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
         Left            =   0
         TabIndex        =   112
         Top             =   10000
         Width           =   615
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   ": KM Next Service"
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
         Left            =   720
         TabIndex        =   111
         Top             =   10000
         Width           =   2415
      End
      Begin VB.Label lbLen 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Char :"
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
         Left            =   10080
         TabIndex        =   90
         Top             =   4638
         Width           =   1335
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   9480
         Shape           =   4  'Rounded Rectangle
         Top             =   7080
         Width           =   975
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   ": KM Next Service"
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
         Index           =   0
         Left            =   3720
         TabIndex        =   74
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   ": Tanggal Next Service"
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
         Left            =   960
         TabIndex        =   73
         Top             =   7440
         Width           =   2415
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   ": Nomor Polisi"
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
         Left            =   960
         TabIndex        =   72
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Label Label29 
         BackStyle       =   0  'Transparent
         Caption         =   ": Alamat Customer"
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
         Left            =   960
         TabIndex        =   71
         Top             =   6960
         Width           =   2415
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   ": Nama Customer"
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
         Left            =   960
         TabIndex        =   70
         Top             =   6720
         Width           =   2415
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "#KM"
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
         Left            =   3120
         TabIndex        =   69
         Top             =   7200
         Width           =   615
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "#TG"
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
         Left            =   240
         TabIndex        =   68
         Top             =   7440
         Width           =   615
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "#NP"
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
         Left            =   240
         TabIndex        =   67
         Top             =   7200
         Width           =   615
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "#AL"
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
         Left            =   240
         TabIndex        =   66
         Top             =   6960
         Width           =   615
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "#NM"
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
         Left            =   240
         TabIndex        =   65
         Top             =   6720
         Width           =   615
      End
      Begin VB.Shape Shape13 
         Height          =   1455
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6600
         Width           =   5295
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000000FF&
         X1              =   720
         X2              =   4440
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   1440
         X2              =   4440
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   2160
         X2              =   4440
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "TEMPLATE SMS"
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
         Left            =   4560
         TabIndex        =   37
         Top             =   120
         Width           =   6735
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Isi Pesan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   5040
         Width           =   840
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Type Template SMS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   13
         Top             =   4668
         Width           =   1935
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   12
         Top             =   4668
         Width           =   75
      End
      Begin VB.Label Label54 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2280
         TabIndex        =   11
         Top             =   5055
         Width           =   75
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   5000
         Y2              =   5000
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   6480
         Y2              =   6480
      End
      Begin VB.Shape ShTrue 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   5520
         Shape           =   4  'Rounded Rectangle
         Top             =   7080
         Width           =   3855
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   10560
         Shape           =   4  'Rounded Rectangle
         Top             =   7080
         Width           =   975
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0FF&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11595
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00BFE8EA&
      Height          =   8175
      Index           =   1
      Left            =   360
      TabIndex        =   35
      Top             =   480
      Width           =   11775
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
         Left            =   8880
         Picture         =   "FormSMS.frx":149FB
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   6840
         Width           =   735
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
         Left            =   5640
         Picture         =   "FormSMS.frx":14E77
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   6840
         Width           =   735
      End
      Begin VB.CommandButton cmdRefreshSe 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   735
         Left            =   9600
         Picture         =   "FormSMS.frx":154C6
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   6840
         Width           =   735
      End
      Begin VB.CommandButton cmdExitSe 
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
         Left            =   10680
         Picture         =   "FormSMS.frx":15B2A
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   6840
         Width           =   735
      End
      Begin VB.TextBox txtSe1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   80
         TabIndex        =   46
         Top             =   4680
         Width           =   9975
      End
      Begin VB.TextBox txtTujuan 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   45
         Top             =   5400
         Width           =   2415
      End
      Begin VB.TextBox txtWaktuSe 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   44
         Top             =   5760
         Width           =   2415
      End
      Begin VB.TextBox txtUserSe 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   30
         TabIndex        =   43
         Top             =   6120
         Width           =   2415
      End
      Begin VB.TextBox txtSe2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1560
         MaxLength       =   80
         TabIndex        =   42
         Top             =   4995
         Width           =   9975
      End
      Begin TrueOleDBGrid70.TDBGrid GridSend 
         Height          =   3975
         Left            =   0
         TabIndex        =   48
         Top             =   480
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   7011
         _LayoutType     =   4
         _RowHeight      =   28
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No."
         Columns(0).DataField=   "id"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Pesan"
         Columns(1).DataField=   "pesan"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Tujuan"
         Columns(2).DataField=   "NoTujuan"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Waktu"
         Columns(3).DataField=   "waktu"
         Columns(3).ExternalEditor=   "tgl"
         Columns(3).ExternalEditor.vt=   8
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "User"
         Columns(4).DataField=   "milik"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1244"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1138"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=8652"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=8546"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3969"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3863"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=3519"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3413"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1667"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1561"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
      Begin VB.TextBox txtNoSe 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   11160
         TabIndex        =   59
         Top             =   480
         Width           =   375
      End
      Begin TDBDate6Ctl.TDBDate Dtt1 
         Height          =   315
         Left            =   1560
         TabIndex        =   81
         Top             =   6960
         Width           =   1695
         _Version        =   65536
         _ExtentX        =   2990
         _ExtentY        =   556
         Calendar        =   "FormSMS.frx":162BA
         Caption         =   "FormSMS.frx":163E6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FormSMS.frx":16452
         Keys            =   "FormSMS.frx":16470
         Spin            =   "FormSMS.frx":164CE
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
      Begin TDBDate6Ctl.TDBDate Dtt2 
         Height          =   315
         Left            =   3720
         TabIndex        =   82
         Top             =   6960
         Width           =   1575
         _Version        =   65536
         _ExtentX        =   2778
         _ExtentY        =   556
         Calendar        =   "FormSMS.frx":164F6
         Caption         =   "FormSMS.frx":16622
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "FormSMS.frx":1668E
         Keys            =   "FormSMS.frx":166AC
         Spin            =   "FormSMS.frx":1670A
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
      Begin VB.Line Line27 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   5400
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Label Label36 
         Caption         =   "Tgl Sent"
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
         Left            =   240
         TabIndex        =   84
         Top             =   6960
         Width           =   975
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1440
         TabIndex        =   83
         Top             =   6960
         Width           =   75
      End
      Begin VB.Line Line24 
         X1              =   3360
         X2              =   3600
         Y1              =   7080
         Y2              =   7080
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   5520
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pesan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   58
         Top             =   4680
         Width           =   555
      End
      Begin VB.Line Line18 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Shape Shape8 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   10560
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   56
         Top             =   4695
         Width           =   75
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tujuan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   55
         Top             =   5400
         Width           =   645
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   54
         Top             =   5400
         Width           =   75
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   53
         Top             =   5760
         Width           =   615
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   52
         Top             =   5760
         Width           =   75
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   51
         Top             =   6120
         Width           =   435
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   50
         Top             =   6120
         Width           =   75
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SENT ITEM"
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
         Left            =   4680
         TabIndex        =   49
         Top             =   120
         Width           =   6615
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         X1              =   2160
         X2              =   4440
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         X1              =   1440
         X2              =   4440
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000FF&
         X1              =   720
         X2              =   4440
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   8760
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0FF&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11595
      End
   End
   Begin VB.Frame fr 
      BackColor       =   &H00BFE8EA&
      Height          =   8175
      Index           =   2
      Left            =   480
      TabIndex        =   36
      Top             =   480
      Width           =   12255
      Begin VB.CommandButton cmdSMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pulsa"
         Height          =   735
         Left            =   360
         Picture         =   "FormSMS.frx":16732
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   6840
         Width           =   855
      End
      Begin VB.CommandButton CmdXlsStatistik 
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
         Left            =   8880
         Picture         =   "FormSMS.frx":16AB4
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   6840
         Width           =   735
      End
      Begin MSChart20Lib.MSChart MSChart 
         Height          =   3375
         Left            =   120
         OleObjectBlob   =   "FormSMS.frx":16F30
         TabIndex        =   64
         Top             =   3120
         Width           =   11415
      End
      Begin VB.CommandButton cmdExitSt 
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
         Left            =   10680
         Picture         =   "FormSMS.frx":18746
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   6840
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Refresh"
         Height          =   735
         Left            =   9600
         Picture         =   "FormSMS.frx":18ED6
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   6840
         Width           =   735
      End
      Begin TrueOleDBGrid70.TDBGrid GridSta 
         Height          =   2415
         Left            =   0
         TabIndex        =   62
         Top             =   480
         Width           =   11595
         _ExtentX        =   20452
         _ExtentY        =   4260
         _LayoutType     =   4
         _RowHeight      =   16
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Users"
         Columns(0).DataField=   "Users"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Total"
         Columns(1).DataField=   "Total"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Jan"
         Columns(2).DataField=   "01"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Feb"
         Columns(3).DataField=   "02"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Mar"
         Columns(4).DataField=   "03"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Apr"
         Columns(5).DataField=   "04"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Mei"
         Columns(6).DataField=   "05"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Jun"
         Columns(7).DataField=   "06"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Jul"
         Columns(8).DataField=   "07"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Ags"
         Columns(9).DataField=   "08"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Sep"
         Columns(10).DataField=   "09"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Okt"
         Columns(11).DataField=   "10"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Nov"
         Columns(12).DataField=   "11"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Des"
         Columns(13).DataField=   "12"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   14
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=14"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=1773"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=1667"
         Splits(0)._ColumnProps(8)=   "Column(1)._ColStyle=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1667"
         Splits(0)._ColumnProps(13)=   "Column(2)._ColStyle=1"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=1773"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=1667"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=1"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1667"
         Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=1"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(25)=   "Column(5).Width=1773"
         Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=1667"
         Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=1"
         Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(30)=   "Column(6).Width=1773"
         Splits(0)._ColumnProps(31)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(32)=   "Column(6)._WidthInPix=1667"
         Splits(0)._ColumnProps(33)=   "Column(6)._ColStyle=1"
         Splits(0)._ColumnProps(34)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(35)=   "Column(7).Width=1773"
         Splits(0)._ColumnProps(36)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(37)=   "Column(7)._WidthInPix=1667"
         Splits(0)._ColumnProps(38)=   "Column(7)._ColStyle=1"
         Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(40)=   "Column(8).Width=1773"
         Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=1667"
         Splits(0)._ColumnProps(43)=   "Column(8)._ColStyle=1"
         Splits(0)._ColumnProps(44)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(45)=   "Column(9).Width=1773"
         Splits(0)._ColumnProps(46)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(9)._WidthInPix=1667"
         Splits(0)._ColumnProps(48)=   "Column(9)._ColStyle=1"
         Splits(0)._ColumnProps(49)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(50)=   "Column(10).Width=1773"
         Splits(0)._ColumnProps(51)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(52)=   "Column(10)._WidthInPix=1667"
         Splits(0)._ColumnProps(53)=   "Column(10)._ColStyle=1"
         Splits(0)._ColumnProps(54)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(55)=   "Column(11).Width=1773"
         Splits(0)._ColumnProps(56)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(11)._WidthInPix=1667"
         Splits(0)._ColumnProps(58)=   "Column(11)._ColStyle=1"
         Splits(0)._ColumnProps(59)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(60)=   "Column(12).Width=1773"
         Splits(0)._ColumnProps(61)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(62)=   "Column(12)._WidthInPix=1667"
         Splits(0)._ColumnProps(63)=   "Column(12)._ColStyle=1"
         Splits(0)._ColumnProps(64)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(65)=   "Column(13).Width=1773"
         Splits(0)._ColumnProps(66)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(13)._WidthInPix=1667"
         Splits(0)._ColumnProps(68)=   "Column(13)._ColStyle=1"
         Splits(0)._ColumnProps(69)=   "Column(13).Order=14"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HD8E9EC&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=86,.parent=13,.alignment=2"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=83,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=84,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=85,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=82,.parent=13,.alignment=2"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=78,.parent=13,.alignment=2"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=75,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=76,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=77,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=71,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=72,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=73,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=67,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=68,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=69,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=63,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=64,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=65,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=2"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13,.alignment=2"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=2"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=50,.parent=13,.alignment=2"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=13,.alignment=2"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=29,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=30,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=31,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=28,.parent=13,.alignment=2"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=25,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=26,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=27,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=46,.parent=13,.alignment=2"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=43,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=44,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=45,.parent=17"
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
      Begin VB.Frame FrPulsa 
         BackColor       =   &H00BFE8EA&
         BorderStyle     =   0  'None
         Height          =   7815
         Left            =   0
         TabIndex        =   92
         Top             =   0
         Visible         =   0   'False
         Width           =   12015
         Begin VB.CommandButton cmdSMS2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pulsa"
            Height          =   735
            Left            =   360
            Picture         =   "FormSMS.frx":1953A
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   6840
            Width           =   855
         End
         Begin VB.CommandButton Command5 
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
            Left            =   10680
            Picture         =   "FormSMS.frx":198BC
            Style           =   1  'Graphical
            TabIndex        =   108
            Top             =   6840
            Width           =   735
         End
         Begin VB.Frame frPasca 
            BackColor       =   &H00BFE8EA&
            BorderStyle     =   0  'None
            Height          =   3255
            Left            =   240
            TabIndex        =   97
            Top             =   2160
            Width           =   7575
            Begin VB.CommandButton cmdEditP 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Edit"
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
               Left            =   3840
               Picture         =   "FormSMS.frx":1A04C
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   2280
               Width           =   735
            End
            Begin VB.CommandButton cmdSaveP 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Save"
               Enabled         =   0   'False
               Height          =   735
               Left            =   4560
               Picture         =   "FormSMS.frx":1A6DE
               Style           =   1  'Graphical
               TabIndex        =   105
               Top             =   2280
               Width           =   735
            End
            Begin VB.CommandButton cmdCancelP 
               BackColor       =   &H00FFFFFF&
               Caption         =   "&Cancel"
               Enabled         =   0   'False
               Height          =   735
               Left            =   5280
               Picture         =   "FormSMS.frx":1ADAA
               Style           =   1  'Graphical
               TabIndex        =   104
               Top             =   2280
               Width           =   735
            End
            Begin TDBNumber6Ctl.TDBNumber txtPul 
               Height          =   525
               Left            =   3720
               TabIndex        =   98
               Top             =   720
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
               _ExtentY        =   926
               Calculator      =   "FormSMS.frx":1B409
               Caption         =   "FormSMS.frx":1B429
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormSMS.frx":1B495
               Keys            =   "FormSMS.frx":1B4B3
               Spin            =   "FormSMS.frx":1B4FD
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   ","
               DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###,##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
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
               MaxValueVT      =   1397882885
               MinValueVT      =   1112342533
            End
            Begin TDBNumber6Ctl.TDBNumber txtTar 
               Height          =   525
               Left            =   3720
               TabIndex        =   101
               Top             =   1320
               Width           =   2415
               _Version        =   65536
               _ExtentX        =   4260
               _ExtentY        =   926
               Calculator      =   "FormSMS.frx":1B525
               Caption         =   "FormSMS.frx":1B545
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   18
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "FormSMS.frx":1B5B1
               Keys            =   "FormSMS.frx":1B5CF
               Spin            =   "FormSMS.frx":1B619
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   -2147483643
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   ","
               DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
               EditMode        =   0
               Enabled         =   0
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###,##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
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
               MaxValueVT      =   1397882885
               MinValueVT      =   1112342533
            End
            Begin VB.Line Line37 
               BorderColor     =   &H000000FF&
               X1              =   0
               X2              =   6360
               Y1              =   1920
               Y2              =   1920
            End
            Begin VB.Line Line36 
               BorderColor     =   &H000000FF&
               X1              =   0
               X2              =   6360
               Y1              =   600
               Y2              =   600
            End
            Begin VB.Label Label45 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Pulsa :"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   0
               TabIndex        =   107
               Top             =   200
               Width           =   1035
            End
            Begin VB.Line Line28 
               BorderColor     =   &H000000FF&
               X1              =   0
               X2              =   6360
               Y1              =   120
               Y2              =   120
            End
            Begin VB.Shape Shape20 
               BorderColor     =   &H00C0C0C0&
               FillColor       =   &H00FFFFFF&
               Height          =   975
               Left            =   3720
               Shape           =   4  'Rounded Rectangle
               Top             =   2160
               Width           =   2415
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3480
               TabIndex        =   103
               Top             =   1320
               Width           =   105
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   ":"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   3480
               TabIndex        =   102
               Top             =   720
               Width           =   105
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tarif SMS Tertinggi"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   240
               TabIndex        =   100
               Top             =   1320
               Width           =   2865
            End
            Begin VB.Label Label38 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Perkiraan Nilai Pulsa"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   240
               TabIndex        =   99
               Top             =   720
               Width           =   3075
            End
         End
         Begin VB.OptionButton OpPas 
            BackColor       =   &H00BFE8EA&
            Caption         =   "Pra Bayar *"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   96
            Top             =   1800
            Width           =   2655
         End
         Begin VB.OptionButton OpPra 
            BackColor       =   &H00BFE8EA&
            Caption         =   "Pasca Bayar "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   95
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Shape Shape21 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   240
            Shape           =   4  'Rounded Rectangle
            Top             =   6720
            Width           =   1095
         End
         Begin VB.Shape Shape22 
            BorderColor     =   &H00C0C0C0&
            FillColor       =   &H00FFFFFF&
            Height          =   975
            Left            =   10560
            Shape           =   4  'Rounded Rectangle
            Top             =   6720
            Width           =   975
         End
         Begin VB.Line Line35 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   3000
            Y1              =   720
            Y2              =   720
         End
         Begin VB.Line Line34 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   3000
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Line Line33 
            BorderColor     =   &H000000FF&
            X1              =   240
            X2              =   3000
            Y1              =   1200
            Y2              =   1200
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   94
            Top             =   800
            Width           =   840
         End
         Begin VB.Line Line32 
            BorderColor     =   &H000000FF&
            X1              =   0
            X2              =   11400
            Y1              =   6600
            Y2              =   6600
         End
         Begin VB.Line Line31 
            BorderColor     =   &H000000FF&
            X1              =   720
            X2              =   4440
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line30 
            BorderColor     =   &H000000FF&
            X1              =   1440
            X2              =   4440
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line29 
            BorderColor     =   &H000000FF&
            X1              =   2160
            X2              =   4440
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label37 
            BackStyle       =   0  'Transparent
            Caption         =   "COUNTER SMS "
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
            Left            =   4680
            TabIndex        =   93
            Top             =   120
            Width           =   6615
         End
         Begin VB.Shape Shape19 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00C0C0FF&
            Height          =   495
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   11595
         End
      End
      Begin VB.Shape Shape18 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Line Line23 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   10560
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "STATISTIK"
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
         Left            =   4680
         TabIndex        =   63
         Top             =   120
         Width           =   6615
      End
      Begin VB.Line Line22 
         BorderColor     =   &H000000FF&
         X1              =   2160
         X2              =   4440
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line21 
         BorderColor     =   &H000000FF&
         X1              =   1440
         X2              =   4440
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line20 
         BorderColor     =   &H000000FF&
         X1              =   720
         X2              =   4440
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line19 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   11520
         Y1              =   6600
         Y2              =   6600
      End
      Begin VB.Shape Shape10 
         BorderColor     =   &H00C0C0C0&
         FillColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   8760
         Shape           =   4  'Rounded Rectangle
         Top             =   6720
         Width           =   1695
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H00404040&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0FF&
         Height          =   495
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11595
      End
   End
   Begin VB.Shape Shape16 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   975
   End
   Begin VB.Image smsX1 
      Height          =   540
      Left            =   480
      Picture         =   "FormSMS.frx":1B641
      Top             =   0
      Width           =   2940
   End
   Begin VB.Image smsX3 
      Height          =   540
      Left            =   6240
      Picture         =   "FormSMS.frx":20DE6
      Top             =   0
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Image smsX2 
      Height          =   540
      Left            =   3360
      Picture         =   "FormSMS.frx":2668F
      Top             =   0
      Width           =   2940
   End
   Begin VB.Image smsX4 
      Height          =   540
      Left            =   9120
      Picture         =   "FormSMS.frx":2C166
      Top             =   0
      Width           =   2940
   End
   Begin VB.Image sms4 
      Height          =   540
      Left            =   9120
      Picture         =   "FormSMS.frx":31FDA
      Top             =   0
      Width           =   2940
   End
   Begin VB.Image sms2 
      Height          =   540
      Left            =   3360
      Picture         =   "FormSMS.frx":37E4E
      Top             =   0
      Width           =   2940
   End
   Begin VB.Image sms3 
      Height          =   540
      Left            =   6240
      Picture         =   "FormSMS.frx":3D925
      Top             =   0
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Image sms1 
      Height          =   540
      Left            =   480
      Picture         =   "FormSMS.frx":431CE
      Top             =   0
      Width           =   2940
   End
End
Attribute VB_Name = "FormSMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================
'Motion Captured Data Provided By Acep Suryanto
'==============================================
Option Explicit

Private OnLoad  As Boolean
Private isOwner As Boolean
Private objTem  As New ClsTemplate
Private rsTem   As ADODB.Recordset
Private objIn   As New ClsTemplate
Private rsIn    As ADODB.Recordset
Private objSe   As New ClsTemplate
Private rsSe    As ADODB.Recordset
Private objSt   As New ClsTemplate
Private rsSt    As ADODB.Recordset

Private rsTemp  As New ADODB.Recordset
Private objPu   As New ClsTemplate
Private rsPul   As ADODB.Recordset

Private NewId   As Integer
Private Const MsgTitle1     As String = "Add New Template SMS"
Private Const MsgTitle2     As String = "Update Template SMS"
Private IsRemainder         As String

Private col     As TrueOleDBGrid70.Column
Private cols    As TrueOleDBGrid70.Columns
Private Urut    As String
Private isAdd As Boolean

Public Property Let OnLoaded(value As Boolean)
   OnLoad = value
End Property

Public Property Let OnOwner(value As Boolean)
   isOwner = value
End Property

Public Property Get Remainder() As String
   Remainder = IsRemainder
End Property

Private Sub cmdExitForm_Click()
   Unload Me
End Sub

Private Sub cmdSearch2_Click()
   Dim Er As New ErrInfo
   If IsNull(Dtt1.value) = True Or IsNull(Dtt2.value) = True Then
        Er.HeadCap = "Send Item SMS"
        Er.InfoCap = "Range Tanggal Send Item Tidak Valid !"
        Beep
        Er.show (vbModal)
        Exit Sub
'      MsgBox "Range Tanggal Send Item Tidak Valid !", vbExclamation, "Send Item SMS"
'      Exit Sub
   End If
   Set rsSe = objSe.FetchDataSendItemRange(Dtt1.value, Dtt2.value)
   If rsSe.RecordCount = 0 Then
        Er.HeadCap = "Send Item SMS"
        Er.InfoCap = "Data Send Item tidak ada !"
        Beep
        Er.show (vbModal)
      'MsgBox "Data Send Item tidak ada !", vbInformation + vbOKOnly, "SMS"
      Call FillDGSendItem
   Else
      Set GridSend.DataSource = rsSe
      rsSe.MoveFirst
      Call BindToControlSend(rsSe)
   End If
End Sub

Private Sub CmdXls_Click() 'send item
On Error GoTo Err_Export
    Dim Er As New ErrInfo
    If rsSe.RecordCount = 0 Then
        Er.HeadCap = "Export To Excel !"
        Er.InfoCap = "List Send Item Kosong !"
        Beep
        Er.show (vbModal)
        Exit Sub
    End If
     Dim nmFile As String
     ComDialogXls.InitDir = App.Path
     ComDialogXls.FileName = nmFile
     ComDialogXls.CancelError = True
     ComDialogXls.filter = "Excel Files (*.xls)|*.xls"
     ComDialogXls.ShowSave
     nmFile = ComDialogXls.FileName
     Me.GridSend.ExportToFile nmFile, True
     
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & nmFile
            Beep
            Er.show (vbModal)
     'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"), vbInformation
Keluar:
End Sub

Private Sub CmdXlsStatistik_Click()
On Error GoTo Err_Export
   Dim Er As New ErrInfo
   If rsSt.RecordCount = 0 Then
        Er.HeadCap = "Export To Excel !"
        Er.InfoCap = "List Data Statistik Kosong !"
        Beep
        Er.show (vbModal)
        Exit Sub
    End If
     Dim nmFile As String
     ComDialogXls.InitDir = App.Path
     ComDialogXls.FileName = nmFile
     ComDialogXls.CancelError = True
     ComDialogXls.filter = "Excel Files (*.xls)|*.xls"
     ComDialogXls.ShowSave
     nmFile = ComDialogXls.FileName
     Me.GridSta.ExportToFile nmFile, True
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & nmFile
            
            Beep
            Er.show (vbModal)
     'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"), vbInformation
Keluar:
End Sub

Private Sub CmdXlsTemp_Click() 'template
On Error GoTo Err_Export
    Dim Er As New ErrInfo
    If rsTem.RecordCount = 0 Then
        Er.HeadCap = "Export To Excel !"
        Er.InfoCap = "List Data Template Kosong !"
        Beep
        Er.show (vbModal)
        Exit Sub
    End If
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
    
     GridTemplate.ExportToFile onVal, True
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            
            Beep
            Er.show (vbModal)
     'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub CmdXlxInbox_Click()
On Error GoTo Err_Export
    Dim Er As New ErrInfo
    If rsIn.RecordCount = 0 Then
        Er.HeadCap = "Export To Excel !"
        Er.InfoCap = "List Data Inbox Kosong !"
        Beep
        Er.show (vbModal)
        Exit Sub
    End If
     Dim nmFile As String
     ComDialogXls.InitDir = App.Path
     ComDialogXls.FileName = nmFile
     ComDialogXls.CancelError = True
     ComDialogXls.filter = "Excel Files (*.xls)|*.xls"
     ComDialogXls.ShowSave
     nmFile = ComDialogXls.FileName
     Me.GridInbox.ExportToFile nmFile, True
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & nmFile
            
            Beep
            Er.show (vbModal)
     'MsgBox "Data sudah di Export ke File Excel.", vbInformation, "Export File >> Excel"
     Exit Sub
Err_Export:
      If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"), vbInformation
Keluar:
End Sub

Private Sub Form_Load()
   isAdd = False
   If OnLoad = False Then Exit Sub
   fr(0).Visible = False
   fr(1).Visible = False
   fr(2).Visible = False
   fr(3).Visible = False
   Call SetPicSMS
   Call smsX4_Click
End Sub

Private Sub SetPicSMS()
   With Me
      .smsX1.Visible = False
      .smsX2.Visible = False
      .smsX3.Visible = False
      .smsX4.Visible = False
   End With
End Sub

Private Sub GridInbox_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsIn.Clone
    If Bookmark > 1 Then rs.Bookmark = Bookmark
    If Nul(rs.Fields(5).value) = "Belum Dibaca" Then RowStyle = "notSent"
End Sub

Private Sub sms1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetPicSMS
   smsX1.Visible = True
End Sub
Private Sub sms2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetPicSMS
   smsX2.Visible = True
End Sub
Private Sub sms3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetPicSMS
   smsX3.Visible = True
End Sub
Private Sub sms4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call SetPicSMS
   smsX4.Visible = True
End Sub


'====================================================================================================
'====================================================================================================

Private Sub smsX4_Click()
   fr(0).Visible = False
   fr(1).Visible = False
   fr(2).Visible = False
   fr(3).Visible = True
    FrPulsa.Visible = False
   Call FillDGTemplate
End Sub

Private Sub FillDGTemplate()
   Call SetOwner
   Set Me.GridTemplate.DataSource = Nothing
   Set rsTem = objTem.FetchDataTemplate
   If rsTem.RecordCount = 0 Then
      Call BindingNullFill
   Else
      Set GridTemplate.DataSource = rsTem
      rsTem.MoveFirst
      Call BindToControlTemplate(rsTem)
      Call OnLoadCommand
   End If
End Sub

Private Sub SetOwner()
   With Me
      .cmdAdd.Visible = True
      .cmdCancel.Visible = True
      .cmdCreate.Visible = True
      .cmdDelete.Visible = True
      .cmdEdit.Visible = True
      .cmdSave.Visible = True
   End With
End Sub

Private Sub BindToControlTemplate(rs As ADODB.Recordset)
   With Me
    Set .txtId.DataSource = rs
        txtId.DataField = rs.Fields(0).Name
    Set .txtType.DataSource = rs
        txtType.DataField = rs.Fields(1).Name
    Set .txtIsi.DataSource = rs
        txtIsi.DataField = rs.Fields(2).Name
   End With
End Sub

Private Sub cmdCancel_Click()
   Call FillDGTemplate
End Sub

Private Sub CmdAdd_Click()
    isAdd = False
   If rsTem.RecordCount = 0 Then
      NewId = 1
      Me.GridTemplate.Enabled = False
      isAdd = True
   Else
      rsTem.MoveLast
      NewId = Val(rsTem.Fields(0).value + 1)
      Me.GridTemplate.Enabled = False
      Me.GridTemplate.MoveFirst
      isAdd = True
   End If
      Call ClearFill
End Sub

Private Sub cmdCreate_Click()
   If txtType.text <> "" Then
      'If txtPesan1.Text <> "" Or txtPesan2.Text <> "" Or txtPesan3.Text <> "" Then
      If txtIsi.text <> "" Then
         If MsgBox("Simpan Template SMS baru ?", vbQuestion + vbYesNo, MsgTitle1) = vbYes Then
            Call objTem.AddNew_TemplateSMS(txtId.text, txtType.text, txtIsi.text)
            Call FillDGTemplate
         End If
         Exit Sub
      End If
   End If
        Dim Er As New ErrInfo
        Er.HeadCap = MsgTitle1
        Er.InfoCap = "Template SMS belum lengkap !"
        Beep
        Er.show (vbModal)
   'MsgBox "Template SMS belum lengkap !", vbInformation + vbOKOnly, MsgTitle1
End Sub

Private Sub cmdEdit_Click()
   Me.GridTemplate.Enabled = False
   Call EnabledFill
End Sub

Private Sub cmdSave_Click()
   If txtType.text <> "" Then
      'If txtPesan1.Text <> "" Or txtPesan2.Text <> "" Or txtPesan3.Text <> "" Then
      If txtIsi.text <> "" Then
         If MsgBox("Simpan perubahan data Template ?", vbQuestion + vbYesNo, MsgTitle2) = vbYes Then
            Call objTem.Update_TemplateSMS(txtId.text, txtType.text, txtIsi.text)
            Call FillDGTemplate
         End If
         Exit Sub
      End If
   End If
        Dim Er As New ErrInfo
        Er.HeadCap = MsgTitle1
        Er.InfoCap = "Template SMS belum lengkap !"
        Beep
        Er.show (vbModal)
   'MsgBox "Template SMS belum lengkap !", vbInformation + vbOKOnly, MsgTitle1
End Sub

Private Sub CmdDelete_Click()
   If MsgBox("Delete Item Template SMS (No." & txtId.text & ") ?", vbQuestion + vbYesNo, "Delete Item Template SMS") = vbYes Then
      Call objTem.Delete_Item_TemplateSMS(Me.txtId.text)
      Call FillDGTemplate
   End If
End Sub
Private Sub cmdExit_Click()
   Unload Me
End Sub


Private Sub GridTemplate_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If rsTem.RecordCount = 0 Then Exit Sub
   If isAdd Then Exit Sub
   Call BindToControlTemplate(rsTem)
End Sub


Private Sub OnLoadCommand()
   With Me
      .GridTemplate.Enabled = True
      .cmdAdd.Enabled = True
      .cmdEdit.Enabled = True
      .cmdCreate.Visible = False
      .cmdSave.Visible = True
      .cmdSave.Enabled = False
      .cmdCancel.Enabled = False
      .cmdDelete.Enabled = True
      .cmdExit.Enabled = True
      .txtType.Enabled = False
      .txtIsi.Enabled = False
'      .txtPesan1.Enabled = False
'      .txtPesan2.Enabled = False
'      .txtPesan3.Enabled = False
   End With
End Sub

Private Sub BindingNullFill()
With Me
      .txtId.text = ""
      .txtType.text = ""
      .txtIsi.text = ""
'      .txtPesan1.Text = ""
'      .txtPesan2.Text = ""
'      .txtPesan3.Text = ""
   .txtId.Enabled = False
   .txtType.Enabled = False
   .txtIsi.Enabled = False
'   .txtPesan1.Enabled = False
'   .txtPesan2.Enabled = False
'   .txtPesan3.Enabled = False
   .cmdAdd.Enabled = True
   .cmdEdit.Enabled = False
   .cmdCreate.Enabled = False
   .cmdSave.Enabled = False
   .cmdCancel.Enabled = False
   .cmdDelete.Enabled = False
   .cmdExit.Enabled = True
End With
End Sub


Private Sub ClearFill()
   With Me
      .cmdAdd.Enabled = False
      .cmdEdit.Enabled = False
      .cmdCreate.Enabled = True
      .cmdCreate.Visible = True
      .cmdSave.Visible = False
      .cmdCancel.Enabled = True
      .cmdDelete.Enabled = False
      .cmdExit.Enabled = False
      .txtId.text = NewId
      .txtType.text = ""
      .txtIsi.text = ""
'      .txtPesan1.Text = ""
'      .txtPesan2.Text = ""
'      .txtPesan3.Text = ""
      .txtType.Enabled = True
      .txtIsi.Enabled = True
'      .txtPesan1.Enabled = True
'      .txtPesan2.Enabled = True
'      .txtPesan3.Enabled = True
      .txtType.SetFocus
   End With
End Sub

Private Sub EnabledFill()
   With Me
      .cmdAdd.Enabled = False
      .cmdEdit.Enabled = False
      .cmdSave.Enabled = True
      .cmdSave.Visible = True
      .cmdCreate.Visible = False
      .cmdCancel.Enabled = True
      .cmdDelete.Enabled = False
      .cmdExit.Enabled = False
      .txtType.Enabled = True
      .txtIsi.Enabled = True
'      .txtPesan1.Enabled = True
'      .txtPesan2.Enabled = True
'      .txtPesan3.Enabled = True
      .txtType.SetFocus
   End With
End Sub



'Private Sub txtPesan1_Change()
'On Error Resume Next
'   If Len(txtPesan1.Text) > 60 Then txtPesan2.SetFocus
'   lbLen.Caption = "Char : " & (Len(txtPesan1.Text) + Len(txtPesan2.Text) + Len(txtPesan3.Text))
'End Sub
'Private Sub txtPesan2_Change()
'On Error Resume Next
'   If Len(txtPesan2.Text) > 60 Then txtPesan3.SetFocus
'   lbLen.Caption = "Char : " & (Len(txtPesan1.Text) + Len(txtPesan2.Text) + Len(txtPesan3.Text))
'End Sub
'Private Sub txtPesan3_Change()
'   lbLen.Caption = "Char : " & (Len(txtPesan1.Text) + Len(txtPesan2.Text) + Len(txtPesan3.Text))
'End Sub

Private Sub txtisi_Change()
   lbLen.Caption = "Char : " & (Len(txtIsi.text))
End Sub


'Private Sub txtPesan1_KeyPress(KeyAscii As Integer)
'   Call EnterPress(KeyAscii)
'End Sub
'Private Sub txtPesan2_KeyPress(KeyAscii As Integer)
'   Call EnterPress(KeyAscii)
'End Sub
'Private Sub txtPesan3_KeyPress(KeyAscii As Integer)
'   Call EnterPress(KeyAscii)
'End Sub

Private Sub txtType_KeyPress(KeyAscii As Integer)
   Call EnterPress(KeyAscii)
End Sub


'====================================================================================================
'====================================================================================================

Private Sub cmdExitIn_Click()
   Unload Me
End Sub

Private Sub smsX1_Click()
    fr(0).Visible = True
    fr(1).Visible = False
    fr(2).Visible = False
    fr(3).Visible = False
    FrPulsa.Visible = False
    Call FillDGInbox
End Sub

Private Sub FillDGInbox()

   Set Me.GridInbox.DataSource = Nothing
   Set rsIn = objIn.FetchDataInbox
   If rsIn.RecordCount = 0 Then
      Call BindingNullFillInbox
      Exit Sub
   Else
      Set GridInbox.DataSource = rsIn
      rsIn.MoveFirst
      Call BindToControlInbox(rsIn)
   End If
End Sub

Private Sub BindingNullFillInbox()
   With Me
      .txtIn1.text = ""
      .txtIn2.text = ""
      .txtIn3.text = ""
      .txtPengirim.text = ""
      .txtWaktu.text = ""
      .txtUser.text = ""
      .txtStatus.text = ""
   End With
End Sub

Private Sub BindToControlInbox(rs As ADODB.Recordset)
   With Me
    .txtNo.text = NullText(rs.Fields(0).value)
    .txtIn1.text = Left(NullText(rs.Fields(1).value), 75)
    .txtIn2.text = Mid(NullText(rs.Fields(1).value), 76, 75)
    .txtIn3.text = Mid(NullText(rs.Fields(1).value), 151, 75)
    .txtPengirim.text = Trim(NullText(rs.Fields(2).value))
    .txtWaktu.text = Trim(NullText(rs.Fields(3).value))
    .txtUser.text = Trim(NullText(rs.Fields(4).value))
    .txtStatus.text = Trim(NullText(rs.Fields(5).value))
    .txtIn1.Enabled = False
    .txtIn2.Enabled = False
    .txtIn3.Enabled = False
    .txtPengirim.Enabled = False
    .txtWaktu.Enabled = False
    .txtUser.Enabled = False
    .txtStatus.Enabled = False
   End With
   
End Sub

Private Sub CmdSearch_Click()
   Dim Er As New ErrInfo
   If IsNull(Dt1.value) = True Or IsNull(Dt2.value) = True Then
        Er.HeadCap = "Inbox SMS"
        Er.InfoCap = "Range Tanggal Inbox Tidak Valid !"
        Beep
        Er.show (vbModal)
        Exit Sub
      'MsgBox "Range Tanggal Inbox Tidak Valid !", vbExclamation, "Inbox SMS"
      'Exit Sub
   End If
   Set rsIn = objIn.FetchDataInboxRange(Dt1, Dt2)
   If rsIn.RecordCount = 0 Then
        Er.HeadCap = "Inbox SMS"
        Er.InfoCap = "Data Tidak Ada !"
        Beep
        Er.show (vbModal)
      'MsgBox "Data Tidak Ada !", vbInformation + vbOKOnly, "Inbox SMS"
      Call FillDGInbox
   Else
      Set GridInbox.DataSource = rsIn
      rsIn.MoveFirst
      Call BindToControlInbox(rsIn)
   End If
End Sub

Private Sub GridInbox_FilterChange()
On Error GoTo err
   Set cols = Me.GridInbox.Columns
   Dim C As Integer
   
   C = GridInbox.col
   GridInbox.HoldFields
   rsIn.filter = getFilter()
   If rsIn.EOF = True Then
      For Each col In GridInbox.Columns
         col.FilterText = ""
      Next col
      rsIn.filter = ""
   End If
   GridInbox.col = C
   GridInbox.EditActive = True
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
'         If col.DataField = "waktu" Then
'            If Len(col.FilterText) <> 10 Then Exit Function
'            tmp = tmp & col.DataField & " > '" & col.FilterText & "' And " & col.DataField & " < '" & CDate(CLng(CDate(col.FilterText)) + 1) & "'"
'         Else
'            tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
'         End If
          tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub GridInbox_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 5 Then rsIn.Sort = GridInbox.Columns(ColIndex).DataField & " " & Urut
4        GridInbox.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form FormSMS_Inbox")

End Sub

Private Sub GridInbox_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If txtStatus.Enabled = True Then Call FillDGInbox
   
   If rsIn.RecordCount <> 0 Then Call BindToControlInbox(rsIn)
End Sub

Private Sub cmdRead_Click()
    txtIn1.Enabled = True
    txtIn2.Enabled = True
    txtIn3.Enabled = True
    txtPengirim.Enabled = True
    txtWaktu.Enabled = True
    txtUser.Enabled = True
    txtStatus.Enabled = True
    txtStatus.text = "Sudah Dibaca"
    objIn.Update_ReadSMS (txtNo.text)
End Sub

Private Sub CmdRefreshLap_Click()
   Call FillDGInbox
End Sub

Private Sub txtIn1_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
Private Sub txtIn2_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
Private Sub txtIn3_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
Private Sub txtPengirim_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
Private Sub txtStatus_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
Private Sub txtWaktu_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
   KeyAscii = 0
End Sub


'====================================================================================================
'====================================================================================================

Private Sub cmdExitSe_Click()
   Unload Me
End Sub

Private Sub smsX2_Click()
   fr(0).Visible = False
   fr(1).Visible = True
   fr(2).Visible = False
   fr(3).Visible = False
    FrPulsa.Visible = False
   Call FillDGSendItem
End Sub

Private Sub FillDGSendItem()
   Set Me.GridSend.DataSource = Nothing
   Set rsSe = objSe.FetchDataSendItem
   If rsSe.RecordCount = 0 Then
      Call BindingNullFillSendItem
      Exit Sub
   Else
      Set GridSend.DataSource = rsSe
      rsSe.MoveFirst
      Call BindToControlSend(rsSe)
   End If
End Sub

Private Sub BindingNullFillSendItem()
   With Me
      .txtSe1.text = ""
      .txtSe2.text = ""
      .txtTujuan.text = ""
      .txtWaktuSe.text = ""
      .txtUserSe.text = ""
   End With
End Sub

Private Sub BindToControlSend(rs As ADODB.Recordset)
   With Me
    .txtNoSe.text = NullText(rs.Fields(0).value)
    .txtSe1.text = Left(NullText(rs.Fields(1).value), 80)
    .txtSe2.text = Mid(NullText(rs.Fields(1).value), 81, 80)
    .txtTujuan.text = Trim(NullText(rs.Fields(2).value))
    .txtWaktuSe.text = Trim(NullText(rs.Fields(3).value))
    .txtUserSe.text = Trim(NullText(rs.Fields(4).value))
   End With
End Sub

Private Sub GridSend_FilterChange()
On Error GoTo err
   Set cols = Me.GridSend.Columns
   Dim C As Integer
   
   C = GridSend.col
   GridSend.HoldFields
   rsSe.filter = getFilter()
   If rsSe.EOF = True Then
      For Each col In GridSend.Columns
         col.FilterText = ""
      Next col
      rsSe.filter = ""
   End If
   GridSend.col = C
   GridSend.EditActive = True
   Exit Sub
err:
   For Each col In cols
      col.FilterText = ""
   Next col
   Resume Next
End Sub

Private Sub GridSend_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 4 Then rsSe.Sort = GridSend.Columns(ColIndex).DataField & " " & Urut
4        GridSend.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form FormSMS_send")
End Sub

Private Sub GridSend_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   If rsSe.RecordCount <> 0 Then Call BindToControlSend(rsSe)
End Sub

Private Sub cmdRefreshSe_Click()
   Call FillDGSendItem
End Sub


'===================================================================================================
'====================================================================================================

Private Sub smsX3_Click()
   fr(0).Visible = False
   fr(1).Visible = False
   fr(2).Visible = True
   fr(3).Visible = False
    FrPulsa.Visible = False
   Call FillDGStatistik
   'Call FillDGStatistikGrafik
End Sub

Private Sub FillDGStatistik()
   CreateTblSt True
   Set Me.GridSta.DataSource = Nothing
   Set rsSt = objSt.FetchDataStatistik
   If rsSt.RecordCount = 0 Then
      Exit Sub
   Else
      Set GridSta.DataSource = rsSt
            rsSt.MoveFirst
            Do Until rsSt.EOF
                rsTemp.AddNew
                rsTemp!Jan = IIf(IsNull(rsSt.Fields(2).value) = True, 0, rsSt.Fields(2).value)
                rsTemp!Feb = IIf(IsNull(rsSt.Fields(3).value) = True, 0, rsSt.Fields(3).value)
                rsTemp!Mar = IIf(IsNull(rsSt.Fields(4).value) = True, 0, rsSt.Fields(4).value)
                rsTemp!Apr = IIf(IsNull(rsSt.Fields(5).value) = True, 0, rsSt.Fields(5).value)
                rsTemp!Mei = IIf(IsNull(rsSt.Fields(6).value) = True, 0, rsSt.Fields(6).value)
                rsTemp!Jun = IIf(IsNull(rsSt.Fields(7).value) = True, 0, rsSt.Fields(7).value)
                rsTemp!Jul = IIf(IsNull(rsSt.Fields(8).value) = True, 0, rsSt.Fields(8).value)
                rsTemp!Ags = IIf(IsNull(rsSt.Fields(9).value) = True, 0, rsSt.Fields(9).value)
                rsTemp!Sep = IIf(IsNull(rsSt.Fields(10).value) = True, 0, rsSt.Fields(10).value)
                rsTemp!Okt = IIf(IsNull(rsSt.Fields(11).value) = True, 0, rsSt.Fields(11).value)
                rsTemp!Nov = IIf(IsNull(rsSt.Fields(12).value) = True, 0, rsSt.Fields(12).value)
                rsTemp!Des = IIf(IsNull(rsSt.Fields(13).value) = True, 0, rsSt.Fields(13).value)

                rsTemp.Update
                rsSt.MoveNext
            Loop
        rsSt.MoveFirst
   End If
End Sub

Private Sub FillDGStatistikGrafik()

   Set Me.MSChart.DataSource = Nothing
   Set rsSt = objSt.FetchDataStatistikGrafik
   If rsSt.RecordCount = 0 Then
      Exit Sub
   Else
      MSChart.ShowLegend = True
      Set MSChart.DataSource = rsSt
   End If
End Sub

Private Sub cmdExitSt_Click()
   Unload Me
End Sub

Private Sub CreateTblSt(Optional Edit As Boolean = False)
1         On Error GoTo vb_error
2         If rsTemp.State Then rsTemp.Close
3         Set rsTemp = Nothing
4         With rsTemp
5             .Fields.Append "Jan", adDouble, 9
6             .Fields.Append "Feb", adDouble, 9
7             .Fields.Append "Mar", adDouble, 9
8             .Fields.Append "Apr", adDouble, 9
9             .Fields.Append "Mei", adDouble, 9
10            .Fields.Append "Jun", adDouble, 9
11            .Fields.Append "Jul", adDouble, 9
12            .Fields.Append "Ags", adDouble, 9
13            .Fields.Append "Sep", adDouble, 9
14            .Fields.Append "Okt", adDouble, 9
15            .Fields.Append "Nov", adDouble, 9
16            .Fields.Append "Des", adDouble, 9
17            .Open
18            Set MSChart.DataSource = rsTemp
19        End With
20        Exit Sub
vb_error:
21        err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub cmdSMS_Click()
    FrPulsa.Visible = True
    Call ViewPulsa
End Sub

Private Sub cmdSMS2_Click()
    FrPulsa.Visible = False
End Sub

Private Sub ViewPulsa()
    Set rsPul = objPu.FetchPulsa
    If rsPul.RecordCount <> 0 Then
        OpPas.value = rsPul.Fields(0).value
        txtPul = IIf(Val(rsPul.Fields(1).value) < 0, 0, rsPul.Fields(1).value)
        txtTar = rsPul.Fields(2).value
    End If
End Sub

Private Sub OpPas_Click()
    Call SetFrPasca_Visible
End Sub
Private Sub OpPra_Click()
    Call SetFrPasca_Visible
    Call cmdCancelP_Click
End Sub
Private Sub SetFrPasca_Visible()
    If OpPas.value = True Then
        frPasca.Visible = True
        Call objPu.Update_TipePulsa(1)
    Else
        frPasca.Visible = False
        Call objPu.Update_TipePulsa(0)
    End If
End Sub

Private Sub cmdEditP_Click()
    With Me
        .cmdEditP.Enabled = False
        .cmdCancelP.Enabled = True
        .cmdSaveP.Enabled = True
        .txtPul.Enabled = True
        .txtTar.Enabled = True
    End With
End Sub

Private Sub cmdSaveP_Click()
    If MsgBox("Simpan Nilai Perkiraan Pulsa ?", vbQuestion + vbYesNo) = vbYes Then
        Call objPu.Update_NilaiPulsa(Val(txtPul), Val(txtTar))
        With Me
            .cmdEditP.Enabled = True
            .cmdCancelP.Enabled = False
            .cmdSaveP.Enabled = False
            .txtPul.Enabled = False
            .txtTar.Enabled = False
        End With
    End If
End Sub

Private Sub cmdCancelP_Click()
    Call ViewPulsa
    With Me
        .cmdEditP.Enabled = True
        .cmdCancelP.Enabled = False
        .cmdSaveP.Enabled = False
        .txtPul.Enabled = False
        .txtTar.Enabled = False
    End With
End Sub

