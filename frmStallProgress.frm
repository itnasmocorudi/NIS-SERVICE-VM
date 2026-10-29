VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Begin VB.Form frmStallProgress 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CONTROL BOARD"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "frmStallProgress.frx":0000
   ScaleHeight     =   9510
   ScaleWidth      =   19080
   Begin TrueOleDBGrid70.TDBGrid listSA 
      Height          =   2415
      Left            =   15960
      TabIndex        =   94
      Top             =   6480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4260
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NAME"
      Columns(0).DataField=   "SA"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14933984
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4419"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4313"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      ColumnHeaders   =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "SA"
      MultipleLines   =   0
      CellTipsWidth   =   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Trebuchet MS"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Trebuchet MS"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Trebuchet MS"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1,.bold=0,.fontsize=900"
      _StyleDefs(37)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=28,.fontname=Chaparral Pro"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Keluar"
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
      Left            =   17640
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9120
      Width           =   1200
   End
   Begin VB.Frame legend 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   375
      Left            =   2400
      TabIndex        =   82
      Top             =   9120
      Width           =   16575
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00C000C0&
         Caption         =   "H - Inspection"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   11100
         TabIndex        =   92
         Top             =   45
         Width           =   1470
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H004D9AE8&
         Caption         =   "B - Booking"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   91
         Top             =   40
         Width           =   1185
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFF80&
         Caption         =   "W - Waiting"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   1320
         TabIndex        =   90
         Top             =   40
         Width           =   1215
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "P - Proses"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   4200
         TabIndex        =   89
         Top             =   40
         Width           =   1035
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         Caption         =   "S - Stopped"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   300
         Left            =   5280
         TabIndex        =   88
         Top             =   40
         Width           =   1230
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H0000FF00&
         Caption         =   "F - Finish"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Left            =   6600
         TabIndex        =   87
         Top             =   40
         Width           =   990
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00DAFEA5&
         Caption         =   "I - Invoicing"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   7680
         TabIndex        =   86
         Top             =   40
         Width           =   1245
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00EAF9FD&
         Caption         =   "C - Closing"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   9000
         TabIndex        =   85
         Top             =   40
         Width           =   1095
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         Caption         =   "- Actual"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   10200
         TabIndex        =   84
         Top             =   40
         Width           =   840
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H0091910D&
         Caption         =   "w - Book Wait"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2640
         TabIndex        =   83
         Top             =   40
         Width           =   1455
      End
   End
   Begin TrueOleDBGrid70.TDBGrid listMech 
      Height          =   6375
      Left            =   15960
      TabIndex        =   66
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   11245
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NAME"
      Columns(0).DataField=   "NAME"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14933984
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=4419"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=4313"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
      PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      ColumnHeaders   =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "IDLE MECHANIC"
      MultipleLines   =   0
      CellTipsWidth   =   0
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Trebuchet MS"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Trebuchet MS"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=975,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Trebuchet MS"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1,.bold=0,.fontsize=900"
      _StyleDefs(37)  =   ":id=28,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(38)  =   ":id=28,.fontname=Chaparral Pro"
      _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(42)  =   "Named:id=33:Normal"
      _StyleDefs(43)  =   ":id=33,.parent=0"
      _StyleDefs(44)  =   "Named:id=34:Heading"
      _StyleDefs(45)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(46)  =   ":id=34,.wraptext=-1"
      _StyleDefs(47)  =   "Named:id=35:Footing"
      _StyleDefs(48)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(49)  =   "Named:id=36:Selected"
      _StyleDefs(50)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(51)  =   "Named:id=37:Caption"
      _StyleDefs(52)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(53)  =   "Named:id=38:HighlightRow"
      _StyleDefs(54)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(55)  =   "Named:id=39:EvenRow"
      _StyleDefs(56)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(57)  =   "Named:id=40:OddRow"
      _StyleDefs(58)  =   ":id=40,.parent=33"
      _StyleDefs(59)  =   "Named:id=41:RecordSelector"
      _StyleDefs(60)  =   ":id=41,.parent=34"
      _StyleDefs(61)  =   "Named:id=42:FilterBar"
      _StyleDefs(62)  =   ":id=42,.parent=33"
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Atur &Job Kemarin"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   9120
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   30000
      Left            =   600
      Top             =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   15495
      Begin VB.Frame frmeStall 
         BackColor       =   &H00C0FFFF&
         Height          =   7935
         Left            =   2280
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   9255
         Begin AIFCmp1.asxPowerBanner StopWash 
            Height          =   375
            Left            =   6480
            Top             =   7320
            Visible         =   0   'False
            Width           =   1550
            _ExtentX        =   2725
            _ExtentY        =   661
            StartColor      =   33023
            EndColor        =   8438015
            FormatString    =   "Stop Washing"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin AIFCmp1.asxPowerBanner StartWash 
            Height          =   375
            Left            =   6480
            Top             =   6960
            Visible         =   0   'False
            Width           =   1550
            _ExtentX        =   2725
            _ExtentY        =   661
            StartColor      =   33023
            EndColor        =   8438015
            FormatString    =   "Start Washing"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin AIFCmp1.asxPowerBanner StopInsp 
            Height          =   375
            Left            =   6480
            Top             =   6600
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            StartColor      =   33023
            EndColor        =   8438015
            FormatString    =   "Stop Inspection"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin AIFCmp1.asxPowerBanner StartInsp 
            Height          =   375
            Left            =   6480
            Top             =   6240
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   661
            StartColor      =   33023
            EndColor        =   8438015
            FormatString    =   "Start Inspection"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin VB.CheckBox chSBNP 
            BackColor       =   &H00C0FFFF&
            Caption         =   "SB NP"
            Height          =   255
            Left            =   2760
            TabIndex        =   74
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chEM 
            BackColor       =   &H00C0FFFF&
            Caption         =   "EM"
            Height          =   255
            Left            =   3720
            TabIndex        =   68
            Top             =   360
            Width           =   615
         End
         Begin TrueOleDBGrid70.TDBGrid gridJobOrder 
            Height          =   2055
            Left            =   120
            TabIndex        =   21
            Top             =   1680
            Width           =   9015
            _ExtentX        =   15901
            _ExtentY        =   3625
            _LayoutType     =   0
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
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "JOB ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
         Begin VB.Frame frmeMekanik 
            BackColor       =   &H00C0FFFF&
            Caption         =   "   Ganti Mekanik   "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   120
            TabIndex        =   44
            Top             =   5280
            Width           =   9015
            Begin VB.ComboBox cboStall 
               Height          =   315
               Left            =   4080
               Style           =   2  'Dropdown List
               TabIndex        =   52
               Top             =   345
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CommandButton GantiOK 
               BackColor       =   &H00DCFCFB&
               Caption         =   "&Proses"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7200
               Style           =   1  'Graphical
               TabIndex        =   47
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton GantiBatal 
               BackColor       =   &H00DCFCFB&
               Caption         =   "&Batal"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7200
               Style           =   1  'Graphical
               TabIndex        =   50
               Top             =   480
               Width           =   1335
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   14
               Left            =   840
               TabIndex        =   45
               Top             =   360
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   15
               Left            =   1800
               TabIndex        =   49
               Top             =   360
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin VB.Frame frmeMechEM 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1815
               Left            =   120
               TabIndex        =   67
               Top             =   600
               Visible         =   0   'False
               Width           =   6495
               Begin MSFlexGridLib.MSFlexGrid gridEM 
                  Height          =   1785
                  Left            =   0
                  TabIndex        =   46
                  Top             =   120
                  Width           =   5865
                  _ExtentX        =   10345
                  _ExtentY        =   3149
                  _Version        =   393216
                  Cols            =   3
                  FixedCols       =   0
                  HighLight       =   2
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Stall Baru"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   3000
               TabIndex        =   53
               Top             =   382
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.Shape Shape1 
               BackStyle       =   1  'Opaque
               Height          =   270
               Left            =   1635
               Top             =   360
               Width           =   60
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mekanik"
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
               Left            =   120
               TabIndex        =   48
               Top             =   405
               Width           =   570
            End
         End
         Begin VB.CommandButton cmdTutup 
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   50
            TabIndex        =   40
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton toDown 
            Caption         =   "v"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   33
            Top             =   2670
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton toRight 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   32
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton toLeft 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   240
            TabIndex        =   31
            Top             =   2400
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.CommandButton toUp 
            Caption         =   "^"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   12
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   360
            TabIndex        =   30
            Top             =   2160
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Frame frmePause 
            BackColor       =   &H00C0FFFF&
            Caption         =   "   Validasi Pause   "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   22
            Top             =   4320
            Width           =   9015
            Begin VB.OptionButton opPause 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Final Check"
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
               Index           =   4
               Left            =   3000
               TabIndex        =   81
               Top             =   240
               Width           =   1215
            End
            Begin VB.OptionButton opPause 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Lainnya"
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
               Index           =   3
               Left            =   4320
               TabIndex        =   38
               Top             =   240
               Value           =   -1  'True
               Width           =   855
            End
            Begin VB.OptionButton opPause 
               BackColor       =   &H00C0FFFF&
               Caption         =   "Add Job"
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
               Index           =   2
               Left            =   1800
               TabIndex        =   37
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton opPause 
               BackColor       =   &H00C0FFFF&
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
               Height          =   195
               Index           =   1
               Left            =   960
               TabIndex        =   36
               Top             =   240
               Width           =   855
            End
            Begin VB.CommandButton cmdBtl 
               BackColor       =   &H00DCFCFB&
               Caption         =   "&Batal"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7200
               Style           =   1  'Graphical
               TabIndex        =   43
               Top             =   600
               Width           =   1335
            End
            Begin VB.CommandButton cmdPause 
               BackColor       =   &H00DCFCFB&
               Caption         =   "&Proses"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7200
               Style           =   1  'Graphical
               TabIndex        =   41
               Top             =   240
               Width           =   1335
            End
            Begin VB.OptionButton opPause 
               BackColor       =   &H00C0FFFF&
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
               Height          =   195
               Index           =   0
               Left            =   240
               TabIndex        =   23
               Top             =   240
               Width           =   855
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   9
               Left            =   240
               TabIndex        =   39
               Top             =   495
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   40
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
         End
         Begin AIFCmp1.asxPowerBanner toStart 
            Height          =   375
            Left            =   120
            Top             =   3840
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            StartColor      =   33023
            EndColor        =   8438015
            FormatString    =   " Start"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   0
            Left            =   1410
            TabIndex        =   5
            Top             =   120
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   1
            Left            =   1410
            TabIndex        =   7
            Top             =   405
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   2
            Left            =   1410
            TabIndex        =   9
            Top             =   705
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   3
            Left            =   6180
            TabIndex        =   11
            Top             =   120
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   6180
            TabIndex        =   13
            Top             =   405
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   5
            Left            =   8460
            TabIndex        =   15
            Top             =   405
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   6
            Left            =   6180
            TabIndex        =   17
            Top             =   720
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99-99-9999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   7
            Left            =   7365
            TabIndex        =   19
            Top             =   705
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   8
            Left            =   8100
            TabIndex        =   20
            Top             =   705
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCC"
            PromptChar      =   ">"
         End
         Begin AIFCmp1.asxPowerBanner toMove 
            Height          =   375
            Left            =   2760
            Top             =   3840
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   661
            StartColor      =   33023
            EndColor        =   8438015
            FormatString    =   " Pindah Stall"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin AIFCmp1.asxPowerBanner toClose 
            Height          =   375
            Left            =   8040
            Top             =   3840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            StartColor      =   8438015
            EndColor        =   33023
            FormatString    =   " Tutup"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin AIFCmp1.asxPowerBanner toFinish 
            Height          =   375
            Left            =   1200
            Top             =   3840
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            StartColor      =   33023
            EndColor        =   8438015
            FormatString    =   " Selesai"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   10
            Left            =   1410
            TabIndex        =   25
            Top             =   1005
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   3165
            TabIndex        =   27
            Top             =   1005
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   25
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   7200
            TabIndex        =   29
            Top             =   1005
            Width           =   1930
            _ExtentX        =   3413
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   13
            Left            =   7980
            TabIndex        =   34
            Top             =   120
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin AIFCmp1.asxPowerBanner gantiMech 
            Height          =   375
            Left            =   4920
            Top             =   3840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   661
            StartColor      =   33023
            EndColor        =   8438015
            FormatString    =   "Ganti Mekanik"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   64
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   16
            Left            =   3720
            TabIndex        =   51
            Top             =   2280
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   17
            Left            =   1470
            TabIndex        =   54
            Top             =   2040
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   18
            Left            =   2610
            TabIndex        =   56
            Top             =   720
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   19
            Left            =   1410
            TabIndex        =   70
            Top             =   1320
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   40
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   20
            Left            =   8460
            TabIndex        =   71
            Top             =   1320
            Width           =   675
            _ExtentX        =   1191
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99:99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   21
            Left            =   7200
            TabIndex        =   73
            Top             =   1320
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   503
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "99-99-9999"
            PromptChar      =   ">"
         End
         Begin AIFCmp1.asxPowerBanner cmdHistory 
            Height          =   375
            Left            =   6840
            Top             =   3840
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            StartColor      =   12582912
            EndColor        =   16777088
            FormatString    =   "HISTORY"
            Orientation     =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   16777215
         End
         Begin VB.Label sgnPre 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pre Diagnosis"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   7680
            TabIndex        =   93
            Top             =   1560
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.Shape shpBook 
            BorderWidth     =   2
            Height          =   495
            Left            =   4460
            Top             =   160
            Width           =   495
         End
         Begin VB.Label lblBook 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   570
            Left            =   4560
            TabIndex        =   76
            Top             =   100
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mekanik"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   405
            Left            =   3840
            TabIndex        =   75
            Top             =   600
            Width           =   1110
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Penyerahan"
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
            Left            =   6180
            TabIndex        =   72
            Top             =   1365
            Width           =   870
         End
         Begin VB.Label Label24 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alasan Stopped"
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
            Left            =   135
            TabIndex        =   69
            Top             =   1365
            Width           =   1125
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SA"
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
            Left            =   2280
            TabIndex        =   57
            Top             =   720
            Width           =   195
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stall"
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
            Left            =   1020
            TabIndex        =   55
            Top             =   2085
            Width           =   300
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. WO"
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
            Left            =   7320
            TabIndex        =   35
            Top             =   165
            Width           =   570
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Warna"
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
            Left            =   6570
            TabIndex        =   28
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Model Kend."
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
            Left            =   375
            TabIndex        =   26
            Top             =   1050
            Width           =   885
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Stopped"
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
            Left            =   5415
            TabIndex        =   18
            Top             =   750
            Width           =   600
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai"
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
            Left            =   7320
            TabIndex        =   16
            Top             =   450
            Width           =   825
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai"
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
            Left            =   5340
            TabIndex        =   14
            Top             =   450
            Width           =   690
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actual Rate"
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
            Left            =   5190
            TabIndex        =   12
            Top             =   165
            Width           =   840
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mekanik"
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
            Left            =   690
            TabIndex        =   10
            Top             =   750
            Width           =   570
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Polisi"
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
            Left            =   615
            TabIndex        =   8
            Top             =   450
            Width           =   645
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. Rangka"
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
            Left            =   420
            TabIndex        =   6
            Top             =   165
            Width           =   840
         End
      End
      Begin VB.Frame frameHistory 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   480
         TabIndex        =   77
         Top             =   2160
         Visible         =   0   'False
         Width           =   10815
         Begin VB.CommandButton tblOther 
            BackColor       =   &H00DCFCFB&
            Caption         =   "&Other"
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
            Left            =   6960
            Style           =   1  'Graphical
            TabIndex        =   80
            Top             =   4560
            Width           =   1815
         End
         Begin VB.CommandButton tblTutup 
            BackColor       =   &H00DCFCFB&
            Caption         =   "Tutup"
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
            Left            =   8880
            Style           =   1  'Graphical
            TabIndex        =   78
            Top             =   4560
            Width           =   1815
         End
         Begin TrueOleDBGrid70.TDBGrid HistoryGrid 
            Height          =   4215
            Left            =   120
            TabIndex        =   79
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   7435
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
            Columns.Count   =   14
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=14"
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
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
            PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Trebuchet MS"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Trebuchet MS"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=Trebuchet MS"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HF1E4BE&"
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
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=62,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=13"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=90,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=87,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=88,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=89,.parent=17"
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
      End
      Begin VB.Frame framSearch 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   1200
         TabIndex        =   62
         Top             =   2400
         Visible         =   0   'False
         Width           =   10815
         Begin VB.CommandButton searchClose 
            BackColor       =   &H00DCFCFB&
            Caption         =   "Tutup"
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
            Left            =   8880
            Style           =   1  'Graphical
            TabIndex        =   63
            Top             =   4560
            Width           =   1815
         End
         Begin TrueOleDBGrid70.TDBGrid gridSearch 
            Height          =   4215
            Left            =   120
            TabIndex        =   64
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   7435
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
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14215660
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
            PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Trebuchet MS"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(12)  =   ":id=2,.fontname=Trebuchet MS"
            _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(15)  =   ":id=3,.fontname=Trebuchet MS"
            _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HF1E4BE&"
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
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
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
      End
      Begin VB.Frame frmePending 
         BackColor       =   &H00808080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5055
         Left            =   1800
         TabIndex        =   58
         Top             =   3000
         Visible         =   0   'False
         Width           =   10815
         Begin VB.CommandButton cmdTutupWait 
            BackColor       =   &H00DCFCFB&
            Caption         =   "Tutup"
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
            Left            =   8880
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   4560
            Width           =   1815
         End
         Begin TrueOleDBGrid70.TDBGrid pendingGrid 
            Height          =   4215
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   10575
            _ExtentX        =   18653
            _ExtentY        =   7435
            _LayoutType     =   0
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
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "WAITING JOB"
            MultipleLines   =   0
            CellTipsWidth   =   0
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
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
      End
      Begin VB.CommandButton cmdRepair 
         BackColor       =   &H00DCFCFB&
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   15000
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton cmdSearch 
         BackColor       =   &H00DCFCFB&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   120
         Width           =   495
      End
      Begin VB.Frame F4 
         Height          =   1215
         Left            =   8880
         TabIndex        =   24
         Top             =   4200
         Visible         =   0   'False
         Width           =   2175
      End
      Begin TrueDBGrid70.TDBGrid gridStall 
         Height          =   8415
         Left            =   0
         TabIndex        =   3
         Top             =   480
         Width           =   15300
         _ExtentX        =   26988
         _ExtentY        =   14843
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
         Columns(33)._VlistStyle=   0
         Columns(33)._MaxComboItems=   5
         Columns(33).DataField=   ""
         Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(34)._VlistStyle=   0
         Columns(34)._MaxComboItems=   5
         Columns(34).DataField=   ""
         Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(35)._VlistStyle=   0
         Columns(35)._MaxComboItems=   5
         Columns(35).DataField=   ""
         Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(36)._VlistStyle=   0
         Columns(36)._MaxComboItems=   5
         Columns(36).DataField=   ""
         Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(37)._VlistStyle=   0
         Columns(37)._MaxComboItems=   5
         Columns(37).DataField=   ""
         Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(38)._VlistStyle=   0
         Columns(38)._MaxComboItems=   5
         Columns(38).DataField=   ""
         Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(39)._VlistStyle=   0
         Columns(39)._MaxComboItems=   5
         Columns(39).DataField=   ""
         Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(40)._VlistStyle=   0
         Columns(40)._MaxComboItems=   5
         Columns(40).DataField=   ""
         Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(41)._VlistStyle=   0
         Columns(41)._MaxComboItems=   5
         Columns(41).DataField=   ""
         Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(42)._VlistStyle=   0
         Columns(42)._MaxComboItems=   5
         Columns(42).DataField=   ""
         Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(43)._VlistStyle=   0
         Columns(43)._MaxComboItems=   5
         Columns(43).DataField=   ""
         Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(44)._VlistStyle=   0
         Columns(44)._MaxComboItems=   5
         Columns(44).DataField=   ""
         Columns(44)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(45)._VlistStyle=   0
         Columns(45)._MaxComboItems=   5
         Columns(45).DataField=   ""
         Columns(45)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(46)._VlistStyle=   0
         Columns(46)._MaxComboItems=   5
         Columns(46).DataField=   ""
         Columns(46)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(47)._VlistStyle=   0
         Columns(47)._MaxComboItems=   5
         Columns(47).DataField=   ""
         Columns(47)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(48)._VlistStyle=   0
         Columns(48)._MaxComboItems=   5
         Columns(48).DataField=   ""
         Columns(48)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(49)._VlistStyle=   0
         Columns(49)._MaxComboItems=   5
         Columns(49).DataField=   ""
         Columns(49)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(50)._VlistStyle=   0
         Columns(50)._MaxComboItems=   5
         Columns(50).DataField=   ""
         Columns(50)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(51)._VlistStyle=   0
         Columns(51)._MaxComboItems=   5
         Columns(51).DataField=   ""
         Columns(51)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(52)._VlistStyle=   0
         Columns(52)._MaxComboItems=   5
         Columns(52).DataField=   ""
         Columns(52)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(53)._VlistStyle=   0
         Columns(53)._MaxComboItems=   5
         Columns(53).DataField=   ""
         Columns(53)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(54)._VlistStyle=   0
         Columns(54)._MaxComboItems=   5
         Columns(54).DataField=   ""
         Columns(54)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(55)._VlistStyle=   0
         Columns(55)._MaxComboItems=   5
         Columns(55).DataField=   ""
         Columns(55)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(56)._VlistStyle=   0
         Columns(56)._MaxComboItems=   5
         Columns(56).DataField=   ""
         Columns(56)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(57)._VlistStyle=   0
         Columns(57)._MaxComboItems=   5
         Columns(57).DataField=   ""
         Columns(57)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(58)._VlistStyle=   0
         Columns(58)._MaxComboItems=   5
         Columns(58).DataField=   ""
         Columns(58)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(59)._VlistStyle=   0
         Columns(59)._MaxComboItems=   5
         Columns(59).DataField=   ""
         Columns(59)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(60)._VlistStyle=   0
         Columns(60)._MaxComboItems=   5
         Columns(60).DataField=   ""
         Columns(60)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(61)._VlistStyle=   0
         Columns(61)._MaxComboItems=   5
         Columns(61).DataField=   ""
         Columns(61)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(62)._VlistStyle=   0
         Columns(62)._MaxComboItems=   5
         Columns(62).DataField=   ""
         Columns(62)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(63)._VlistStyle=   0
         Columns(63)._MaxComboItems=   5
         Columns(63).DataField=   ""
         Columns(63)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(64)._VlistStyle=   0
         Columns(64)._MaxComboItems=   5
         Columns(64).DataField=   ""
         Columns(64)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(65)._VlistStyle=   0
         Columns(65)._MaxComboItems=   5
         Columns(65).DataField=   ""
         Columns(65)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(66)._VlistStyle=   0
         Columns(66)._MaxComboItems=   5
         Columns(66).DataField=   ""
         Columns(66)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(67)._VlistStyle=   0
         Columns(67)._MaxComboItems=   5
         Columns(67).DataField=   ""
         Columns(67)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(68)._VlistStyle=   0
         Columns(68)._MaxComboItems=   5
         Columns(68).DataField=   ""
         Columns(68)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(69)._VlistStyle=   0
         Columns(69)._MaxComboItems=   5
         Columns(69).DataField=   ""
         Columns(69)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(70)._VlistStyle=   0
         Columns(70)._MaxComboItems=   5
         Columns(70).DataField=   ""
         Columns(70)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(71)._VlistStyle=   0
         Columns(71)._MaxComboItems=   5
         Columns(71).DataField=   ""
         Columns(71)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(72)._VlistStyle=   0
         Columns(72)._MaxComboItems=   5
         Columns(72).DataField=   ""
         Columns(72)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(73)._VlistStyle=   0
         Columns(73)._MaxComboItems=   5
         Columns(73).DataField=   ""
         Columns(73)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(74)._VlistStyle=   0
         Columns(74)._MaxComboItems=   5
         Columns(74).DataField=   ""
         Columns(74)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(75)._VlistStyle=   0
         Columns(75)._MaxComboItems=   5
         Columns(75).DataField=   ""
         Columns(75)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(76)._VlistStyle=   0
         Columns(76)._MaxComboItems=   5
         Columns(76).DataField=   ""
         Columns(76)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(77)._VlistStyle=   0
         Columns(77)._MaxComboItems=   5
         Columns(77).DataField=   ""
         Columns(77)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(78)._VlistStyle=   0
         Columns(78)._MaxComboItems=   5
         Columns(78).DataField=   ""
         Columns(78)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(79)._VlistStyle=   0
         Columns(79)._MaxComboItems=   5
         Columns(79).DataField=   ""
         Columns(79)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(80)._VlistStyle=   0
         Columns(80)._MaxComboItems=   5
         Columns(80).DataField=   ""
         Columns(80)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(81)._VlistStyle=   0
         Columns(81)._MaxComboItems=   5
         Columns(81).DataField=   ""
         Columns(81)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(82)._VlistStyle=   0
         Columns(82)._MaxComboItems=   5
         Columns(82).DataField=   ""
         Columns(82)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(83)._VlistStyle=   0
         Columns(83)._MaxComboItems=   5
         Columns(83).DataField=   ""
         Columns(83)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(84)._VlistStyle=   0
         Columns(84)._MaxComboItems=   5
         Columns(84).DataField=   ""
         Columns(84)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(85)._VlistStyle=   0
         Columns(85)._MaxComboItems=   5
         Columns(85).DataField=   ""
         Columns(85)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(86)._VlistStyle=   0
         Columns(86)._MaxComboItems=   5
         Columns(86).DataField=   ""
         Columns(86)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(87)._VlistStyle=   0
         Columns(87)._MaxComboItems=   5
         Columns(87).DataField=   ""
         Columns(87)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(88)._VlistStyle=   0
         Columns(88)._MaxComboItems=   5
         Columns(88).DataField=   ""
         Columns(88)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(89)._VlistStyle=   0
         Columns(89)._MaxComboItems=   5
         Columns(89).DataField=   ""
         Columns(89)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(90)._VlistStyle=   0
         Columns(90)._MaxComboItems=   5
         Columns(90).DataField=   ""
         Columns(90)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(91)._VlistStyle=   0
         Columns(91)._MaxComboItems=   5
         Columns(91).DataField=   ""
         Columns(91)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(92)._VlistStyle=   0
         Columns(92)._MaxComboItems=   5
         Columns(92).DataField=   ""
         Columns(92)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(93)._VlistStyle=   0
         Columns(93)._MaxComboItems=   5
         Columns(93).DataField=   ""
         Columns(93)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(94)._VlistStyle=   0
         Columns(94)._MaxComboItems=   5
         Columns(94).DataField=   ""
         Columns(94)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(95)._VlistStyle=   0
         Columns(95)._MaxComboItems=   5
         Columns(95).DataField=   ""
         Columns(95)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(96)._VlistStyle=   0
         Columns(96)._MaxComboItems=   5
         Columns(96).DataField=   ""
         Columns(96)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(97)._VlistStyle=   0
         Columns(97)._MaxComboItems=   5
         Columns(97).DataField=   ""
         Columns(97)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(98)._VlistStyle=   0
         Columns(98)._MaxComboItems=   5
         Columns(98).DataField=   ""
         Columns(98)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(99)._VlistStyle=   0
         Columns(99)._MaxComboItems=   5
         Columns(99).DataField=   ""
         Columns(99)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(100)._VlistStyle=   0
         Columns(100)._MaxComboItems=   5
         Columns(100).DataField=   ""
         Columns(100)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(101)._VlistStyle=   0
         Columns(101)._MaxComboItems=   5
         Columns(101).DataField=   ""
         Columns(101)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(102)._VlistStyle=   0
         Columns(102)._MaxComboItems=   5
         Columns(102).DataField=   ""
         Columns(102)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(103)._VlistStyle=   0
         Columns(103)._MaxComboItems=   5
         Columns(103).DataField=   ""
         Columns(103)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(104)._VlistStyle=   0
         Columns(104)._MaxComboItems=   5
         Columns(104).DataField=   ""
         Columns(104)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(105)._VlistStyle=   0
         Columns(105)._MaxComboItems=   5
         Columns(105).DataField=   ""
         Columns(105)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(106)._VlistStyle=   0
         Columns(106)._MaxComboItems=   5
         Columns(106).DataField=   ""
         Columns(106)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(107)._VlistStyle=   0
         Columns(107)._MaxComboItems=   5
         Columns(107).DataField=   ""
         Columns(107)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(108)._VlistStyle=   0
         Columns(108)._MaxComboItems=   5
         Columns(108).DataField=   ""
         Columns(108)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(109)._VlistStyle=   0
         Columns(109)._MaxComboItems=   5
         Columns(109).DataField=   ""
         Columns(109)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(110)._VlistStyle=   0
         Columns(110)._MaxComboItems=   5
         Columns(110).DataField=   ""
         Columns(110)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(111)._VlistStyle=   0
         Columns(111)._MaxComboItems=   5
         Columns(111).DataField=   ""
         Columns(111)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(112)._VlistStyle=   0
         Columns(112)._MaxComboItems=   5
         Columns(112).DataField=   ""
         Columns(112)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(113)._VlistStyle=   0
         Columns(113)._MaxComboItems=   5
         Columns(113).DataField=   ""
         Columns(113)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(114)._VlistStyle=   0
         Columns(114)._MaxComboItems=   5
         Columns(114).DataField=   ""
         Columns(114)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(115)._VlistStyle=   0
         Columns(115)._MaxComboItems=   5
         Columns(115).DataField=   ""
         Columns(115)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(116)._VlistStyle=   0
         Columns(116)._MaxComboItems=   5
         Columns(116).DataField=   ""
         Columns(116)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(117)._VlistStyle=   0
         Columns(117)._MaxComboItems=   5
         Columns(117).DataField=   ""
         Columns(117)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(118)._VlistStyle=   0
         Columns(118)._MaxComboItems=   5
         Columns(118).DataField=   ""
         Columns(118)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(119)._VlistStyle=   0
         Columns(119)._MaxComboItems=   5
         Columns(119).DataField=   ""
         Columns(119)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(120)._VlistStyle=   0
         Columns(120)._MaxComboItems=   5
         Columns(120).DataField=   ""
         Columns(120)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(121)._VlistStyle=   0
         Columns(121)._MaxComboItems=   5
         Columns(121).DataField=   ""
         Columns(121)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(122)._VlistStyle=   0
         Columns(122)._MaxComboItems=   5
         Columns(122).DataField=   ""
         Columns(122)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(123)._VlistStyle=   0
         Columns(123)._MaxComboItems=   5
         Columns(123).DataField=   ""
         Columns(123)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(124)._VlistStyle=   0
         Columns(124)._MaxComboItems=   5
         Columns(124).DataField=   ""
         Columns(124)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(125)._VlistStyle=   0
         Columns(125)._MaxComboItems=   5
         Columns(125).DataField=   ""
         Columns(125)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(126)._VlistStyle=   0
         Columns(126)._MaxComboItems=   5
         Columns(126).DataField=   ""
         Columns(126)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(127)._VlistStyle=   0
         Columns(127)._MaxComboItems=   5
         Columns(127).DataField=   ""
         Columns(127)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(128)._VlistStyle=   0
         Columns(128)._MaxComboItems=   5
         Columns(128).DataField=   ""
         Columns(128)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(129)._VlistStyle=   0
         Columns(129)._MaxComboItems=   5
         Columns(129).DataField=   ""
         Columns(129)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(130)._VlistStyle=   0
         Columns(130)._MaxComboItems=   5
         Columns(130).DataField=   ""
         Columns(130)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(131)._VlistStyle=   0
         Columns(131)._MaxComboItems=   5
         Columns(131).DataField=   ""
         Columns(131)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(132)._VlistStyle=   0
         Columns(132)._MaxComboItems=   5
         Columns(132).DataField=   ""
         Columns(132)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(133)._VlistStyle=   0
         Columns(133)._MaxComboItems=   5
         Columns(133).DataField=   ""
         Columns(133)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(134)._VlistStyle=   0
         Columns(134)._MaxComboItems=   5
         Columns(134).DataField=   ""
         Columns(134)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(135)._VlistStyle=   0
         Columns(135)._MaxComboItems=   5
         Columns(135).DataField=   ""
         Columns(135)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(136)._VlistStyle=   0
         Columns(136)._MaxComboItems=   5
         Columns(136).DataField=   ""
         Columns(136)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(137)._VlistStyle=   0
         Columns(137)._MaxComboItems=   5
         Columns(137).DataField=   ""
         Columns(137)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(138)._VlistStyle=   0
         Columns(138)._MaxComboItems=   5
         Columns(138).DataField=   ""
         Columns(138)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(139)._VlistStyle=   0
         Columns(139)._MaxComboItems=   5
         Columns(139).DataField=   ""
         Columns(139)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(140)._VlistStyle=   0
         Columns(140)._MaxComboItems=   5
         Columns(140).DataField=   ""
         Columns(140)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(141)._VlistStyle=   0
         Columns(141)._MaxComboItems=   5
         Columns(141).DataField=   ""
         Columns(141)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(142)._VlistStyle=   0
         Columns(142)._MaxComboItems=   5
         Columns(142).DataField=   ""
         Columns(142)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(143)._VlistStyle=   0
         Columns(143)._MaxComboItems=   5
         Columns(143).DataField=   ""
         Columns(143)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(144)._VlistStyle=   0
         Columns(144)._MaxComboItems=   5
         Columns(144).DataField=   ""
         Columns(144)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(145)._VlistStyle=   0
         Columns(145)._MaxComboItems=   5
         Columns(145).DataField=   ""
         Columns(145)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(146)._VlistStyle=   0
         Columns(146)._MaxComboItems=   5
         Columns(146).DataField=   ""
         Columns(146)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(147)._VlistStyle=   0
         Columns(147)._MaxComboItems=   5
         Columns(147).DataField=   ""
         Columns(147)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(148)._VlistStyle=   0
         Columns(148)._MaxComboItems=   5
         Columns(148).DataField=   ""
         Columns(148)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(149)._VlistStyle=   0
         Columns(149)._MaxComboItems=   5
         Columns(149).DataField=   ""
         Columns(149)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(150)._VlistStyle=   0
         Columns(150)._MaxComboItems=   5
         Columns(150).DataField=   ""
         Columns(150)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   151
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowFocus=   0   'False
         Splits(0).SizeMode=   2
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).AllowColSelect=   0   'False
         Splits(0).AllowRowSelect=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14869218
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=151"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2646"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2514"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=4366"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=4366"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4233"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=4366"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=4233"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=4366"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=4233"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=4366"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=4233"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=4366"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=4233"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=4366"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=4233"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=4366"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=4233"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(37)=   "Column(9).Width=4366"
         Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=4233"
         Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(41)=   "Column(10).Width=4366"
         Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=4233"
         Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(45)=   "Column(11).Width=4366"
         Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=4233"
         Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(49)=   "Column(12).Width=4366"
         Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=4233"
         Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(53)=   "Column(13).Width=4366"
         Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=4233"
         Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(57)=   "Column(14).Width=4366"
         Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=4233"
         Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(61)=   "Column(15).Width=4366"
         Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=4233"
         Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(65)=   "Column(16).Width=4366"
         Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=4233"
         Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(69)=   "Column(17).Width=4366"
         Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=4233"
         Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
         Splits(0)._ColumnProps(73)=   "Column(18).Width=4366"
         Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
         Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=4233"
         Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
         Splits(0)._ColumnProps(77)=   "Column(19).Width=4366"
         Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
         Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=4233"
         Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
         Splits(0)._ColumnProps(81)=   "Column(20).Width=4366"
         Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=4233"
         Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
         Splits(0)._ColumnProps(85)=   "Column(21).Width=4366"
         Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
         Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=4233"
         Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
         Splits(0)._ColumnProps(89)=   "Column(22).Width=4366"
         Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
         Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=4233"
         Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
         Splits(0)._ColumnProps(93)=   "Column(23).Width=4366"
         Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
         Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=4233"
         Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
         Splits(0)._ColumnProps(97)=   "Column(24).Width=4366"
         Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
         Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=4233"
         Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
         Splits(0)._ColumnProps(101)=   "Column(25).Width=4366"
         Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
         Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=4233"
         Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
         Splits(0)._ColumnProps(105)=   "Column(26).Width=4366"
         Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
         Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=4233"
         Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
         Splits(0)._ColumnProps(109)=   "Column(27).Width=4366"
         Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
         Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=4233"
         Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
         Splits(0)._ColumnProps(113)=   "Column(28).Width=4366"
         Splits(0)._ColumnProps(114)=   "Column(28).DividerColor=0"
         Splits(0)._ColumnProps(115)=   "Column(28)._WidthInPix=4233"
         Splits(0)._ColumnProps(116)=   "Column(28).Order=29"
         Splits(0)._ColumnProps(117)=   "Column(29).Width=4366"
         Splits(0)._ColumnProps(118)=   "Column(29).DividerColor=0"
         Splits(0)._ColumnProps(119)=   "Column(29)._WidthInPix=4233"
         Splits(0)._ColumnProps(120)=   "Column(29).Order=30"
         Splits(0)._ColumnProps(121)=   "Column(30).Width=4366"
         Splits(0)._ColumnProps(122)=   "Column(30).DividerColor=0"
         Splits(0)._ColumnProps(123)=   "Column(30)._WidthInPix=4233"
         Splits(0)._ColumnProps(124)=   "Column(30).Order=31"
         Splits(0)._ColumnProps(125)=   "Column(31).Width=4366"
         Splits(0)._ColumnProps(126)=   "Column(31).DividerColor=0"
         Splits(0)._ColumnProps(127)=   "Column(31)._WidthInPix=4233"
         Splits(0)._ColumnProps(128)=   "Column(31).Order=32"
         Splits(0)._ColumnProps(129)=   "Column(32).Width=4366"
         Splits(0)._ColumnProps(130)=   "Column(32).DividerColor=0"
         Splits(0)._ColumnProps(131)=   "Column(32)._WidthInPix=4233"
         Splits(0)._ColumnProps(132)=   "Column(32).Order=33"
         Splits(0)._ColumnProps(133)=   "Column(33).Width=4366"
         Splits(0)._ColumnProps(134)=   "Column(33).DividerColor=0"
         Splits(0)._ColumnProps(135)=   "Column(33)._WidthInPix=4233"
         Splits(0)._ColumnProps(136)=   "Column(33).Order=34"
         Splits(0)._ColumnProps(137)=   "Column(34).Width=4366"
         Splits(0)._ColumnProps(138)=   "Column(34).DividerColor=0"
         Splits(0)._ColumnProps(139)=   "Column(34)._WidthInPix=4233"
         Splits(0)._ColumnProps(140)=   "Column(34).Order=35"
         Splits(0)._ColumnProps(141)=   "Column(35).Width=4366"
         Splits(0)._ColumnProps(142)=   "Column(35).DividerColor=0"
         Splits(0)._ColumnProps(143)=   "Column(35)._WidthInPix=4233"
         Splits(0)._ColumnProps(144)=   "Column(35).Order=36"
         Splits(0)._ColumnProps(145)=   "Column(36).Width=4366"
         Splits(0)._ColumnProps(146)=   "Column(36).DividerColor=0"
         Splits(0)._ColumnProps(147)=   "Column(36)._WidthInPix=4233"
         Splits(0)._ColumnProps(148)=   "Column(36).Order=37"
         Splits(0)._ColumnProps(149)=   "Column(37).Width=4366"
         Splits(0)._ColumnProps(150)=   "Column(37).DividerColor=0"
         Splits(0)._ColumnProps(151)=   "Column(37)._WidthInPix=4233"
         Splits(0)._ColumnProps(152)=   "Column(37).Order=38"
         Splits(0)._ColumnProps(153)=   "Column(38).Width=4366"
         Splits(0)._ColumnProps(154)=   "Column(38).DividerColor=0"
         Splits(0)._ColumnProps(155)=   "Column(38)._WidthInPix=4233"
         Splits(0)._ColumnProps(156)=   "Column(38).Order=39"
         Splits(0)._ColumnProps(157)=   "Column(39).Width=4366"
         Splits(0)._ColumnProps(158)=   "Column(39).DividerColor=0"
         Splits(0)._ColumnProps(159)=   "Column(39)._WidthInPix=4233"
         Splits(0)._ColumnProps(160)=   "Column(39).Order=40"
         Splits(0)._ColumnProps(161)=   "Column(40).Width=4366"
         Splits(0)._ColumnProps(162)=   "Column(40).DividerColor=0"
         Splits(0)._ColumnProps(163)=   "Column(40)._WidthInPix=4233"
         Splits(0)._ColumnProps(164)=   "Column(40).Order=41"
         Splits(0)._ColumnProps(165)=   "Column(41).Width=4366"
         Splits(0)._ColumnProps(166)=   "Column(41).DividerColor=0"
         Splits(0)._ColumnProps(167)=   "Column(41)._WidthInPix=4233"
         Splits(0)._ColumnProps(168)=   "Column(41).Order=42"
         Splits(0)._ColumnProps(169)=   "Column(42).Width=4366"
         Splits(0)._ColumnProps(170)=   "Column(42).DividerColor=0"
         Splits(0)._ColumnProps(171)=   "Column(42)._WidthInPix=4233"
         Splits(0)._ColumnProps(172)=   "Column(42).Order=43"
         Splits(0)._ColumnProps(173)=   "Column(43).Width=4366"
         Splits(0)._ColumnProps(174)=   "Column(43).DividerColor=0"
         Splits(0)._ColumnProps(175)=   "Column(43)._WidthInPix=4233"
         Splits(0)._ColumnProps(176)=   "Column(43).Order=44"
         Splits(0)._ColumnProps(177)=   "Column(44).Width=4366"
         Splits(0)._ColumnProps(178)=   "Column(44).DividerColor=0"
         Splits(0)._ColumnProps(179)=   "Column(44)._WidthInPix=4233"
         Splits(0)._ColumnProps(180)=   "Column(44).Order=45"
         Splits(0)._ColumnProps(181)=   "Column(45).Width=4366"
         Splits(0)._ColumnProps(182)=   "Column(45).DividerColor=0"
         Splits(0)._ColumnProps(183)=   "Column(45)._WidthInPix=4233"
         Splits(0)._ColumnProps(184)=   "Column(45).Order=46"
         Splits(0)._ColumnProps(185)=   "Column(46).Width=4366"
         Splits(0)._ColumnProps(186)=   "Column(46).DividerColor=0"
         Splits(0)._ColumnProps(187)=   "Column(46)._WidthInPix=4233"
         Splits(0)._ColumnProps(188)=   "Column(46).Order=47"
         Splits(0)._ColumnProps(189)=   "Column(47).Width=4366"
         Splits(0)._ColumnProps(190)=   "Column(47).DividerColor=0"
         Splits(0)._ColumnProps(191)=   "Column(47)._WidthInPix=4233"
         Splits(0)._ColumnProps(192)=   "Column(47).Order=48"
         Splits(0)._ColumnProps(193)=   "Column(48).Width=4366"
         Splits(0)._ColumnProps(194)=   "Column(48).DividerColor=0"
         Splits(0)._ColumnProps(195)=   "Column(48)._WidthInPix=4233"
         Splits(0)._ColumnProps(196)=   "Column(48).Order=49"
         Splits(0)._ColumnProps(197)=   "Column(49).Width=4366"
         Splits(0)._ColumnProps(198)=   "Column(49).DividerColor=0"
         Splits(0)._ColumnProps(199)=   "Column(49)._WidthInPix=4233"
         Splits(0)._ColumnProps(200)=   "Column(49).Order=50"
         Splits(0)._ColumnProps(201)=   "Column(50).Width=4366"
         Splits(0)._ColumnProps(202)=   "Column(50).DividerColor=0"
         Splits(0)._ColumnProps(203)=   "Column(50)._WidthInPix=4233"
         Splits(0)._ColumnProps(204)=   "Column(50).Order=51"
         Splits(0)._ColumnProps(205)=   "Column(51).Width=4366"
         Splits(0)._ColumnProps(206)=   "Column(51).DividerColor=0"
         Splits(0)._ColumnProps(207)=   "Column(51)._WidthInPix=4233"
         Splits(0)._ColumnProps(208)=   "Column(51).Order=52"
         Splits(0)._ColumnProps(209)=   "Column(52).Width=4366"
         Splits(0)._ColumnProps(210)=   "Column(52).DividerColor=0"
         Splits(0)._ColumnProps(211)=   "Column(52)._WidthInPix=4233"
         Splits(0)._ColumnProps(212)=   "Column(52).Order=53"
         Splits(0)._ColumnProps(213)=   "Column(53).Width=4366"
         Splits(0)._ColumnProps(214)=   "Column(53).DividerColor=0"
         Splits(0)._ColumnProps(215)=   "Column(53)._WidthInPix=4233"
         Splits(0)._ColumnProps(216)=   "Column(53).Order=54"
         Splits(0)._ColumnProps(217)=   "Column(54).Width=4366"
         Splits(0)._ColumnProps(218)=   "Column(54).DividerColor=0"
         Splits(0)._ColumnProps(219)=   "Column(54)._WidthInPix=4233"
         Splits(0)._ColumnProps(220)=   "Column(54).Order=55"
         Splits(0)._ColumnProps(221)=   "Column(55).Width=4366"
         Splits(0)._ColumnProps(222)=   "Column(55).DividerColor=0"
         Splits(0)._ColumnProps(223)=   "Column(55)._WidthInPix=4233"
         Splits(0)._ColumnProps(224)=   "Column(55).Order=56"
         Splits(0)._ColumnProps(225)=   "Column(56).Width=4366"
         Splits(0)._ColumnProps(226)=   "Column(56).DividerColor=0"
         Splits(0)._ColumnProps(227)=   "Column(56)._WidthInPix=4233"
         Splits(0)._ColumnProps(228)=   "Column(56).Order=57"
         Splits(0)._ColumnProps(229)=   "Column(57).Width=4366"
         Splits(0)._ColumnProps(230)=   "Column(57).DividerColor=0"
         Splits(0)._ColumnProps(231)=   "Column(57)._WidthInPix=4233"
         Splits(0)._ColumnProps(232)=   "Column(57).Order=58"
         Splits(0)._ColumnProps(233)=   "Column(58).Width=4366"
         Splits(0)._ColumnProps(234)=   "Column(58).DividerColor=0"
         Splits(0)._ColumnProps(235)=   "Column(58)._WidthInPix=4233"
         Splits(0)._ColumnProps(236)=   "Column(58).Order=59"
         Splits(0)._ColumnProps(237)=   "Column(59).Width=4366"
         Splits(0)._ColumnProps(238)=   "Column(59).DividerColor=0"
         Splits(0)._ColumnProps(239)=   "Column(59)._WidthInPix=4233"
         Splits(0)._ColumnProps(240)=   "Column(59).Order=60"
         Splits(0)._ColumnProps(241)=   "Column(60).Width=4366"
         Splits(0)._ColumnProps(242)=   "Column(60).DividerColor=0"
         Splits(0)._ColumnProps(243)=   "Column(60)._WidthInPix=4233"
         Splits(0)._ColumnProps(244)=   "Column(60).Order=61"
         Splits(0)._ColumnProps(245)=   "Column(61).Width=4366"
         Splits(0)._ColumnProps(246)=   "Column(61).DividerColor=0"
         Splits(0)._ColumnProps(247)=   "Column(61)._WidthInPix=4233"
         Splits(0)._ColumnProps(248)=   "Column(61).Order=62"
         Splits(0)._ColumnProps(249)=   "Column(62).Width=4366"
         Splits(0)._ColumnProps(250)=   "Column(62).DividerColor=0"
         Splits(0)._ColumnProps(251)=   "Column(62)._WidthInPix=4233"
         Splits(0)._ColumnProps(252)=   "Column(62).Order=63"
         Splits(0)._ColumnProps(253)=   "Column(63).Width=4366"
         Splits(0)._ColumnProps(254)=   "Column(63).DividerColor=0"
         Splits(0)._ColumnProps(255)=   "Column(63)._WidthInPix=4233"
         Splits(0)._ColumnProps(256)=   "Column(63).Order=64"
         Splits(0)._ColumnProps(257)=   "Column(64).Width=4366"
         Splits(0)._ColumnProps(258)=   "Column(64).DividerColor=0"
         Splits(0)._ColumnProps(259)=   "Column(64)._WidthInPix=4233"
         Splits(0)._ColumnProps(260)=   "Column(64).Order=65"
         Splits(0)._ColumnProps(261)=   "Column(65).Width=4366"
         Splits(0)._ColumnProps(262)=   "Column(65).DividerColor=0"
         Splits(0)._ColumnProps(263)=   "Column(65)._WidthInPix=4233"
         Splits(0)._ColumnProps(264)=   "Column(65).Order=66"
         Splits(0)._ColumnProps(265)=   "Column(66).Width=4366"
         Splits(0)._ColumnProps(266)=   "Column(66).DividerColor=0"
         Splits(0)._ColumnProps(267)=   "Column(66)._WidthInPix=4233"
         Splits(0)._ColumnProps(268)=   "Column(66).Order=67"
         Splits(0)._ColumnProps(269)=   "Column(67).Width=4366"
         Splits(0)._ColumnProps(270)=   "Column(67).DividerColor=0"
         Splits(0)._ColumnProps(271)=   "Column(67)._WidthInPix=4233"
         Splits(0)._ColumnProps(272)=   "Column(67).Order=68"
         Splits(0)._ColumnProps(273)=   "Column(68).Width=4366"
         Splits(0)._ColumnProps(274)=   "Column(68).DividerColor=0"
         Splits(0)._ColumnProps(275)=   "Column(68)._WidthInPix=4233"
         Splits(0)._ColumnProps(276)=   "Column(68).Order=69"
         Splits(0)._ColumnProps(277)=   "Column(69).Width=4366"
         Splits(0)._ColumnProps(278)=   "Column(69).DividerColor=0"
         Splits(0)._ColumnProps(279)=   "Column(69)._WidthInPix=4233"
         Splits(0)._ColumnProps(280)=   "Column(69).Order=70"
         Splits(0)._ColumnProps(281)=   "Column(70).Width=4366"
         Splits(0)._ColumnProps(282)=   "Column(70).DividerColor=0"
         Splits(0)._ColumnProps(283)=   "Column(70)._WidthInPix=4233"
         Splits(0)._ColumnProps(284)=   "Column(70).Order=71"
         Splits(0)._ColumnProps(285)=   "Column(71).Width=4366"
         Splits(0)._ColumnProps(286)=   "Column(71).DividerColor=0"
         Splits(0)._ColumnProps(287)=   "Column(71)._WidthInPix=4233"
         Splits(0)._ColumnProps(288)=   "Column(71).Order=72"
         Splits(0)._ColumnProps(289)=   "Column(72).Width=4366"
         Splits(0)._ColumnProps(290)=   "Column(72).DividerColor=0"
         Splits(0)._ColumnProps(291)=   "Column(72)._WidthInPix=4233"
         Splits(0)._ColumnProps(292)=   "Column(72).Order=73"
         Splits(0)._ColumnProps(293)=   "Column(73).Width=4366"
         Splits(0)._ColumnProps(294)=   "Column(73).DividerColor=0"
         Splits(0)._ColumnProps(295)=   "Column(73)._WidthInPix=4233"
         Splits(0)._ColumnProps(296)=   "Column(73).Order=74"
         Splits(0)._ColumnProps(297)=   "Column(74).Width=4366"
         Splits(0)._ColumnProps(298)=   "Column(74).DividerColor=0"
         Splits(0)._ColumnProps(299)=   "Column(74)._WidthInPix=4233"
         Splits(0)._ColumnProps(300)=   "Column(74).Order=75"
         Splits(0)._ColumnProps(301)=   "Column(75).Width=4366"
         Splits(0)._ColumnProps(302)=   "Column(75).DividerColor=0"
         Splits(0)._ColumnProps(303)=   "Column(75)._WidthInPix=4233"
         Splits(0)._ColumnProps(304)=   "Column(75).Order=76"
         Splits(0)._ColumnProps(305)=   "Column(76).Width=4366"
         Splits(0)._ColumnProps(306)=   "Column(76).DividerColor=0"
         Splits(0)._ColumnProps(307)=   "Column(76)._WidthInPix=4233"
         Splits(0)._ColumnProps(308)=   "Column(76).Order=77"
         Splits(0)._ColumnProps(309)=   "Column(77).Width=4366"
         Splits(0)._ColumnProps(310)=   "Column(77).DividerColor=0"
         Splits(0)._ColumnProps(311)=   "Column(77)._WidthInPix=4233"
         Splits(0)._ColumnProps(312)=   "Column(77).Order=78"
         Splits(0)._ColumnProps(313)=   "Column(78).Width=4366"
         Splits(0)._ColumnProps(314)=   "Column(78).DividerColor=0"
         Splits(0)._ColumnProps(315)=   "Column(78)._WidthInPix=4233"
         Splits(0)._ColumnProps(316)=   "Column(78).Order=79"
         Splits(0)._ColumnProps(317)=   "Column(79).Width=4366"
         Splits(0)._ColumnProps(318)=   "Column(79).DividerColor=0"
         Splits(0)._ColumnProps(319)=   "Column(79)._WidthInPix=4233"
         Splits(0)._ColumnProps(320)=   "Column(79).Order=80"
         Splits(0)._ColumnProps(321)=   "Column(80).Width=4366"
         Splits(0)._ColumnProps(322)=   "Column(80).DividerColor=0"
         Splits(0)._ColumnProps(323)=   "Column(80)._WidthInPix=4233"
         Splits(0)._ColumnProps(324)=   "Column(80).Order=81"
         Splits(0)._ColumnProps(325)=   "Column(81).Width=4366"
         Splits(0)._ColumnProps(326)=   "Column(81).DividerColor=0"
         Splits(0)._ColumnProps(327)=   "Column(81)._WidthInPix=4233"
         Splits(0)._ColumnProps(328)=   "Column(81).Order=82"
         Splits(0)._ColumnProps(329)=   "Column(82).Width=4366"
         Splits(0)._ColumnProps(330)=   "Column(82).DividerColor=0"
         Splits(0)._ColumnProps(331)=   "Column(82)._WidthInPix=4233"
         Splits(0)._ColumnProps(332)=   "Column(82).Order=83"
         Splits(0)._ColumnProps(333)=   "Column(83).Width=4366"
         Splits(0)._ColumnProps(334)=   "Column(83).DividerColor=0"
         Splits(0)._ColumnProps(335)=   "Column(83)._WidthInPix=4233"
         Splits(0)._ColumnProps(336)=   "Column(83).Order=84"
         Splits(0)._ColumnProps(337)=   "Column(84).Width=4366"
         Splits(0)._ColumnProps(338)=   "Column(84).DividerColor=0"
         Splits(0)._ColumnProps(339)=   "Column(84)._WidthInPix=4233"
         Splits(0)._ColumnProps(340)=   "Column(84).Order=85"
         Splits(0)._ColumnProps(341)=   "Column(85).Width=4366"
         Splits(0)._ColumnProps(342)=   "Column(85).DividerColor=0"
         Splits(0)._ColumnProps(343)=   "Column(85)._WidthInPix=4233"
         Splits(0)._ColumnProps(344)=   "Column(85).Order=86"
         Splits(0)._ColumnProps(345)=   "Column(86).Width=4366"
         Splits(0)._ColumnProps(346)=   "Column(86).DividerColor=0"
         Splits(0)._ColumnProps(347)=   "Column(86)._WidthInPix=4233"
         Splits(0)._ColumnProps(348)=   "Column(86).Order=87"
         Splits(0)._ColumnProps(349)=   "Column(87).Width=4366"
         Splits(0)._ColumnProps(350)=   "Column(87).DividerColor=0"
         Splits(0)._ColumnProps(351)=   "Column(87)._WidthInPix=4233"
         Splits(0)._ColumnProps(352)=   "Column(87).Order=88"
         Splits(0)._ColumnProps(353)=   "Column(88).Width=4366"
         Splits(0)._ColumnProps(354)=   "Column(88).DividerColor=0"
         Splits(0)._ColumnProps(355)=   "Column(88)._WidthInPix=4233"
         Splits(0)._ColumnProps(356)=   "Column(88).Order=89"
         Splits(0)._ColumnProps(357)=   "Column(89).Width=4366"
         Splits(0)._ColumnProps(358)=   "Column(89).DividerColor=0"
         Splits(0)._ColumnProps(359)=   "Column(89)._WidthInPix=4233"
         Splits(0)._ColumnProps(360)=   "Column(89).Order=90"
         Splits(0)._ColumnProps(361)=   "Column(90).Width=4366"
         Splits(0)._ColumnProps(362)=   "Column(90).DividerColor=0"
         Splits(0)._ColumnProps(363)=   "Column(90)._WidthInPix=4233"
         Splits(0)._ColumnProps(364)=   "Column(90).Order=91"
         Splits(0)._ColumnProps(365)=   "Column(91).Width=4366"
         Splits(0)._ColumnProps(366)=   "Column(91).DividerColor=0"
         Splits(0)._ColumnProps(367)=   "Column(91)._WidthInPix=4233"
         Splits(0)._ColumnProps(368)=   "Column(91).Order=92"
         Splits(0)._ColumnProps(369)=   "Column(92).Width=4366"
         Splits(0)._ColumnProps(370)=   "Column(92).DividerColor=0"
         Splits(0)._ColumnProps(371)=   "Column(92)._WidthInPix=4233"
         Splits(0)._ColumnProps(372)=   "Column(92).Order=93"
         Splits(0)._ColumnProps(373)=   "Column(93).Width=4366"
         Splits(0)._ColumnProps(374)=   "Column(93).DividerColor=0"
         Splits(0)._ColumnProps(375)=   "Column(93)._WidthInPix=4233"
         Splits(0)._ColumnProps(376)=   "Column(93).Order=94"
         Splits(0)._ColumnProps(377)=   "Column(94).Width=4366"
         Splits(0)._ColumnProps(378)=   "Column(94).DividerColor=0"
         Splits(0)._ColumnProps(379)=   "Column(94)._WidthInPix=4233"
         Splits(0)._ColumnProps(380)=   "Column(94).Order=95"
         Splits(0)._ColumnProps(381)=   "Column(95).Width=4366"
         Splits(0)._ColumnProps(382)=   "Column(95).DividerColor=0"
         Splits(0)._ColumnProps(383)=   "Column(95)._WidthInPix=4233"
         Splits(0)._ColumnProps(384)=   "Column(95).Order=96"
         Splits(0)._ColumnProps(385)=   "Column(96).Width=4366"
         Splits(0)._ColumnProps(386)=   "Column(96).DividerColor=0"
         Splits(0)._ColumnProps(387)=   "Column(96)._WidthInPix=4233"
         Splits(0)._ColumnProps(388)=   "Column(96).Order=97"
         Splits(0)._ColumnProps(389)=   "Column(97).Width=4366"
         Splits(0)._ColumnProps(390)=   "Column(97).DividerColor=0"
         Splits(0)._ColumnProps(391)=   "Column(97)._WidthInPix=4233"
         Splits(0)._ColumnProps(392)=   "Column(97).Order=98"
         Splits(0)._ColumnProps(393)=   "Column(98).Width=4366"
         Splits(0)._ColumnProps(394)=   "Column(98).DividerColor=0"
         Splits(0)._ColumnProps(395)=   "Column(98)._WidthInPix=4233"
         Splits(0)._ColumnProps(396)=   "Column(98).Order=99"
         Splits(0)._ColumnProps(397)=   "Column(99).Width=4366"
         Splits(0)._ColumnProps(398)=   "Column(99).DividerColor=0"
         Splits(0)._ColumnProps(399)=   "Column(99)._WidthInPix=4233"
         Splits(0)._ColumnProps(400)=   "Column(99).Order=100"
         Splits(0)._ColumnProps(401)=   "Column(100).Width=4366"
         Splits(0)._ColumnProps(402)=   "Column(100).DividerColor=0"
         Splits(0)._ColumnProps(403)=   "Column(100)._WidthInPix=4233"
         Splits(0)._ColumnProps(404)=   "Column(100).Order=101"
         Splits(0)._ColumnProps(405)=   "Column(101).Width=4366"
         Splits(0)._ColumnProps(406)=   "Column(101).DividerColor=0"
         Splits(0)._ColumnProps(407)=   "Column(101)._WidthInPix=4233"
         Splits(0)._ColumnProps(408)=   "Column(101).Order=102"
         Splits(0)._ColumnProps(409)=   "Column(102).Width=4366"
         Splits(0)._ColumnProps(410)=   "Column(102).DividerColor=0"
         Splits(0)._ColumnProps(411)=   "Column(102)._WidthInPix=4233"
         Splits(0)._ColumnProps(412)=   "Column(102).Order=103"
         Splits(0)._ColumnProps(413)=   "Column(103).Width=4366"
         Splits(0)._ColumnProps(414)=   "Column(103).DividerColor=0"
         Splits(0)._ColumnProps(415)=   "Column(103)._WidthInPix=4233"
         Splits(0)._ColumnProps(416)=   "Column(103).Order=104"
         Splits(0)._ColumnProps(417)=   "Column(104).Width=4366"
         Splits(0)._ColumnProps(418)=   "Column(104).DividerColor=0"
         Splits(0)._ColumnProps(419)=   "Column(104)._WidthInPix=4233"
         Splits(0)._ColumnProps(420)=   "Column(104).Order=105"
         Splits(0)._ColumnProps(421)=   "Column(105).Width=4366"
         Splits(0)._ColumnProps(422)=   "Column(105).DividerColor=0"
         Splits(0)._ColumnProps(423)=   "Column(105)._WidthInPix=4233"
         Splits(0)._ColumnProps(424)=   "Column(105).Order=106"
         Splits(0)._ColumnProps(425)=   "Column(106).Width=4366"
         Splits(0)._ColumnProps(426)=   "Column(106).DividerColor=0"
         Splits(0)._ColumnProps(427)=   "Column(106)._WidthInPix=4233"
         Splits(0)._ColumnProps(428)=   "Column(106).Order=107"
         Splits(0)._ColumnProps(429)=   "Column(107).Width=4366"
         Splits(0)._ColumnProps(430)=   "Column(107).DividerColor=0"
         Splits(0)._ColumnProps(431)=   "Column(107)._WidthInPix=4233"
         Splits(0)._ColumnProps(432)=   "Column(107).Order=108"
         Splits(0)._ColumnProps(433)=   "Column(108).Width=4366"
         Splits(0)._ColumnProps(434)=   "Column(108).DividerColor=0"
         Splits(0)._ColumnProps(435)=   "Column(108)._WidthInPix=4233"
         Splits(0)._ColumnProps(436)=   "Column(108).Order=109"
         Splits(0)._ColumnProps(437)=   "Column(109).Width=4366"
         Splits(0)._ColumnProps(438)=   "Column(109).DividerColor=0"
         Splits(0)._ColumnProps(439)=   "Column(109)._WidthInPix=4233"
         Splits(0)._ColumnProps(440)=   "Column(109).Order=110"
         Splits(0)._ColumnProps(441)=   "Column(110).Width=4366"
         Splits(0)._ColumnProps(442)=   "Column(110).DividerColor=0"
         Splits(0)._ColumnProps(443)=   "Column(110)._WidthInPix=4233"
         Splits(0)._ColumnProps(444)=   "Column(110).Order=111"
         Splits(0)._ColumnProps(445)=   "Column(111).Width=4366"
         Splits(0)._ColumnProps(446)=   "Column(111).DividerColor=0"
         Splits(0)._ColumnProps(447)=   "Column(111)._WidthInPix=4233"
         Splits(0)._ColumnProps(448)=   "Column(111).Order=112"
         Splits(0)._ColumnProps(449)=   "Column(112).Width=4366"
         Splits(0)._ColumnProps(450)=   "Column(112).DividerColor=0"
         Splits(0)._ColumnProps(451)=   "Column(112)._WidthInPix=4233"
         Splits(0)._ColumnProps(452)=   "Column(112).Order=113"
         Splits(0)._ColumnProps(453)=   "Column(113).Width=4366"
         Splits(0)._ColumnProps(454)=   "Column(113).DividerColor=0"
         Splits(0)._ColumnProps(455)=   "Column(113)._WidthInPix=4233"
         Splits(0)._ColumnProps(456)=   "Column(113).Order=114"
         Splits(0)._ColumnProps(457)=   "Column(114).Width=4366"
         Splits(0)._ColumnProps(458)=   "Column(114).DividerColor=0"
         Splits(0)._ColumnProps(459)=   "Column(114)._WidthInPix=4233"
         Splits(0)._ColumnProps(460)=   "Column(114).Order=115"
         Splits(0)._ColumnProps(461)=   "Column(115).Width=4366"
         Splits(0)._ColumnProps(462)=   "Column(115).DividerColor=0"
         Splits(0)._ColumnProps(463)=   "Column(115)._WidthInPix=4233"
         Splits(0)._ColumnProps(464)=   "Column(115).Order=116"
         Splits(0)._ColumnProps(465)=   "Column(116).Width=4366"
         Splits(0)._ColumnProps(466)=   "Column(116).DividerColor=0"
         Splits(0)._ColumnProps(467)=   "Column(116)._WidthInPix=4233"
         Splits(0)._ColumnProps(468)=   "Column(116).Order=117"
         Splits(0)._ColumnProps(469)=   "Column(117).Width=4366"
         Splits(0)._ColumnProps(470)=   "Column(117).DividerColor=0"
         Splits(0)._ColumnProps(471)=   "Column(117)._WidthInPix=4233"
         Splits(0)._ColumnProps(472)=   "Column(117).Order=118"
         Splits(0)._ColumnProps(473)=   "Column(118).Width=4366"
         Splits(0)._ColumnProps(474)=   "Column(118).DividerColor=0"
         Splits(0)._ColumnProps(475)=   "Column(118)._WidthInPix=4233"
         Splits(0)._ColumnProps(476)=   "Column(118).Order=119"
         Splits(0)._ColumnProps(477)=   "Column(119).Width=4366"
         Splits(0)._ColumnProps(478)=   "Column(119).DividerColor=0"
         Splits(0)._ColumnProps(479)=   "Column(119)._WidthInPix=4233"
         Splits(0)._ColumnProps(480)=   "Column(119).Order=120"
         Splits(0)._ColumnProps(481)=   "Column(120).Width=4366"
         Splits(0)._ColumnProps(482)=   "Column(120).DividerColor=0"
         Splits(0)._ColumnProps(483)=   "Column(120)._WidthInPix=4233"
         Splits(0)._ColumnProps(484)=   "Column(120).Order=121"
         Splits(0)._ColumnProps(485)=   "Column(121).Width=4366"
         Splits(0)._ColumnProps(486)=   "Column(121).DividerColor=0"
         Splits(0)._ColumnProps(487)=   "Column(121)._WidthInPix=4233"
         Splits(0)._ColumnProps(488)=   "Column(121).Order=122"
         Splits(0)._ColumnProps(489)=   "Column(122).Width=4366"
         Splits(0)._ColumnProps(490)=   "Column(122).DividerColor=0"
         Splits(0)._ColumnProps(491)=   "Column(122)._WidthInPix=4233"
         Splits(0)._ColumnProps(492)=   "Column(122).Order=123"
         Splits(0)._ColumnProps(493)=   "Column(123).Width=4366"
         Splits(0)._ColumnProps(494)=   "Column(123).DividerColor=0"
         Splits(0)._ColumnProps(495)=   "Column(123)._WidthInPix=4233"
         Splits(0)._ColumnProps(496)=   "Column(123).Order=124"
         Splits(0)._ColumnProps(497)=   "Column(124).Width=4366"
         Splits(0)._ColumnProps(498)=   "Column(124).DividerColor=0"
         Splits(0)._ColumnProps(499)=   "Column(124)._WidthInPix=4233"
         Splits(0)._ColumnProps(500)=   "Column(124).Order=125"
         Splits(0)._ColumnProps(501)=   "Column(125).Width=4366"
         Splits(0)._ColumnProps(502)=   "Column(125).DividerColor=0"
         Splits(0)._ColumnProps(503)=   "Column(125)._WidthInPix=4233"
         Splits(0)._ColumnProps(504)=   "Column(125).Order=126"
         Splits(0)._ColumnProps(505)=   "Column(126).Width=4366"
         Splits(0)._ColumnProps(506)=   "Column(126).DividerColor=0"
         Splits(0)._ColumnProps(507)=   "Column(126)._WidthInPix=4233"
         Splits(0)._ColumnProps(508)=   "Column(126).Order=127"
         Splits(0)._ColumnProps(509)=   "Column(127).Width=4366"
         Splits(0)._ColumnProps(510)=   "Column(127).DividerColor=0"
         Splits(0)._ColumnProps(511)=   "Column(127)._WidthInPix=4233"
         Splits(0)._ColumnProps(512)=   "Column(127).Order=128"
         Splits(0)._ColumnProps(513)=   "Column(128).Width=4366"
         Splits(0)._ColumnProps(514)=   "Column(128).DividerColor=0"
         Splits(0)._ColumnProps(515)=   "Column(128)._WidthInPix=4233"
         Splits(0)._ColumnProps(516)=   "Column(128).Order=129"
         Splits(0)._ColumnProps(517)=   "Column(129).Width=4366"
         Splits(0)._ColumnProps(518)=   "Column(129).DividerColor=0"
         Splits(0)._ColumnProps(519)=   "Column(129)._WidthInPix=4233"
         Splits(0)._ColumnProps(520)=   "Column(129).Order=130"
         Splits(0)._ColumnProps(521)=   "Column(130).Width=4366"
         Splits(0)._ColumnProps(522)=   "Column(130).DividerColor=0"
         Splits(0)._ColumnProps(523)=   "Column(130)._WidthInPix=4233"
         Splits(0)._ColumnProps(524)=   "Column(130).Order=131"
         Splits(0)._ColumnProps(525)=   "Column(131).Width=4366"
         Splits(0)._ColumnProps(526)=   "Column(131).DividerColor=0"
         Splits(0)._ColumnProps(527)=   "Column(131)._WidthInPix=4233"
         Splits(0)._ColumnProps(528)=   "Column(131).Order=132"
         Splits(0)._ColumnProps(529)=   "Column(132).Width=4366"
         Splits(0)._ColumnProps(530)=   "Column(132).DividerColor=0"
         Splits(0)._ColumnProps(531)=   "Column(132)._WidthInPix=4233"
         Splits(0)._ColumnProps(532)=   "Column(132).Order=133"
         Splits(0)._ColumnProps(533)=   "Column(133).Width=4366"
         Splits(0)._ColumnProps(534)=   "Column(133).DividerColor=0"
         Splits(0)._ColumnProps(535)=   "Column(133)._WidthInPix=4233"
         Splits(0)._ColumnProps(536)=   "Column(133).Order=134"
         Splits(0)._ColumnProps(537)=   "Column(134).Width=4366"
         Splits(0)._ColumnProps(538)=   "Column(134).DividerColor=0"
         Splits(0)._ColumnProps(539)=   "Column(134)._WidthInPix=4233"
         Splits(0)._ColumnProps(540)=   "Column(134).Order=135"
         Splits(0)._ColumnProps(541)=   "Column(135).Width=4366"
         Splits(0)._ColumnProps(542)=   "Column(135).DividerColor=0"
         Splits(0)._ColumnProps(543)=   "Column(135)._WidthInPix=4233"
         Splits(0)._ColumnProps(544)=   "Column(135).Order=136"
         Splits(0)._ColumnProps(545)=   "Column(136).Width=4366"
         Splits(0)._ColumnProps(546)=   "Column(136).DividerColor=0"
         Splits(0)._ColumnProps(547)=   "Column(136)._WidthInPix=4233"
         Splits(0)._ColumnProps(548)=   "Column(136).Order=137"
         Splits(0)._ColumnProps(549)=   "Column(137).Width=4366"
         Splits(0)._ColumnProps(550)=   "Column(137).DividerColor=0"
         Splits(0)._ColumnProps(551)=   "Column(137)._WidthInPix=4233"
         Splits(0)._ColumnProps(552)=   "Column(137).Order=138"
         Splits(0)._ColumnProps(553)=   "Column(138).Width=4366"
         Splits(0)._ColumnProps(554)=   "Column(138).DividerColor=0"
         Splits(0)._ColumnProps(555)=   "Column(138)._WidthInPix=4233"
         Splits(0)._ColumnProps(556)=   "Column(138).Order=139"
         Splits(0)._ColumnProps(557)=   "Column(139).Width=4366"
         Splits(0)._ColumnProps(558)=   "Column(139).DividerColor=0"
         Splits(0)._ColumnProps(559)=   "Column(139)._WidthInPix=4233"
         Splits(0)._ColumnProps(560)=   "Column(139).Order=140"
         Splits(0)._ColumnProps(561)=   "Column(140).Width=4366"
         Splits(0)._ColumnProps(562)=   "Column(140).DividerColor=0"
         Splits(0)._ColumnProps(563)=   "Column(140)._WidthInPix=4233"
         Splits(0)._ColumnProps(564)=   "Column(140).Order=141"
         Splits(0)._ColumnProps(565)=   "Column(141).Width=4366"
         Splits(0)._ColumnProps(566)=   "Column(141).DividerColor=0"
         Splits(0)._ColumnProps(567)=   "Column(141)._WidthInPix=4233"
         Splits(0)._ColumnProps(568)=   "Column(141).Order=142"
         Splits(0)._ColumnProps(569)=   "Column(142).Width=4366"
         Splits(0)._ColumnProps(570)=   "Column(142).DividerColor=0"
         Splits(0)._ColumnProps(571)=   "Column(142)._WidthInPix=4233"
         Splits(0)._ColumnProps(572)=   "Column(142).Order=143"
         Splits(0)._ColumnProps(573)=   "Column(143).Width=4366"
         Splits(0)._ColumnProps(574)=   "Column(143).DividerColor=0"
         Splits(0)._ColumnProps(575)=   "Column(143)._WidthInPix=4233"
         Splits(0)._ColumnProps(576)=   "Column(143).Order=144"
         Splits(0)._ColumnProps(577)=   "Column(144).Width=4366"
         Splits(0)._ColumnProps(578)=   "Column(144).DividerColor=0"
         Splits(0)._ColumnProps(579)=   "Column(144)._WidthInPix=4233"
         Splits(0)._ColumnProps(580)=   "Column(144).Order=145"
         Splits(0)._ColumnProps(581)=   "Column(145).Width=4366"
         Splits(0)._ColumnProps(582)=   "Column(145).DividerColor=0"
         Splits(0)._ColumnProps(583)=   "Column(145)._WidthInPix=4233"
         Splits(0)._ColumnProps(584)=   "Column(145).Order=146"
         Splits(0)._ColumnProps(585)=   "Column(146).Width=4366"
         Splits(0)._ColumnProps(586)=   "Column(146).DividerColor=0"
         Splits(0)._ColumnProps(587)=   "Column(146)._WidthInPix=4233"
         Splits(0)._ColumnProps(588)=   "Column(146).Order=147"
         Splits(0)._ColumnProps(589)=   "Column(147).Width=4366"
         Splits(0)._ColumnProps(590)=   "Column(147).DividerColor=0"
         Splits(0)._ColumnProps(591)=   "Column(147)._WidthInPix=4233"
         Splits(0)._ColumnProps(592)=   "Column(147).Order=148"
         Splits(0)._ColumnProps(593)=   "Column(148).Width=4366"
         Splits(0)._ColumnProps(594)=   "Column(148).DividerColor=0"
         Splits(0)._ColumnProps(595)=   "Column(148)._WidthInPix=4233"
         Splits(0)._ColumnProps(596)=   "Column(148).Order=149"
         Splits(0)._ColumnProps(597)=   "Column(149).Width=4366"
         Splits(0)._ColumnProps(598)=   "Column(149).DividerColor=0"
         Splits(0)._ColumnProps(599)=   "Column(149)._WidthInPix=4233"
         Splits(0)._ColumnProps(600)=   "Column(149).Order=150"
         Splits(0)._ColumnProps(601)=   "Column(150).Width=4366"
         Splits(0)._ColumnProps(602)=   "Column(150).DividerColor=0"
         Splits(0)._ColumnProps(603)=   "Column(150)._WidthInPix=4233"
         Splits(0)._ColumnProps(604)=   "Column(150).Order=151"
         Splits(1)._UserFlags=   0
         Splits(1).AllowRowSizing=   0   'False
         Splits(1).RecordSelectorWidth=   503
         Splits(1).AllowColSelect=   0   'False
         Splits(1).AlternatingRowStyle=   -1  'True
         Splits(1).DividerColor=   14869218
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=151"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=2646"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2514"
         Splits(1)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(5)=   "Column(1).Width=4366"
         Splits(1)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(7)=   "Column(1)._WidthInPix=4233"
         Splits(1)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(9)=   "Column(2).Width=4366"
         Splits(1)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(11)=   "Column(2)._WidthInPix=4233"
         Splits(1)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(13)=   "Column(3).Width=4366"
         Splits(1)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(15)=   "Column(3)._WidthInPix=4233"
         Splits(1)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(1)._ColumnProps(17)=   "Column(4).Width=4366"
         Splits(1)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(1)._ColumnProps(19)=   "Column(4)._WidthInPix=4233"
         Splits(1)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(1)._ColumnProps(21)=   "Column(5).Width=4366"
         Splits(1)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(1)._ColumnProps(23)=   "Column(5)._WidthInPix=4233"
         Splits(1)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(1)._ColumnProps(25)=   "Column(6).Width=4366"
         Splits(1)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(1)._ColumnProps(27)=   "Column(6)._WidthInPix=4233"
         Splits(1)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(1)._ColumnProps(29)=   "Column(7).Width=4366"
         Splits(1)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(1)._ColumnProps(31)=   "Column(7)._WidthInPix=4233"
         Splits(1)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(1)._ColumnProps(33)=   "Column(8).Width=4366"
         Splits(1)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(1)._ColumnProps(35)=   "Column(8)._WidthInPix=4233"
         Splits(1)._ColumnProps(36)=   "Column(8).Order=9"
         Splits(1)._ColumnProps(37)=   "Column(9).Width=4366"
         Splits(1)._ColumnProps(38)=   "Column(9).DividerColor=0"
         Splits(1)._ColumnProps(39)=   "Column(9)._WidthInPix=4233"
         Splits(1)._ColumnProps(40)=   "Column(9).Order=10"
         Splits(1)._ColumnProps(41)=   "Column(10).Width=4366"
         Splits(1)._ColumnProps(42)=   "Column(10).DividerColor=0"
         Splits(1)._ColumnProps(43)=   "Column(10)._WidthInPix=4233"
         Splits(1)._ColumnProps(44)=   "Column(10).Order=11"
         Splits(1)._ColumnProps(45)=   "Column(11).Width=4366"
         Splits(1)._ColumnProps(46)=   "Column(11).DividerColor=0"
         Splits(1)._ColumnProps(47)=   "Column(11)._WidthInPix=4233"
         Splits(1)._ColumnProps(48)=   "Column(11).Order=12"
         Splits(1)._ColumnProps(49)=   "Column(12).Width=4366"
         Splits(1)._ColumnProps(50)=   "Column(12).DividerColor=0"
         Splits(1)._ColumnProps(51)=   "Column(12)._WidthInPix=4233"
         Splits(1)._ColumnProps(52)=   "Column(12).Order=13"
         Splits(1)._ColumnProps(53)=   "Column(13).Width=4366"
         Splits(1)._ColumnProps(54)=   "Column(13).DividerColor=0"
         Splits(1)._ColumnProps(55)=   "Column(13)._WidthInPix=4233"
         Splits(1)._ColumnProps(56)=   "Column(13).Order=14"
         Splits(1)._ColumnProps(57)=   "Column(14).Width=4366"
         Splits(1)._ColumnProps(58)=   "Column(14).DividerColor=0"
         Splits(1)._ColumnProps(59)=   "Column(14)._WidthInPix=4233"
         Splits(1)._ColumnProps(60)=   "Column(14).Order=15"
         Splits(1)._ColumnProps(61)=   "Column(15).Width=4366"
         Splits(1)._ColumnProps(62)=   "Column(15).DividerColor=0"
         Splits(1)._ColumnProps(63)=   "Column(15)._WidthInPix=4233"
         Splits(1)._ColumnProps(64)=   "Column(15).Order=16"
         Splits(1)._ColumnProps(65)=   "Column(16).Width=4366"
         Splits(1)._ColumnProps(66)=   "Column(16).DividerColor=0"
         Splits(1)._ColumnProps(67)=   "Column(16)._WidthInPix=4233"
         Splits(1)._ColumnProps(68)=   "Column(16).Order=17"
         Splits(1)._ColumnProps(69)=   "Column(17).Width=4366"
         Splits(1)._ColumnProps(70)=   "Column(17).DividerColor=0"
         Splits(1)._ColumnProps(71)=   "Column(17)._WidthInPix=4233"
         Splits(1)._ColumnProps(72)=   "Column(17).Order=18"
         Splits(1)._ColumnProps(73)=   "Column(18).Width=4366"
         Splits(1)._ColumnProps(74)=   "Column(18).DividerColor=0"
         Splits(1)._ColumnProps(75)=   "Column(18)._WidthInPix=4233"
         Splits(1)._ColumnProps(76)=   "Column(18).Order=19"
         Splits(1)._ColumnProps(77)=   "Column(19).Width=4366"
         Splits(1)._ColumnProps(78)=   "Column(19).DividerColor=0"
         Splits(1)._ColumnProps(79)=   "Column(19)._WidthInPix=4233"
         Splits(1)._ColumnProps(80)=   "Column(19).Order=20"
         Splits(1)._ColumnProps(81)=   "Column(20).Width=4366"
         Splits(1)._ColumnProps(82)=   "Column(20).DividerColor=0"
         Splits(1)._ColumnProps(83)=   "Column(20)._WidthInPix=4233"
         Splits(1)._ColumnProps(84)=   "Column(20).Order=21"
         Splits(1)._ColumnProps(85)=   "Column(21).Width=4366"
         Splits(1)._ColumnProps(86)=   "Column(21).DividerColor=0"
         Splits(1)._ColumnProps(87)=   "Column(21)._WidthInPix=4233"
         Splits(1)._ColumnProps(88)=   "Column(21).Order=22"
         Splits(1)._ColumnProps(89)=   "Column(22).Width=4366"
         Splits(1)._ColumnProps(90)=   "Column(22).DividerColor=0"
         Splits(1)._ColumnProps(91)=   "Column(22)._WidthInPix=4233"
         Splits(1)._ColumnProps(92)=   "Column(22).Order=23"
         Splits(1)._ColumnProps(93)=   "Column(23).Width=4366"
         Splits(1)._ColumnProps(94)=   "Column(23).DividerColor=0"
         Splits(1)._ColumnProps(95)=   "Column(23)._WidthInPix=4233"
         Splits(1)._ColumnProps(96)=   "Column(23).Order=24"
         Splits(1)._ColumnProps(97)=   "Column(24).Width=4366"
         Splits(1)._ColumnProps(98)=   "Column(24).DividerColor=0"
         Splits(1)._ColumnProps(99)=   "Column(24)._WidthInPix=4233"
         Splits(1)._ColumnProps(100)=   "Column(24).Order=25"
         Splits(1)._ColumnProps(101)=   "Column(25).Width=4366"
         Splits(1)._ColumnProps(102)=   "Column(25).DividerColor=0"
         Splits(1)._ColumnProps(103)=   "Column(25)._WidthInPix=4233"
         Splits(1)._ColumnProps(104)=   "Column(25).Order=26"
         Splits(1)._ColumnProps(105)=   "Column(26).Width=4366"
         Splits(1)._ColumnProps(106)=   "Column(26).DividerColor=0"
         Splits(1)._ColumnProps(107)=   "Column(26)._WidthInPix=4233"
         Splits(1)._ColumnProps(108)=   "Column(26).Order=27"
         Splits(1)._ColumnProps(109)=   "Column(27).Width=4366"
         Splits(1)._ColumnProps(110)=   "Column(27).DividerColor=0"
         Splits(1)._ColumnProps(111)=   "Column(27)._WidthInPix=4233"
         Splits(1)._ColumnProps(112)=   "Column(27).Order=28"
         Splits(1)._ColumnProps(113)=   "Column(28).Width=4366"
         Splits(1)._ColumnProps(114)=   "Column(28).DividerColor=0"
         Splits(1)._ColumnProps(115)=   "Column(28)._WidthInPix=4233"
         Splits(1)._ColumnProps(116)=   "Column(28).Order=29"
         Splits(1)._ColumnProps(117)=   "Column(29).Width=4366"
         Splits(1)._ColumnProps(118)=   "Column(29).DividerColor=0"
         Splits(1)._ColumnProps(119)=   "Column(29)._WidthInPix=4233"
         Splits(1)._ColumnProps(120)=   "Column(29).Order=30"
         Splits(1)._ColumnProps(121)=   "Column(30).Width=4366"
         Splits(1)._ColumnProps(122)=   "Column(30).DividerColor=0"
         Splits(1)._ColumnProps(123)=   "Column(30)._WidthInPix=4233"
         Splits(1)._ColumnProps(124)=   "Column(30).Order=31"
         Splits(1)._ColumnProps(125)=   "Column(31).Width=4366"
         Splits(1)._ColumnProps(126)=   "Column(31).DividerColor=0"
         Splits(1)._ColumnProps(127)=   "Column(31)._WidthInPix=4233"
         Splits(1)._ColumnProps(128)=   "Column(31).Order=32"
         Splits(1)._ColumnProps(129)=   "Column(32).Width=4366"
         Splits(1)._ColumnProps(130)=   "Column(32).DividerColor=0"
         Splits(1)._ColumnProps(131)=   "Column(32)._WidthInPix=4233"
         Splits(1)._ColumnProps(132)=   "Column(32).Order=33"
         Splits(1)._ColumnProps(133)=   "Column(33).Width=4366"
         Splits(1)._ColumnProps(134)=   "Column(33).DividerColor=0"
         Splits(1)._ColumnProps(135)=   "Column(33)._WidthInPix=4233"
         Splits(1)._ColumnProps(136)=   "Column(33).Order=34"
         Splits(1)._ColumnProps(137)=   "Column(34).Width=4366"
         Splits(1)._ColumnProps(138)=   "Column(34).DividerColor=0"
         Splits(1)._ColumnProps(139)=   "Column(34)._WidthInPix=4233"
         Splits(1)._ColumnProps(140)=   "Column(34).Order=35"
         Splits(1)._ColumnProps(141)=   "Column(35).Width=4366"
         Splits(1)._ColumnProps(142)=   "Column(35).DividerColor=0"
         Splits(1)._ColumnProps(143)=   "Column(35)._WidthInPix=4233"
         Splits(1)._ColumnProps(144)=   "Column(35).Order=36"
         Splits(1)._ColumnProps(145)=   "Column(36).Width=4366"
         Splits(1)._ColumnProps(146)=   "Column(36).DividerColor=0"
         Splits(1)._ColumnProps(147)=   "Column(36)._WidthInPix=4233"
         Splits(1)._ColumnProps(148)=   "Column(36).Order=37"
         Splits(1)._ColumnProps(149)=   "Column(37).Width=4366"
         Splits(1)._ColumnProps(150)=   "Column(37).DividerColor=0"
         Splits(1)._ColumnProps(151)=   "Column(37)._WidthInPix=4233"
         Splits(1)._ColumnProps(152)=   "Column(37).Order=38"
         Splits(1)._ColumnProps(153)=   "Column(38).Width=4366"
         Splits(1)._ColumnProps(154)=   "Column(38).DividerColor=0"
         Splits(1)._ColumnProps(155)=   "Column(38)._WidthInPix=4233"
         Splits(1)._ColumnProps(156)=   "Column(38).Order=39"
         Splits(1)._ColumnProps(157)=   "Column(39).Width=4366"
         Splits(1)._ColumnProps(158)=   "Column(39).DividerColor=0"
         Splits(1)._ColumnProps(159)=   "Column(39)._WidthInPix=4233"
         Splits(1)._ColumnProps(160)=   "Column(39).Order=40"
         Splits(1)._ColumnProps(161)=   "Column(40).Width=4366"
         Splits(1)._ColumnProps(162)=   "Column(40).DividerColor=0"
         Splits(1)._ColumnProps(163)=   "Column(40)._WidthInPix=4233"
         Splits(1)._ColumnProps(164)=   "Column(40).Order=41"
         Splits(1)._ColumnProps(165)=   "Column(41).Width=4366"
         Splits(1)._ColumnProps(166)=   "Column(41).DividerColor=0"
         Splits(1)._ColumnProps(167)=   "Column(41)._WidthInPix=4233"
         Splits(1)._ColumnProps(168)=   "Column(41).Order=42"
         Splits(1)._ColumnProps(169)=   "Column(42).Width=4366"
         Splits(1)._ColumnProps(170)=   "Column(42).DividerColor=0"
         Splits(1)._ColumnProps(171)=   "Column(42)._WidthInPix=4233"
         Splits(1)._ColumnProps(172)=   "Column(42).Order=43"
         Splits(1)._ColumnProps(173)=   "Column(43).Width=4366"
         Splits(1)._ColumnProps(174)=   "Column(43).DividerColor=0"
         Splits(1)._ColumnProps(175)=   "Column(43)._WidthInPix=4233"
         Splits(1)._ColumnProps(176)=   "Column(43).Order=44"
         Splits(1)._ColumnProps(177)=   "Column(44).Width=4366"
         Splits(1)._ColumnProps(178)=   "Column(44).DividerColor=0"
         Splits(1)._ColumnProps(179)=   "Column(44)._WidthInPix=4233"
         Splits(1)._ColumnProps(180)=   "Column(44).Order=45"
         Splits(1)._ColumnProps(181)=   "Column(45).Width=4366"
         Splits(1)._ColumnProps(182)=   "Column(45).DividerColor=0"
         Splits(1)._ColumnProps(183)=   "Column(45)._WidthInPix=4233"
         Splits(1)._ColumnProps(184)=   "Column(45).Order=46"
         Splits(1)._ColumnProps(185)=   "Column(46).Width=4366"
         Splits(1)._ColumnProps(186)=   "Column(46).DividerColor=0"
         Splits(1)._ColumnProps(187)=   "Column(46)._WidthInPix=4233"
         Splits(1)._ColumnProps(188)=   "Column(46).Order=47"
         Splits(1)._ColumnProps(189)=   "Column(47).Width=4366"
         Splits(1)._ColumnProps(190)=   "Column(47).DividerColor=0"
         Splits(1)._ColumnProps(191)=   "Column(47)._WidthInPix=4233"
         Splits(1)._ColumnProps(192)=   "Column(47).Order=48"
         Splits(1)._ColumnProps(193)=   "Column(48).Width=4366"
         Splits(1)._ColumnProps(194)=   "Column(48).DividerColor=0"
         Splits(1)._ColumnProps(195)=   "Column(48)._WidthInPix=4233"
         Splits(1)._ColumnProps(196)=   "Column(48).Order=49"
         Splits(1)._ColumnProps(197)=   "Column(49).Width=4366"
         Splits(1)._ColumnProps(198)=   "Column(49).DividerColor=0"
         Splits(1)._ColumnProps(199)=   "Column(49)._WidthInPix=4233"
         Splits(1)._ColumnProps(200)=   "Column(49).Order=50"
         Splits(1)._ColumnProps(201)=   "Column(50).Width=4366"
         Splits(1)._ColumnProps(202)=   "Column(50).DividerColor=0"
         Splits(1)._ColumnProps(203)=   "Column(50)._WidthInPix=4233"
         Splits(1)._ColumnProps(204)=   "Column(50).Order=51"
         Splits(1)._ColumnProps(205)=   "Column(51).Width=4366"
         Splits(1)._ColumnProps(206)=   "Column(51).DividerColor=0"
         Splits(1)._ColumnProps(207)=   "Column(51)._WidthInPix=4233"
         Splits(1)._ColumnProps(208)=   "Column(51).Order=52"
         Splits(1)._ColumnProps(209)=   "Column(52).Width=4366"
         Splits(1)._ColumnProps(210)=   "Column(52).DividerColor=0"
         Splits(1)._ColumnProps(211)=   "Column(52)._WidthInPix=4233"
         Splits(1)._ColumnProps(212)=   "Column(52).Order=53"
         Splits(1)._ColumnProps(213)=   "Column(53).Width=4366"
         Splits(1)._ColumnProps(214)=   "Column(53).DividerColor=0"
         Splits(1)._ColumnProps(215)=   "Column(53)._WidthInPix=4233"
         Splits(1)._ColumnProps(216)=   "Column(53).Order=54"
         Splits(1)._ColumnProps(217)=   "Column(54).Width=4366"
         Splits(1)._ColumnProps(218)=   "Column(54).DividerColor=0"
         Splits(1)._ColumnProps(219)=   "Column(54)._WidthInPix=4233"
         Splits(1)._ColumnProps(220)=   "Column(54).Order=55"
         Splits(1)._ColumnProps(221)=   "Column(55).Width=4366"
         Splits(1)._ColumnProps(222)=   "Column(55).DividerColor=0"
         Splits(1)._ColumnProps(223)=   "Column(55)._WidthInPix=4233"
         Splits(1)._ColumnProps(224)=   "Column(55).Order=56"
         Splits(1)._ColumnProps(225)=   "Column(56).Width=4366"
         Splits(1)._ColumnProps(226)=   "Column(56).DividerColor=0"
         Splits(1)._ColumnProps(227)=   "Column(56)._WidthInPix=4233"
         Splits(1)._ColumnProps(228)=   "Column(56).Order=57"
         Splits(1)._ColumnProps(229)=   "Column(57).Width=4366"
         Splits(1)._ColumnProps(230)=   "Column(57).DividerColor=0"
         Splits(1)._ColumnProps(231)=   "Column(57)._WidthInPix=4233"
         Splits(1)._ColumnProps(232)=   "Column(57).Order=58"
         Splits(1)._ColumnProps(233)=   "Column(58).Width=4366"
         Splits(1)._ColumnProps(234)=   "Column(58).DividerColor=0"
         Splits(1)._ColumnProps(235)=   "Column(58)._WidthInPix=4233"
         Splits(1)._ColumnProps(236)=   "Column(58).Order=59"
         Splits(1)._ColumnProps(237)=   "Column(59).Width=4366"
         Splits(1)._ColumnProps(238)=   "Column(59).DividerColor=0"
         Splits(1)._ColumnProps(239)=   "Column(59)._WidthInPix=4233"
         Splits(1)._ColumnProps(240)=   "Column(59).Order=60"
         Splits(1)._ColumnProps(241)=   "Column(60).Width=4366"
         Splits(1)._ColumnProps(242)=   "Column(60).DividerColor=0"
         Splits(1)._ColumnProps(243)=   "Column(60)._WidthInPix=4233"
         Splits(1)._ColumnProps(244)=   "Column(60).Order=61"
         Splits(1)._ColumnProps(245)=   "Column(61).Width=4366"
         Splits(1)._ColumnProps(246)=   "Column(61).DividerColor=0"
         Splits(1)._ColumnProps(247)=   "Column(61)._WidthInPix=4233"
         Splits(1)._ColumnProps(248)=   "Column(61).Order=62"
         Splits(1)._ColumnProps(249)=   "Column(62).Width=4366"
         Splits(1)._ColumnProps(250)=   "Column(62).DividerColor=0"
         Splits(1)._ColumnProps(251)=   "Column(62)._WidthInPix=4233"
         Splits(1)._ColumnProps(252)=   "Column(62).Order=63"
         Splits(1)._ColumnProps(253)=   "Column(63).Width=4366"
         Splits(1)._ColumnProps(254)=   "Column(63).DividerColor=0"
         Splits(1)._ColumnProps(255)=   "Column(63)._WidthInPix=4233"
         Splits(1)._ColumnProps(256)=   "Column(63).Order=64"
         Splits(1)._ColumnProps(257)=   "Column(64).Width=4366"
         Splits(1)._ColumnProps(258)=   "Column(64).DividerColor=0"
         Splits(1)._ColumnProps(259)=   "Column(64)._WidthInPix=4233"
         Splits(1)._ColumnProps(260)=   "Column(64).Order=65"
         Splits(1)._ColumnProps(261)=   "Column(65).Width=4366"
         Splits(1)._ColumnProps(262)=   "Column(65).DividerColor=0"
         Splits(1)._ColumnProps(263)=   "Column(65)._WidthInPix=4233"
         Splits(1)._ColumnProps(264)=   "Column(65).Order=66"
         Splits(1)._ColumnProps(265)=   "Column(66).Width=4366"
         Splits(1)._ColumnProps(266)=   "Column(66).DividerColor=0"
         Splits(1)._ColumnProps(267)=   "Column(66)._WidthInPix=4233"
         Splits(1)._ColumnProps(268)=   "Column(66).Order=67"
         Splits(1)._ColumnProps(269)=   "Column(67).Width=4366"
         Splits(1)._ColumnProps(270)=   "Column(67).DividerColor=0"
         Splits(1)._ColumnProps(271)=   "Column(67)._WidthInPix=4233"
         Splits(1)._ColumnProps(272)=   "Column(67).Order=68"
         Splits(1)._ColumnProps(273)=   "Column(68).Width=4366"
         Splits(1)._ColumnProps(274)=   "Column(68).DividerColor=0"
         Splits(1)._ColumnProps(275)=   "Column(68)._WidthInPix=4233"
         Splits(1)._ColumnProps(276)=   "Column(68).Order=69"
         Splits(1)._ColumnProps(277)=   "Column(69).Width=4366"
         Splits(1)._ColumnProps(278)=   "Column(69).DividerColor=0"
         Splits(1)._ColumnProps(279)=   "Column(69)._WidthInPix=4233"
         Splits(1)._ColumnProps(280)=   "Column(69).Order=70"
         Splits(1)._ColumnProps(281)=   "Column(70).Width=4366"
         Splits(1)._ColumnProps(282)=   "Column(70).DividerColor=0"
         Splits(1)._ColumnProps(283)=   "Column(70)._WidthInPix=4233"
         Splits(1)._ColumnProps(284)=   "Column(70).Order=71"
         Splits(1)._ColumnProps(285)=   "Column(71).Width=4366"
         Splits(1)._ColumnProps(286)=   "Column(71).DividerColor=0"
         Splits(1)._ColumnProps(287)=   "Column(71)._WidthInPix=4233"
         Splits(1)._ColumnProps(288)=   "Column(71).Order=72"
         Splits(1)._ColumnProps(289)=   "Column(72).Width=4366"
         Splits(1)._ColumnProps(290)=   "Column(72).DividerColor=0"
         Splits(1)._ColumnProps(291)=   "Column(72)._WidthInPix=4233"
         Splits(1)._ColumnProps(292)=   "Column(72).Order=73"
         Splits(1)._ColumnProps(293)=   "Column(73).Width=4366"
         Splits(1)._ColumnProps(294)=   "Column(73).DividerColor=0"
         Splits(1)._ColumnProps(295)=   "Column(73)._WidthInPix=4233"
         Splits(1)._ColumnProps(296)=   "Column(73).Order=74"
         Splits(1)._ColumnProps(297)=   "Column(74).Width=4366"
         Splits(1)._ColumnProps(298)=   "Column(74).DividerColor=0"
         Splits(1)._ColumnProps(299)=   "Column(74)._WidthInPix=4233"
         Splits(1)._ColumnProps(300)=   "Column(74).Order=75"
         Splits(1)._ColumnProps(301)=   "Column(75).Width=4366"
         Splits(1)._ColumnProps(302)=   "Column(75).DividerColor=0"
         Splits(1)._ColumnProps(303)=   "Column(75)._WidthInPix=4233"
         Splits(1)._ColumnProps(304)=   "Column(75).Order=76"
         Splits(1)._ColumnProps(305)=   "Column(76).Width=4366"
         Splits(1)._ColumnProps(306)=   "Column(76).DividerColor=0"
         Splits(1)._ColumnProps(307)=   "Column(76)._WidthInPix=4233"
         Splits(1)._ColumnProps(308)=   "Column(76).Order=77"
         Splits(1)._ColumnProps(309)=   "Column(77).Width=4366"
         Splits(1)._ColumnProps(310)=   "Column(77).DividerColor=0"
         Splits(1)._ColumnProps(311)=   "Column(77)._WidthInPix=4233"
         Splits(1)._ColumnProps(312)=   "Column(77).Order=78"
         Splits(1)._ColumnProps(313)=   "Column(78).Width=4366"
         Splits(1)._ColumnProps(314)=   "Column(78).DividerColor=0"
         Splits(1)._ColumnProps(315)=   "Column(78)._WidthInPix=4233"
         Splits(1)._ColumnProps(316)=   "Column(78).Order=79"
         Splits(1)._ColumnProps(317)=   "Column(79).Width=4366"
         Splits(1)._ColumnProps(318)=   "Column(79).DividerColor=0"
         Splits(1)._ColumnProps(319)=   "Column(79)._WidthInPix=4233"
         Splits(1)._ColumnProps(320)=   "Column(79).Order=80"
         Splits(1)._ColumnProps(321)=   "Column(80).Width=4366"
         Splits(1)._ColumnProps(322)=   "Column(80).DividerColor=0"
         Splits(1)._ColumnProps(323)=   "Column(80)._WidthInPix=4233"
         Splits(1)._ColumnProps(324)=   "Column(80).Order=81"
         Splits(1)._ColumnProps(325)=   "Column(81).Width=4366"
         Splits(1)._ColumnProps(326)=   "Column(81).DividerColor=0"
         Splits(1)._ColumnProps(327)=   "Column(81)._WidthInPix=4233"
         Splits(1)._ColumnProps(328)=   "Column(81).Order=82"
         Splits(1)._ColumnProps(329)=   "Column(82).Width=4366"
         Splits(1)._ColumnProps(330)=   "Column(82).DividerColor=0"
         Splits(1)._ColumnProps(331)=   "Column(82)._WidthInPix=4233"
         Splits(1)._ColumnProps(332)=   "Column(82).Order=83"
         Splits(1)._ColumnProps(333)=   "Column(83).Width=4366"
         Splits(1)._ColumnProps(334)=   "Column(83).DividerColor=0"
         Splits(1)._ColumnProps(335)=   "Column(83)._WidthInPix=4233"
         Splits(1)._ColumnProps(336)=   "Column(83).Order=84"
         Splits(1)._ColumnProps(337)=   "Column(84).Width=4366"
         Splits(1)._ColumnProps(338)=   "Column(84).DividerColor=0"
         Splits(1)._ColumnProps(339)=   "Column(84)._WidthInPix=4233"
         Splits(1)._ColumnProps(340)=   "Column(84).Order=85"
         Splits(1)._ColumnProps(341)=   "Column(85).Width=4366"
         Splits(1)._ColumnProps(342)=   "Column(85).DividerColor=0"
         Splits(1)._ColumnProps(343)=   "Column(85)._WidthInPix=4233"
         Splits(1)._ColumnProps(344)=   "Column(85).Order=86"
         Splits(1)._ColumnProps(345)=   "Column(86).Width=4366"
         Splits(1)._ColumnProps(346)=   "Column(86).DividerColor=0"
         Splits(1)._ColumnProps(347)=   "Column(86)._WidthInPix=4233"
         Splits(1)._ColumnProps(348)=   "Column(86).Order=87"
         Splits(1)._ColumnProps(349)=   "Column(87).Width=4366"
         Splits(1)._ColumnProps(350)=   "Column(87).DividerColor=0"
         Splits(1)._ColumnProps(351)=   "Column(87)._WidthInPix=4233"
         Splits(1)._ColumnProps(352)=   "Column(87).Order=88"
         Splits(1)._ColumnProps(353)=   "Column(88).Width=4366"
         Splits(1)._ColumnProps(354)=   "Column(88).DividerColor=0"
         Splits(1)._ColumnProps(355)=   "Column(88)._WidthInPix=4233"
         Splits(1)._ColumnProps(356)=   "Column(88).Order=89"
         Splits(1)._ColumnProps(357)=   "Column(89).Width=4366"
         Splits(1)._ColumnProps(358)=   "Column(89).DividerColor=0"
         Splits(1)._ColumnProps(359)=   "Column(89)._WidthInPix=4233"
         Splits(1)._ColumnProps(360)=   "Column(89).Order=90"
         Splits(1)._ColumnProps(361)=   "Column(90).Width=4366"
         Splits(1)._ColumnProps(362)=   "Column(90).DividerColor=0"
         Splits(1)._ColumnProps(363)=   "Column(90)._WidthInPix=4233"
         Splits(1)._ColumnProps(364)=   "Column(90).Order=91"
         Splits(1)._ColumnProps(365)=   "Column(91).Width=4366"
         Splits(1)._ColumnProps(366)=   "Column(91).DividerColor=0"
         Splits(1)._ColumnProps(367)=   "Column(91)._WidthInPix=4233"
         Splits(1)._ColumnProps(368)=   "Column(91).Order=92"
         Splits(1)._ColumnProps(369)=   "Column(92).Width=4366"
         Splits(1)._ColumnProps(370)=   "Column(92).DividerColor=0"
         Splits(1)._ColumnProps(371)=   "Column(92)._WidthInPix=4233"
         Splits(1)._ColumnProps(372)=   "Column(92).Order=93"
         Splits(1)._ColumnProps(373)=   "Column(93).Width=4366"
         Splits(1)._ColumnProps(374)=   "Column(93).DividerColor=0"
         Splits(1)._ColumnProps(375)=   "Column(93)._WidthInPix=4233"
         Splits(1)._ColumnProps(376)=   "Column(93).Order=94"
         Splits(1)._ColumnProps(377)=   "Column(94).Width=4366"
         Splits(1)._ColumnProps(378)=   "Column(94).DividerColor=0"
         Splits(1)._ColumnProps(379)=   "Column(94)._WidthInPix=4233"
         Splits(1)._ColumnProps(380)=   "Column(94).Order=95"
         Splits(1)._ColumnProps(381)=   "Column(95).Width=4366"
         Splits(1)._ColumnProps(382)=   "Column(95).DividerColor=0"
         Splits(1)._ColumnProps(383)=   "Column(95)._WidthInPix=4233"
         Splits(1)._ColumnProps(384)=   "Column(95).Order=96"
         Splits(1)._ColumnProps(385)=   "Column(96).Width=4366"
         Splits(1)._ColumnProps(386)=   "Column(96).DividerColor=0"
         Splits(1)._ColumnProps(387)=   "Column(96)._WidthInPix=4233"
         Splits(1)._ColumnProps(388)=   "Column(96).Order=97"
         Splits(1)._ColumnProps(389)=   "Column(97).Width=4366"
         Splits(1)._ColumnProps(390)=   "Column(97).DividerColor=0"
         Splits(1)._ColumnProps(391)=   "Column(97)._WidthInPix=4233"
         Splits(1)._ColumnProps(392)=   "Column(97).Order=98"
         Splits(1)._ColumnProps(393)=   "Column(98).Width=4366"
         Splits(1)._ColumnProps(394)=   "Column(98).DividerColor=0"
         Splits(1)._ColumnProps(395)=   "Column(98)._WidthInPix=4233"
         Splits(1)._ColumnProps(396)=   "Column(98).Order=99"
         Splits(1)._ColumnProps(397)=   "Column(99).Width=4366"
         Splits(1)._ColumnProps(398)=   "Column(99).DividerColor=0"
         Splits(1)._ColumnProps(399)=   "Column(99)._WidthInPix=4233"
         Splits(1)._ColumnProps(400)=   "Column(99).Order=100"
         Splits(1)._ColumnProps(401)=   "Column(100).Width=4366"
         Splits(1)._ColumnProps(402)=   "Column(100).DividerColor=0"
         Splits(1)._ColumnProps(403)=   "Column(100)._WidthInPix=4233"
         Splits(1)._ColumnProps(404)=   "Column(100).Order=101"
         Splits(1)._ColumnProps(405)=   "Column(101).Width=4366"
         Splits(1)._ColumnProps(406)=   "Column(101).DividerColor=0"
         Splits(1)._ColumnProps(407)=   "Column(101)._WidthInPix=4233"
         Splits(1)._ColumnProps(408)=   "Column(101).Order=102"
         Splits(1)._ColumnProps(409)=   "Column(102).Width=4366"
         Splits(1)._ColumnProps(410)=   "Column(102).DividerColor=0"
         Splits(1)._ColumnProps(411)=   "Column(102)._WidthInPix=4233"
         Splits(1)._ColumnProps(412)=   "Column(102).Order=103"
         Splits(1)._ColumnProps(413)=   "Column(103).Width=4366"
         Splits(1)._ColumnProps(414)=   "Column(103).DividerColor=0"
         Splits(1)._ColumnProps(415)=   "Column(103)._WidthInPix=4233"
         Splits(1)._ColumnProps(416)=   "Column(103).Order=104"
         Splits(1)._ColumnProps(417)=   "Column(104).Width=4366"
         Splits(1)._ColumnProps(418)=   "Column(104).DividerColor=0"
         Splits(1)._ColumnProps(419)=   "Column(104)._WidthInPix=4233"
         Splits(1)._ColumnProps(420)=   "Column(104).Order=105"
         Splits(1)._ColumnProps(421)=   "Column(105).Width=4366"
         Splits(1)._ColumnProps(422)=   "Column(105).DividerColor=0"
         Splits(1)._ColumnProps(423)=   "Column(105)._WidthInPix=4233"
         Splits(1)._ColumnProps(424)=   "Column(105).Order=106"
         Splits(1)._ColumnProps(425)=   "Column(106).Width=4366"
         Splits(1)._ColumnProps(426)=   "Column(106).DividerColor=0"
         Splits(1)._ColumnProps(427)=   "Column(106)._WidthInPix=4233"
         Splits(1)._ColumnProps(428)=   "Column(106).Order=107"
         Splits(1)._ColumnProps(429)=   "Column(107).Width=4366"
         Splits(1)._ColumnProps(430)=   "Column(107).DividerColor=0"
         Splits(1)._ColumnProps(431)=   "Column(107)._WidthInPix=4233"
         Splits(1)._ColumnProps(432)=   "Column(107).Order=108"
         Splits(1)._ColumnProps(433)=   "Column(108).Width=4366"
         Splits(1)._ColumnProps(434)=   "Column(108).DividerColor=0"
         Splits(1)._ColumnProps(435)=   "Column(108)._WidthInPix=4233"
         Splits(1)._ColumnProps(436)=   "Column(108).Order=109"
         Splits(1)._ColumnProps(437)=   "Column(109).Width=4366"
         Splits(1)._ColumnProps(438)=   "Column(109).DividerColor=0"
         Splits(1)._ColumnProps(439)=   "Column(109)._WidthInPix=4233"
         Splits(1)._ColumnProps(440)=   "Column(109).Order=110"
         Splits(1)._ColumnProps(441)=   "Column(110).Width=4366"
         Splits(1)._ColumnProps(442)=   "Column(110).DividerColor=0"
         Splits(1)._ColumnProps(443)=   "Column(110)._WidthInPix=4233"
         Splits(1)._ColumnProps(444)=   "Column(110).Order=111"
         Splits(1)._ColumnProps(445)=   "Column(111).Width=4366"
         Splits(1)._ColumnProps(446)=   "Column(111).DividerColor=0"
         Splits(1)._ColumnProps(447)=   "Column(111)._WidthInPix=4233"
         Splits(1)._ColumnProps(448)=   "Column(111).Order=112"
         Splits(1)._ColumnProps(449)=   "Column(112).Width=4366"
         Splits(1)._ColumnProps(450)=   "Column(112).DividerColor=0"
         Splits(1)._ColumnProps(451)=   "Column(112)._WidthInPix=4233"
         Splits(1)._ColumnProps(452)=   "Column(112).Order=113"
         Splits(1)._ColumnProps(453)=   "Column(113).Width=4366"
         Splits(1)._ColumnProps(454)=   "Column(113).DividerColor=0"
         Splits(1)._ColumnProps(455)=   "Column(113)._WidthInPix=4233"
         Splits(1)._ColumnProps(456)=   "Column(113).Order=114"
         Splits(1)._ColumnProps(457)=   "Column(114).Width=4366"
         Splits(1)._ColumnProps(458)=   "Column(114).DividerColor=0"
         Splits(1)._ColumnProps(459)=   "Column(114)._WidthInPix=4233"
         Splits(1)._ColumnProps(460)=   "Column(114).Order=115"
         Splits(1)._ColumnProps(461)=   "Column(115).Width=4366"
         Splits(1)._ColumnProps(462)=   "Column(115).DividerColor=0"
         Splits(1)._ColumnProps(463)=   "Column(115)._WidthInPix=4233"
         Splits(1)._ColumnProps(464)=   "Column(115).Order=116"
         Splits(1)._ColumnProps(465)=   "Column(116).Width=4366"
         Splits(1)._ColumnProps(466)=   "Column(116).DividerColor=0"
         Splits(1)._ColumnProps(467)=   "Column(116)._WidthInPix=4233"
         Splits(1)._ColumnProps(468)=   "Column(116).Order=117"
         Splits(1)._ColumnProps(469)=   "Column(117).Width=4366"
         Splits(1)._ColumnProps(470)=   "Column(117).DividerColor=0"
         Splits(1)._ColumnProps(471)=   "Column(117)._WidthInPix=4233"
         Splits(1)._ColumnProps(472)=   "Column(117).Order=118"
         Splits(1)._ColumnProps(473)=   "Column(118).Width=4366"
         Splits(1)._ColumnProps(474)=   "Column(118).DividerColor=0"
         Splits(1)._ColumnProps(475)=   "Column(118)._WidthInPix=4233"
         Splits(1)._ColumnProps(476)=   "Column(118).Order=119"
         Splits(1)._ColumnProps(477)=   "Column(119).Width=4366"
         Splits(1)._ColumnProps(478)=   "Column(119).DividerColor=0"
         Splits(1)._ColumnProps(479)=   "Column(119)._WidthInPix=4233"
         Splits(1)._ColumnProps(480)=   "Column(119).Order=120"
         Splits(1)._ColumnProps(481)=   "Column(120).Width=4366"
         Splits(1)._ColumnProps(482)=   "Column(120).DividerColor=0"
         Splits(1)._ColumnProps(483)=   "Column(120)._WidthInPix=4233"
         Splits(1)._ColumnProps(484)=   "Column(120).Order=121"
         Splits(1)._ColumnProps(485)=   "Column(121).Width=4366"
         Splits(1)._ColumnProps(486)=   "Column(121).DividerColor=0"
         Splits(1)._ColumnProps(487)=   "Column(121)._WidthInPix=4233"
         Splits(1)._ColumnProps(488)=   "Column(121).Order=122"
         Splits(1)._ColumnProps(489)=   "Column(122).Width=4366"
         Splits(1)._ColumnProps(490)=   "Column(122).DividerColor=0"
         Splits(1)._ColumnProps(491)=   "Column(122)._WidthInPix=4233"
         Splits(1)._ColumnProps(492)=   "Column(122).Order=123"
         Splits(1)._ColumnProps(493)=   "Column(123).Width=4366"
         Splits(1)._ColumnProps(494)=   "Column(123).DividerColor=0"
         Splits(1)._ColumnProps(495)=   "Column(123)._WidthInPix=4233"
         Splits(1)._ColumnProps(496)=   "Column(123).Order=124"
         Splits(1)._ColumnProps(497)=   "Column(124).Width=4366"
         Splits(1)._ColumnProps(498)=   "Column(124).DividerColor=0"
         Splits(1)._ColumnProps(499)=   "Column(124)._WidthInPix=4233"
         Splits(1)._ColumnProps(500)=   "Column(124).Order=125"
         Splits(1)._ColumnProps(501)=   "Column(125).Width=4366"
         Splits(1)._ColumnProps(502)=   "Column(125).DividerColor=0"
         Splits(1)._ColumnProps(503)=   "Column(125)._WidthInPix=4233"
         Splits(1)._ColumnProps(504)=   "Column(125).Order=126"
         Splits(1)._ColumnProps(505)=   "Column(126).Width=4366"
         Splits(1)._ColumnProps(506)=   "Column(126).DividerColor=0"
         Splits(1)._ColumnProps(507)=   "Column(126)._WidthInPix=4233"
         Splits(1)._ColumnProps(508)=   "Column(126).Order=127"
         Splits(1)._ColumnProps(509)=   "Column(127).Width=4366"
         Splits(1)._ColumnProps(510)=   "Column(127).DividerColor=0"
         Splits(1)._ColumnProps(511)=   "Column(127)._WidthInPix=4233"
         Splits(1)._ColumnProps(512)=   "Column(127).Order=128"
         Splits(1)._ColumnProps(513)=   "Column(128).Width=4366"
         Splits(1)._ColumnProps(514)=   "Column(128).DividerColor=0"
         Splits(1)._ColumnProps(515)=   "Column(128)._WidthInPix=4233"
         Splits(1)._ColumnProps(516)=   "Column(128).Order=129"
         Splits(1)._ColumnProps(517)=   "Column(129).Width=4366"
         Splits(1)._ColumnProps(518)=   "Column(129).DividerColor=0"
         Splits(1)._ColumnProps(519)=   "Column(129)._WidthInPix=4233"
         Splits(1)._ColumnProps(520)=   "Column(129).Order=130"
         Splits(1)._ColumnProps(521)=   "Column(130).Width=4366"
         Splits(1)._ColumnProps(522)=   "Column(130).DividerColor=0"
         Splits(1)._ColumnProps(523)=   "Column(130)._WidthInPix=4233"
         Splits(1)._ColumnProps(524)=   "Column(130).Order=131"
         Splits(1)._ColumnProps(525)=   "Column(131).Width=4366"
         Splits(1)._ColumnProps(526)=   "Column(131).DividerColor=0"
         Splits(1)._ColumnProps(527)=   "Column(131)._WidthInPix=4233"
         Splits(1)._ColumnProps(528)=   "Column(131).Order=132"
         Splits(1)._ColumnProps(529)=   "Column(132).Width=4366"
         Splits(1)._ColumnProps(530)=   "Column(132).DividerColor=0"
         Splits(1)._ColumnProps(531)=   "Column(132)._WidthInPix=4233"
         Splits(1)._ColumnProps(532)=   "Column(132).Order=133"
         Splits(1)._ColumnProps(533)=   "Column(133).Width=4366"
         Splits(1)._ColumnProps(534)=   "Column(133).DividerColor=0"
         Splits(1)._ColumnProps(535)=   "Column(133)._WidthInPix=4233"
         Splits(1)._ColumnProps(536)=   "Column(133).Order=134"
         Splits(1)._ColumnProps(537)=   "Column(134).Width=4366"
         Splits(1)._ColumnProps(538)=   "Column(134).DividerColor=0"
         Splits(1)._ColumnProps(539)=   "Column(134)._WidthInPix=4233"
         Splits(1)._ColumnProps(540)=   "Column(134).Order=135"
         Splits(1)._ColumnProps(541)=   "Column(135).Width=4366"
         Splits(1)._ColumnProps(542)=   "Column(135).DividerColor=0"
         Splits(1)._ColumnProps(543)=   "Column(135)._WidthInPix=4233"
         Splits(1)._ColumnProps(544)=   "Column(135).Order=136"
         Splits(1)._ColumnProps(545)=   "Column(136).Width=4366"
         Splits(1)._ColumnProps(546)=   "Column(136).DividerColor=0"
         Splits(1)._ColumnProps(547)=   "Column(136)._WidthInPix=4233"
         Splits(1)._ColumnProps(548)=   "Column(136).Order=137"
         Splits(1)._ColumnProps(549)=   "Column(137).Width=4366"
         Splits(1)._ColumnProps(550)=   "Column(137).DividerColor=0"
         Splits(1)._ColumnProps(551)=   "Column(137)._WidthInPix=4233"
         Splits(1)._ColumnProps(552)=   "Column(137).Order=138"
         Splits(1)._ColumnProps(553)=   "Column(138).Width=4366"
         Splits(1)._ColumnProps(554)=   "Column(138).DividerColor=0"
         Splits(1)._ColumnProps(555)=   "Column(138)._WidthInPix=4233"
         Splits(1)._ColumnProps(556)=   "Column(138).Order=139"
         Splits(1)._ColumnProps(557)=   "Column(139).Width=4366"
         Splits(1)._ColumnProps(558)=   "Column(139).DividerColor=0"
         Splits(1)._ColumnProps(559)=   "Column(139)._WidthInPix=4233"
         Splits(1)._ColumnProps(560)=   "Column(139).Order=140"
         Splits(1)._ColumnProps(561)=   "Column(140).Width=4366"
         Splits(1)._ColumnProps(562)=   "Column(140).DividerColor=0"
         Splits(1)._ColumnProps(563)=   "Column(140)._WidthInPix=4233"
         Splits(1)._ColumnProps(564)=   "Column(140).Order=141"
         Splits(1)._ColumnProps(565)=   "Column(141).Width=4366"
         Splits(1)._ColumnProps(566)=   "Column(141).DividerColor=0"
         Splits(1)._ColumnProps(567)=   "Column(141)._WidthInPix=4233"
         Splits(1)._ColumnProps(568)=   "Column(141).Order=142"
         Splits(1)._ColumnProps(569)=   "Column(142).Width=4366"
         Splits(1)._ColumnProps(570)=   "Column(142).DividerColor=0"
         Splits(1)._ColumnProps(571)=   "Column(142)._WidthInPix=4233"
         Splits(1)._ColumnProps(572)=   "Column(142).Order=143"
         Splits(1)._ColumnProps(573)=   "Column(143).Width=4366"
         Splits(1)._ColumnProps(574)=   "Column(143).DividerColor=0"
         Splits(1)._ColumnProps(575)=   "Column(143)._WidthInPix=4233"
         Splits(1)._ColumnProps(576)=   "Column(143).Order=144"
         Splits(1)._ColumnProps(577)=   "Column(144).Width=4366"
         Splits(1)._ColumnProps(578)=   "Column(144).DividerColor=0"
         Splits(1)._ColumnProps(579)=   "Column(144)._WidthInPix=4233"
         Splits(1)._ColumnProps(580)=   "Column(144).Order=145"
         Splits(1)._ColumnProps(581)=   "Column(145).Width=4366"
         Splits(1)._ColumnProps(582)=   "Column(145).DividerColor=0"
         Splits(1)._ColumnProps(583)=   "Column(145)._WidthInPix=4233"
         Splits(1)._ColumnProps(584)=   "Column(145).Order=146"
         Splits(1)._ColumnProps(585)=   "Column(146).Width=4366"
         Splits(1)._ColumnProps(586)=   "Column(146).DividerColor=0"
         Splits(1)._ColumnProps(587)=   "Column(146)._WidthInPix=4233"
         Splits(1)._ColumnProps(588)=   "Column(146).Order=147"
         Splits(1)._ColumnProps(589)=   "Column(147).Width=4366"
         Splits(1)._ColumnProps(590)=   "Column(147).DividerColor=0"
         Splits(1)._ColumnProps(591)=   "Column(147)._WidthInPix=4233"
         Splits(1)._ColumnProps(592)=   "Column(147).Order=148"
         Splits(1)._ColumnProps(593)=   "Column(148).Width=4366"
         Splits(1)._ColumnProps(594)=   "Column(148).DividerColor=0"
         Splits(1)._ColumnProps(595)=   "Column(148)._WidthInPix=4233"
         Splits(1)._ColumnProps(596)=   "Column(148).Order=149"
         Splits(1)._ColumnProps(597)=   "Column(149).Width=4366"
         Splits(1)._ColumnProps(598)=   "Column(149).DividerColor=0"
         Splits(1)._ColumnProps(599)=   "Column(149)._WidthInPix=4233"
         Splits(1)._ColumnProps(600)=   "Column(149).Order=150"
         Splits(1)._ColumnProps(601)=   "Column(150).Width=4366"
         Splits(1)._ColumnProps(602)=   "Column(150).DividerColor=0"
         Splits(1)._ColumnProps(603)=   "Column(150)._WidthInPix=4233"
         Splits(1)._ColumnProps(604)=   "Column(150).Order=151"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=14.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Georgia"
         PrintInfos(0).PageFooterFont=   "Size=14.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Georgia"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MousePointer    =   16
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12577002
         RowDividerColor =   14869218
         RowSubDividerColor=   14869218
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=44,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=1425,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Georgia"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=43,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=52,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=44,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=45,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=46,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=48,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=47,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=49,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=50,.parent=9,.bgcolor=&HEAEAEA&"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=51,.parent=10,.bgcolor=&HFFFFFF&"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=53,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=54,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=2154,.parent=43"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=2151,.parent=44"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=2152,.parent=45"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=2153,.parent=47"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=2146,.parent=43"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=2143,.parent=44"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=2144,.parent=45"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=2145,.parent=47"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=2138,.parent=43"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=2135,.parent=44"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=2136,.parent=45"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=2137,.parent=47"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=2130,.parent=43"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=2127,.parent=44"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=2128,.parent=45"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=2129,.parent=47"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=2122,.parent=43"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=2119,.parent=44"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=2120,.parent=45"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=2121,.parent=47"
         _StyleDefs(52)  =   "Splits(0).Columns(5).Style:id=2114,.parent=43"
         _StyleDefs(53)  =   "Splits(0).Columns(5).HeadingStyle:id=2111,.parent=44"
         _StyleDefs(54)  =   "Splits(0).Columns(5).FooterStyle:id=2112,.parent=45"
         _StyleDefs(55)  =   "Splits(0).Columns(5).EditorStyle:id=2113,.parent=47"
         _StyleDefs(56)  =   "Splits(0).Columns(6).Style:id=2106,.parent=43"
         _StyleDefs(57)  =   "Splits(0).Columns(6).HeadingStyle:id=2103,.parent=44"
         _StyleDefs(58)  =   "Splits(0).Columns(6).FooterStyle:id=2104,.parent=45"
         _StyleDefs(59)  =   "Splits(0).Columns(6).EditorStyle:id=2105,.parent=47"
         _StyleDefs(60)  =   "Splits(0).Columns(7).Style:id=2098,.parent=43"
         _StyleDefs(61)  =   "Splits(0).Columns(7).HeadingStyle:id=2095,.parent=44"
         _StyleDefs(62)  =   "Splits(0).Columns(7).FooterStyle:id=2096,.parent=45"
         _StyleDefs(63)  =   "Splits(0).Columns(7).EditorStyle:id=2097,.parent=47"
         _StyleDefs(64)  =   "Splits(0).Columns(8).Style:id=2090,.parent=43"
         _StyleDefs(65)  =   "Splits(0).Columns(8).HeadingStyle:id=2087,.parent=44"
         _StyleDefs(66)  =   "Splits(0).Columns(8).FooterStyle:id=2088,.parent=45"
         _StyleDefs(67)  =   "Splits(0).Columns(8).EditorStyle:id=2089,.parent=47"
         _StyleDefs(68)  =   "Splits(0).Columns(9).Style:id=2082,.parent=43"
         _StyleDefs(69)  =   "Splits(0).Columns(9).HeadingStyle:id=2079,.parent=44"
         _StyleDefs(70)  =   "Splits(0).Columns(9).FooterStyle:id=2080,.parent=45"
         _StyleDefs(71)  =   "Splits(0).Columns(9).EditorStyle:id=2081,.parent=47"
         _StyleDefs(72)  =   "Splits(0).Columns(10).Style:id=2074,.parent=43"
         _StyleDefs(73)  =   "Splits(0).Columns(10).HeadingStyle:id=2071,.parent=44"
         _StyleDefs(74)  =   "Splits(0).Columns(10).FooterStyle:id=2072,.parent=45"
         _StyleDefs(75)  =   "Splits(0).Columns(10).EditorStyle:id=2073,.parent=47"
         _StyleDefs(76)  =   "Splits(0).Columns(11).Style:id=2066,.parent=43"
         _StyleDefs(77)  =   "Splits(0).Columns(11).HeadingStyle:id=2063,.parent=44"
         _StyleDefs(78)  =   "Splits(0).Columns(11).FooterStyle:id=2064,.parent=45"
         _StyleDefs(79)  =   "Splits(0).Columns(11).EditorStyle:id=2065,.parent=47"
         _StyleDefs(80)  =   "Splits(0).Columns(12).Style:id=2058,.parent=43"
         _StyleDefs(81)  =   "Splits(0).Columns(12).HeadingStyle:id=2055,.parent=44"
         _StyleDefs(82)  =   "Splits(0).Columns(12).FooterStyle:id=2056,.parent=45"
         _StyleDefs(83)  =   "Splits(0).Columns(12).EditorStyle:id=2057,.parent=47"
         _StyleDefs(84)  =   "Splits(0).Columns(13).Style:id=2050,.parent=43"
         _StyleDefs(85)  =   "Splits(0).Columns(13).HeadingStyle:id=2047,.parent=44"
         _StyleDefs(86)  =   "Splits(0).Columns(13).FooterStyle:id=2048,.parent=45"
         _StyleDefs(87)  =   "Splits(0).Columns(13).EditorStyle:id=2049,.parent=47"
         _StyleDefs(88)  =   "Splits(0).Columns(14).Style:id=2042,.parent=43"
         _StyleDefs(89)  =   "Splits(0).Columns(14).HeadingStyle:id=2039,.parent=44"
         _StyleDefs(90)  =   "Splits(0).Columns(14).FooterStyle:id=2040,.parent=45"
         _StyleDefs(91)  =   "Splits(0).Columns(14).EditorStyle:id=2041,.parent=47"
         _StyleDefs(92)  =   "Splits(0).Columns(15).Style:id=2034,.parent=43"
         _StyleDefs(93)  =   "Splits(0).Columns(15).HeadingStyle:id=2031,.parent=44"
         _StyleDefs(94)  =   "Splits(0).Columns(15).FooterStyle:id=2032,.parent=45"
         _StyleDefs(95)  =   "Splits(0).Columns(15).EditorStyle:id=2033,.parent=47"
         _StyleDefs(96)  =   "Splits(0).Columns(16).Style:id=2026,.parent=43"
         _StyleDefs(97)  =   "Splits(0).Columns(16).HeadingStyle:id=2023,.parent=44"
         _StyleDefs(98)  =   "Splits(0).Columns(16).FooterStyle:id=2024,.parent=45"
         _StyleDefs(99)  =   "Splits(0).Columns(16).EditorStyle:id=2025,.parent=47"
         _StyleDefs(100) =   "Splits(0).Columns(17).Style:id=2018,.parent=43"
         _StyleDefs(101) =   "Splits(0).Columns(17).HeadingStyle:id=2015,.parent=44"
         _StyleDefs(102) =   "Splits(0).Columns(17).FooterStyle:id=2016,.parent=45"
         _StyleDefs(103) =   "Splits(0).Columns(17).EditorStyle:id=2017,.parent=47"
         _StyleDefs(104) =   "Splits(0).Columns(18).Style:id=2010,.parent=43"
         _StyleDefs(105) =   "Splits(0).Columns(18).HeadingStyle:id=2007,.parent=44"
         _StyleDefs(106) =   "Splits(0).Columns(18).FooterStyle:id=2008,.parent=45"
         _StyleDefs(107) =   "Splits(0).Columns(18).EditorStyle:id=2009,.parent=47"
         _StyleDefs(108) =   "Splits(0).Columns(19).Style:id=2002,.parent=43"
         _StyleDefs(109) =   "Splits(0).Columns(19).HeadingStyle:id=1999,.parent=44"
         _StyleDefs(110) =   "Splits(0).Columns(19).FooterStyle:id=2000,.parent=45"
         _StyleDefs(111) =   "Splits(0).Columns(19).EditorStyle:id=2001,.parent=47"
         _StyleDefs(112) =   "Splits(0).Columns(20).Style:id=1994,.parent=43"
         _StyleDefs(113) =   "Splits(0).Columns(20).HeadingStyle:id=1991,.parent=44"
         _StyleDefs(114) =   "Splits(0).Columns(20).FooterStyle:id=1992,.parent=45"
         _StyleDefs(115) =   "Splits(0).Columns(20).EditorStyle:id=1993,.parent=47"
         _StyleDefs(116) =   "Splits(0).Columns(21).Style:id=1986,.parent=43"
         _StyleDefs(117) =   "Splits(0).Columns(21).HeadingStyle:id=1983,.parent=44"
         _StyleDefs(118) =   "Splits(0).Columns(21).FooterStyle:id=1984,.parent=45"
         _StyleDefs(119) =   "Splits(0).Columns(21).EditorStyle:id=1985,.parent=47"
         _StyleDefs(120) =   "Splits(0).Columns(22).Style:id=1978,.parent=43"
         _StyleDefs(121) =   "Splits(0).Columns(22).HeadingStyle:id=1975,.parent=44"
         _StyleDefs(122) =   "Splits(0).Columns(22).FooterStyle:id=1976,.parent=45"
         _StyleDefs(123) =   "Splits(0).Columns(22).EditorStyle:id=1977,.parent=47"
         _StyleDefs(124) =   "Splits(0).Columns(23).Style:id=1970,.parent=43"
         _StyleDefs(125) =   "Splits(0).Columns(23).HeadingStyle:id=1967,.parent=44"
         _StyleDefs(126) =   "Splits(0).Columns(23).FooterStyle:id=1968,.parent=45"
         _StyleDefs(127) =   "Splits(0).Columns(23).EditorStyle:id=1969,.parent=47"
         _StyleDefs(128) =   "Splits(0).Columns(24).Style:id=1962,.parent=43"
         _StyleDefs(129) =   "Splits(0).Columns(24).HeadingStyle:id=1959,.parent=44"
         _StyleDefs(130) =   "Splits(0).Columns(24).FooterStyle:id=1960,.parent=45"
         _StyleDefs(131) =   "Splits(0).Columns(24).EditorStyle:id=1961,.parent=47"
         _StyleDefs(132) =   "Splits(0).Columns(25).Style:id=1954,.parent=43"
         _StyleDefs(133) =   "Splits(0).Columns(25).HeadingStyle:id=1951,.parent=44"
         _StyleDefs(134) =   "Splits(0).Columns(25).FooterStyle:id=1952,.parent=45"
         _StyleDefs(135) =   "Splits(0).Columns(25).EditorStyle:id=1953,.parent=47"
         _StyleDefs(136) =   "Splits(0).Columns(26).Style:id=1946,.parent=43"
         _StyleDefs(137) =   "Splits(0).Columns(26).HeadingStyle:id=1943,.parent=44"
         _StyleDefs(138) =   "Splits(0).Columns(26).FooterStyle:id=1944,.parent=45"
         _StyleDefs(139) =   "Splits(0).Columns(26).EditorStyle:id=1945,.parent=47"
         _StyleDefs(140) =   "Splits(0).Columns(27).Style:id=1938,.parent=43"
         _StyleDefs(141) =   "Splits(0).Columns(27).HeadingStyle:id=1935,.parent=44"
         _StyleDefs(142) =   "Splits(0).Columns(27).FooterStyle:id=1936,.parent=45"
         _StyleDefs(143) =   "Splits(0).Columns(27).EditorStyle:id=1937,.parent=47"
         _StyleDefs(144) =   "Splits(0).Columns(28).Style:id=1930,.parent=43"
         _StyleDefs(145) =   "Splits(0).Columns(28).HeadingStyle:id=1927,.parent=44"
         _StyleDefs(146) =   "Splits(0).Columns(28).FooterStyle:id=1928,.parent=45"
         _StyleDefs(147) =   "Splits(0).Columns(28).EditorStyle:id=1929,.parent=47"
         _StyleDefs(148) =   "Splits(0).Columns(29).Style:id=1922,.parent=43"
         _StyleDefs(149) =   "Splits(0).Columns(29).HeadingStyle:id=1919,.parent=44"
         _StyleDefs(150) =   "Splits(0).Columns(29).FooterStyle:id=1920,.parent=45"
         _StyleDefs(151) =   "Splits(0).Columns(29).EditorStyle:id=1921,.parent=47"
         _StyleDefs(152) =   "Splits(0).Columns(30).Style:id=1914,.parent=43"
         _StyleDefs(153) =   "Splits(0).Columns(30).HeadingStyle:id=1911,.parent=44"
         _StyleDefs(154) =   "Splits(0).Columns(30).FooterStyle:id=1912,.parent=45"
         _StyleDefs(155) =   "Splits(0).Columns(30).EditorStyle:id=1913,.parent=47"
         _StyleDefs(156) =   "Splits(0).Columns(31).Style:id=1906,.parent=43"
         _StyleDefs(157) =   "Splits(0).Columns(31).HeadingStyle:id=1903,.parent=44"
         _StyleDefs(158) =   "Splits(0).Columns(31).FooterStyle:id=1904,.parent=45"
         _StyleDefs(159) =   "Splits(0).Columns(31).EditorStyle:id=1905,.parent=47"
         _StyleDefs(160) =   "Splits(0).Columns(32).Style:id=1898,.parent=43"
         _StyleDefs(161) =   "Splits(0).Columns(32).HeadingStyle:id=1895,.parent=44"
         _StyleDefs(162) =   "Splits(0).Columns(32).FooterStyle:id=1896,.parent=45"
         _StyleDefs(163) =   "Splits(0).Columns(32).EditorStyle:id=1897,.parent=47"
         _StyleDefs(164) =   "Splits(0).Columns(33).Style:id=1890,.parent=43"
         _StyleDefs(165) =   "Splits(0).Columns(33).HeadingStyle:id=1887,.parent=44"
         _StyleDefs(166) =   "Splits(0).Columns(33).FooterStyle:id=1888,.parent=45"
         _StyleDefs(167) =   "Splits(0).Columns(33).EditorStyle:id=1889,.parent=47"
         _StyleDefs(168) =   "Splits(0).Columns(34).Style:id=1882,.parent=43"
         _StyleDefs(169) =   "Splits(0).Columns(34).HeadingStyle:id=1879,.parent=44"
         _StyleDefs(170) =   "Splits(0).Columns(34).FooterStyle:id=1880,.parent=45"
         _StyleDefs(171) =   "Splits(0).Columns(34).EditorStyle:id=1881,.parent=47"
         _StyleDefs(172) =   "Splits(0).Columns(35).Style:id=1874,.parent=43"
         _StyleDefs(173) =   "Splits(0).Columns(35).HeadingStyle:id=1871,.parent=44"
         _StyleDefs(174) =   "Splits(0).Columns(35).FooterStyle:id=1872,.parent=45"
         _StyleDefs(175) =   "Splits(0).Columns(35).EditorStyle:id=1873,.parent=47"
         _StyleDefs(176) =   "Splits(0).Columns(36).Style:id=1866,.parent=43"
         _StyleDefs(177) =   "Splits(0).Columns(36).HeadingStyle:id=1863,.parent=44"
         _StyleDefs(178) =   "Splits(0).Columns(36).FooterStyle:id=1864,.parent=45"
         _StyleDefs(179) =   "Splits(0).Columns(36).EditorStyle:id=1865,.parent=47"
         _StyleDefs(180) =   "Splits(0).Columns(37).Style:id=1858,.parent=43"
         _StyleDefs(181) =   "Splits(0).Columns(37).HeadingStyle:id=1855,.parent=44"
         _StyleDefs(182) =   "Splits(0).Columns(37).FooterStyle:id=1856,.parent=45"
         _StyleDefs(183) =   "Splits(0).Columns(37).EditorStyle:id=1857,.parent=47"
         _StyleDefs(184) =   "Splits(0).Columns(38).Style:id=1850,.parent=43"
         _StyleDefs(185) =   "Splits(0).Columns(38).HeadingStyle:id=1847,.parent=44"
         _StyleDefs(186) =   "Splits(0).Columns(38).FooterStyle:id=1848,.parent=45"
         _StyleDefs(187) =   "Splits(0).Columns(38).EditorStyle:id=1849,.parent=47"
         _StyleDefs(188) =   "Splits(0).Columns(39).Style:id=1842,.parent=43"
         _StyleDefs(189) =   "Splits(0).Columns(39).HeadingStyle:id=1839,.parent=44"
         _StyleDefs(190) =   "Splits(0).Columns(39).FooterStyle:id=1840,.parent=45"
         _StyleDefs(191) =   "Splits(0).Columns(39).EditorStyle:id=1841,.parent=47"
         _StyleDefs(192) =   "Splits(0).Columns(40).Style:id=1834,.parent=43"
         _StyleDefs(193) =   "Splits(0).Columns(40).HeadingStyle:id=1831,.parent=44"
         _StyleDefs(194) =   "Splits(0).Columns(40).FooterStyle:id=1832,.parent=45"
         _StyleDefs(195) =   "Splits(0).Columns(40).EditorStyle:id=1833,.parent=47"
         _StyleDefs(196) =   "Splits(0).Columns(41).Style:id=1826,.parent=43"
         _StyleDefs(197) =   "Splits(0).Columns(41).HeadingStyle:id=1823,.parent=44"
         _StyleDefs(198) =   "Splits(0).Columns(41).FooterStyle:id=1824,.parent=45"
         _StyleDefs(199) =   "Splits(0).Columns(41).EditorStyle:id=1825,.parent=47"
         _StyleDefs(200) =   "Splits(0).Columns(42).Style:id=1818,.parent=43"
         _StyleDefs(201) =   "Splits(0).Columns(42).HeadingStyle:id=1815,.parent=44"
         _StyleDefs(202) =   "Splits(0).Columns(42).FooterStyle:id=1816,.parent=45"
         _StyleDefs(203) =   "Splits(0).Columns(42).EditorStyle:id=1817,.parent=47"
         _StyleDefs(204) =   "Splits(0).Columns(43).Style:id=1810,.parent=43"
         _StyleDefs(205) =   "Splits(0).Columns(43).HeadingStyle:id=1807,.parent=44"
         _StyleDefs(206) =   "Splits(0).Columns(43).FooterStyle:id=1808,.parent=45"
         _StyleDefs(207) =   "Splits(0).Columns(43).EditorStyle:id=1809,.parent=47"
         _StyleDefs(208) =   "Splits(0).Columns(44).Style:id=1802,.parent=43"
         _StyleDefs(209) =   "Splits(0).Columns(44).HeadingStyle:id=1799,.parent=44"
         _StyleDefs(210) =   "Splits(0).Columns(44).FooterStyle:id=1800,.parent=45"
         _StyleDefs(211) =   "Splits(0).Columns(44).EditorStyle:id=1801,.parent=47"
         _StyleDefs(212) =   "Splits(0).Columns(45).Style:id=1794,.parent=43"
         _StyleDefs(213) =   "Splits(0).Columns(45).HeadingStyle:id=1791,.parent=44"
         _StyleDefs(214) =   "Splits(0).Columns(45).FooterStyle:id=1792,.parent=45"
         _StyleDefs(215) =   "Splits(0).Columns(45).EditorStyle:id=1793,.parent=47"
         _StyleDefs(216) =   "Splits(0).Columns(46).Style:id=1786,.parent=43"
         _StyleDefs(217) =   "Splits(0).Columns(46).HeadingStyle:id=1783,.parent=44"
         _StyleDefs(218) =   "Splits(0).Columns(46).FooterStyle:id=1784,.parent=45"
         _StyleDefs(219) =   "Splits(0).Columns(46).EditorStyle:id=1785,.parent=47"
         _StyleDefs(220) =   "Splits(0).Columns(47).Style:id=1778,.parent=43"
         _StyleDefs(221) =   "Splits(0).Columns(47).HeadingStyle:id=1775,.parent=44"
         _StyleDefs(222) =   "Splits(0).Columns(47).FooterStyle:id=1776,.parent=45"
         _StyleDefs(223) =   "Splits(0).Columns(47).EditorStyle:id=1777,.parent=47"
         _StyleDefs(224) =   "Splits(0).Columns(48).Style:id=1770,.parent=43"
         _StyleDefs(225) =   "Splits(0).Columns(48).HeadingStyle:id=1767,.parent=44"
         _StyleDefs(226) =   "Splits(0).Columns(48).FooterStyle:id=1768,.parent=45"
         _StyleDefs(227) =   "Splits(0).Columns(48).EditorStyle:id=1769,.parent=47"
         _StyleDefs(228) =   "Splits(0).Columns(49).Style:id=1762,.parent=43"
         _StyleDefs(229) =   "Splits(0).Columns(49).HeadingStyle:id=1759,.parent=44"
         _StyleDefs(230) =   "Splits(0).Columns(49).FooterStyle:id=1760,.parent=45"
         _StyleDefs(231) =   "Splits(0).Columns(49).EditorStyle:id=1761,.parent=47"
         _StyleDefs(232) =   "Splits(0).Columns(50).Style:id=1754,.parent=43"
         _StyleDefs(233) =   "Splits(0).Columns(50).HeadingStyle:id=1751,.parent=44"
         _StyleDefs(234) =   "Splits(0).Columns(50).FooterStyle:id=1752,.parent=45"
         _StyleDefs(235) =   "Splits(0).Columns(50).EditorStyle:id=1753,.parent=47"
         _StyleDefs(236) =   "Splits(0).Columns(51).Style:id=1746,.parent=43"
         _StyleDefs(237) =   "Splits(0).Columns(51).HeadingStyle:id=1743,.parent=44"
         _StyleDefs(238) =   "Splits(0).Columns(51).FooterStyle:id=1744,.parent=45"
         _StyleDefs(239) =   "Splits(0).Columns(51).EditorStyle:id=1745,.parent=47"
         _StyleDefs(240) =   "Splits(0).Columns(52).Style:id=1738,.parent=43"
         _StyleDefs(241) =   "Splits(0).Columns(52).HeadingStyle:id=1735,.parent=44"
         _StyleDefs(242) =   "Splits(0).Columns(52).FooterStyle:id=1736,.parent=45"
         _StyleDefs(243) =   "Splits(0).Columns(52).EditorStyle:id=1737,.parent=47"
         _StyleDefs(244) =   "Splits(0).Columns(53).Style:id=1730,.parent=43"
         _StyleDefs(245) =   "Splits(0).Columns(53).HeadingStyle:id=1727,.parent=44"
         _StyleDefs(246) =   "Splits(0).Columns(53).FooterStyle:id=1728,.parent=45"
         _StyleDefs(247) =   "Splits(0).Columns(53).EditorStyle:id=1729,.parent=47"
         _StyleDefs(248) =   "Splits(0).Columns(54).Style:id=1722,.parent=43"
         _StyleDefs(249) =   "Splits(0).Columns(54).HeadingStyle:id=1719,.parent=44"
         _StyleDefs(250) =   "Splits(0).Columns(54).FooterStyle:id=1720,.parent=45"
         _StyleDefs(251) =   "Splits(0).Columns(54).EditorStyle:id=1721,.parent=47"
         _StyleDefs(252) =   "Splits(0).Columns(55).Style:id=1714,.parent=43"
         _StyleDefs(253) =   "Splits(0).Columns(55).HeadingStyle:id=1711,.parent=44"
         _StyleDefs(254) =   "Splits(0).Columns(55).FooterStyle:id=1712,.parent=45"
         _StyleDefs(255) =   "Splits(0).Columns(55).EditorStyle:id=1713,.parent=47"
         _StyleDefs(256) =   "Splits(0).Columns(56).Style:id=1706,.parent=43"
         _StyleDefs(257) =   "Splits(0).Columns(56).HeadingStyle:id=1703,.parent=44"
         _StyleDefs(258) =   "Splits(0).Columns(56).FooterStyle:id=1704,.parent=45"
         _StyleDefs(259) =   "Splits(0).Columns(56).EditorStyle:id=1705,.parent=47"
         _StyleDefs(260) =   "Splits(0).Columns(57).Style:id=1698,.parent=43"
         _StyleDefs(261) =   "Splits(0).Columns(57).HeadingStyle:id=1695,.parent=44"
         _StyleDefs(262) =   "Splits(0).Columns(57).FooterStyle:id=1696,.parent=45"
         _StyleDefs(263) =   "Splits(0).Columns(57).EditorStyle:id=1697,.parent=47"
         _StyleDefs(264) =   "Splits(0).Columns(58).Style:id=1690,.parent=43"
         _StyleDefs(265) =   "Splits(0).Columns(58).HeadingStyle:id=1687,.parent=44"
         _StyleDefs(266) =   "Splits(0).Columns(58).FooterStyle:id=1688,.parent=45"
         _StyleDefs(267) =   "Splits(0).Columns(58).EditorStyle:id=1689,.parent=47"
         _StyleDefs(268) =   "Splits(0).Columns(59).Style:id=1682,.parent=43"
         _StyleDefs(269) =   "Splits(0).Columns(59).HeadingStyle:id=1679,.parent=44"
         _StyleDefs(270) =   "Splits(0).Columns(59).FooterStyle:id=1680,.parent=45"
         _StyleDefs(271) =   "Splits(0).Columns(59).EditorStyle:id=1681,.parent=47"
         _StyleDefs(272) =   "Splits(0).Columns(60).Style:id=1674,.parent=43"
         _StyleDefs(273) =   "Splits(0).Columns(60).HeadingStyle:id=1671,.parent=44"
         _StyleDefs(274) =   "Splits(0).Columns(60).FooterStyle:id=1672,.parent=45"
         _StyleDefs(275) =   "Splits(0).Columns(60).EditorStyle:id=1673,.parent=47"
         _StyleDefs(276) =   "Splits(0).Columns(61).Style:id=1666,.parent=43"
         _StyleDefs(277) =   "Splits(0).Columns(61).HeadingStyle:id=1663,.parent=44"
         _StyleDefs(278) =   "Splits(0).Columns(61).FooterStyle:id=1664,.parent=45"
         _StyleDefs(279) =   "Splits(0).Columns(61).EditorStyle:id=1665,.parent=47"
         _StyleDefs(280) =   "Splits(0).Columns(62).Style:id=1658,.parent=43"
         _StyleDefs(281) =   "Splits(0).Columns(62).HeadingStyle:id=1655,.parent=44"
         _StyleDefs(282) =   "Splits(0).Columns(62).FooterStyle:id=1656,.parent=45"
         _StyleDefs(283) =   "Splits(0).Columns(62).EditorStyle:id=1657,.parent=47"
         _StyleDefs(284) =   "Splits(0).Columns(63).Style:id=1650,.parent=43"
         _StyleDefs(285) =   "Splits(0).Columns(63).HeadingStyle:id=1647,.parent=44"
         _StyleDefs(286) =   "Splits(0).Columns(63).FooterStyle:id=1648,.parent=45"
         _StyleDefs(287) =   "Splits(0).Columns(63).EditorStyle:id=1649,.parent=47"
         _StyleDefs(288) =   "Splits(0).Columns(64).Style:id=1642,.parent=43"
         _StyleDefs(289) =   "Splits(0).Columns(64).HeadingStyle:id=1639,.parent=44"
         _StyleDefs(290) =   "Splits(0).Columns(64).FooterStyle:id=1640,.parent=45"
         _StyleDefs(291) =   "Splits(0).Columns(64).EditorStyle:id=1641,.parent=47"
         _StyleDefs(292) =   "Splits(0).Columns(65).Style:id=1634,.parent=43"
         _StyleDefs(293) =   "Splits(0).Columns(65).HeadingStyle:id=1631,.parent=44"
         _StyleDefs(294) =   "Splits(0).Columns(65).FooterStyle:id=1632,.parent=45"
         _StyleDefs(295) =   "Splits(0).Columns(65).EditorStyle:id=1633,.parent=47"
         _StyleDefs(296) =   "Splits(0).Columns(66).Style:id=1626,.parent=43"
         _StyleDefs(297) =   "Splits(0).Columns(66).HeadingStyle:id=1623,.parent=44"
         _StyleDefs(298) =   "Splits(0).Columns(66).FooterStyle:id=1624,.parent=45"
         _StyleDefs(299) =   "Splits(0).Columns(66).EditorStyle:id=1625,.parent=47"
         _StyleDefs(300) =   "Splits(0).Columns(67).Style:id=1618,.parent=43"
         _StyleDefs(301) =   "Splits(0).Columns(67).HeadingStyle:id=1615,.parent=44"
         _StyleDefs(302) =   "Splits(0).Columns(67).FooterStyle:id=1616,.parent=45"
         _StyleDefs(303) =   "Splits(0).Columns(67).EditorStyle:id=1617,.parent=47"
         _StyleDefs(304) =   "Splits(0).Columns(68).Style:id=1610,.parent=43"
         _StyleDefs(305) =   "Splits(0).Columns(68).HeadingStyle:id=1607,.parent=44"
         _StyleDefs(306) =   "Splits(0).Columns(68).FooterStyle:id=1608,.parent=45"
         _StyleDefs(307) =   "Splits(0).Columns(68).EditorStyle:id=1609,.parent=47"
         _StyleDefs(308) =   "Splits(0).Columns(69).Style:id=1602,.parent=43"
         _StyleDefs(309) =   "Splits(0).Columns(69).HeadingStyle:id=1599,.parent=44"
         _StyleDefs(310) =   "Splits(0).Columns(69).FooterStyle:id=1600,.parent=45"
         _StyleDefs(311) =   "Splits(0).Columns(69).EditorStyle:id=1601,.parent=47"
         _StyleDefs(312) =   "Splits(0).Columns(70).Style:id=1594,.parent=43"
         _StyleDefs(313) =   "Splits(0).Columns(70).HeadingStyle:id=1591,.parent=44"
         _StyleDefs(314) =   "Splits(0).Columns(70).FooterStyle:id=1592,.parent=45"
         _StyleDefs(315) =   "Splits(0).Columns(70).EditorStyle:id=1593,.parent=47"
         _StyleDefs(316) =   "Splits(0).Columns(71).Style:id=1586,.parent=43"
         _StyleDefs(317) =   "Splits(0).Columns(71).HeadingStyle:id=1583,.parent=44"
         _StyleDefs(318) =   "Splits(0).Columns(71).FooterStyle:id=1584,.parent=45"
         _StyleDefs(319) =   "Splits(0).Columns(71).EditorStyle:id=1585,.parent=47"
         _StyleDefs(320) =   "Splits(0).Columns(72).Style:id=1578,.parent=43"
         _StyleDefs(321) =   "Splits(0).Columns(72).HeadingStyle:id=1575,.parent=44"
         _StyleDefs(322) =   "Splits(0).Columns(72).FooterStyle:id=1576,.parent=45"
         _StyleDefs(323) =   "Splits(0).Columns(72).EditorStyle:id=1577,.parent=47"
         _StyleDefs(324) =   "Splits(0).Columns(73).Style:id=1570,.parent=43"
         _StyleDefs(325) =   "Splits(0).Columns(73).HeadingStyle:id=1567,.parent=44"
         _StyleDefs(326) =   "Splits(0).Columns(73).FooterStyle:id=1568,.parent=45"
         _StyleDefs(327) =   "Splits(0).Columns(73).EditorStyle:id=1569,.parent=47"
         _StyleDefs(328) =   "Splits(0).Columns(74).Style:id=1562,.parent=43"
         _StyleDefs(329) =   "Splits(0).Columns(74).HeadingStyle:id=1559,.parent=44"
         _StyleDefs(330) =   "Splits(0).Columns(74).FooterStyle:id=1560,.parent=45"
         _StyleDefs(331) =   "Splits(0).Columns(74).EditorStyle:id=1561,.parent=47"
         _StyleDefs(332) =   "Splits(0).Columns(75).Style:id=1554,.parent=43"
         _StyleDefs(333) =   "Splits(0).Columns(75).HeadingStyle:id=1551,.parent=44"
         _StyleDefs(334) =   "Splits(0).Columns(75).FooterStyle:id=1552,.parent=45"
         _StyleDefs(335) =   "Splits(0).Columns(75).EditorStyle:id=1553,.parent=47"
         _StyleDefs(336) =   "Splits(0).Columns(76).Style:id=1546,.parent=43"
         _StyleDefs(337) =   "Splits(0).Columns(76).HeadingStyle:id=1543,.parent=44"
         _StyleDefs(338) =   "Splits(0).Columns(76).FooterStyle:id=1544,.parent=45"
         _StyleDefs(339) =   "Splits(0).Columns(76).EditorStyle:id=1545,.parent=47"
         _StyleDefs(340) =   "Splits(0).Columns(77).Style:id=1538,.parent=43"
         _StyleDefs(341) =   "Splits(0).Columns(77).HeadingStyle:id=1535,.parent=44"
         _StyleDefs(342) =   "Splits(0).Columns(77).FooterStyle:id=1536,.parent=45"
         _StyleDefs(343) =   "Splits(0).Columns(77).EditorStyle:id=1537,.parent=47"
         _StyleDefs(344) =   "Splits(0).Columns(78).Style:id=1530,.parent=43"
         _StyleDefs(345) =   "Splits(0).Columns(78).HeadingStyle:id=1527,.parent=44"
         _StyleDefs(346) =   "Splits(0).Columns(78).FooterStyle:id=1528,.parent=45"
         _StyleDefs(347) =   "Splits(0).Columns(78).EditorStyle:id=1529,.parent=47"
         _StyleDefs(348) =   "Splits(0).Columns(79).Style:id=1522,.parent=43"
         _StyleDefs(349) =   "Splits(0).Columns(79).HeadingStyle:id=1519,.parent=44"
         _StyleDefs(350) =   "Splits(0).Columns(79).FooterStyle:id=1520,.parent=45"
         _StyleDefs(351) =   "Splits(0).Columns(79).EditorStyle:id=1521,.parent=47"
         _StyleDefs(352) =   "Splits(0).Columns(80).Style:id=1514,.parent=43"
         _StyleDefs(353) =   "Splits(0).Columns(80).HeadingStyle:id=1511,.parent=44"
         _StyleDefs(354) =   "Splits(0).Columns(80).FooterStyle:id=1512,.parent=45"
         _StyleDefs(355) =   "Splits(0).Columns(80).EditorStyle:id=1513,.parent=47"
         _StyleDefs(356) =   "Splits(0).Columns(81).Style:id=1506,.parent=43"
         _StyleDefs(357) =   "Splits(0).Columns(81).HeadingStyle:id=1503,.parent=44"
         _StyleDefs(358) =   "Splits(0).Columns(81).FooterStyle:id=1504,.parent=45"
         _StyleDefs(359) =   "Splits(0).Columns(81).EditorStyle:id=1505,.parent=47"
         _StyleDefs(360) =   "Splits(0).Columns(82).Style:id=1498,.parent=43"
         _StyleDefs(361) =   "Splits(0).Columns(82).HeadingStyle:id=1495,.parent=44"
         _StyleDefs(362) =   "Splits(0).Columns(82).FooterStyle:id=1496,.parent=45"
         _StyleDefs(363) =   "Splits(0).Columns(82).EditorStyle:id=1497,.parent=47"
         _StyleDefs(364) =   "Splits(0).Columns(83).Style:id=1490,.parent=43"
         _StyleDefs(365) =   "Splits(0).Columns(83).HeadingStyle:id=1487,.parent=44"
         _StyleDefs(366) =   "Splits(0).Columns(83).FooterStyle:id=1488,.parent=45"
         _StyleDefs(367) =   "Splits(0).Columns(83).EditorStyle:id=1489,.parent=47"
         _StyleDefs(368) =   "Splits(0).Columns(84).Style:id=1482,.parent=43"
         _StyleDefs(369) =   "Splits(0).Columns(84).HeadingStyle:id=1479,.parent=44"
         _StyleDefs(370) =   "Splits(0).Columns(84).FooterStyle:id=1480,.parent=45"
         _StyleDefs(371) =   "Splits(0).Columns(84).EditorStyle:id=1481,.parent=47"
         _StyleDefs(372) =   "Splits(0).Columns(85).Style:id=1474,.parent=43"
         _StyleDefs(373) =   "Splits(0).Columns(85).HeadingStyle:id=1471,.parent=44"
         _StyleDefs(374) =   "Splits(0).Columns(85).FooterStyle:id=1472,.parent=45"
         _StyleDefs(375) =   "Splits(0).Columns(85).EditorStyle:id=1473,.parent=47"
         _StyleDefs(376) =   "Splits(0).Columns(86).Style:id=1466,.parent=43"
         _StyleDefs(377) =   "Splits(0).Columns(86).HeadingStyle:id=1463,.parent=44"
         _StyleDefs(378) =   "Splits(0).Columns(86).FooterStyle:id=1464,.parent=45"
         _StyleDefs(379) =   "Splits(0).Columns(86).EditorStyle:id=1465,.parent=47"
         _StyleDefs(380) =   "Splits(0).Columns(87).Style:id=1458,.parent=43"
         _StyleDefs(381) =   "Splits(0).Columns(87).HeadingStyle:id=1455,.parent=44"
         _StyleDefs(382) =   "Splits(0).Columns(87).FooterStyle:id=1456,.parent=45"
         _StyleDefs(383) =   "Splits(0).Columns(87).EditorStyle:id=1457,.parent=47"
         _StyleDefs(384) =   "Splits(0).Columns(88).Style:id=1450,.parent=43"
         _StyleDefs(385) =   "Splits(0).Columns(88).HeadingStyle:id=1447,.parent=44"
         _StyleDefs(386) =   "Splits(0).Columns(88).FooterStyle:id=1448,.parent=45"
         _StyleDefs(387) =   "Splits(0).Columns(88).EditorStyle:id=1449,.parent=47"
         _StyleDefs(388) =   "Splits(0).Columns(89).Style:id=1442,.parent=43"
         _StyleDefs(389) =   "Splits(0).Columns(89).HeadingStyle:id=1439,.parent=44"
         _StyleDefs(390) =   "Splits(0).Columns(89).FooterStyle:id=1440,.parent=45"
         _StyleDefs(391) =   "Splits(0).Columns(89).EditorStyle:id=1441,.parent=47"
         _StyleDefs(392) =   "Splits(0).Columns(90).Style:id=1434,.parent=43"
         _StyleDefs(393) =   "Splits(0).Columns(90).HeadingStyle:id=1431,.parent=44"
         _StyleDefs(394) =   "Splits(0).Columns(90).FooterStyle:id=1432,.parent=45"
         _StyleDefs(395) =   "Splits(0).Columns(90).EditorStyle:id=1433,.parent=47"
         _StyleDefs(396) =   "Splits(0).Columns(91).Style:id=1426,.parent=43"
         _StyleDefs(397) =   "Splits(0).Columns(91).HeadingStyle:id=1423,.parent=44"
         _StyleDefs(398) =   "Splits(0).Columns(91).FooterStyle:id=1424,.parent=45"
         _StyleDefs(399) =   "Splits(0).Columns(91).EditorStyle:id=1425,.parent=47"
         _StyleDefs(400) =   "Splits(0).Columns(92).Style:id=1418,.parent=43"
         _StyleDefs(401) =   "Splits(0).Columns(92).HeadingStyle:id=1415,.parent=44"
         _StyleDefs(402) =   "Splits(0).Columns(92).FooterStyle:id=1416,.parent=45"
         _StyleDefs(403) =   "Splits(0).Columns(92).EditorStyle:id=1417,.parent=47"
         _StyleDefs(404) =   "Splits(0).Columns(93).Style:id=1410,.parent=43"
         _StyleDefs(405) =   "Splits(0).Columns(93).HeadingStyle:id=1407,.parent=44"
         _StyleDefs(406) =   "Splits(0).Columns(93).FooterStyle:id=1408,.parent=45"
         _StyleDefs(407) =   "Splits(0).Columns(93).EditorStyle:id=1409,.parent=47"
         _StyleDefs(408) =   "Splits(0).Columns(94).Style:id=1402,.parent=43"
         _StyleDefs(409) =   "Splits(0).Columns(94).HeadingStyle:id=1399,.parent=44"
         _StyleDefs(410) =   "Splits(0).Columns(94).FooterStyle:id=1400,.parent=45"
         _StyleDefs(411) =   "Splits(0).Columns(94).EditorStyle:id=1401,.parent=47"
         _StyleDefs(412) =   "Splits(0).Columns(95).Style:id=1394,.parent=43"
         _StyleDefs(413) =   "Splits(0).Columns(95).HeadingStyle:id=1391,.parent=44"
         _StyleDefs(414) =   "Splits(0).Columns(95).FooterStyle:id=1392,.parent=45"
         _StyleDefs(415) =   "Splits(0).Columns(95).EditorStyle:id=1393,.parent=47"
         _StyleDefs(416) =   "Splits(0).Columns(96).Style:id=1386,.parent=43"
         _StyleDefs(417) =   "Splits(0).Columns(96).HeadingStyle:id=1383,.parent=44"
         _StyleDefs(418) =   "Splits(0).Columns(96).FooterStyle:id=1384,.parent=45"
         _StyleDefs(419) =   "Splits(0).Columns(96).EditorStyle:id=1385,.parent=47"
         _StyleDefs(420) =   "Splits(0).Columns(97).Style:id=1378,.parent=43"
         _StyleDefs(421) =   "Splits(0).Columns(97).HeadingStyle:id=1375,.parent=44"
         _StyleDefs(422) =   "Splits(0).Columns(97).FooterStyle:id=1376,.parent=45"
         _StyleDefs(423) =   "Splits(0).Columns(97).EditorStyle:id=1377,.parent=47"
         _StyleDefs(424) =   "Splits(0).Columns(98).Style:id=1370,.parent=43"
         _StyleDefs(425) =   "Splits(0).Columns(98).HeadingStyle:id=1367,.parent=44"
         _StyleDefs(426) =   "Splits(0).Columns(98).FooterStyle:id=1368,.parent=45"
         _StyleDefs(427) =   "Splits(0).Columns(98).EditorStyle:id=1369,.parent=47"
         _StyleDefs(428) =   "Splits(0).Columns(99).Style:id=1362,.parent=43"
         _StyleDefs(429) =   "Splits(0).Columns(99).HeadingStyle:id=1359,.parent=44"
         _StyleDefs(430) =   "Splits(0).Columns(99).FooterStyle:id=1360,.parent=45"
         _StyleDefs(431) =   "Splits(0).Columns(99).EditorStyle:id=1361,.parent=47"
         _StyleDefs(432) =   "Splits(0).Columns(100).Style:id=1354,.parent=43"
         _StyleDefs(433) =   "Splits(0).Columns(100).HeadingStyle:id=1351,.parent=44"
         _StyleDefs(434) =   "Splits(0).Columns(100).FooterStyle:id=1352,.parent=45"
         _StyleDefs(435) =   "Splits(0).Columns(100).EditorStyle:id=1353,.parent=47"
         _StyleDefs(436) =   "Splits(0).Columns(101).Style:id=1346,.parent=43"
         _StyleDefs(437) =   "Splits(0).Columns(101).HeadingStyle:id=1343,.parent=44"
         _StyleDefs(438) =   "Splits(0).Columns(101).FooterStyle:id=1344,.parent=45"
         _StyleDefs(439) =   "Splits(0).Columns(101).EditorStyle:id=1345,.parent=47"
         _StyleDefs(440) =   "Splits(0).Columns(102).Style:id=1338,.parent=43"
         _StyleDefs(441) =   "Splits(0).Columns(102).HeadingStyle:id=1335,.parent=44"
         _StyleDefs(442) =   "Splits(0).Columns(102).FooterStyle:id=1336,.parent=45"
         _StyleDefs(443) =   "Splits(0).Columns(102).EditorStyle:id=1337,.parent=47"
         _StyleDefs(444) =   "Splits(0).Columns(103).Style:id=1330,.parent=43"
         _StyleDefs(445) =   "Splits(0).Columns(103).HeadingStyle:id=1327,.parent=44"
         _StyleDefs(446) =   "Splits(0).Columns(103).FooterStyle:id=1328,.parent=45"
         _StyleDefs(447) =   "Splits(0).Columns(103).EditorStyle:id=1329,.parent=47"
         _StyleDefs(448) =   "Splits(0).Columns(104).Style:id=1322,.parent=43"
         _StyleDefs(449) =   "Splits(0).Columns(104).HeadingStyle:id=1319,.parent=44"
         _StyleDefs(450) =   "Splits(0).Columns(104).FooterStyle:id=1320,.parent=45"
         _StyleDefs(451) =   "Splits(0).Columns(104).EditorStyle:id=1321,.parent=47"
         _StyleDefs(452) =   "Splits(0).Columns(105).Style:id=1314,.parent=43"
         _StyleDefs(453) =   "Splits(0).Columns(105).HeadingStyle:id=1311,.parent=44"
         _StyleDefs(454) =   "Splits(0).Columns(105).FooterStyle:id=1312,.parent=45"
         _StyleDefs(455) =   "Splits(0).Columns(105).EditorStyle:id=1313,.parent=47"
         _StyleDefs(456) =   "Splits(0).Columns(106).Style:id=1306,.parent=43"
         _StyleDefs(457) =   "Splits(0).Columns(106).HeadingStyle:id=1303,.parent=44"
         _StyleDefs(458) =   "Splits(0).Columns(106).FooterStyle:id=1304,.parent=45"
         _StyleDefs(459) =   "Splits(0).Columns(106).EditorStyle:id=1305,.parent=47"
         _StyleDefs(460) =   "Splits(0).Columns(107).Style:id=1298,.parent=43"
         _StyleDefs(461) =   "Splits(0).Columns(107).HeadingStyle:id=1295,.parent=44"
         _StyleDefs(462) =   "Splits(0).Columns(107).FooterStyle:id=1296,.parent=45"
         _StyleDefs(463) =   "Splits(0).Columns(107).EditorStyle:id=1297,.parent=47"
         _StyleDefs(464) =   "Splits(0).Columns(108).Style:id=1290,.parent=43"
         _StyleDefs(465) =   "Splits(0).Columns(108).HeadingStyle:id=1287,.parent=44"
         _StyleDefs(466) =   "Splits(0).Columns(108).FooterStyle:id=1288,.parent=45"
         _StyleDefs(467) =   "Splits(0).Columns(108).EditorStyle:id=1289,.parent=47"
         _StyleDefs(468) =   "Splits(0).Columns(109).Style:id=1282,.parent=43"
         _StyleDefs(469) =   "Splits(0).Columns(109).HeadingStyle:id=1279,.parent=44"
         _StyleDefs(470) =   "Splits(0).Columns(109).FooterStyle:id=1280,.parent=45"
         _StyleDefs(471) =   "Splits(0).Columns(109).EditorStyle:id=1281,.parent=47"
         _StyleDefs(472) =   "Splits(0).Columns(110).Style:id=1274,.parent=43"
         _StyleDefs(473) =   "Splits(0).Columns(110).HeadingStyle:id=1271,.parent=44"
         _StyleDefs(474) =   "Splits(0).Columns(110).FooterStyle:id=1272,.parent=45"
         _StyleDefs(475) =   "Splits(0).Columns(110).EditorStyle:id=1273,.parent=47"
         _StyleDefs(476) =   "Splits(0).Columns(111).Style:id=1266,.parent=43"
         _StyleDefs(477) =   "Splits(0).Columns(111).HeadingStyle:id=1263,.parent=44"
         _StyleDefs(478) =   "Splits(0).Columns(111).FooterStyle:id=1264,.parent=45"
         _StyleDefs(479) =   "Splits(0).Columns(111).EditorStyle:id=1265,.parent=47"
         _StyleDefs(480) =   "Splits(0).Columns(112).Style:id=1258,.parent=43"
         _StyleDefs(481) =   "Splits(0).Columns(112).HeadingStyle:id=1255,.parent=44"
         _StyleDefs(482) =   "Splits(0).Columns(112).FooterStyle:id=1256,.parent=45"
         _StyleDefs(483) =   "Splits(0).Columns(112).EditorStyle:id=1257,.parent=47"
         _StyleDefs(484) =   "Splits(0).Columns(113).Style:id=1250,.parent=43"
         _StyleDefs(485) =   "Splits(0).Columns(113).HeadingStyle:id=1247,.parent=44"
         _StyleDefs(486) =   "Splits(0).Columns(113).FooterStyle:id=1248,.parent=45"
         _StyleDefs(487) =   "Splits(0).Columns(113).EditorStyle:id=1249,.parent=47"
         _StyleDefs(488) =   "Splits(0).Columns(114).Style:id=1242,.parent=43"
         _StyleDefs(489) =   "Splits(0).Columns(114).HeadingStyle:id=1239,.parent=44"
         _StyleDefs(490) =   "Splits(0).Columns(114).FooterStyle:id=1240,.parent=45"
         _StyleDefs(491) =   "Splits(0).Columns(114).EditorStyle:id=1241,.parent=47"
         _StyleDefs(492) =   "Splits(0).Columns(115).Style:id=1234,.parent=43"
         _StyleDefs(493) =   "Splits(0).Columns(115).HeadingStyle:id=1231,.parent=44"
         _StyleDefs(494) =   "Splits(0).Columns(115).FooterStyle:id=1232,.parent=45"
         _StyleDefs(495) =   "Splits(0).Columns(115).EditorStyle:id=1233,.parent=47"
         _StyleDefs(496) =   "Splits(0).Columns(116).Style:id=1226,.parent=43"
         _StyleDefs(497) =   "Splits(0).Columns(116).HeadingStyle:id=1223,.parent=44"
         _StyleDefs(498) =   "Splits(0).Columns(116).FooterStyle:id=1224,.parent=45"
         _StyleDefs(499) =   "Splits(0).Columns(116).EditorStyle:id=1225,.parent=47"
         _StyleDefs(500) =   "Splits(0).Columns(117).Style:id=1218,.parent=43"
         _StyleDefs(501) =   "Splits(0).Columns(117).HeadingStyle:id=1215,.parent=44"
         _StyleDefs(502) =   "Splits(0).Columns(117).FooterStyle:id=1216,.parent=45"
         _StyleDefs(503) =   "Splits(0).Columns(117).EditorStyle:id=1217,.parent=47"
         _StyleDefs(504) =   "Splits(0).Columns(118).Style:id=1210,.parent=43"
         _StyleDefs(505) =   "Splits(0).Columns(118).HeadingStyle:id=1207,.parent=44"
         _StyleDefs(506) =   "Splits(0).Columns(118).FooterStyle:id=1208,.parent=45"
         _StyleDefs(507) =   "Splits(0).Columns(118).EditorStyle:id=1209,.parent=47"
         _StyleDefs(508) =   "Splits(0).Columns(119).Style:id=1202,.parent=43"
         _StyleDefs(509) =   "Splits(0).Columns(119).HeadingStyle:id=1199,.parent=44"
         _StyleDefs(510) =   "Splits(0).Columns(119).FooterStyle:id=1200,.parent=45"
         _StyleDefs(511) =   "Splits(0).Columns(119).EditorStyle:id=1201,.parent=47"
         _StyleDefs(512) =   "Splits(0).Columns(120).Style:id=1194,.parent=43"
         _StyleDefs(513) =   "Splits(0).Columns(120).HeadingStyle:id=1191,.parent=44"
         _StyleDefs(514) =   "Splits(0).Columns(120).FooterStyle:id=1192,.parent=45"
         _StyleDefs(515) =   "Splits(0).Columns(120).EditorStyle:id=1193,.parent=47"
         _StyleDefs(516) =   "Splits(0).Columns(121).Style:id=1186,.parent=43"
         _StyleDefs(517) =   "Splits(0).Columns(121).HeadingStyle:id=1183,.parent=44"
         _StyleDefs(518) =   "Splits(0).Columns(121).FooterStyle:id=1184,.parent=45"
         _StyleDefs(519) =   "Splits(0).Columns(121).EditorStyle:id=1185,.parent=47"
         _StyleDefs(520) =   "Splits(0).Columns(122).Style:id=1178,.parent=43"
         _StyleDefs(521) =   "Splits(0).Columns(122).HeadingStyle:id=1175,.parent=44"
         _StyleDefs(522) =   "Splits(0).Columns(122).FooterStyle:id=1176,.parent=45"
         _StyleDefs(523) =   "Splits(0).Columns(122).EditorStyle:id=1177,.parent=47"
         _StyleDefs(524) =   "Splits(0).Columns(123).Style:id=1170,.parent=43"
         _StyleDefs(525) =   "Splits(0).Columns(123).HeadingStyle:id=1167,.parent=44"
         _StyleDefs(526) =   "Splits(0).Columns(123).FooterStyle:id=1168,.parent=45"
         _StyleDefs(527) =   "Splits(0).Columns(123).EditorStyle:id=1169,.parent=47"
         _StyleDefs(528) =   "Splits(0).Columns(124).Style:id=1162,.parent=43"
         _StyleDefs(529) =   "Splits(0).Columns(124).HeadingStyle:id=1159,.parent=44"
         _StyleDefs(530) =   "Splits(0).Columns(124).FooterStyle:id=1160,.parent=45"
         _StyleDefs(531) =   "Splits(0).Columns(124).EditorStyle:id=1161,.parent=47"
         _StyleDefs(532) =   "Splits(0).Columns(125).Style:id=882,.parent=43"
         _StyleDefs(533) =   "Splits(0).Columns(125).HeadingStyle:id=879,.parent=44"
         _StyleDefs(534) =   "Splits(0).Columns(125).FooterStyle:id=880,.parent=45"
         _StyleDefs(535) =   "Splits(0).Columns(125).EditorStyle:id=881,.parent=47"
         _StyleDefs(536) =   "Splits(0).Columns(126).Style:id=874,.parent=43"
         _StyleDefs(537) =   "Splits(0).Columns(126).HeadingStyle:id=871,.parent=44"
         _StyleDefs(538) =   "Splits(0).Columns(126).FooterStyle:id=872,.parent=45"
         _StyleDefs(539) =   "Splits(0).Columns(126).EditorStyle:id=873,.parent=47"
         _StyleDefs(540) =   "Splits(0).Columns(127).Style:id=866,.parent=43"
         _StyleDefs(541) =   "Splits(0).Columns(127).HeadingStyle:id=863,.parent=44"
         _StyleDefs(542) =   "Splits(0).Columns(127).FooterStyle:id=864,.parent=45"
         _StyleDefs(543) =   "Splits(0).Columns(127).EditorStyle:id=865,.parent=47"
         _StyleDefs(544) =   "Splits(0).Columns(128).Style:id=858,.parent=43"
         _StyleDefs(545) =   "Splits(0).Columns(128).HeadingStyle:id=855,.parent=44"
         _StyleDefs(546) =   "Splits(0).Columns(128).FooterStyle:id=856,.parent=45"
         _StyleDefs(547) =   "Splits(0).Columns(128).EditorStyle:id=857,.parent=47"
         _StyleDefs(548) =   "Splits(0).Columns(129).Style:id=850,.parent=43"
         _StyleDefs(549) =   "Splits(0).Columns(129).HeadingStyle:id=847,.parent=44"
         _StyleDefs(550) =   "Splits(0).Columns(129).FooterStyle:id=848,.parent=45"
         _StyleDefs(551) =   "Splits(0).Columns(129).EditorStyle:id=849,.parent=47"
         _StyleDefs(552) =   "Splits(0).Columns(130).Style:id=842,.parent=43"
         _StyleDefs(553) =   "Splits(0).Columns(130).HeadingStyle:id=839,.parent=44"
         _StyleDefs(554) =   "Splits(0).Columns(130).FooterStyle:id=840,.parent=45"
         _StyleDefs(555) =   "Splits(0).Columns(130).EditorStyle:id=841,.parent=47"
         _StyleDefs(556) =   "Splits(0).Columns(131).Style:id=834,.parent=43"
         _StyleDefs(557) =   "Splits(0).Columns(131).HeadingStyle:id=831,.parent=44"
         _StyleDefs(558) =   "Splits(0).Columns(131).FooterStyle:id=832,.parent=45"
         _StyleDefs(559) =   "Splits(0).Columns(131).EditorStyle:id=833,.parent=47"
         _StyleDefs(560) =   "Splits(0).Columns(132).Style:id=826,.parent=43"
         _StyleDefs(561) =   "Splits(0).Columns(132).HeadingStyle:id=823,.parent=44"
         _StyleDefs(562) =   "Splits(0).Columns(132).FooterStyle:id=824,.parent=45"
         _StyleDefs(563) =   "Splits(0).Columns(132).EditorStyle:id=825,.parent=47"
         _StyleDefs(564) =   "Splits(0).Columns(133).Style:id=818,.parent=43"
         _StyleDefs(565) =   "Splits(0).Columns(133).HeadingStyle:id=815,.parent=44"
         _StyleDefs(566) =   "Splits(0).Columns(133).FooterStyle:id=816,.parent=45"
         _StyleDefs(567) =   "Splits(0).Columns(133).EditorStyle:id=817,.parent=47"
         _StyleDefs(568) =   "Splits(0).Columns(134).Style:id=810,.parent=43"
         _StyleDefs(569) =   "Splits(0).Columns(134).HeadingStyle:id=807,.parent=44"
         _StyleDefs(570) =   "Splits(0).Columns(134).FooterStyle:id=808,.parent=45"
         _StyleDefs(571) =   "Splits(0).Columns(134).EditorStyle:id=809,.parent=47"
         _StyleDefs(572) =   "Splits(0).Columns(135).Style:id=802,.parent=43"
         _StyleDefs(573) =   "Splits(0).Columns(135).HeadingStyle:id=799,.parent=44"
         _StyleDefs(574) =   "Splits(0).Columns(135).FooterStyle:id=800,.parent=45"
         _StyleDefs(575) =   "Splits(0).Columns(135).EditorStyle:id=801,.parent=47"
         _StyleDefs(576) =   "Splits(0).Columns(136).Style:id=898,.parent=43"
         _StyleDefs(577) =   "Splits(0).Columns(136).HeadingStyle:id=895,.parent=44"
         _StyleDefs(578) =   "Splits(0).Columns(136).FooterStyle:id=896,.parent=45"
         _StyleDefs(579) =   "Splits(0).Columns(136).EditorStyle:id=897,.parent=47"
         _StyleDefs(580) =   "Splits(0).Columns(137).Style:id=906,.parent=43"
         _StyleDefs(581) =   "Splits(0).Columns(137).HeadingStyle:id=903,.parent=44"
         _StyleDefs(582) =   "Splits(0).Columns(137).FooterStyle:id=904,.parent=45"
         _StyleDefs(583) =   "Splits(0).Columns(137).EditorStyle:id=905,.parent=47"
         _StyleDefs(584) =   "Splits(0).Columns(138).Style:id=914,.parent=43"
         _StyleDefs(585) =   "Splits(0).Columns(138).HeadingStyle:id=911,.parent=44"
         _StyleDefs(586) =   "Splits(0).Columns(138).FooterStyle:id=912,.parent=45"
         _StyleDefs(587) =   "Splits(0).Columns(138).EditorStyle:id=913,.parent=47"
         _StyleDefs(588) =   "Splits(0).Columns(139).Style:id=922,.parent=43"
         _StyleDefs(589) =   "Splits(0).Columns(139).HeadingStyle:id=919,.parent=44"
         _StyleDefs(590) =   "Splits(0).Columns(139).FooterStyle:id=920,.parent=45"
         _StyleDefs(591) =   "Splits(0).Columns(139).EditorStyle:id=921,.parent=47"
         _StyleDefs(592) =   "Splits(0).Columns(140).Style:id=930,.parent=43"
         _StyleDefs(593) =   "Splits(0).Columns(140).HeadingStyle:id=927,.parent=44"
         _StyleDefs(594) =   "Splits(0).Columns(140).FooterStyle:id=928,.parent=45"
         _StyleDefs(595) =   "Splits(0).Columns(140).EditorStyle:id=929,.parent=47"
         _StyleDefs(596) =   "Splits(0).Columns(141).Style:id=938,.parent=43"
         _StyleDefs(597) =   "Splits(0).Columns(141).HeadingStyle:id=935,.parent=44"
         _StyleDefs(598) =   "Splits(0).Columns(141).FooterStyle:id=936,.parent=45"
         _StyleDefs(599) =   "Splits(0).Columns(141).EditorStyle:id=937,.parent=47"
         _StyleDefs(600) =   "Splits(0).Columns(142).Style:id=946,.parent=43"
         _StyleDefs(601) =   "Splits(0).Columns(142).HeadingStyle:id=943,.parent=44"
         _StyleDefs(602) =   "Splits(0).Columns(142).FooterStyle:id=944,.parent=45"
         _StyleDefs(603) =   "Splits(0).Columns(142).EditorStyle:id=945,.parent=47"
         _StyleDefs(604) =   "Splits(0).Columns(143).Style:id=954,.parent=43"
         _StyleDefs(605) =   "Splits(0).Columns(143).HeadingStyle:id=951,.parent=44"
         _StyleDefs(606) =   "Splits(0).Columns(143).FooterStyle:id=952,.parent=45"
         _StyleDefs(607) =   "Splits(0).Columns(143).EditorStyle:id=953,.parent=47"
         _StyleDefs(608) =   "Splits(0).Columns(144).Style:id=962,.parent=43"
         _StyleDefs(609) =   "Splits(0).Columns(144).HeadingStyle:id=959,.parent=44"
         _StyleDefs(610) =   "Splits(0).Columns(144).FooterStyle:id=960,.parent=45"
         _StyleDefs(611) =   "Splits(0).Columns(144).EditorStyle:id=961,.parent=47"
         _StyleDefs(612) =   "Splits(0).Columns(145).Style:id=970,.parent=43"
         _StyleDefs(613) =   "Splits(0).Columns(145).HeadingStyle:id=967,.parent=44"
         _StyleDefs(614) =   "Splits(0).Columns(145).FooterStyle:id=968,.parent=45"
         _StyleDefs(615) =   "Splits(0).Columns(145).EditorStyle:id=969,.parent=47"
         _StyleDefs(616) =   "Splits(0).Columns(146).Style:id=978,.parent=43"
         _StyleDefs(617) =   "Splits(0).Columns(146).HeadingStyle:id=975,.parent=44"
         _StyleDefs(618) =   "Splits(0).Columns(146).FooterStyle:id=976,.parent=45"
         _StyleDefs(619) =   "Splits(0).Columns(146).EditorStyle:id=977,.parent=47"
         _StyleDefs(620) =   "Splits(0).Columns(147).Style:id=986,.parent=43"
         _StyleDefs(621) =   "Splits(0).Columns(147).HeadingStyle:id=983,.parent=44"
         _StyleDefs(622) =   "Splits(0).Columns(147).FooterStyle:id=984,.parent=45"
         _StyleDefs(623) =   "Splits(0).Columns(147).EditorStyle:id=985,.parent=47"
         _StyleDefs(624) =   "Splits(0).Columns(148).Style:id=994,.parent=43"
         _StyleDefs(625) =   "Splits(0).Columns(148).HeadingStyle:id=991,.parent=44"
         _StyleDefs(626) =   "Splits(0).Columns(148).FooterStyle:id=992,.parent=45"
         _StyleDefs(627) =   "Splits(0).Columns(148).EditorStyle:id=993,.parent=47"
         _StyleDefs(628) =   "Splits(0).Columns(149).Style:id=1002,.parent=43"
         _StyleDefs(629) =   "Splits(0).Columns(149).HeadingStyle:id=999,.parent=44"
         _StyleDefs(630) =   "Splits(0).Columns(149).FooterStyle:id=1000,.parent=45"
         _StyleDefs(631) =   "Splits(0).Columns(149).EditorStyle:id=1001,.parent=47"
         _StyleDefs(632) =   "Splits(0).Columns(150).Style:id=1010,.parent=43"
         _StyleDefs(633) =   "Splits(0).Columns(150).HeadingStyle:id=1007,.parent=44"
         _StyleDefs(634) =   "Splits(0).Columns(150).FooterStyle:id=1008,.parent=45"
         _StyleDefs(635) =   "Splits(0).Columns(150).EditorStyle:id=1009,.parent=47"
         _StyleDefs(636) =   "Splits(1).Style:id=13,.parent=1"
         _StyleDefs(637) =   "Splits(1).CaptionStyle:id=22,.parent=4"
         _StyleDefs(638) =   "Splits(1).HeadingStyle:id=14,.parent=2"
         _StyleDefs(639) =   "Splits(1).FooterStyle:id=15,.parent=3"
         _StyleDefs(640) =   "Splits(1).InactiveStyle:id=16,.parent=5"
         _StyleDefs(641) =   "Splits(1).SelectedStyle:id=18,.parent=6"
         _StyleDefs(642) =   "Splits(1).EditorStyle:id=17,.parent=7"
         _StyleDefs(643) =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(644) =   "Splits(1).EvenRowStyle:id=20,.parent=9,.bgcolor=&HEAEAEA&"
         _StyleDefs(645) =   "Splits(1).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
         _StyleDefs(646) =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(647) =   "Splits(1).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(648) =   "Splits(1).Columns(0).Style:id=2158,.parent=13"
         _StyleDefs(649) =   "Splits(1).Columns(0).HeadingStyle:id=2155,.parent=14"
         _StyleDefs(650) =   "Splits(1).Columns(0).FooterStyle:id=2156,.parent=15"
         _StyleDefs(651) =   "Splits(1).Columns(0).EditorStyle:id=2157,.parent=17"
         _StyleDefs(652) =   "Splits(1).Columns(1).Style:id=2150,.parent=13"
         _StyleDefs(653) =   "Splits(1).Columns(1).HeadingStyle:id=2147,.parent=14"
         _StyleDefs(654) =   "Splits(1).Columns(1).FooterStyle:id=2148,.parent=15"
         _StyleDefs(655) =   "Splits(1).Columns(1).EditorStyle:id=2149,.parent=17"
         _StyleDefs(656) =   "Splits(1).Columns(2).Style:id=2142,.parent=13"
         _StyleDefs(657) =   "Splits(1).Columns(2).HeadingStyle:id=2139,.parent=14"
         _StyleDefs(658) =   "Splits(1).Columns(2).FooterStyle:id=2140,.parent=15"
         _StyleDefs(659) =   "Splits(1).Columns(2).EditorStyle:id=2141,.parent=17"
         _StyleDefs(660) =   "Splits(1).Columns(3).Style:id=2134,.parent=13"
         _StyleDefs(661) =   "Splits(1).Columns(3).HeadingStyle:id=2131,.parent=14"
         _StyleDefs(662) =   "Splits(1).Columns(3).FooterStyle:id=2132,.parent=15"
         _StyleDefs(663) =   "Splits(1).Columns(3).EditorStyle:id=2133,.parent=17"
         _StyleDefs(664) =   "Splits(1).Columns(4).Style:id=2126,.parent=13"
         _StyleDefs(665) =   "Splits(1).Columns(4).HeadingStyle:id=2123,.parent=14"
         _StyleDefs(666) =   "Splits(1).Columns(4).FooterStyle:id=2124,.parent=15"
         _StyleDefs(667) =   "Splits(1).Columns(4).EditorStyle:id=2125,.parent=17"
         _StyleDefs(668) =   "Splits(1).Columns(5).Style:id=2118,.parent=13"
         _StyleDefs(669) =   "Splits(1).Columns(5).HeadingStyle:id=2115,.parent=14"
         _StyleDefs(670) =   "Splits(1).Columns(5).FooterStyle:id=2116,.parent=15"
         _StyleDefs(671) =   "Splits(1).Columns(5).EditorStyle:id=2117,.parent=17"
         _StyleDefs(672) =   "Splits(1).Columns(6).Style:id=2110,.parent=13"
         _StyleDefs(673) =   "Splits(1).Columns(6).HeadingStyle:id=2107,.parent=14"
         _StyleDefs(674) =   "Splits(1).Columns(6).FooterStyle:id=2108,.parent=15"
         _StyleDefs(675) =   "Splits(1).Columns(6).EditorStyle:id=2109,.parent=17"
         _StyleDefs(676) =   "Splits(1).Columns(7).Style:id=2102,.parent=13"
         _StyleDefs(677) =   "Splits(1).Columns(7).HeadingStyle:id=2099,.parent=14"
         _StyleDefs(678) =   "Splits(1).Columns(7).FooterStyle:id=2100,.parent=15"
         _StyleDefs(679) =   "Splits(1).Columns(7).EditorStyle:id=2101,.parent=17"
         _StyleDefs(680) =   "Splits(1).Columns(8).Style:id=2094,.parent=13"
         _StyleDefs(681) =   "Splits(1).Columns(8).HeadingStyle:id=2091,.parent=14"
         _StyleDefs(682) =   "Splits(1).Columns(8).FooterStyle:id=2092,.parent=15"
         _StyleDefs(683) =   "Splits(1).Columns(8).EditorStyle:id=2093,.parent=17"
         _StyleDefs(684) =   "Splits(1).Columns(9).Style:id=2086,.parent=13"
         _StyleDefs(685) =   "Splits(1).Columns(9).HeadingStyle:id=2083,.parent=14"
         _StyleDefs(686) =   "Splits(1).Columns(9).FooterStyle:id=2084,.parent=15"
         _StyleDefs(687) =   "Splits(1).Columns(9).EditorStyle:id=2085,.parent=17"
         _StyleDefs(688) =   "Splits(1).Columns(10).Style:id=2078,.parent=13"
         _StyleDefs(689) =   "Splits(1).Columns(10).HeadingStyle:id=2075,.parent=14"
         _StyleDefs(690) =   "Splits(1).Columns(10).FooterStyle:id=2076,.parent=15"
         _StyleDefs(691) =   "Splits(1).Columns(10).EditorStyle:id=2077,.parent=17"
         _StyleDefs(692) =   "Splits(1).Columns(11).Style:id=2070,.parent=13"
         _StyleDefs(693) =   "Splits(1).Columns(11).HeadingStyle:id=2067,.parent=14"
         _StyleDefs(694) =   "Splits(1).Columns(11).FooterStyle:id=2068,.parent=15"
         _StyleDefs(695) =   "Splits(1).Columns(11).EditorStyle:id=2069,.parent=17"
         _StyleDefs(696) =   "Splits(1).Columns(12).Style:id=2062,.parent=13"
         _StyleDefs(697) =   "Splits(1).Columns(12).HeadingStyle:id=2059,.parent=14"
         _StyleDefs(698) =   "Splits(1).Columns(12).FooterStyle:id=2060,.parent=15"
         _StyleDefs(699) =   "Splits(1).Columns(12).EditorStyle:id=2061,.parent=17"
         _StyleDefs(700) =   "Splits(1).Columns(13).Style:id=2054,.parent=13"
         _StyleDefs(701) =   "Splits(1).Columns(13).HeadingStyle:id=2051,.parent=14"
         _StyleDefs(702) =   "Splits(1).Columns(13).FooterStyle:id=2052,.parent=15"
         _StyleDefs(703) =   "Splits(1).Columns(13).EditorStyle:id=2053,.parent=17"
         _StyleDefs(704) =   "Splits(1).Columns(14).Style:id=2046,.parent=13"
         _StyleDefs(705) =   "Splits(1).Columns(14).HeadingStyle:id=2043,.parent=14"
         _StyleDefs(706) =   "Splits(1).Columns(14).FooterStyle:id=2044,.parent=15"
         _StyleDefs(707) =   "Splits(1).Columns(14).EditorStyle:id=2045,.parent=17"
         _StyleDefs(708) =   "Splits(1).Columns(15).Style:id=2038,.parent=13"
         _StyleDefs(709) =   "Splits(1).Columns(15).HeadingStyle:id=2035,.parent=14"
         _StyleDefs(710) =   "Splits(1).Columns(15).FooterStyle:id=2036,.parent=15"
         _StyleDefs(711) =   "Splits(1).Columns(15).EditorStyle:id=2037,.parent=17"
         _StyleDefs(712) =   "Splits(1).Columns(16).Style:id=2030,.parent=13"
         _StyleDefs(713) =   "Splits(1).Columns(16).HeadingStyle:id=2027,.parent=14"
         _StyleDefs(714) =   "Splits(1).Columns(16).FooterStyle:id=2028,.parent=15"
         _StyleDefs(715) =   "Splits(1).Columns(16).EditorStyle:id=2029,.parent=17"
         _StyleDefs(716) =   "Splits(1).Columns(17).Style:id=2022,.parent=13"
         _StyleDefs(717) =   "Splits(1).Columns(17).HeadingStyle:id=2019,.parent=14"
         _StyleDefs(718) =   "Splits(1).Columns(17).FooterStyle:id=2020,.parent=15"
         _StyleDefs(719) =   "Splits(1).Columns(17).EditorStyle:id=2021,.parent=17"
         _StyleDefs(720) =   "Splits(1).Columns(18).Style:id=2014,.parent=13"
         _StyleDefs(721) =   "Splits(1).Columns(18).HeadingStyle:id=2011,.parent=14"
         _StyleDefs(722) =   "Splits(1).Columns(18).FooterStyle:id=2012,.parent=15"
         _StyleDefs(723) =   "Splits(1).Columns(18).EditorStyle:id=2013,.parent=17"
         _StyleDefs(724) =   "Splits(1).Columns(19).Style:id=2006,.parent=13"
         _StyleDefs(725) =   "Splits(1).Columns(19).HeadingStyle:id=2003,.parent=14"
         _StyleDefs(726) =   "Splits(1).Columns(19).FooterStyle:id=2004,.parent=15"
         _StyleDefs(727) =   "Splits(1).Columns(19).EditorStyle:id=2005,.parent=17"
         _StyleDefs(728) =   "Splits(1).Columns(20).Style:id=1998,.parent=13"
         _StyleDefs(729) =   "Splits(1).Columns(20).HeadingStyle:id=1995,.parent=14"
         _StyleDefs(730) =   "Splits(1).Columns(20).FooterStyle:id=1996,.parent=15"
         _StyleDefs(731) =   "Splits(1).Columns(20).EditorStyle:id=1997,.parent=17"
         _StyleDefs(732) =   "Splits(1).Columns(21).Style:id=1990,.parent=13"
         _StyleDefs(733) =   "Splits(1).Columns(21).HeadingStyle:id=1987,.parent=14"
         _StyleDefs(734) =   "Splits(1).Columns(21).FooterStyle:id=1988,.parent=15"
         _StyleDefs(735) =   "Splits(1).Columns(21).EditorStyle:id=1989,.parent=17"
         _StyleDefs(736) =   "Splits(1).Columns(22).Style:id=1982,.parent=13"
         _StyleDefs(737) =   "Splits(1).Columns(22).HeadingStyle:id=1979,.parent=14"
         _StyleDefs(738) =   "Splits(1).Columns(22).FooterStyle:id=1980,.parent=15"
         _StyleDefs(739) =   "Splits(1).Columns(22).EditorStyle:id=1981,.parent=17"
         _StyleDefs(740) =   "Splits(1).Columns(23).Style:id=1974,.parent=13"
         _StyleDefs(741) =   "Splits(1).Columns(23).HeadingStyle:id=1971,.parent=14"
         _StyleDefs(742) =   "Splits(1).Columns(23).FooterStyle:id=1972,.parent=15"
         _StyleDefs(743) =   "Splits(1).Columns(23).EditorStyle:id=1973,.parent=17"
         _StyleDefs(744) =   "Splits(1).Columns(24).Style:id=1966,.parent=13"
         _StyleDefs(745) =   "Splits(1).Columns(24).HeadingStyle:id=1963,.parent=14"
         _StyleDefs(746) =   "Splits(1).Columns(24).FooterStyle:id=1964,.parent=15"
         _StyleDefs(747) =   "Splits(1).Columns(24).EditorStyle:id=1965,.parent=17"
         _StyleDefs(748) =   "Splits(1).Columns(25).Style:id=1958,.parent=13"
         _StyleDefs(749) =   "Splits(1).Columns(25).HeadingStyle:id=1955,.parent=14"
         _StyleDefs(750) =   "Splits(1).Columns(25).FooterStyle:id=1956,.parent=15"
         _StyleDefs(751) =   "Splits(1).Columns(25).EditorStyle:id=1957,.parent=17"
         _StyleDefs(752) =   "Splits(1).Columns(26).Style:id=1950,.parent=13"
         _StyleDefs(753) =   "Splits(1).Columns(26).HeadingStyle:id=1947,.parent=14"
         _StyleDefs(754) =   "Splits(1).Columns(26).FooterStyle:id=1948,.parent=15"
         _StyleDefs(755) =   "Splits(1).Columns(26).EditorStyle:id=1949,.parent=17"
         _StyleDefs(756) =   "Splits(1).Columns(27).Style:id=1942,.parent=13"
         _StyleDefs(757) =   "Splits(1).Columns(27).HeadingStyle:id=1939,.parent=14"
         _StyleDefs(758) =   "Splits(1).Columns(27).FooterStyle:id=1940,.parent=15"
         _StyleDefs(759) =   "Splits(1).Columns(27).EditorStyle:id=1941,.parent=17"
         _StyleDefs(760) =   "Splits(1).Columns(28).Style:id=1934,.parent=13"
         _StyleDefs(761) =   "Splits(1).Columns(28).HeadingStyle:id=1931,.parent=14"
         _StyleDefs(762) =   "Splits(1).Columns(28).FooterStyle:id=1932,.parent=15"
         _StyleDefs(763) =   "Splits(1).Columns(28).EditorStyle:id=1933,.parent=17"
         _StyleDefs(764) =   "Splits(1).Columns(29).Style:id=1926,.parent=13"
         _StyleDefs(765) =   "Splits(1).Columns(29).HeadingStyle:id=1923,.parent=14"
         _StyleDefs(766) =   "Splits(1).Columns(29).FooterStyle:id=1924,.parent=15"
         _StyleDefs(767) =   "Splits(1).Columns(29).EditorStyle:id=1925,.parent=17"
         _StyleDefs(768) =   "Splits(1).Columns(30).Style:id=1918,.parent=13"
         _StyleDefs(769) =   "Splits(1).Columns(30).HeadingStyle:id=1915,.parent=14"
         _StyleDefs(770) =   "Splits(1).Columns(30).FooterStyle:id=1916,.parent=15"
         _StyleDefs(771) =   "Splits(1).Columns(30).EditorStyle:id=1917,.parent=17"
         _StyleDefs(772) =   "Splits(1).Columns(31).Style:id=1910,.parent=13"
         _StyleDefs(773) =   "Splits(1).Columns(31).HeadingStyle:id=1907,.parent=14"
         _StyleDefs(774) =   "Splits(1).Columns(31).FooterStyle:id=1908,.parent=15"
         _StyleDefs(775) =   "Splits(1).Columns(31).EditorStyle:id=1909,.parent=17"
         _StyleDefs(776) =   "Splits(1).Columns(32).Style:id=1902,.parent=13"
         _StyleDefs(777) =   "Splits(1).Columns(32).HeadingStyle:id=1899,.parent=14"
         _StyleDefs(778) =   "Splits(1).Columns(32).FooterStyle:id=1900,.parent=15"
         _StyleDefs(779) =   "Splits(1).Columns(32).EditorStyle:id=1901,.parent=17"
         _StyleDefs(780) =   "Splits(1).Columns(33).Style:id=1894,.parent=13"
         _StyleDefs(781) =   "Splits(1).Columns(33).HeadingStyle:id=1891,.parent=14"
         _StyleDefs(782) =   "Splits(1).Columns(33).FooterStyle:id=1892,.parent=15"
         _StyleDefs(783) =   "Splits(1).Columns(33).EditorStyle:id=1893,.parent=17"
         _StyleDefs(784) =   "Splits(1).Columns(34).Style:id=1886,.parent=13"
         _StyleDefs(785) =   "Splits(1).Columns(34).HeadingStyle:id=1883,.parent=14"
         _StyleDefs(786) =   "Splits(1).Columns(34).FooterStyle:id=1884,.parent=15"
         _StyleDefs(787) =   "Splits(1).Columns(34).EditorStyle:id=1885,.parent=17"
         _StyleDefs(788) =   "Splits(1).Columns(35).Style:id=1878,.parent=13"
         _StyleDefs(789) =   "Splits(1).Columns(35).HeadingStyle:id=1875,.parent=14"
         _StyleDefs(790) =   "Splits(1).Columns(35).FooterStyle:id=1876,.parent=15"
         _StyleDefs(791) =   "Splits(1).Columns(35).EditorStyle:id=1877,.parent=17"
         _StyleDefs(792) =   "Splits(1).Columns(36).Style:id=1870,.parent=13"
         _StyleDefs(793) =   "Splits(1).Columns(36).HeadingStyle:id=1867,.parent=14"
         _StyleDefs(794) =   "Splits(1).Columns(36).FooterStyle:id=1868,.parent=15"
         _StyleDefs(795) =   "Splits(1).Columns(36).EditorStyle:id=1869,.parent=17"
         _StyleDefs(796) =   "Splits(1).Columns(37).Style:id=1862,.parent=13"
         _StyleDefs(797) =   "Splits(1).Columns(37).HeadingStyle:id=1859,.parent=14"
         _StyleDefs(798) =   "Splits(1).Columns(37).FooterStyle:id=1860,.parent=15"
         _StyleDefs(799) =   "Splits(1).Columns(37).EditorStyle:id=1861,.parent=17"
         _StyleDefs(800) =   "Splits(1).Columns(38).Style:id=1854,.parent=13"
         _StyleDefs(801) =   "Splits(1).Columns(38).HeadingStyle:id=1851,.parent=14"
         _StyleDefs(802) =   "Splits(1).Columns(38).FooterStyle:id=1852,.parent=15"
         _StyleDefs(803) =   "Splits(1).Columns(38).EditorStyle:id=1853,.parent=17"
         _StyleDefs(804) =   "Splits(1).Columns(39).Style:id=1846,.parent=13"
         _StyleDefs(805) =   "Splits(1).Columns(39).HeadingStyle:id=1843,.parent=14"
         _StyleDefs(806) =   "Splits(1).Columns(39).FooterStyle:id=1844,.parent=15"
         _StyleDefs(807) =   "Splits(1).Columns(39).EditorStyle:id=1845,.parent=17"
         _StyleDefs(808) =   "Splits(1).Columns(40).Style:id=1838,.parent=13"
         _StyleDefs(809) =   "Splits(1).Columns(40).HeadingStyle:id=1835,.parent=14"
         _StyleDefs(810) =   "Splits(1).Columns(40).FooterStyle:id=1836,.parent=15"
         _StyleDefs(811) =   "Splits(1).Columns(40).EditorStyle:id=1837,.parent=17"
         _StyleDefs(812) =   "Splits(1).Columns(41).Style:id=1830,.parent=13"
         _StyleDefs(813) =   "Splits(1).Columns(41).HeadingStyle:id=1827,.parent=14"
         _StyleDefs(814) =   "Splits(1).Columns(41).FooterStyle:id=1828,.parent=15"
         _StyleDefs(815) =   "Splits(1).Columns(41).EditorStyle:id=1829,.parent=17"
         _StyleDefs(816) =   "Splits(1).Columns(42).Style:id=1822,.parent=13"
         _StyleDefs(817) =   "Splits(1).Columns(42).HeadingStyle:id=1819,.parent=14"
         _StyleDefs(818) =   "Splits(1).Columns(42).FooterStyle:id=1820,.parent=15"
         _StyleDefs(819) =   "Splits(1).Columns(42).EditorStyle:id=1821,.parent=17"
         _StyleDefs(820) =   "Splits(1).Columns(43).Style:id=1814,.parent=13"
         _StyleDefs(821) =   "Splits(1).Columns(43).HeadingStyle:id=1811,.parent=14"
         _StyleDefs(822) =   "Splits(1).Columns(43).FooterStyle:id=1812,.parent=15"
         _StyleDefs(823) =   "Splits(1).Columns(43).EditorStyle:id=1813,.parent=17"
         _StyleDefs(824) =   "Splits(1).Columns(44).Style:id=1806,.parent=13"
         _StyleDefs(825) =   "Splits(1).Columns(44).HeadingStyle:id=1803,.parent=14"
         _StyleDefs(826) =   "Splits(1).Columns(44).FooterStyle:id=1804,.parent=15"
         _StyleDefs(827) =   "Splits(1).Columns(44).EditorStyle:id=1805,.parent=17"
         _StyleDefs(828) =   "Splits(1).Columns(45).Style:id=1798,.parent=13"
         _StyleDefs(829) =   "Splits(1).Columns(45).HeadingStyle:id=1795,.parent=14"
         _StyleDefs(830) =   "Splits(1).Columns(45).FooterStyle:id=1796,.parent=15"
         _StyleDefs(831) =   "Splits(1).Columns(45).EditorStyle:id=1797,.parent=17"
         _StyleDefs(832) =   "Splits(1).Columns(46).Style:id=1790,.parent=13"
         _StyleDefs(833) =   "Splits(1).Columns(46).HeadingStyle:id=1787,.parent=14"
         _StyleDefs(834) =   "Splits(1).Columns(46).FooterStyle:id=1788,.parent=15"
         _StyleDefs(835) =   "Splits(1).Columns(46).EditorStyle:id=1789,.parent=17"
         _StyleDefs(836) =   "Splits(1).Columns(47).Style:id=1782,.parent=13"
         _StyleDefs(837) =   "Splits(1).Columns(47).HeadingStyle:id=1779,.parent=14"
         _StyleDefs(838) =   "Splits(1).Columns(47).FooterStyle:id=1780,.parent=15"
         _StyleDefs(839) =   "Splits(1).Columns(47).EditorStyle:id=1781,.parent=17"
         _StyleDefs(840) =   "Splits(1).Columns(48).Style:id=1774,.parent=13"
         _StyleDefs(841) =   "Splits(1).Columns(48).HeadingStyle:id=1771,.parent=14"
         _StyleDefs(842) =   "Splits(1).Columns(48).FooterStyle:id=1772,.parent=15"
         _StyleDefs(843) =   "Splits(1).Columns(48).EditorStyle:id=1773,.parent=17"
         _StyleDefs(844) =   "Splits(1).Columns(49).Style:id=1766,.parent=13"
         _StyleDefs(845) =   "Splits(1).Columns(49).HeadingStyle:id=1763,.parent=14"
         _StyleDefs(846) =   "Splits(1).Columns(49).FooterStyle:id=1764,.parent=15"
         _StyleDefs(847) =   "Splits(1).Columns(49).EditorStyle:id=1765,.parent=17"
         _StyleDefs(848) =   "Splits(1).Columns(50).Style:id=1758,.parent=13"
         _StyleDefs(849) =   "Splits(1).Columns(50).HeadingStyle:id=1755,.parent=14"
         _StyleDefs(850) =   "Splits(1).Columns(50).FooterStyle:id=1756,.parent=15"
         _StyleDefs(851) =   "Splits(1).Columns(50).EditorStyle:id=1757,.parent=17"
         _StyleDefs(852) =   "Splits(1).Columns(51).Style:id=1750,.parent=13"
         _StyleDefs(853) =   "Splits(1).Columns(51).HeadingStyle:id=1747,.parent=14"
         _StyleDefs(854) =   "Splits(1).Columns(51).FooterStyle:id=1748,.parent=15"
         _StyleDefs(855) =   "Splits(1).Columns(51).EditorStyle:id=1749,.parent=17"
         _StyleDefs(856) =   "Splits(1).Columns(52).Style:id=1742,.parent=13"
         _StyleDefs(857) =   "Splits(1).Columns(52).HeadingStyle:id=1739,.parent=14"
         _StyleDefs(858) =   "Splits(1).Columns(52).FooterStyle:id=1740,.parent=15"
         _StyleDefs(859) =   "Splits(1).Columns(52).EditorStyle:id=1741,.parent=17"
         _StyleDefs(860) =   "Splits(1).Columns(53).Style:id=1734,.parent=13"
         _StyleDefs(861) =   "Splits(1).Columns(53).HeadingStyle:id=1731,.parent=14"
         _StyleDefs(862) =   "Splits(1).Columns(53).FooterStyle:id=1732,.parent=15"
         _StyleDefs(863) =   "Splits(1).Columns(53).EditorStyle:id=1733,.parent=17"
         _StyleDefs(864) =   "Splits(1).Columns(54).Style:id=1726,.parent=13"
         _StyleDefs(865) =   "Splits(1).Columns(54).HeadingStyle:id=1723,.parent=14"
         _StyleDefs(866) =   "Splits(1).Columns(54).FooterStyle:id=1724,.parent=15"
         _StyleDefs(867) =   "Splits(1).Columns(54).EditorStyle:id=1725,.parent=17"
         _StyleDefs(868) =   "Splits(1).Columns(55).Style:id=1718,.parent=13"
         _StyleDefs(869) =   "Splits(1).Columns(55).HeadingStyle:id=1715,.parent=14"
         _StyleDefs(870) =   "Splits(1).Columns(55).FooterStyle:id=1716,.parent=15"
         _StyleDefs(871) =   "Splits(1).Columns(55).EditorStyle:id=1717,.parent=17"
         _StyleDefs(872) =   "Splits(1).Columns(56).Style:id=1710,.parent=13"
         _StyleDefs(873) =   "Splits(1).Columns(56).HeadingStyle:id=1707,.parent=14"
         _StyleDefs(874) =   "Splits(1).Columns(56).FooterStyle:id=1708,.parent=15"
         _StyleDefs(875) =   "Splits(1).Columns(56).EditorStyle:id=1709,.parent=17"
         _StyleDefs(876) =   "Splits(1).Columns(57).Style:id=1702,.parent=13"
         _StyleDefs(877) =   "Splits(1).Columns(57).HeadingStyle:id=1699,.parent=14"
         _StyleDefs(878) =   "Splits(1).Columns(57).FooterStyle:id=1700,.parent=15"
         _StyleDefs(879) =   "Splits(1).Columns(57).EditorStyle:id=1701,.parent=17"
         _StyleDefs(880) =   "Splits(1).Columns(58).Style:id=1694,.parent=13"
         _StyleDefs(881) =   "Splits(1).Columns(58).HeadingStyle:id=1691,.parent=14"
         _StyleDefs(882) =   "Splits(1).Columns(58).FooterStyle:id=1692,.parent=15"
         _StyleDefs(883) =   "Splits(1).Columns(58).EditorStyle:id=1693,.parent=17"
         _StyleDefs(884) =   "Splits(1).Columns(59).Style:id=1686,.parent=13"
         _StyleDefs(885) =   "Splits(1).Columns(59).HeadingStyle:id=1683,.parent=14"
         _StyleDefs(886) =   "Splits(1).Columns(59).FooterStyle:id=1684,.parent=15"
         _StyleDefs(887) =   "Splits(1).Columns(59).EditorStyle:id=1685,.parent=17"
         _StyleDefs(888) =   "Splits(1).Columns(60).Style:id=1678,.parent=13"
         _StyleDefs(889) =   "Splits(1).Columns(60).HeadingStyle:id=1675,.parent=14"
         _StyleDefs(890) =   "Splits(1).Columns(60).FooterStyle:id=1676,.parent=15"
         _StyleDefs(891) =   "Splits(1).Columns(60).EditorStyle:id=1677,.parent=17"
         _StyleDefs(892) =   "Splits(1).Columns(61).Style:id=1670,.parent=13"
         _StyleDefs(893) =   "Splits(1).Columns(61).HeadingStyle:id=1667,.parent=14"
         _StyleDefs(894) =   "Splits(1).Columns(61).FooterStyle:id=1668,.parent=15"
         _StyleDefs(895) =   "Splits(1).Columns(61).EditorStyle:id=1669,.parent=17"
         _StyleDefs(896) =   "Splits(1).Columns(62).Style:id=1662,.parent=13"
         _StyleDefs(897) =   "Splits(1).Columns(62).HeadingStyle:id=1659,.parent=14"
         _StyleDefs(898) =   "Splits(1).Columns(62).FooterStyle:id=1660,.parent=15"
         _StyleDefs(899) =   "Splits(1).Columns(62).EditorStyle:id=1661,.parent=17"
         _StyleDefs(900) =   "Splits(1).Columns(63).Style:id=1654,.parent=13"
         _StyleDefs(901) =   "Splits(1).Columns(63).HeadingStyle:id=1651,.parent=14"
         _StyleDefs(902) =   "Splits(1).Columns(63).FooterStyle:id=1652,.parent=15"
         _StyleDefs(903) =   "Splits(1).Columns(63).EditorStyle:id=1653,.parent=17"
         _StyleDefs(904) =   "Splits(1).Columns(64).Style:id=1646,.parent=13"
         _StyleDefs(905) =   "Splits(1).Columns(64).HeadingStyle:id=1643,.parent=14"
         _StyleDefs(906) =   "Splits(1).Columns(64).FooterStyle:id=1644,.parent=15"
         _StyleDefs(907) =   "Splits(1).Columns(64).EditorStyle:id=1645,.parent=17"
         _StyleDefs(908) =   "Splits(1).Columns(65).Style:id=1638,.parent=13"
         _StyleDefs(909) =   "Splits(1).Columns(65).HeadingStyle:id=1635,.parent=14"
         _StyleDefs(910) =   "Splits(1).Columns(65).FooterStyle:id=1636,.parent=15"
         _StyleDefs(911) =   "Splits(1).Columns(65).EditorStyle:id=1637,.parent=17"
         _StyleDefs(912) =   "Splits(1).Columns(66).Style:id=1630,.parent=13"
         _StyleDefs(913) =   "Splits(1).Columns(66).HeadingStyle:id=1627,.parent=14"
         _StyleDefs(914) =   "Splits(1).Columns(66).FooterStyle:id=1628,.parent=15"
         _StyleDefs(915) =   "Splits(1).Columns(66).EditorStyle:id=1629,.parent=17"
         _StyleDefs(916) =   "Splits(1).Columns(67).Style:id=1622,.parent=13"
         _StyleDefs(917) =   "Splits(1).Columns(67).HeadingStyle:id=1619,.parent=14"
         _StyleDefs(918) =   "Splits(1).Columns(67).FooterStyle:id=1620,.parent=15"
         _StyleDefs(919) =   "Splits(1).Columns(67).EditorStyle:id=1621,.parent=17"
         _StyleDefs(920) =   "Splits(1).Columns(68).Style:id=1614,.parent=13"
         _StyleDefs(921) =   "Splits(1).Columns(68).HeadingStyle:id=1611,.parent=14"
         _StyleDefs(922) =   "Splits(1).Columns(68).FooterStyle:id=1612,.parent=15"
         _StyleDefs(923) =   "Splits(1).Columns(68).EditorStyle:id=1613,.parent=17"
         _StyleDefs(924) =   "Splits(1).Columns(69).Style:id=1606,.parent=13"
         _StyleDefs(925) =   "Splits(1).Columns(69).HeadingStyle:id=1603,.parent=14"
         _StyleDefs(926) =   "Splits(1).Columns(69).FooterStyle:id=1604,.parent=15"
         _StyleDefs(927) =   "Splits(1).Columns(69).EditorStyle:id=1605,.parent=17"
         _StyleDefs(928) =   "Splits(1).Columns(70).Style:id=1598,.parent=13"
         _StyleDefs(929) =   "Splits(1).Columns(70).HeadingStyle:id=1595,.parent=14"
         _StyleDefs(930) =   "Splits(1).Columns(70).FooterStyle:id=1596,.parent=15"
         _StyleDefs(931) =   "Splits(1).Columns(70).EditorStyle:id=1597,.parent=17"
         _StyleDefs(932) =   "Splits(1).Columns(71).Style:id=1590,.parent=13"
         _StyleDefs(933) =   "Splits(1).Columns(71).HeadingStyle:id=1587,.parent=14"
         _StyleDefs(934) =   "Splits(1).Columns(71).FooterStyle:id=1588,.parent=15"
         _StyleDefs(935) =   "Splits(1).Columns(71).EditorStyle:id=1589,.parent=17"
         _StyleDefs(936) =   "Splits(1).Columns(72).Style:id=1582,.parent=13"
         _StyleDefs(937) =   "Splits(1).Columns(72).HeadingStyle:id=1579,.parent=14"
         _StyleDefs(938) =   "Splits(1).Columns(72).FooterStyle:id=1580,.parent=15"
         _StyleDefs(939) =   "Splits(1).Columns(72).EditorStyle:id=1581,.parent=17"
         _StyleDefs(940) =   "Splits(1).Columns(73).Style:id=1574,.parent=13"
         _StyleDefs(941) =   "Splits(1).Columns(73).HeadingStyle:id=1571,.parent=14"
         _StyleDefs(942) =   "Splits(1).Columns(73).FooterStyle:id=1572,.parent=15"
         _StyleDefs(943) =   "Splits(1).Columns(73).EditorStyle:id=1573,.parent=17"
         _StyleDefs(944) =   "Splits(1).Columns(74).Style:id=1566,.parent=13"
         _StyleDefs(945) =   "Splits(1).Columns(74).HeadingStyle:id=1563,.parent=14"
         _StyleDefs(946) =   "Splits(1).Columns(74).FooterStyle:id=1564,.parent=15"
         _StyleDefs(947) =   "Splits(1).Columns(74).EditorStyle:id=1565,.parent=17"
         _StyleDefs(948) =   "Splits(1).Columns(75).Style:id=1558,.parent=13"
         _StyleDefs(949) =   "Splits(1).Columns(75).HeadingStyle:id=1555,.parent=14"
         _StyleDefs(950) =   "Splits(1).Columns(75).FooterStyle:id=1556,.parent=15"
         _StyleDefs(951) =   "Splits(1).Columns(75).EditorStyle:id=1557,.parent=17"
         _StyleDefs(952) =   "Splits(1).Columns(76).Style:id=1550,.parent=13"
         _StyleDefs(953) =   "Splits(1).Columns(76).HeadingStyle:id=1547,.parent=14"
         _StyleDefs(954) =   "Splits(1).Columns(76).FooterStyle:id=1548,.parent=15"
         _StyleDefs(955) =   "Splits(1).Columns(76).EditorStyle:id=1549,.parent=17"
         _StyleDefs(956) =   "Splits(1).Columns(77).Style:id=1542,.parent=13"
         _StyleDefs(957) =   "Splits(1).Columns(77).HeadingStyle:id=1539,.parent=14"
         _StyleDefs(958) =   "Splits(1).Columns(77).FooterStyle:id=1540,.parent=15"
         _StyleDefs(959) =   "Splits(1).Columns(77).EditorStyle:id=1541,.parent=17"
         _StyleDefs(960) =   "Splits(1).Columns(78).Style:id=1534,.parent=13"
         _StyleDefs(961) =   "Splits(1).Columns(78).HeadingStyle:id=1531,.parent=14"
         _StyleDefs(962) =   "Splits(1).Columns(78).FooterStyle:id=1532,.parent=15"
         _StyleDefs(963) =   "Splits(1).Columns(78).EditorStyle:id=1533,.parent=17"
         _StyleDefs(964) =   "Splits(1).Columns(79).Style:id=1526,.parent=13"
         _StyleDefs(965) =   "Splits(1).Columns(79).HeadingStyle:id=1523,.parent=14"
         _StyleDefs(966) =   "Splits(1).Columns(79).FooterStyle:id=1524,.parent=15"
         _StyleDefs(967) =   "Splits(1).Columns(79).EditorStyle:id=1525,.parent=17"
         _StyleDefs(968) =   "Splits(1).Columns(80).Style:id=1518,.parent=13"
         _StyleDefs(969) =   "Splits(1).Columns(80).HeadingStyle:id=1515,.parent=14"
         _StyleDefs(970) =   "Splits(1).Columns(80).FooterStyle:id=1516,.parent=15"
         _StyleDefs(971) =   "Splits(1).Columns(80).EditorStyle:id=1517,.parent=17"
         _StyleDefs(972) =   "Splits(1).Columns(81).Style:id=1510,.parent=13"
         _StyleDefs(973) =   "Splits(1).Columns(81).HeadingStyle:id=1507,.parent=14"
         _StyleDefs(974) =   "Splits(1).Columns(81).FooterStyle:id=1508,.parent=15"
         _StyleDefs(975) =   "Splits(1).Columns(81).EditorStyle:id=1509,.parent=17"
         _StyleDefs(976) =   "Splits(1).Columns(82).Style:id=1502,.parent=13"
         _StyleDefs(977) =   "Splits(1).Columns(82).HeadingStyle:id=1499,.parent=14"
         _StyleDefs(978) =   "Splits(1).Columns(82).FooterStyle:id=1500,.parent=15"
         _StyleDefs(979) =   "Splits(1).Columns(82).EditorStyle:id=1501,.parent=17"
         _StyleDefs(980) =   "Splits(1).Columns(83).Style:id=1494,.parent=13"
         _StyleDefs(981) =   "Splits(1).Columns(83).HeadingStyle:id=1491,.parent=14"
         _StyleDefs(982) =   "Splits(1).Columns(83).FooterStyle:id=1492,.parent=15"
         _StyleDefs(983) =   "Splits(1).Columns(83).EditorStyle:id=1493,.parent=17"
         _StyleDefs(984) =   "Splits(1).Columns(84).Style:id=1486,.parent=13"
         _StyleDefs(985) =   "Splits(1).Columns(84).HeadingStyle:id=1483,.parent=14"
         _StyleDefs(986) =   "Splits(1).Columns(84).FooterStyle:id=1484,.parent=15"
         _StyleDefs(987) =   "Splits(1).Columns(84).EditorStyle:id=1485,.parent=17"
         _StyleDefs(988) =   "Splits(1).Columns(85).Style:id=1478,.parent=13"
         _StyleDefs(989) =   "Splits(1).Columns(85).HeadingStyle:id=1475,.parent=14"
         _StyleDefs(990) =   "Splits(1).Columns(85).FooterStyle:id=1476,.parent=15"
         _StyleDefs(991) =   "Splits(1).Columns(85).EditorStyle:id=1477,.parent=17"
         _StyleDefs(992) =   "Splits(1).Columns(86).Style:id=1470,.parent=13"
         _StyleDefs(993) =   "Splits(1).Columns(86).HeadingStyle:id=1467,.parent=14"
         _StyleDefs(994) =   "Splits(1).Columns(86).FooterStyle:id=1468,.parent=15"
         _StyleDefs(995) =   "Splits(1).Columns(86).EditorStyle:id=1469,.parent=17"
         _StyleDefs(996) =   "Splits(1).Columns(87).Style:id=1462,.parent=13"
         _StyleDefs(997) =   "Splits(1).Columns(87).HeadingStyle:id=1459,.parent=14"
         _StyleDefs(998) =   "Splits(1).Columns(87).FooterStyle:id=1460,.parent=15"
         _StyleDefs(999) =   "Splits(1).Columns(87).EditorStyle:id=1461,.parent=17"
         _StyleDefs(1000)=   "Splits(1).Columns(88).Style:id=1454,.parent=13"
         _StyleDefs(1001)=   "Splits(1).Columns(88).HeadingStyle:id=1451,.parent=14"
         _StyleDefs(1002)=   "Splits(1).Columns(88).FooterStyle:id=1452,.parent=15"
         _StyleDefs(1003)=   "Splits(1).Columns(88).EditorStyle:id=1453,.parent=17"
         _StyleDefs(1004)=   "Splits(1).Columns(89).Style:id=1446,.parent=13"
         _StyleDefs(1005)=   "Splits(1).Columns(89).HeadingStyle:id=1443,.parent=14"
         _StyleDefs(1006)=   "Splits(1).Columns(89).FooterStyle:id=1444,.parent=15"
         _StyleDefs(1007)=   "Splits(1).Columns(89).EditorStyle:id=1445,.parent=17"
         _StyleDefs(1008)=   "Splits(1).Columns(90).Style:id=1438,.parent=13"
         _StyleDefs(1009)=   "Splits(1).Columns(90).HeadingStyle:id=1435,.parent=14"
         _StyleDefs(1010)=   "Splits(1).Columns(90).FooterStyle:id=1436,.parent=15"
         _StyleDefs(1011)=   "Splits(1).Columns(90).EditorStyle:id=1437,.parent=17"
         _StyleDefs(1012)=   "Splits(1).Columns(91).Style:id=1430,.parent=13"
         _StyleDefs(1013)=   "Splits(1).Columns(91).HeadingStyle:id=1427,.parent=14"
         _StyleDefs(1014)=   "Splits(1).Columns(91).FooterStyle:id=1428,.parent=15"
         _StyleDefs(1015)=   "Splits(1).Columns(91).EditorStyle:id=1429,.parent=17"
         _StyleDefs(1016)=   "Splits(1).Columns(92).Style:id=1422,.parent=13"
         _StyleDefs(1017)=   "Splits(1).Columns(92).HeadingStyle:id=1419,.parent=14"
         _StyleDefs(1018)=   "Splits(1).Columns(92).FooterStyle:id=1420,.parent=15"
         _StyleDefs(1019)=   "Splits(1).Columns(92).EditorStyle:id=1421,.parent=17"
         _StyleDefs(1020)=   "Splits(1).Columns(93).Style:id=1414,.parent=13"
         _StyleDefs(1021)=   "Splits(1).Columns(93).HeadingStyle:id=1411,.parent=14"
         _StyleDefs(1022)=   "Splits(1).Columns(93).FooterStyle:id=1412,.parent=15"
         _StyleDefs(1023)=   "Splits(1).Columns(93).EditorStyle:id=1413,.parent=17"
         _StyleDefs(1024)=   "Splits(1).Columns(94).Style:id=1406,.parent=13"
         _StyleDefs(1025)=   "Splits(1).Columns(94).HeadingStyle:id=1403,.parent=14"
         _StyleDefs(1026)=   "Splits(1).Columns(94).FooterStyle:id=1404,.parent=15"
         _StyleDefs(1027)=   "Splits(1).Columns(94).EditorStyle:id=1405,.parent=17"
         _StyleDefs(1028)=   "Splits(1).Columns(95).Style:id=1398,.parent=13"
         _StyleDefs(1029)=   "Splits(1).Columns(95).HeadingStyle:id=1395,.parent=14"
         _StyleDefs(1030)=   "Splits(1).Columns(95).FooterStyle:id=1396,.parent=15"
         _StyleDefs(1031)=   "Splits(1).Columns(95).EditorStyle:id=1397,.parent=17"
         _StyleDefs(1032)=   "Splits(1).Columns(96).Style:id=1390,.parent=13"
         _StyleDefs(1033)=   "Splits(1).Columns(96).HeadingStyle:id=1387,.parent=14"
         _StyleDefs(1034)=   "Splits(1).Columns(96).FooterStyle:id=1388,.parent=15"
         _StyleDefs(1035)=   "Splits(1).Columns(96).EditorStyle:id=1389,.parent=17"
         _StyleDefs(1036)=   "Splits(1).Columns(97).Style:id=1382,.parent=13"
         _StyleDefs(1037)=   "Splits(1).Columns(97).HeadingStyle:id=1379,.parent=14"
         _StyleDefs(1038)=   "Splits(1).Columns(97).FooterStyle:id=1380,.parent=15"
         _StyleDefs(1039)=   "Splits(1).Columns(97).EditorStyle:id=1381,.parent=17"
         _StyleDefs(1040)=   "Splits(1).Columns(98).Style:id=1374,.parent=13"
         _StyleDefs(1041)=   "Splits(1).Columns(98).HeadingStyle:id=1371,.parent=14"
         _StyleDefs(1042)=   "Splits(1).Columns(98).FooterStyle:id=1372,.parent=15"
         _StyleDefs(1043)=   "Splits(1).Columns(98).EditorStyle:id=1373,.parent=17"
         _StyleDefs(1044)=   "Splits(1).Columns(99).Style:id=1366,.parent=13"
         _StyleDefs(1045)=   "Splits(1).Columns(99).HeadingStyle:id=1363,.parent=14"
         _StyleDefs(1046)=   "Splits(1).Columns(99).FooterStyle:id=1364,.parent=15"
         _StyleDefs(1047)=   "Splits(1).Columns(99).EditorStyle:id=1365,.parent=17"
         _StyleDefs(1048)=   "Splits(1).Columns(100).Style:id=1358,.parent=13"
         _StyleDefs(1049)=   "Splits(1).Columns(100).HeadingStyle:id=1355,.parent=14"
         _StyleDefs(1050)=   "Splits(1).Columns(100).FooterStyle:id=1356,.parent=15"
         _StyleDefs(1051)=   "Splits(1).Columns(100).EditorStyle:id=1357,.parent=17"
         _StyleDefs(1052)=   "Splits(1).Columns(101).Style:id=1350,.parent=13"
         _StyleDefs(1053)=   "Splits(1).Columns(101).HeadingStyle:id=1347,.parent=14"
         _StyleDefs(1054)=   "Splits(1).Columns(101).FooterStyle:id=1348,.parent=15"
         _StyleDefs(1055)=   "Splits(1).Columns(101).EditorStyle:id=1349,.parent=17"
         _StyleDefs(1056)=   "Splits(1).Columns(102).Style:id=1342,.parent=13"
         _StyleDefs(1057)=   "Splits(1).Columns(102).HeadingStyle:id=1339,.parent=14"
         _StyleDefs(1058)=   "Splits(1).Columns(102).FooterStyle:id=1340,.parent=15"
         _StyleDefs(1059)=   "Splits(1).Columns(102).EditorStyle:id=1341,.parent=17"
         _StyleDefs(1060)=   "Splits(1).Columns(103).Style:id=1334,.parent=13"
         _StyleDefs(1061)=   "Splits(1).Columns(103).HeadingStyle:id=1331,.parent=14"
         _StyleDefs(1062)=   "Splits(1).Columns(103).FooterStyle:id=1332,.parent=15"
         _StyleDefs(1063)=   "Splits(1).Columns(103).EditorStyle:id=1333,.parent=17"
         _StyleDefs(1064)=   "Splits(1).Columns(104).Style:id=1326,.parent=13"
         _StyleDefs(1065)=   "Splits(1).Columns(104).HeadingStyle:id=1323,.parent=14"
         _StyleDefs(1066)=   "Splits(1).Columns(104).FooterStyle:id=1324,.parent=15"
         _StyleDefs(1067)=   "Splits(1).Columns(104).EditorStyle:id=1325,.parent=17"
         _StyleDefs(1068)=   "Splits(1).Columns(105).Style:id=1318,.parent=13"
         _StyleDefs(1069)=   "Splits(1).Columns(105).HeadingStyle:id=1315,.parent=14"
         _StyleDefs(1070)=   "Splits(1).Columns(105).FooterStyle:id=1316,.parent=15"
         _StyleDefs(1071)=   "Splits(1).Columns(105).EditorStyle:id=1317,.parent=17"
         _StyleDefs(1072)=   "Splits(1).Columns(106).Style:id=1310,.parent=13"
         _StyleDefs(1073)=   "Splits(1).Columns(106).HeadingStyle:id=1307,.parent=14"
         _StyleDefs(1074)=   "Splits(1).Columns(106).FooterStyle:id=1308,.parent=15"
         _StyleDefs(1075)=   "Splits(1).Columns(106).EditorStyle:id=1309,.parent=17"
         _StyleDefs(1076)=   "Splits(1).Columns(107).Style:id=1302,.parent=13"
         _StyleDefs(1077)=   "Splits(1).Columns(107).HeadingStyle:id=1299,.parent=14"
         _StyleDefs(1078)=   "Splits(1).Columns(107).FooterStyle:id=1300,.parent=15"
         _StyleDefs(1079)=   "Splits(1).Columns(107).EditorStyle:id=1301,.parent=17"
         _StyleDefs(1080)=   "Splits(1).Columns(108).Style:id=1294,.parent=13"
         _StyleDefs(1081)=   "Splits(1).Columns(108).HeadingStyle:id=1291,.parent=14"
         _StyleDefs(1082)=   "Splits(1).Columns(108).FooterStyle:id=1292,.parent=15"
         _StyleDefs(1083)=   "Splits(1).Columns(108).EditorStyle:id=1293,.parent=17"
         _StyleDefs(1084)=   "Splits(1).Columns(109).Style:id=1286,.parent=13"
         _StyleDefs(1085)=   "Splits(1).Columns(109).HeadingStyle:id=1283,.parent=14"
         _StyleDefs(1086)=   "Splits(1).Columns(109).FooterStyle:id=1284,.parent=15"
         _StyleDefs(1087)=   "Splits(1).Columns(109).EditorStyle:id=1285,.parent=17"
         _StyleDefs(1088)=   "Splits(1).Columns(110).Style:id=1278,.parent=13"
         _StyleDefs(1089)=   "Splits(1).Columns(110).HeadingStyle:id=1275,.parent=14"
         _StyleDefs(1090)=   "Splits(1).Columns(110).FooterStyle:id=1276,.parent=15"
         _StyleDefs(1091)=   "Splits(1).Columns(110).EditorStyle:id=1277,.parent=17"
         _StyleDefs(1092)=   "Splits(1).Columns(111).Style:id=1270,.parent=13"
         _StyleDefs(1093)=   "Splits(1).Columns(111).HeadingStyle:id=1267,.parent=14"
         _StyleDefs(1094)=   "Splits(1).Columns(111).FooterStyle:id=1268,.parent=15"
         _StyleDefs(1095)=   "Splits(1).Columns(111).EditorStyle:id=1269,.parent=17"
         _StyleDefs(1096)=   "Splits(1).Columns(112).Style:id=1262,.parent=13"
         _StyleDefs(1097)=   "Splits(1).Columns(112).HeadingStyle:id=1259,.parent=14"
         _StyleDefs(1098)=   "Splits(1).Columns(112).FooterStyle:id=1260,.parent=15"
         _StyleDefs(1099)=   "Splits(1).Columns(112).EditorStyle:id=1261,.parent=17"
         _StyleDefs(1100)=   "Splits(1).Columns(113).Style:id=1254,.parent=13"
         _StyleDefs(1101)=   "Splits(1).Columns(113).HeadingStyle:id=1251,.parent=14"
         _StyleDefs(1102)=   "Splits(1).Columns(113).FooterStyle:id=1252,.parent=15"
         _StyleDefs(1103)=   "Splits(1).Columns(113).EditorStyle:id=1253,.parent=17"
         _StyleDefs(1104)=   "Splits(1).Columns(114).Style:id=1246,.parent=13"
         _StyleDefs(1105)=   "Splits(1).Columns(114).HeadingStyle:id=1243,.parent=14"
         _StyleDefs(1106)=   "Splits(1).Columns(114).FooterStyle:id=1244,.parent=15"
         _StyleDefs(1107)=   "Splits(1).Columns(114).EditorStyle:id=1245,.parent=17"
         _StyleDefs(1108)=   "Splits(1).Columns(115).Style:id=1238,.parent=13"
         _StyleDefs(1109)=   "Splits(1).Columns(115).HeadingStyle:id=1235,.parent=14"
         _StyleDefs(1110)=   "Splits(1).Columns(115).FooterStyle:id=1236,.parent=15"
         _StyleDefs(1111)=   "Splits(1).Columns(115).EditorStyle:id=1237,.parent=17"
         _StyleDefs(1112)=   "Splits(1).Columns(116).Style:id=1230,.parent=13"
         _StyleDefs(1113)=   "Splits(1).Columns(116).HeadingStyle:id=1227,.parent=14"
         _StyleDefs(1114)=   "Splits(1).Columns(116).FooterStyle:id=1228,.parent=15"
         _StyleDefs(1115)=   "Splits(1).Columns(116).EditorStyle:id=1229,.parent=17"
         _StyleDefs(1116)=   "Splits(1).Columns(117).Style:id=1222,.parent=13"
         _StyleDefs(1117)=   "Splits(1).Columns(117).HeadingStyle:id=1219,.parent=14"
         _StyleDefs(1118)=   "Splits(1).Columns(117).FooterStyle:id=1220,.parent=15"
         _StyleDefs(1119)=   "Splits(1).Columns(117).EditorStyle:id=1221,.parent=17"
         _StyleDefs(1120)=   "Splits(1).Columns(118).Style:id=1214,.parent=13"
         _StyleDefs(1121)=   "Splits(1).Columns(118).HeadingStyle:id=1211,.parent=14"
         _StyleDefs(1122)=   "Splits(1).Columns(118).FooterStyle:id=1212,.parent=15"
         _StyleDefs(1123)=   "Splits(1).Columns(118).EditorStyle:id=1213,.parent=17"
         _StyleDefs(1124)=   "Splits(1).Columns(119).Style:id=1206,.parent=13"
         _StyleDefs(1125)=   "Splits(1).Columns(119).HeadingStyle:id=1203,.parent=14"
         _StyleDefs(1126)=   "Splits(1).Columns(119).FooterStyle:id=1204,.parent=15"
         _StyleDefs(1127)=   "Splits(1).Columns(119).EditorStyle:id=1205,.parent=17"
         _StyleDefs(1128)=   "Splits(1).Columns(120).Style:id=1198,.parent=13"
         _StyleDefs(1129)=   "Splits(1).Columns(120).HeadingStyle:id=1195,.parent=14"
         _StyleDefs(1130)=   "Splits(1).Columns(120).FooterStyle:id=1196,.parent=15"
         _StyleDefs(1131)=   "Splits(1).Columns(120).EditorStyle:id=1197,.parent=17"
         _StyleDefs(1132)=   "Splits(1).Columns(121).Style:id=1190,.parent=13"
         _StyleDefs(1133)=   "Splits(1).Columns(121).HeadingStyle:id=1187,.parent=14"
         _StyleDefs(1134)=   "Splits(1).Columns(121).FooterStyle:id=1188,.parent=15"
         _StyleDefs(1135)=   "Splits(1).Columns(121).EditorStyle:id=1189,.parent=17"
         _StyleDefs(1136)=   "Splits(1).Columns(122).Style:id=1182,.parent=13"
         _StyleDefs(1137)=   "Splits(1).Columns(122).HeadingStyle:id=1179,.parent=14"
         _StyleDefs(1138)=   "Splits(1).Columns(122).FooterStyle:id=1180,.parent=15"
         _StyleDefs(1139)=   "Splits(1).Columns(122).EditorStyle:id=1181,.parent=17"
         _StyleDefs(1140)=   "Splits(1).Columns(123).Style:id=1174,.parent=13"
         _StyleDefs(1141)=   "Splits(1).Columns(123).HeadingStyle:id=1171,.parent=14"
         _StyleDefs(1142)=   "Splits(1).Columns(123).FooterStyle:id=1172,.parent=15"
         _StyleDefs(1143)=   "Splits(1).Columns(123).EditorStyle:id=1173,.parent=17"
         _StyleDefs(1144)=   "Splits(1).Columns(124).Style:id=1166,.parent=13"
         _StyleDefs(1145)=   "Splits(1).Columns(124).HeadingStyle:id=1163,.parent=14"
         _StyleDefs(1146)=   "Splits(1).Columns(124).FooterStyle:id=1164,.parent=15"
         _StyleDefs(1147)=   "Splits(1).Columns(124).EditorStyle:id=1165,.parent=17"
         _StyleDefs(1148)=   "Splits(1).Columns(125).Style:id=886,.parent=13"
         _StyleDefs(1149)=   "Splits(1).Columns(125).HeadingStyle:id=883,.parent=14"
         _StyleDefs(1150)=   "Splits(1).Columns(125).FooterStyle:id=884,.parent=15"
         _StyleDefs(1151)=   "Splits(1).Columns(125).EditorStyle:id=885,.parent=17"
         _StyleDefs(1152)=   "Splits(1).Columns(126).Style:id=878,.parent=13"
         _StyleDefs(1153)=   "Splits(1).Columns(126).HeadingStyle:id=875,.parent=14"
         _StyleDefs(1154)=   "Splits(1).Columns(126).FooterStyle:id=876,.parent=15"
         _StyleDefs(1155)=   "Splits(1).Columns(126).EditorStyle:id=877,.parent=17"
         _StyleDefs(1156)=   "Splits(1).Columns(127).Style:id=870,.parent=13"
         _StyleDefs(1157)=   "Splits(1).Columns(127).HeadingStyle:id=867,.parent=14"
         _StyleDefs(1158)=   "Splits(1).Columns(127).FooterStyle:id=868,.parent=15"
         _StyleDefs(1159)=   "Splits(1).Columns(127).EditorStyle:id=869,.parent=17"
         _StyleDefs(1160)=   "Splits(1).Columns(128).Style:id=862,.parent=13"
         _StyleDefs(1161)=   "Splits(1).Columns(128).HeadingStyle:id=859,.parent=14"
         _StyleDefs(1162)=   "Splits(1).Columns(128).FooterStyle:id=860,.parent=15"
         _StyleDefs(1163)=   "Splits(1).Columns(128).EditorStyle:id=861,.parent=17"
         _StyleDefs(1164)=   "Splits(1).Columns(129).Style:id=854,.parent=13"
         _StyleDefs(1165)=   "Splits(1).Columns(129).HeadingStyle:id=851,.parent=14"
         _StyleDefs(1166)=   "Splits(1).Columns(129).FooterStyle:id=852,.parent=15"
         _StyleDefs(1167)=   "Splits(1).Columns(129).EditorStyle:id=853,.parent=17"
         _StyleDefs(1168)=   "Splits(1).Columns(130).Style:id=846,.parent=13"
         _StyleDefs(1169)=   "Splits(1).Columns(130).HeadingStyle:id=843,.parent=14"
         _StyleDefs(1170)=   "Splits(1).Columns(130).FooterStyle:id=844,.parent=15"
         _StyleDefs(1171)=   "Splits(1).Columns(130).EditorStyle:id=845,.parent=17"
         _StyleDefs(1172)=   "Splits(1).Columns(131).Style:id=838,.parent=13"
         _StyleDefs(1173)=   "Splits(1).Columns(131).HeadingStyle:id=835,.parent=14"
         _StyleDefs(1174)=   "Splits(1).Columns(131).FooterStyle:id=836,.parent=15"
         _StyleDefs(1175)=   "Splits(1).Columns(131).EditorStyle:id=837,.parent=17"
         _StyleDefs(1176)=   "Splits(1).Columns(132).Style:id=830,.parent=13"
         _StyleDefs(1177)=   "Splits(1).Columns(132).HeadingStyle:id=827,.parent=14"
         _StyleDefs(1178)=   "Splits(1).Columns(132).FooterStyle:id=828,.parent=15"
         _StyleDefs(1179)=   "Splits(1).Columns(132).EditorStyle:id=829,.parent=17"
         _StyleDefs(1180)=   "Splits(1).Columns(133).Style:id=822,.parent=13"
         _StyleDefs(1181)=   "Splits(1).Columns(133).HeadingStyle:id=819,.parent=14"
         _StyleDefs(1182)=   "Splits(1).Columns(133).FooterStyle:id=820,.parent=15"
         _StyleDefs(1183)=   "Splits(1).Columns(133).EditorStyle:id=821,.parent=17"
         _StyleDefs(1184)=   "Splits(1).Columns(134).Style:id=814,.parent=13"
         _StyleDefs(1185)=   "Splits(1).Columns(134).HeadingStyle:id=811,.parent=14"
         _StyleDefs(1186)=   "Splits(1).Columns(134).FooterStyle:id=812,.parent=15"
         _StyleDefs(1187)=   "Splits(1).Columns(134).EditorStyle:id=813,.parent=17"
         _StyleDefs(1188)=   "Splits(1).Columns(135).Style:id=806,.parent=13"
         _StyleDefs(1189)=   "Splits(1).Columns(135).HeadingStyle:id=803,.parent=14"
         _StyleDefs(1190)=   "Splits(1).Columns(135).FooterStyle:id=804,.parent=15"
         _StyleDefs(1191)=   "Splits(1).Columns(135).EditorStyle:id=805,.parent=17"
         _StyleDefs(1192)=   "Splits(1).Columns(136).Style:id=902,.parent=13"
         _StyleDefs(1193)=   "Splits(1).Columns(136).HeadingStyle:id=899,.parent=14"
         _StyleDefs(1194)=   "Splits(1).Columns(136).FooterStyle:id=900,.parent=15"
         _StyleDefs(1195)=   "Splits(1).Columns(136).EditorStyle:id=901,.parent=17"
         _StyleDefs(1196)=   "Splits(1).Columns(137).Style:id=910,.parent=13"
         _StyleDefs(1197)=   "Splits(1).Columns(137).HeadingStyle:id=907,.parent=14"
         _StyleDefs(1198)=   "Splits(1).Columns(137).FooterStyle:id=908,.parent=15"
         _StyleDefs(1199)=   "Splits(1).Columns(137).EditorStyle:id=909,.parent=17"
         _StyleDefs(1200)=   "Splits(1).Columns(138).Style:id=918,.parent=13"
         _StyleDefs(1201)=   "Splits(1).Columns(138).HeadingStyle:id=915,.parent=14"
         _StyleDefs(1202)=   "Splits(1).Columns(138).FooterStyle:id=916,.parent=15"
         _StyleDefs(1203)=   "Splits(1).Columns(138).EditorStyle:id=917,.parent=17"
         _StyleDefs(1204)=   "Splits(1).Columns(139).Style:id=926,.parent=13"
         _StyleDefs(1205)=   "Splits(1).Columns(139).HeadingStyle:id=923,.parent=14"
         _StyleDefs(1206)=   "Splits(1).Columns(139).FooterStyle:id=924,.parent=15"
         _StyleDefs(1207)=   "Splits(1).Columns(139).EditorStyle:id=925,.parent=17"
         _StyleDefs(1208)=   "Splits(1).Columns(140).Style:id=934,.parent=13"
         _StyleDefs(1209)=   "Splits(1).Columns(140).HeadingStyle:id=931,.parent=14"
         _StyleDefs(1210)=   "Splits(1).Columns(140).FooterStyle:id=932,.parent=15"
         _StyleDefs(1211)=   "Splits(1).Columns(140).EditorStyle:id=933,.parent=17"
         _StyleDefs(1212)=   "Splits(1).Columns(141).Style:id=942,.parent=13"
         _StyleDefs(1213)=   "Splits(1).Columns(141).HeadingStyle:id=939,.parent=14"
         _StyleDefs(1214)=   "Splits(1).Columns(141).FooterStyle:id=940,.parent=15"
         _StyleDefs(1215)=   "Splits(1).Columns(141).EditorStyle:id=941,.parent=17"
         _StyleDefs(1216)=   "Splits(1).Columns(142).Style:id=950,.parent=13"
         _StyleDefs(1217)=   "Splits(1).Columns(142).HeadingStyle:id=947,.parent=14"
         _StyleDefs(1218)=   "Splits(1).Columns(142).FooterStyle:id=948,.parent=15"
         _StyleDefs(1219)=   "Splits(1).Columns(142).EditorStyle:id=949,.parent=17"
         _StyleDefs(1220)=   "Splits(1).Columns(143).Style:id=958,.parent=13"
         _StyleDefs(1221)=   "Splits(1).Columns(143).HeadingStyle:id=955,.parent=14"
         _StyleDefs(1222)=   "Splits(1).Columns(143).FooterStyle:id=956,.parent=15"
         _StyleDefs(1223)=   "Splits(1).Columns(143).EditorStyle:id=957,.parent=17"
         _StyleDefs(1224)=   "Splits(1).Columns(144).Style:id=966,.parent=13"
         _StyleDefs(1225)=   "Splits(1).Columns(144).HeadingStyle:id=963,.parent=14"
         _StyleDefs(1226)=   "Splits(1).Columns(144).FooterStyle:id=964,.parent=15"
         _StyleDefs(1227)=   "Splits(1).Columns(144).EditorStyle:id=965,.parent=17"
         _StyleDefs(1228)=   "Splits(1).Columns(145).Style:id=974,.parent=13"
         _StyleDefs(1229)=   "Splits(1).Columns(145).HeadingStyle:id=971,.parent=14"
         _StyleDefs(1230)=   "Splits(1).Columns(145).FooterStyle:id=972,.parent=15"
         _StyleDefs(1231)=   "Splits(1).Columns(145).EditorStyle:id=973,.parent=17"
         _StyleDefs(1232)=   "Splits(1).Columns(146).Style:id=982,.parent=13"
         _StyleDefs(1233)=   "Splits(1).Columns(146).HeadingStyle:id=979,.parent=14"
         _StyleDefs(1234)=   "Splits(1).Columns(146).FooterStyle:id=980,.parent=15"
         _StyleDefs(1235)=   "Splits(1).Columns(146).EditorStyle:id=981,.parent=17"
         _StyleDefs(1236)=   "Splits(1).Columns(147).Style:id=990,.parent=13"
         _StyleDefs(1237)=   "Splits(1).Columns(147).HeadingStyle:id=987,.parent=14"
         _StyleDefs(1238)=   "Splits(1).Columns(147).FooterStyle:id=988,.parent=15"
         _StyleDefs(1239)=   "Splits(1).Columns(147).EditorStyle:id=989,.parent=17"
         _StyleDefs(1240)=   "Splits(1).Columns(148).Style:id=998,.parent=13"
         _StyleDefs(1241)=   "Splits(1).Columns(148).HeadingStyle:id=995,.parent=14"
         _StyleDefs(1242)=   "Splits(1).Columns(148).FooterStyle:id=996,.parent=15"
         _StyleDefs(1243)=   "Splits(1).Columns(148).EditorStyle:id=997,.parent=17"
         _StyleDefs(1244)=   "Splits(1).Columns(149).Style:id=1006,.parent=13"
         _StyleDefs(1245)=   "Splits(1).Columns(149).HeadingStyle:id=1003,.parent=14"
         _StyleDefs(1246)=   "Splits(1).Columns(149).FooterStyle:id=1004,.parent=15"
         _StyleDefs(1247)=   "Splits(1).Columns(149).EditorStyle:id=1005,.parent=17"
         _StyleDefs(1248)=   "Splits(1).Columns(150).Style:id=1014,.parent=13"
         _StyleDefs(1249)=   "Splits(1).Columns(150).HeadingStyle:id=1011,.parent=14"
         _StyleDefs(1250)=   "Splits(1).Columns(150).FooterStyle:id=1012,.parent=15"
         _StyleDefs(1251)=   "Splits(1).Columns(150).EditorStyle:id=1013,.parent=17"
         _StyleDefs(1252)=   "Named:id=33:Normal"
         _StyleDefs(1253)=   ":id=33,.parent=0"
         _StyleDefs(1254)=   "Named:id=34:Heading"
         _StyleDefs(1255)=   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(1256)=   ":id=34,.wraptext=-1"
         _StyleDefs(1257)=   "Named:id=35:Footing"
         _StyleDefs(1258)=   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(1259)=   "Named:id=36:Selected"
         _StyleDefs(1260)=   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(1261)=   "Named:id=37:Caption"
         _StyleDefs(1262)=   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(1263)=   "Named:id=38:HighlightRow"
         _StyleDefs(1264)=   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(1265)=   "Named:id=39:EvenRow"
         _StyleDefs(1266)=   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(1267)=   "Named:id=40:OddRow"
         _StyleDefs(1268)=   ":id=40,.parent=33"
         _StyleDefs(1269)=   "Named:id=41:RecordSelector"
         _StyleDefs(1270)=   ":id=41,.parent=34"
         _StyleDefs(1271)=   "Named:id=42:FilterBar"
         _StyleDefs(1272)=   ":id=42,.parent=33"
      End
      Begin VB.Label LblJudul 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   15135
      End
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmStallProgress.frx":0342
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmStallProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 19 Maret 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim row As Integer                          'baris gridStall
Dim coll As Integer                         'kolom gridstall
Dim colSkrg As Integer                      'kolom posisi jam sekarang
Dim nTgl As String                          'tanggal sekarang
Dim njamAwl As Integer                      'mulai jam kerja
Dim njamAk As Integer                       'akhir jam kerja
Dim jmlStall As Integer                     'jumlah Stall
Dim jmlTHS As Integer                       'jumlah THS
Dim jmlAll As Integer                       'jumlah All
Dim stl As String                           'penanda no stall
Dim PRC As Boolean                          'penanda SBE
Dim sBold As Boolean                        'penanda bold
Dim C As TrueOleDBGrid70.Columns            'bantuan untuk filter
Dim xy As Boolean
Dim xy1 As Boolean
Dim batalLoad As Boolean
Dim objProgress As New BClsProgress
Dim batasAwal As Integer
Dim baris As Integer
'Dim FOlama As String

Dim rsPending As ADODB.Recordset
Dim rsSearch2 As ADODB.Recordset
Dim rsStall As ADODB.Recordset           'recordset untuk stall
Dim rsJob As ADODB.Recordset             'recordset untuk job Order
Dim rsMech As ADODB.Recordset            'recordset untuk mekanik idle
Dim rsSA As ADODB.Recordset              'recordset untuk SA
Dim cols As TrueOleDBGrid70.Columns
'Dim myArray As XArrayDB

Const NilaiAwal = 1
Const AkhirPos = 150
Const maxEM = 5

Private Sub hapusForm()
    For i = 0 To 21
        Select Case i
            Case 1
                Text(i) = ""
                Text(i).Mask = ""
            Case 3, 15, 16
                Text(i) = 0
            Case 4, 5, 7
                Text(i) = ""
            Case 6
                Text(i) = ""
                Text(i).Mask = "99-99-9999"
            Case Else
                Text(i) = ""
        End Select
    Next
End Sub

Private Sub createPending()
    
    With rsPending
        If .State Then .Close
        .Fields.Append "NOWO", adChar, 10
        .Fields.Append "NOPOL", adChar, 15
        .Fields.Append "CUST", adChar, 50
        .Fields.Append "STALL", adChar, 15
        .Fields.Append "JAMMULAI", adChar, 5
        .Fields.Append "MECH", adChar, 5
        .Fields.Append "REP", adChar, 50
        .Open
    End With
    
    Set pendingGrid.DataSource = rsPending
    
End Sub

Private Sub SetWaktu(jamAwl As Integer, jamAkhir As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim nilaiAW As Integer
    
    j = 0
    batasAwal = IIf(jamAwl Mod 60 = 0, 6, Round((60 - (jamAwl Mod 60)) / 10)) ' + 1)'
    For i = NilaiAwal To ((jamAkhir - jamAwl) / 10) + IIf((jamAwl Mod 60) Mod 2 = 1, 3, 2) '2
        nilaiAW = NilaiAwal - (6 - batasAwal)
        Select Case i
            Case Is <= batasAwal '+ 1
                If i = NilaiAwal + (j * batasAwal) Then
                    judul = Val(jamAwl / 60) + j
                    j = j + 1
                Else
                    judul = "."
                End If
            Case Else
                If i = nilaiAW + (j * 6) Then
                    judul = Val(jamAwl / 60) + j
                    j = j + 1
                Else
                    judul = "."
                End If
        End Select
        
        With gridStall
            .Splits(1).Columns(i).Caption = judul
            .Splits(1).Columns(i).Width = 350
            .Splits(1).Columns(i).Alignment = dbgCenter
        End With
    Next
    
End Sub

Private Sub setStall(jamAwl As Integer, jamAkhir As Integer)
Dim aI As Integer

    gridStall.RowHeight = gHeight
    gridStall.Style.Font.Bold = True
    gridStall.Style.Font.Size = gFont
    gridStall.Splits(0).Columns(0).Width = 2200
    jmlStall = objO.FetchData_Integer("SELECT FN_JML_STALL FROM T_PROFIL")
    jmlTHS = objO.FetchData_Integer("SELECT FN_JML_THS FROM T_PROFIL")
    
    jmlAll = jmlStall + jmlTHS
    MyArray.ReDim 1, (jmlAll * 2) + 1, 0, ((jamAkhir - jamAwl) / 10) + 2
    MyArray.Clear
    aI = 1
    For i = 1 To (jmlAll * 2)
            If i Mod 2 = 1 Then
                If aI <= jmlStall Then
                    MyArray.value(i, 0) = "Stall " & aI
                ElseIf aI - jmlStall - 1 <= jmlTHS Then
                    MyArray.value(i, 0) = "THS " & (aI - jmlStall)
                End If
                aI = aI + 1
            End If
    Next
    gridStall.Array = MyArray
End Sub

Private Sub waktuSkrg(jamAwl As Integer, jamAkhir As Integer)
Dim jamSkrg As String
Dim hourS As Integer
Dim rsSearch As ADODB.Recordset
    
    jamSkrg = Format(isnow, "HH:NN")
    
    hourS = (CDbl(Left(jamSkrg, InStr(jamSkrg, ":") - 1)) * 60) + CDbl(Right(jamSkrg, InStr(jamSkrg, ":") - 1)) - (jamAwl - (jamAwl Mod 10))
    colSkrg = Round((hourS / 10) + IIf((hourS Mod 10) >= 0 And (hourS Mod 10) < 6, 1, 0), 0)  'Round((hourS / 10) + 0.5, 0) + 1 'Val((hourS / 10) + 0.01) + 1
    
    If colSkrg <= ((jamAkhir - jamAwl) / 10) Then
        For i = 1 To (jmlAll * 2) + 1
            Set rsSearch = objO.FetchData("SELECT SUBSTRING(NOSTALL,7,2) AS STL,* FROM STALL WHERE TANGGAL = CONVERT(varchar(10), GETDATE(),101) AND JAMMULAI < '" & Trim(jamSkrg) & "' AND SUBSTRING(NOSTALL,7,2) = '" & IIf(i = 1, 1, Round((i / 2) + 0.1, 0)) & "' AND (FLAG = 'N') " _
                                        & "ORDER BY CAST(SUBSTRING(NOSTALL,7,2) AS INTEGER)")
            If Not rsSearch.EOF Then
                If i Mod 2 = 1 Then
                    MyArray.value(i, colSkrg) = "?"
                Else
                    MyArray.value(i, colSkrg) = "|"
                End If
            Else
                MyArray.value(i, colSkrg) = "|"
            End If
        Next
        gridStall.Array = MyArray
    End If
End Sub

Private Sub loadStall(nTanggal As String, jamAwl As Integer, jamAkhir As Integer)
    Dim rsAStl As ADODB.Recordset
    Dim ProgressBook As New TrueDBGrid70.Style
    Dim ProgressWait As New TrueDBGrid70.Style
    Dim ProgressWaitBok As New TrueDBGrid70.Style
    Dim ProgressOn As New TrueDBGrid70.Style
    Dim ProgressPause As New TrueDBGrid70.Style
    Dim ProgressFinish As New TrueDBGrid70.Style
    Dim ProgressJam As New TrueDBGrid70.Style
    Dim ProgressProses As New TrueDBGrid70.Style
    Dim ProgressInvoice As New TrueDBGrid70.Style
    Dim ProgressKerja As New TrueDBGrid70.Style
    Dim ProgressClosing As New TrueDBGrid70.Style
    Dim ProgressPending1 As New TrueDBGrid70.Style
    
    ''' TAMBAHAN FINAL CHECK & WASHING
    Dim ProgressWaitFinal As New TrueDBGrid70.Style
    Dim ProgressFinal As New TrueDBGrid70.Style
    Dim ProgressFinishFinal As New TrueDBGrid70.Style
    Dim ProgressWaitWash As New TrueDBGrid70.Style
    Dim ProgressWash As New TrueDBGrid70.Style

    Dim tnd As String
    Dim nStl As String
    Dim rteBaru As Double

    gridStall.EvenRowStyle.BackColor = &HFFFFFF   '&H80000005
    gridStall.OddRowStyle.BackColor = &HFFFFFF    '&H80000005

    ' ###  UNIT BOOKING ###
    ProgressBook.BackColor = &H4D9AE8
    ProgressBook.ForeColor = &HFFFFFF
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressBook, "^B"

    ' ###  UNIT NEXT JOB / BELUM PROSES ###
    ProgressWait.BackColor = &HFFFF80   'vbCyan
    ProgressWait.ForeColor = &H800000
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressWait, "^W"
    
    ' ###  UNIT NEXT JOB / BELUM PROSES KHUSUS BOOKING ###
    ProgressWaitBok.BackColor = &H91910D
    ProgressWaitBok.ForeColor = vbWhite
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressWaitBok, "^w"

    ' ### UNIT SEDANG DIKERJAKAN ###
    ProgressOn.BackColor = vbRed
    ProgressOn.ForeColor = vbYellow
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressOn, "^P"

    ' ### UNIT PAUSE ###
    ProgressPause.BackColor = vbYellow
    ProgressPause.ForeColor = &H40&
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressPause, "^S"

    ' ### UNIT FINISH ###
    ProgressFinish.BackColor = vbGreen
    ProgressFinish.ForeColor = &H4000&              '&HFFFFFF
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressFinish, "^F"

    ' ### WAKTU SEKARANG ###
    ProgressJam.BackColor = &H4000
    ProgressJam.ForeColor = &H4000      '&H1E8036
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressJam, "^|"

    ' ### PROSES ACTUAL PENGERJAAN ###
    ProgressProses.BackColor = &H757575        'vbGrayed
    ProgressProses.ForeColor = vbWhite
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressProses, "^-"

    ' ### PROSES INVOICE ###
    ProgressInvoice.BackColor = &HDAFEA5
    ProgressInvoice.ForeColor = vbBlack
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressInvoice, "^I"

    ' ### PROSES CLOSING ###
    ProgressClosing.BackColor = &HEAF9FD
    ProgressClosing.ForeColor = vbBlack
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressClosing, "^C"

    ' ### PROSES JEJAK ###
    ProgressKerja.BackColor = &HB79891
    ProgressKerja.ForeColor = vbYellow
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressKerja, "^X"

    ' ### PROSES PENDING ###
    ProgressPending1.BackColor = &H46E8AF
    ProgressPending1.ForeColor = vbRed
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressPending1, "^?"
    
    ' ### PROSES WAIT FINAL INSPC ###
    ProgressWaitFinal.BackColor = &HDC6BC0
    ProgressWaitFinal.ForeColor = vbWhite
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressWaitFinal, "^H"
    
    ' ### PROSES FINAL INSPC ###
    ProgressFinal.BackColor = &H409FFF
    ProgressFinal.ForeColor = &H408000
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressFinal, "^Q"
    
    ' ### PROSES STOP FINAL INSPC ###
    ProgressFinishFinal.BackColor = &H96A0E7
    ProgressFinishFinal.ForeColor = vbRed
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressFinishFinal, "^O"
    
    ' ### PROSES WAITING WASHING ###
    ProgressWaitWash.BackColor = &H80
    ProgressWaitWash.ForeColor = vbWhite
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressWaitWash, "^G"
    
    ' ### PROSES WASHING ###
    ProgressWash.BackColor = &H36E9BC
    ProgressWash.ForeColor = vbBlack
    gridStall.Splits(1).AddRegexCellStyle dbgAllCells, ProgressWash, "^A"

    gridStall.AlternatingRowStyle = True

    Call clearMyarray(njamAwl, njamAk) ', 1, 1)
    Call setStall(njamAwl, njamAk)

    'PERBAIKAN STALL
    Call objO.UpdateData("UPDATE STALL SET JAMSELESAI = '23:59' WHERE JAMSELESAI >= '24:00'")
    Call objO.UpdateData("UPDATE STALL SET JAMSELESAI = CONVERT(varchar(5), DATEADD(n,(RATE)*60,JAMMULAI),8) WHERE TANGGAL >= CONVERT(varchar(10), GETDATE(),101) AND JAMSELESAI <> CONVERT(varchar(5), DATEADD(n,(RATE)*60,JAMMULAI),8)")
    Call objO.UpdateData("UPDATE STALL SET JAMSELESAI = '23:59' WHERE JAMSELESAI < JAMMULAI ")
    Call objO.UpdateData("UPDATE STALL SET RATE = 0.1 WHERE RATE <= 0")
    
    Set rsAStl = objO.FetchData("SELECT * FROM STALL WHERE (RIGHT(JAMSELESAI,2) = 60 OR RIGHT(JAMSELESAI,2) = '') AND TANGGAL = '" & nTanggal & "'")
    If Not rsAStl.EOF Then
        Call objO.UpdateData("UPDATE STALL SET JAMSELESAI = CAST(CAST(LEFT(JAMSELESAI,2) AS INTEGER)+1 AS CHAR(2)) + ':00' WHERE (RIGHT(JAMSELESAI,2) = 60 OR RIGHT(JAMSELESAI,2) = '') AND TANGGAL = '" & nTanggal & "'")
    End If
    Set rsAStl = objO.FetchData("SELECT * FROM STALL WHERE DATEDIFF(n,JAMMULAI,JAMSELESAI) < 0 AND TANGGAL = '" & nTanggal & "'")
    If Not rsAStl.EOF Then
        Do While Not rsAStl.EOF
            Call objO.UpdateData("UPDATE STALL SET JAMSELESAI = '" & Left(hitungWaktu(CDbl(rsAStl("RATE")), Trim(rsAStl("JAMMULAI")), Trim(rsAStl("TANGGAL")), "progress"), 5) & "' WHERE KETERANGAN = '" & Trim(rsAStl("KETERANGAN")) & "'")
            rsAStl.MoveNext
        Loop
    End If
    
    ' DARI TABEL STALL
    Dim rsBooK As ADODB.Recordset
    Set rsAStl = objProgress.takeUnit(nTanggal, 1)
    If Not rsAStl.EOF Then
        Do While Not rsAStl.EOF
            If Trim(rsAStl("FLAG")) = "" Then
                tnd = "B"
            ElseIf Trim(rsAStl("FLAG")) = "N" Then
                tnd = "W"
                Dim NoWo As String
                NoWo = Mid(Trim(rsAStl("KETERANGAN")), Len(Left(Trim(rsAStl("KETERANGAN")), InStr(Trim(rsAStl("KETERANGAN")), " |") + 2)), (Len(Left(Trim(rsAStl("KETERANGAN")), InStr(Trim(rsAStl("KETERANGAN")), "| "))) - 2 - Len(Left(Trim(rsAStl("KETERANGAN")), InStr(Trim(rsAStl("KETERANGAN")), " |")))))
                Set rsBooK = objO.FetchData("SELECT * FROM JOBDTA WHERE NO_WO = '" & Trim(NoWo) & "'")
                If Not rsBooK.EOF Then
                    If Trim(rsBooK("NOFAK")) <> "" Then
                        tnd = "w"
                    End If
                End If
                If rsBooK.State <> adStateClosed Then rsBooK.Close
            ElseIf rsAStl("FLAG") = "S" Then
                tnd = "P"
            ElseIf rsAStl("FLAG") = "P" Then
                tnd = "S"
            ElseIf rsAStl("FLAG") = "H" Then
                tnd = "H"
            ElseIf rsAStl("FLAG") = "Q" Then
                tnd = "Q"
            ElseIf rsAStl("FLAG") = "O" Then
                tnd = "O"
            ElseIf rsAStl("FLAG") = "G" Then
                tnd = "G"
            ElseIf rsAStl("FLAG") = "A" Then
                tnd = "A"
            ElseIf rsAStl("FLAG") = "F" Then
                tnd = "F"
            ElseIf rsAStl("FLAG") = "I" Then
                tnd = "I"
            ElseIf rsAStl("FLAG") = "K" Then
                tnd = "C"
            End If
            If Len(Trim(rsAStl!noStall)) > 6 Then
                nStl = CDbl(Mid(Trim(rsAStl("nostall")), 7, Len(Trim(rsAStl("nostall"))) - 6))  '+ 1
            Else
                nStl = CDbl(Mid(Trim(rsAStl("nostall")), 5, Len(Trim(rsAStl("nostall"))) - 4)) + jmlStall
            End If

            Dim S As New TrueDBGrid70.Style
            If sBold Then
                S.Font.Bold = False
                sBold = False
            Else
                S.Font.Bold = True
                sBold = True
            End If
            gridStall.Splits(1).AddCellStyle dbgCurrentCell, S

            If CDbl(rsAStl!Rate * 60) > CDbl(jamAkhir - hitungMenit(Format(Trim(rsAStl("JAMMULAI")), "HH:NN"))) Then
                rteBaru = (CDbl(jamAkhir - hitungMenit(Format(Trim(rsAStl("JAMMULAI")), "HH:NN"))) / 10) - 1
            Else
                rteBaru = CDbl(rsAStl!Rate * 6)
            End If

            Call SetBook(Left(Trim(rsAStl!jamMulai), 2), Right(Trim(rsAStl!jamMulai), 2), Val(rteBaru), (CDbl(nStl) * 2 - 1), tnd)
            rsAStl.MoveNext
        Loop
    End If
    'DARI TEMPPROGRESS - WAKTU ACTUAL MEKANIK
    Set rsAStl = objProgress.takeUnit(nTanggal, 2)
    If Not rsAStl.EOF Then
        Do While Not rsAStl.EOF
            tnd = "-"
            If Len(Trim(rsAStl!noStall)) > 6 Then
                nStl = CDbl(Mid(Trim(rsAStl("nostall")), 7, Len(Trim(rsAStl("nostall"))) - 6)) '+ 1
            Else
                nStl = CDbl(Mid(Trim(rsAStl("nostall")), 5, Len(Trim(rsAStl("nostall"))) - 4)) + jmlStall '+ 1
            End If
            If Trim(rsAStl("FLAGP")) = "" Then
                Call SetBook(Left(pembulatanJam(Trim(rsAStl!Start)), 2), Right(pembulatanJam(Trim(rsAStl!Start)), 2), hitungPanjang(rsAStl("START"), Format(isnow, "HH:NN")), CDbl(nStl) * 2, tnd)
                Call setMech(Trim(rsAStl("MECH")), Trim(rsAStl("NOORDER")), CDbl(nStl) * 2)
            Else
                Call SetBook(Left(pembulatanJam(Trim(rsAStl!Start)), 2), Right(pembulatanJam(Trim(rsAStl!Start)), 2), hitungPanjang(rsAStl("START"), rsAStl("PAUSE")), CDbl(nStl) * 2, tnd)
                Call setMech("", "", CDbl(nStl) * 2)
            End If
            rsAStl.MoveNext
        Loop
    End If

    ' DARI HELPPROGRESS
    Set rsAStl = objProgress.takeUnit(nTanggal, 3)
    If Not rsAStl.EOF Then
        Do While Not rsAStl.EOF
            If Trim(rsAStl("FLAG")) = "X" Then
                tnd = "X"
            End If
            If Len(Trim(rsAStl!noStall)) > 6 Then
                nStl = CDbl(Mid(Trim(rsAStl("nostall")), 7, Len(Trim(rsAStl("nostall"))) - 6))  '+ 1
            Else
                nStl = CDbl(Mid(Trim(rsAStl("nostall")), 5, Len(Trim(rsAStl("nostall"))) - 4)) + jmlStall
            End If
            Call SetBook(Left(Trim(rsAStl!jamMulai), 2), Right(Trim(rsAStl!jamMulai), 2), hitungPanjang(Trim(rsAStl("JAMMULAI")), Left(hitungWaktu(rsAStl("RATE"), rsAStl("JAMMULAI"), rsAStl("TANGGAL")), 5)), (CDbl(nStl) * 2 - 1), tnd) ', CDbl(rsAStl!Rate * 6), (CDbl(nStl) * 2 - 1), tnd)
            rsAStl.MoveNext
        Loop
    End If
    Call waktuSkrg(jamAwl, jamAkhir)
    Call idleMech
    Call saPreview
End Sub

Private Sub idleMech()
    Set rsMech = objO.FetchData("SELECT B.NAME + ' [' + CAST((SELECT COUNT(DISTINCT NOORDER) FROM TEMPPROGRESS WHERE FLAGP <> 'M' AND TANGGAL = CONVERT(varchar(10), GETDATE(),101) AND MECH = A.MECH) AS varchar(5)) + ']  ' + " & _
                "ISNULL((SELECT TOP 1 PAUSE FROM TEMPPROGRESS WHERE MECH = A.MECH AND FLAGP <> 'M' AND TANGGAL = CONVERT(varchar(10), GETDATE(),101) ORDER BY PAUSE DESC),'') AS NAME FROM T_MEKANIK A INNER JOIN MEKANIK B ON A.MECH = B.MECH WHERE MASUK = 1 AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND JAMMULAI <= '" & Format(isnow, "HH:NN") & "' AND JAMSELESAI >= '" & Format(isnow, "HH:NN") & "' AND A.MECH NOT IN (SELECT MECH FROM TEMPPROGRESS WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAGP = '') ORDER BY ISNULL((SELECT TOP 1 PAUSE FROM TEMPPROGRESS WHERE MECH = A.MECH AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' ORDER BY PAUSE DESC),''), B.MECH")
    If Not rsMech.EOF Then
        Set ListMech.DataSource = rsMech
    End If
End Sub

Private Sub saPreview()
    Set rsSA = objO.FetchData("SELECT FS_NM_PETUGAS + '  [' + CAST((SELECT COUNT(DISTINCT NO_WO) FROM JOBDTA WHERE FRONT = A.FS_KD_PETUGAS AND DATE = CONVERT(varchar(10), GETDATE(),101)) AS varchar(5)) + '] ' AS SA " & _
           "FROM T_PETUGAS A WHERE TIPE = 'SA' ORDER BY (SELECT COUNT(DISTINCT NO_WO) FROM JOBDTA WHERE FRONT = A.FS_KD_PETUGAS AND DATE = CONVERT(varchar(10), GETDATE(),101)) DESC, FS_NM_PETUGAS")
    If Not rsSA.EOF Then
        Set listSA.DataSource = rsSA
    End If
End Sub

Private Sub setMech(mekanik As String, nOrder As String, Pos As Integer)
Dim rsHlp As ADODB.Recordset
Dim nPol As String
Dim txtTampil As String

    txtTampil = Trim(mekanik)
    Set rsHlp = objProgress.takeMech(nOrder)
    If Not rsHlp.EOF Then
        If Trim(rsHlp("code")) = 1 Then
            txtTampil = "EM"
        End If
        nPol = Trim(rsHlp("NO_POLISI"))
        txtTampil = txtTampil & " (" & nPol & ") "  'txtTampil & IIf(txtTampil <> "", " (" & nPol & ") ", "")
    End If
    rsHlp.Close
    
    MyArray.value(Pos, 0) = txtTampil
    gridStall.DataSource = MyArray
End Sub

Private Function hitungPanjang(posAwal As String, posAkhir As String) As Integer
Dim jAw As Integer
Dim jAk As Integer
Dim jHsl As Integer
    
    jAw = (CDbl(Left(posAwal, InStr(posAwal, ":") - 1)) * 60) + CDbl(Mid(posAwal, InStr(posAwal, ":") + 1, 2))
    jAk = (CDbl(Left(posAkhir, InStr(posAkhir, ":") - 1)) * 60) + CDbl(Mid(posAkhir, InStr(posAkhir, ":") + 1, 2))
    jHsl = jAk - jAw

    hitungPanjang = Round((jHsl / 10) - 0.5, 0) '- 1
End Function

Private Sub SettingGrid(jamAwl As Integer, jamAkhir As Integer)
Dim nAwal As Integer
    
    nAwal = (((jamAkhir - jamAwl) / 10) + 2)
    For K = nAwal To AkhirPos   '(((jamAkhir - jamAwl) / 10) + 2)
        With gridStall
            .Splits(1).Columns(K).Width = 0
            .Splits(1).Columns(K).Visible = False
        End With
    Next
End Sub

Private Sub SetBook(ByVal JamAw As Integer, ByVal MenitAw As Integer, ByVal Panjang As Integer, ByVal Pos As Integer, sig As String)
    On Error GoTo err
    
    Dim i As Integer
    Dim PosisiKotak As Integer
    Dim PanjangKotak As Integer
    Dim TempPanjang As Integer
    
    If JamAw >= (njamAwl / 60) Or JamAw <= (njamAk / 60) Then
        
        PosisiKotak = Round(((((JamAw * 60) + MenitAw) - (njamAwl + (njamAwl Mod 10))) / 10), 0) + IIf(((((JamAw * 60) + MenitAw) - (njamAwl + (njamAwl Mod 10))) Mod 10) >= 0 And ((((JamAw * 60) + MenitAw) - (njamAwl + (njamAwl Mod 10))) Mod 10) < 6, 1, 0) '+ 0.5
        PanjangKotak = Panjang '- 1
        
        For i = 0 To PanjangKotak
            MyArray.value(Pos, PosisiKotak + i) = sig
        Next
        
        gridStall.Array = MyArray
    End If
    
    Exit Sub
err:
    MsgBox "Garis Flate Rate terlalu panjang !", vbOKOnly, "Error"
End Sub

Private Sub createTblJob(stat As Boolean)
    
    If stat Then
        If rsJob.State Then rsJob.Close
        With rsJob
            .Fields.Append "GROUP", adChar, 5
            .Fields.Append "JOB", adChar, 50
            .Fields.Append "KET", adChar, 50
            .Fields.Append "RATE", adDouble
            .Open
        End With
        
        Set gridJobOrder.DataSource = rsJob
        Call headColumns
    End If
End Sub

Private Sub headColumns()
    
    With gridJobOrder
        .Columns(0).Caption = "GROUP"
        .Columns(0).Width = 800
        .Columns(0).DataField = "GROUP"
        .Columns(0).Locked = True
        .Columns(1).Caption = "JOB ORDER"
        .Columns(1).Width = 4000
        .Columns(1).DataField = "JOB"
        .Columns(1).Locked = True
        .Columns(2).Caption = "KETERANGAN JOB"
        .Columns(2).Width = 2500
        .Columns(2).DataField = "KET"
        .Columns(2).Locked = True
        .Columns(3).Caption = "RATE"
        .Columns(3).Width = 800
        .Columns(3).DataField = "RATE"
        .Columns(3).Locked = True
    End With
End Sub

Private Sub headColumnsPending()
    
    With pendingGrid
        .Columns(0).Caption = "No. WO"
        .Columns(0).Width = 800
        .Columns(0).DataField = "NOWO"
        .Columns(0).Locked = True
        .Columns(1).Caption = "No. Polisi"
        .Columns(1).Width = 1000
        .Columns(1).DataField = "NOPOL"
        .Columns(1).Locked = True
        .Columns(2).Caption = "Customer"
        .Columns(2).Width = 2500
        .Columns(2).DataField = "CUST"
        .Columns(2).Locked = True
        .Columns(3).Caption = "Stall"
        .Columns(3).Width = 800
        .Columns(3).DataField = "STALL"
        .Columns(3).Locked = True
        .Columns(4).Caption = "Jam Mulai"
        .Columns(4).Width = 1200
        .Columns(4).DataField = "JAMMULAI"
        .Columns(4).Locked = True
        .Columns(5).Caption = "Mekanik"
        .Columns(5).Width = 800
        .Columns(5).DataField = "MECH"
        .Columns(5).Locked = True
        .Columns(6).Caption = "Pekerjaan"
        .Columns(6).Width = 3000
        .Columns(6).DataField = "REP"
        .Columns(6).Locked = True
    End With
End Sub

Private Sub clearMyarray(jmAw As Integer, jmAk As Integer)
    For i = 0 To (jmAk - jmAw) / 10 + 2
        For j = 1 To (jmlAll * 2) + 1
            MyArray.value(j, i) = ""
        Next
    Next
End Sub

Private Sub cmdBtl_Click()
    Text(9) = ""
    toStart.FormatString = "Pause"
    toStart.Visible = True
    toFinish.Left = 1200
    toFinish.Visible = True
    frmeStall.Height = toStart.Top + toStart.Height + 125
End Sub

Private Sub cmdHistory_Click()
    Call tblOther_Click
End Sub

Private Sub headGrid()
    With HistoryGrid
        .Columns(0).Caption = "Tanggal"
        .Columns(0).Width = 2000
        .Columns(0).DataField = "tgl"
        .Columns(0).Locked = True
        .Columns(1).Caption = "No. WO"
        .Columns(1).Width = 1200
        .Columns(1).DataField = "nowo"
        .Columns(1).Locked = True
        .Columns(2).Caption = "No. Invoice"
        .Columns(2).Width = 1200
        .Columns(2).DataField = "noinv"
        .Columns(2).Locked = True
        .Columns(3).Caption = "KM"
        .Columns(3).Width = 1000
        .Columns(3).DataField = "km"
        .Columns(3).Locked = True
        .Columns(4).Caption = "SA"
        .Columns(4).Width = 800
        .Columns(4).DataField = "frt"
        .Columns(4).Locked = True
        .Columns(5).Caption = "Mekanik"
        .Columns(5).Width = 800
        .Columns(5).DataField = "mech"
        .Columns(5).Locked = True
        .Columns(6).Caption = "Dealer"
        .Columns(6).Width = 1200
        .Columns(6).DataField = "kddealer"
        .Columns(6).Locked = True
        .Columns(7).Caption = "Tipe"
        .Columns(7).Width = 1000
        .Columns(7).DataField = "tp"
        .Columns(7).Locked = True
        .Columns(8).Caption = "Pekerjaan"
        .Columns(8).Width = 8000
        .Columns(8).DataField = "jobh"
        .Columns(8).Locked = True
        .Columns(9).Caption = "Keluhan"
        .Columns(9).Width = 6000
        .Columns(9).DataField = "keluh"
        .Columns(9).Locked = True
        .Columns(10).Caption = "Permintaan"
        .Columns(10).Width = 6000
        .Columns(10).DataField = "minta"
        .Columns(10).Locked = True
        .Columns(11).Caption = "Pesan"
        .Columns(11).Width = 6000
        .Columns(11).DataField = "pesan"
        .Columns(11).Locked = True
        .Columns(12).Caption = "Saran"
        .Columns(12).Width = 6000
        .Columns(12).DataField = "saran"
        .Columns(12).Locked = True
        .Columns(13).Caption = "Catatan"
        .Columns(13).Width = 6000
        .Columns(13).DataField = "catat"
        .Columns(13).Locked = True
    End With
End Sub

Private Sub cmdKeluar_Click()
    'WARNING PINDAH STALL
    Unload Me
End Sub

Private Sub cmdPause_Click()
Dim vPass As String             'Pilihan Pause

    If Trim(Text(9)) <> "" And Len(Trim(Text(9))) >= 7 Then
        vPass = IIf(opPause(0).value, opPause(0).Caption, IIf(opPause(1).value, opPause(1).Caption, IIf(opPause(2).value, opPause(2).Caption, IIf(opPause(3).value, opPause(3).Caption, opPause(4).Caption))))
        
        If chEM = Unchecked Then
            Call objProgress.pauseProses(Trim(Text(13)), Trim(Text(2)), stl, vPass & " | " & Text(9), Trim(Text(0)), nTgl, IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), CDbl(Text(3)), Trim(Text(5))) ', Format(isnow, "HH:NN"), CDbl(Text(3))) '
        Else
            Call objProgress.pauseProses2(Trim(Text(13)), stl, vPass & " | " & Text(9), Trim(Text(0)), nTgl, IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), CDbl(Text(3)), Trim(Text(5))) ', Format(isnow, "HH:NN"), CDbl(Text(3))) '
        End If
        frmeStall.Visible = False
        
        ' REFRESH GRID
        Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
        gridStall.Refresh
    Else
        MsgBox "Alasan harus diisi !!!" & vbCr & "Minimal 7 karakter"
    End If
End Sub

Private Sub cmdRepair_Click()
    Call objO.DeleteData("DELETE HELPPROGRESS WHERE TANGGAL = CONVERT(varchar(10), GETDATE(),101)")
    ' BOOKING
    Call objO.UpdateData("UPDATE STALL SET FLAG = 'X' WHERE TANGGAL < '" & Format(DateAdd("d", -1, isnow), "MM/DD/YYYY") & "' AND FLAG = 'K' AND LEFT(SUBSTRING(KETERANGAN, CHARINDEX(' |',KETERANGAN) + 2, CHARINDEX('| ',KETERANGAN) - 3 - LEN(LEFT(KETERANGAN, CHARINDEX(' |',KETERANGAN)))),2) = '" & Right(Year(isnow), 2) & "'")
    Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
    gridStall.Refresh
End Sub

Private Sub CmdSearch_Click()
    framSearch.Visible = True
    Call searchStall
End Sub

Private Sub searchStall()
    Set rsSearch2 = objO.FetchData("SELECT DISTINCT LEFT(KETERANGAN, CHARINDEX(' |', KETERANGAN)) AS NO_POLISI, SUBSTRING(KETERANGAN, CHARINDEX(' |', KETERANGAN)+2, CHARINDEX('| ', KETERANGAN) - CHARINDEX(' |', KETERANGAN)-2) AS NO_WO, " _
                               & "SUBSTRING(KETERANGAN, CHARINDEX('| ', KETERANGAN)+2, LEN(KETERANGAN) - CHARINDEX('| ', KETERANGAN)) AS CUSTOMER, NOSTALL, JAMMULAI, " _
                               & "POSISI = CASE FLAG WHEN 'S' THEN 'PROSES' WHEN 'P' THEN 'PAUSE' WHEN 'N' THEN 'WAITING' WHEN 'F' THEN 'FINISH' WHEN 'I' THEN 'INVOICING' WHEN 'K' THEN 'CLOSING' WHEN '' THEN 'BOOKING' WHEN 'H' THEN 'WAITING QC' WHEN 'Q' THEN 'QC' WHEN 'G' THEN 'WAITING WASHING' WHEN 'A' THEN 'WASHING' ELSE 'OTHER' END FROM STALL WHERE TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' and FLAG <> 'X'")
    If Not rsSearch2.EOF Then
    
        For Each col In gridSearch.Columns
            col.FilterText = ""
        Next col
        rsSearch2.filter = ""
        gridSearch.DataSource = rsSearch2
        Call doHeadColSearch
    End If
End Sub

Private Sub doHeadColSearch()
    
    With gridSearch
        .Columns(0).Width = 1200
        .Columns(0).DataField = "NO_POLISI"
        .Columns(0).Caption = "No. Polisi"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 1200
        .Columns(1).DataField = "NO_WO"
        .Columns(1).Caption = "No. Order"
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 2000
        .Columns(2).DataField = "CUSTOMER"
        .Columns(2).Caption = "Customer"
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 1200
        .Columns(3).DataField = "NOSTALL"
        .Columns(3).Caption = "Stall"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 800
        .Columns(4).DataField = "JAMMULAI"
        .Columns(4).Caption = "Mulai"
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 1500
        .Columns(5).DataField = "POSISI"
        .Columns(5).Caption = "Kondisi"
        .Splits(0).Columns(5).Locked = True
    End With
End Sub

Private Sub cmdTutup_Click()
    frmeStall.Visible = False
End Sub

Private Sub cmdTutupWait_Click()
    frmePending.Visible = False
End Sub

Private Sub Command1_Click()
    frmAturProgress.Show vbModal
End Sub

Private Sub Form_Activate()
    If batalLoad Then
        Unload Me
    Else
        Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
    End If
End Sub

Private Sub Form_Load()
        
    nTgl = Format(isnow, "DD/MM/yyyy")
    LblJudul.Caption = "CONTROL BOARD ... HARI : " & Format(nTgl, "dddd") & ", TANGGAL : " & Format(nTgl, "dd mmmm yyyy")
    Me.Tag = ""
    batalLoad = True
    
    With gridStall.Splits(0)
        .SizeMode = dbgNumberOfColumns
        .Size = 1
        .AllowSizing = False
        For i = 1 To AkhirPos
            .Columns(i).Visible = False
        Next
    End With
    
    Set rsStall = objProgress.takeStallTime(Trim(IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)))
    If Not rsStall.EOF Then
        If Trim(rsStall("WAKTUMULAI")) <> "" And Trim(rsStall("WAKTUSELESAI")) <> "" Then
            Dim wkMulai As String
            Dim wkSelesai As String
            wkMulai = Trim(rsStall("waktuMulai"))
            wkSelesai = Trim(rsStall("waktuSelesai"))
            
            njamAwl = (CDbl(Left(wkMulai, 2)) * 60) + CDbl(Mid(wkMulai, InStr(rsStall("waktuMulai"), ":") + 1, 2))
            njamAk = (CDbl(Left(wkSelesai, 2)) * 60) + CDbl(Mid(wkSelesai, InStr(rsStall("waktuSelesai"), ":") + 1, 2))
            
            If Format(isnow, "HH:NN") >= Format(Left(wkMulai, 2), "0#") & ":" & Format(CDbl(Mid(wkMulai, InStr(rsStall("waktuMulai"), ":") + 1, 2)), "0#") And Format(isnow, "HH:NN") <= Format(Left(Trim(rsStall!waktuSelesai), 2), "0#") & ":" & Format(CDbl(Mid(wkSelesai, InStr(rsStall("waktuSelesai"), ":") + 1, 2)), "0#") Then
                sBold = True
                batalLoad = False
                stallAktif = True
                                
                Call SetWaktu(njamAwl, njamAk)
                Call SettingGrid(njamAwl, njamAk)
                Call setStall(njamAwl, njamAk)
                
                frmeStall.Height = toStart.Top + toStart.Height + 125
                frmeStall.Visible = False
                gridStall.Splits(1).Columns(0).Visible = False
                Timer1.Enabled = True
            Else
                MsgBox "Jam Sistem tidak sesuai dengan jam Stall."
                batalLoad = True
                Exit Sub
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
    
    '' ATUR HORISONTAL
    ListMech.Left = Me.Width - ListMech.Width - 100
    Frame1.Width = ListMech.Left - 200
    gridStall.Width = Frame1.Width - 200
    LblJudul.Width = gridStall.Width
    LblJudul.Left = gridStall.Left
    
    '' ATUR VERTIKAL
    Command1.Top = Me.Height - 600 - Command1.Height
    Frame1.Height = Command1.Top - 100
    ListMech.Height = Frame1.Height * (5 / 8)
    listSA.Top = ListMech.Top + ListMech.Height + 50
    listSA.Height = (Frame1.Height * (3 / 8)) - 150
    listSA.Left = ListMech.Left
    listSA.Width = ListMech.Width
    legend.Top = Me.Height - 600 - legend.Height
    legend.Width = cmdKeluar.Left - Command1.Left - Command1.Width - 100
    gridStall.Height = Frame1.Height - cmdSearch.Height - 200
    
    cmdKeluar.Left = ListMech.Left + ListMech.Width - 1200
    cmdKeluar.Top = Me.Height - 600 - cmdKeluar.Height
    
    cmdRepair.Left = gridStall.Width - cmdRepair.Width
End Sub

Private Sub frmeStall_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xy = True
End Sub

Private Sub frmeStall_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If xy Then
        frmeStall.Left = Val(frmeStall.Left) + X
        frmeStall.Top = Val(frmeStall.Top) + Y
    End If
End Sub

Private Sub frmeStall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xy = False
End Sub

Private Sub frameHistory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xy1 = True
End Sub

Private Sub frameHistory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If xy1 Then
        frameHistory.Left = Val(frameHistory.Left) + X
        frameHistory.Top = Val(frameHistory.Top) + Y
    End If
End Sub

Private Sub frameHistory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    xy1 = False
End Sub

Private Sub GantiBatal_Click()
    Text(14) = ""
    frmeStall.Height = toStart.Top + toStart.Height + 125
End Sub

Private Sub gantiMech_Click()
    Text(14).Enabled = True
    Text(14) = ""
    Text(14).SelLength = Len(Text(14))
    Call isiCombo
    toStart.Visible = False
    toFinish.Left = 1200
    toFinish.Visible = False
    toMove.Visible = False
    frmePause.Visible = False
    frmeMekanik.Visible = True
    frmeMekanik.Top = toStart.Top + toStart.Height + 150
    
    If chEM = Checked Then
        frmeMekanik.Height = Text(14).Top + frmeMechEM.Height + 500
        frmeMechEM.Visible = True
        If baris = 1 Then
            Call aturGridMechEM
        End If
    Else
        frmeMekanik.Height = Text(14).Top + Text(14).Height + 200
        frmeMechEM.Visible = False
    End If
    frmeStall.Height = frmeStall.Height + frmeMekanik.Height + 300
End Sub

Private Sub aturGridMechEM()
    gridEM.Clear
    baris = 1
    gridEM.Rows = maxEM + 1
    gridEM.cols = 2
    gridEM.ColWidth(0) = 2500
    gridEM.ColWidth(1) = 800
    gridEM.ColAlignment(1) = flexAlignLeftCenter
    gridEM.TextMatrix(0, 0) = "Mekanik EM"
End Sub

Private Sub isiCombo()
Dim hlp As String
Dim rsHlp As ADODB.Recordset
    
     
    Set rsHlp = objO.FetchData("SELECT FN_JML_STALL FROM T_PROFIL")
    If Not rsHlp.EOF Then
        cboStall.Clear
        For i = 1 To rsHlp("FN_JML_STALL")
            hlp = "Stall " & i
            cboStall.AddItem hlp
        Next
    End If
End Sub

Private Sub GantiOK_Click()
Dim namaMech As String

    If chEM = Checked Then
        If Trim(Text(14)) = "" Then
            Text(14) = Trim(gridEM.TextMatrix(1, 1))
        End If
        
        namaMech = "'"
        For i = 1 To maxEM
            If Trim(gridEM.TextMatrix(i, 1)) <> "" Then
                namaMech = namaMech & Trim(gridEM.TextMatrix(i, 1)) & "','"
            End If
        Next
        If namaMech <> "'" Then namaMech = namaMech & "'"
    End If
    
    If Trim(Text(14)) <> "" Then
        If chEM = Unchecked Then
            If objProgress.changeMech(Trim(Text(13)), nTgl, Trim(Text(0)), stl, Trim(Text(14)), Trim(Text(3)), IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), Trim(Text(5))) Then
                'Trim(cboStall),
                Text(2) = Trim(Text(14))
            End If
        Else
            If objProgress.changeMech2(Trim(Text(13)), nTgl, Trim(Text(0)), stl, Trim(namaMech), CDbl(Text(3)), IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), Trim(Text(5))) Then
                Text(2) = Trim(Text(14))
            End If
        End If
        Me.Tag = ""
        Text(14) = ""
        Text(15) = 0
        frmeStall.Visible = False
        
        ' REFRESH GRID
            Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
            gridStall.Refresh
    End If
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
    For Each col In cols
        If Trim(col.FilterText) <> "" Then
            n = n + 1
            If n > 1 Then
                tmp = tmp & " AND "
            End If
            tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        
        End If
    Next col
    getFilter = tmp
End Function

Private Sub gridEM_DblClick()
Dim a As Integer

    If baris > 1 Then
        a = gridEM.row
        
        With gridEM
            Text(14) = .TextMatrix(a, 1)
            
            For i = a To baris
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
            Next
        End With
        baris = baris - 1
    End If
End Sub

Private Sub gridEM_GotFocus()
Dim rsMekanik As ADODB.Recordset
Dim bolehTambah As Boolean
Dim namaMekanik As String
    
    If chEM = Checked Then
        If baris <= maxEM Then
        
            bolehTambah = True
            For i = 0 To baris
                If gridEM.TextMatrix(i, 1) <> "" Then
                    If gridEM.TextMatrix(i, 1) = Trim(Text(14)) Then
                        bolehTambah = False
                    End If
                End If
            Next
            
            'LIHAT GROUP MEKANIK DAN NAMA MEKANIK
            Set rsMekanik = objO.FetchData("SELECT * FROM MEKANIK WHERE MECH = '" & Trim(Text(14)) & "'")
            If Not rsMekanik.EOF Then
                namaMekanik = Trim(rsMekanik("NAME"))
            End If
            
            If bolehTambah Then
                If Trim(Text(14)) <> "" Then
                    With gridEM
                        .TextMatrix(baris - 1, 0) = Trim(namaMekanik)
                        .TextMatrix(baris - 1, 1) = Trim(Text(14))
                    End With
                    baris = baris + 1
                    Text(14) = ""
                    Text(15) = 0
                    Text(14).SetFocus
                End If
            Else
                MsgBox "Mekanik sudah diinput."
                Text(14) = ""
                Text(15) = 0
                Text(14).SetFocus
            End If
        End If
    End If
End Sub

Private Sub gridSearch_FilterChange()
    On Error GoTo errHandler
    Set cols = gridSearch.Columns
    Dim C As Integer
    
    C = gridSearch.col
    gridSearch.HoldFields
    rsSearch2.filter = getFilter()
    If rsSearch2.EOF Then
        For Each col In gridSearch.Columns
            col.FilterText = ""
        Next col
        rsSearch2.filter = ""
    End If
    gridSearch.col = C
    gridSearch.EditActive = True
    Exit Sub
    
errHandler:
    For Each col In cols
        col.FilterText = ""
    Next col
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub gridStall_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    F4.Top = Y
    F4.Left = X
End Sub

Private Sub gridStall_RowColChange(LastRow As Variant, ByVal LastCol As Integer)

Dim jmS As String
Dim rsU As ADODB.Recordset
Dim rsUt As ADODB.Recordset
Dim rsJ As ADODB.Recordset
Dim nFlag As String
Dim noOrd As String
Dim Txt As String
Dim namaMekanik As String
    
    coll = gridStall.col  'LastCol + 1
    row = gridStall.row
    lblStatus = ""
    
    With gridStall
        sgnPre.Visible = False
        frmeStall.Visible = False
        If coll > 0 Then
        stl = gridStall.Splits(0).Columns(0)
            If UCase(.Splits(1).Columns(coll)) <> "" And gridStall.Splits(0).Columns(0) <> "" Then
                
                Call hapusForm
                Me.Tag = ""
                
                ' ISI DATA UNIT KE FRAME
                Dim jamskr As Integer
                Dim menitskr As Integer
                jamskr = Round(njamAwl / 60) + IIf(coll <= batasAwal, Val((coll - 1) / batasAwal), Val((batasAwal / batasAwal) + (coll - batasAwal - 1) / 6))
                menitskr = IIf(coll <= batasAwal, ((6 - batasAwal) + ((coll - 1) Mod batasAwal)) * 10, (((coll - batasAwal) - 1) Mod 6) * 10)
                jmS = Format(jamskr, "0#") & ":" & Format(menitskr, "0#")
                Set rsU = objProgress.takeDesc(Trim(stl), jmS, IIf(Trim(UCase(.Splits(1).Columns(coll))) = "W", "N", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "P", "S", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "S", "P", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "F", "F", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "I", "I", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "B", "", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "|", "", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "H", "H", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "Q", "Q", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "O", "O", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "G", "G", IIf(Trim(UCase(.Splits(1).Columns(coll))) = "A", "A", "K")))))))))))), isnow)
                If Not rsU.EOF Then
                    If CDbl(rsU.RecordCount) > 0 Then
                            Txt = Trim(rsU("KETERANGAN"))
                            noOrd = Left(Mid(Txt, InStr(Txt, "|") + 1, Len(Txt) - InStr(Txt, "|")), InStr(Mid(Txt, InStr(Txt, "|") + 1, Len(Txt) - InStr(Txt, "|")), "|") - 1)
                            nFlag = IIf(Trim(rsU("FLAG")) = "N", "W", IIf(Trim(rsU("FLAG")) = "S", "P", IIf(Trim(rsU("FLAG")) = "P", "S", IIf(Trim(rsU("FLAG")) = "F", "F", IIf(Trim(rsU("FLAG")) = "I", "I", IIf(Trim(rsU("FLAG")) = "", "B", IIf(Trim(rsU("FLAG")) = "H", "H", IIf(Trim(rsU("FLAG")) = "Q", "Q", IIf(Trim(rsU("FLAG")) = "O", "O", IIf(Trim(rsU("FLAG")) = "G", "G", IIf(Trim(rsU("FLAG")) = "A", "A", "C")))))))))))
                            
                            If nFlag <> UCase(.Splits(1).Columns(coll)) Then
                                If UCase(.Splits(1).Columns(coll)) = "|" Then
                                    .Splits(1).Columns(coll) = nFlag
                                    GoTo lanjut
                                ElseIf UCase(.Splits(1).Columns(coll)) = "?" Then
                                    GoTo WSTATE
                                ElseIf UCase(.Splits(1).Columns(coll)) <> "?" Then
                                    Exit Sub
                                End If
                            End If
                            
lanjut:
                            ' CARI DI JOBDTA
                            Set rsUt = objProgress.takeDetailJob(noOrd, 1)
                            If Not rsUt.EOF Then
                                chEM.Enabled = True
                                chEM = IIf(rsUt("CODE") = 1, Checked, Unchecked)
                                chEM.Enabled = False
                                chSBNP.Enabled = True
                                chSBNP = IIf(rsUt("KD_K") = 1, Checked, Unchecked)
                                chSBNP.Enabled = False
                                Text(0) = rsUt("NO_CHASSIS")
                                Text(1) = rsUt("NO_POLISI")
                                Text(2) = IIf(Trim(rsU("MECH")) <> "", Trim(rsU("MECH")), Trim(rsUt("MECH")))
                                Text(3) = IIf(chEM = Checked, "0,5", IIf(chSBNP = Checked, "1", CDbl(rsUt("ACTRATE1")) + CDbl(rsUt("ACTRATE2")) + CDbl(rsUt("ACTRATE3")) + CDbl(rsUt("ACTRATE4")) + CDbl(rsUt("ACTRATE5")) + CDbl(rsUt("ACTRATE6")) + CDbl(rsUt("ACTRATE7")) + CDbl(rsUt("ACTRATE8"))))
                                Text(4) = rsU("JAMMULAI")
                                Text(5) = rsU("JAMSELESAI")
                                Text(10) = rsUt("MDL")
                                Text(11) = rsUt("FMDL")
                                Text(12) = rsUt("COLOR")
                                Text(18) = rsUt("FRONT")
                                lblStatus = IIf(rsUt("TUNGGUSVC") = 0, "DITUNGGU", "DITINGGAL")
                                lblBook.Visible = False
                                shpBook.Visible = False
                                If Trim(rsUt("NOFAK")) <> "" Then
                                    lblBook.Visible = True
                                    shpBook.Visible = True
                                End If
                                Text(21) = rsUt("DATESERAH")
                                Text(20) = rsUt("TIMESERAH")
                                Label10.Caption = "No. WO"
                                Text(13) = noOrd
                                PRC = IIf(rsUt("FS_RC") <> "", True, False)
                                namaMekanik = Trim(rsUt("SPK")) & ",''"
                                Dim i As Integer
                                Call aturGridMechEM
                                For i = 1 To jmlMekanikEM(namaMekanik)
                                    Set rsJ = objO.FetchData("SELECT * FROM MEKANIK WHERE MECH = '" & Trim(pecahMekanik(namaMekanik, i)) & "'")
                                    If Not rsJ.EOF Then
                                        gridEM.TextMatrix(i, 0) = Trim(rsJ("NAME"))
                                        gridEM.TextMatrix(i, 1) = pecahMekanik(namaMekanik, i)
                                    End If
                                    If rsJ.State <> adStateClosed Then rsJ.Close
                                    baris = baris + 1
                                Next
                                
                                If rsJob.State <> adStateClosed Then
                                    rsJob.Close
                                    Set rsJ = Nothing
                                End If
                                For i = 1 To 7
                                    tmp2 = "KLPKERJA" & i
                                    tmp = "REP" & i
                                    tmp1 = "ACTRATE" & i
                                    If rsUt(tmp) <> "" Then
                                        Set rsJ = objProgress.takeDeffJob(Trim(rsUt(tmp2)), Trim(rsUt(tmp)), Trim(rsUt("MDL")))
                                        If Not rsJ.EOF Then
                                            Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                                            With rsJob
                                                .AddNew
                                                !Group = rsJ("JOB")
                                                !Job = rsJ("FS_ITEM")
                                                !ket = rsJ("KETERANGAN")
                                                !Rate = rsUt(tmp1)
                                                .Update
                                            End With
                                            Set gridJobOrder.DataSource = rsJob
                                            Call headColumns
                                        End If
                                        If rsJ.State <> adStateClosed Then rsJ.Close
                                    End If
                                Next
                            End If
                            rsUt.Close
                            
                            ' CARI DI ADDITIONAL JOB
                            Set rsUt = objO.FetchData("SELECT * FROM JOBADD2 WHERE NO_WO = '" & noOrd & "'")
                            If Not rsUt.EOF Then
                                Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                                Do While Not rsUt.EOF
                                    With rsJob
                                        .AddNew
                                        !Group = Trim(rsUt("KLPKERJA"))
                                        !Job = Trim(rsUt("REP"))
                                        !ket = ""
                                        !Rate = CDbl(Trim(rsUt("ACTRTE")))
                                        .Update
                                    End With
                                    Text(3) = Text(3) + CDbl(rsUt("ACTRTE"))
                                    rsUt.MoveNext
                                Loop
                            End If
                            rsUt.Close
                            
                            ' CARI DI SUBLET
                            Set rsUt = objO.FetchData("SELECT * FROM JOBSUBLET WHERE NO_WO = '" & noOrd & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "'")
                            If Not rsUt.EOF Then
                                Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                                Do While Not rsUt.EOF
                                    With rsJob
                                        .AddNew
                                        !Group = "SUB"
                                        !Job = Trim(rsUt("NAMABARANG"))
                                        !ket = Trim(rsUt("KODESATUAN"))
                                        !Rate = 0
                                        .Update
                                    End With
                                    rsUt.MoveNext
                                Loop
                            End If
                            rsUt.Close
                            
                            ' CARI DI BOOKING
                            Set rsUt = objProgress.searchFromBooking(noOrd)
                            If Not rsUt.EOF Then
                                Text(0) = rsUt("NO_CHASSIS")
                                Text(1) = rsUt("NO_POLISI")
                                Text(2) = rsUt("MECH")
                                Text(3).Mask = ""
                                Text(3) = rsU("RATE")
                                Text(4) = rsU("JAMMULAI")
                                Text(5) = rsU("JAMSELESAI")
                                Text(10) = rsUt("MDL")
                                Text(11) = rsUt("FMDL")
                                Text(12) = rsUt("COLOR")
                                lblStatus = Trim(rsUt("FLAGSVC"))
                                Label10.Caption = "No. Book"
                                Text(13) = noOrd
                                
                                If rsJob.State <> adStateClosed Then
                                    rsJob.Close
                                    Set rsJ = Nothing
                                End If
                                For i = 1 To 7
                                    tmp = "REP" & i
                                    tmp1 = "RTE" & i
                                    Dim rsSearch As ADODB.Recordset
                                    Set rsSearch = objO.FetchData("SELECT * FROM ACTRATE WHERE FS_ITEM = '" & Trim(rsUt(tmp)) & "'")
                                    If Not rsSearch.EOF Then
                                        tmp2 = Trim(rsSearch("JOB"))
                                    End If
                                    rsSearch.Close
                                    
                                    If rsUt(tmp) <> "" Then
                                        Set rsJ = objProgress.takeDeffJob(Trim(tmp2), Trim(rsUt(tmp)), Trim(rsUt("MDL")))
                                        If Not rsJ.EOF Then
                                            Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                                            With rsJob
                                                .AddNew
                                                !Group = rsJ("JOB")
                                                !Job = rsJ("FS_ITEM")
                                                !ket = rsJ("KETERANGAN")
                                                !Rate = rsJ("RATE")
                                                .Update
                                            End With
                                        End If
                                        If rsJ.State <> adStateClosed Then rsJ.Close
                                    End If
                                Next
                            End If
                            
                            ' CARI DI diagDTA
                            Set rsUt = objProgress.takeDetailDiag(noOrd, 1)
                            If Not rsUt.EOF Then
                                sgnPre.Visible = True
                                chEM.Enabled = True
                                chEM = IIf(rsUt("CODE") = 1, Checked, Unchecked)
                                chEM.Enabled = False
                                chSBNP.Enabled = True
                                chSBNP = IIf(rsUt("KD_K") = 1, Checked, Unchecked)
                                chSBNP.Enabled = False
                                Text(0) = rsUt("NO_CHASSIS")
                                Text(1) = rsUt("NO_POLISI")
                                Text(2) = IIf(Trim(rsU("MECH")) <> "", Trim(rsU("MECH")), Trim(rsUt("MECH")))
                                Text(3) = IIf(chEM = Checked, "0,5", IIf(chSBNP = Checked, "1", CDbl(rsUt("ACTRATE1")) + CDbl(rsUt("ACTRATE2")) + CDbl(rsUt("ACTRATE3")) + CDbl(rsUt("ACTRATE4")) + CDbl(rsUt("ACTRATE5")) + CDbl(rsUt("ACTRATE6")) + CDbl(rsUt("ACTRATE7")) + CDbl(rsUt("ACTRATE8"))))
                                Text(4) = rsU("JAMMULAI")
                                Text(5) = rsU("JAMSELESAI")
                                Text(10) = rsUt("MDL")
                                Text(11) = rsUt("FMDL")
                                Text(12) = rsUt("COLOR")
                                Text(18) = rsUt("FRONT")
                                lblStatus = IIf(rsUt("TUNGGUSVC") = 0, "DITUNGGU", "DITINGGAL")
                                lblBook.Visible = False
                                shpBook.Visible = False
                                If Trim(rsUt("NOFAK")) <> "" Then
                                    lblBook.Visible = True
                                    shpBook.Visible = True
                                End If
                                Text(21) = rsUt("DATESERAH")
                                Text(20) = rsUt("TIMESERAH")
                                Label10.Caption = "No. WO"
                                Text(13) = noOrd
                                PRC = IIf(rsUt("FS_RC") <> "", True, False)
                                namaMekanik = Trim(rsUt("SPK")) & ",''"
                                Call aturGridMechEM
                                For i = 1 To jmlMekanikEM(namaMekanik)
                                    Set rsJ = objO.FetchData("SELECT * FROM MEKANIK WHERE MECH = '" & Trim(pecahMekanik(namaMekanik, i)) & "'")
                                    If Not rsJ.EOF Then
                                        gridEM.TextMatrix(i, 0) = Trim(rsJ("NAME"))
                                        gridEM.TextMatrix(i, 1) = pecahMekanik(namaMekanik, i)
                                    End If
                                    If rsJ.State <> adStateClosed Then rsJ.Close
                                    baris = baris + 1
                                Next

                                If rsJob.State <> adStateClosed Then
                                    rsJob.Close
                                    Set rsJ = Nothing
                                End If
                                For i = 1 To 7
                                    tmp2 = "KLPKERJA" & i
                                    tmp = "REP" & i
                                    tmp1 = "ACTRATE" & i
                                    If rsUt(tmp) <> "" Then
                                        Set rsJ = objProgress.takeDeffJob(Trim(rsUt(tmp2)), Trim(rsUt(tmp)), Trim(rsUt("MDL")))
                                        If Not rsJ.EOF Then
                                            Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                                            With rsJob
                                                .AddNew
                                                !Group = rsJ("JOB")
                                                !Job = rsJ("FS_ITEM")
                                                !ket = rsJ("KETERANGAN")
                                                !Rate = rsUt(tmp1)
                                                .Update
                                            End With
                                            Set gridJobOrder.DataSource = rsJob
                                            Call headColumns
                                        End If
                                        If rsJ.State <> adStateClosed Then rsJ.Close
                                    End If
                                Next
                            End If
                            If rsUt.State <> adStateClosed Then rsUt.Close
                            
                            ' CARI DI ADDITIONAL JOB
                            Set rsUt = objO.FetchData("SELECT * FROM diagADD2 WHERE NO_diag = '" & noOrd & "'")
                            If Not rsUt.EOF Then
                                Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                                Do While Not rsUt.EOF
                                    With rsJob
                                        .AddNew
                                        !Group = Trim(rsUt("KLPKERJA"))
                                        !Job = Trim(rsUt("REP"))
                                        !ket = ""
                                        !Rate = CDbl(Trim(rsUt("ACTRTE")))
                                        .Update
                                    End With
                                    Text(3) = Text(3) + CDbl(rsUt("ACTRTE"))
                                    rsUt.MoveNext
                                Loop
                            End If
                            If rsUt.State <> adStateClosed Then rsUt.Close
                            
                            ' LIHAT PAUSE STATUS DI PROGRESS
                            Set rsUt = objProgress.takePauseState(Trim(noOrd), Trim(Text(0)), nTgl)
                            If Not rsUt.EOF Then
                                Text(6) = rsUt("TANGGAL")
                                Text(7) = rsUt("JAM")
                                Text(8) = rsUt("USERADD")
                                Text(19) = Trim(Mid(rsUt("KETERANGAN"), InStr(rsUt("KETERANGAN"), "|") + 1, Len(rsUt("KETERANGAN")) - InStr(rsUt("KETERANGAN"), "|")))
                            End If
                            rsUt.Close
                        
                        frmeStall.Height = toStart.Top + toStart.Height + 120
                        toStart.Visible = False
                        toFinish.Visible = False
                        toMove.Visible = False
                        gantiMech.Visible = False
                        StartInsp.Visible = False
                        StopInsp.Visible = False
                        StartWash.Visible = False
                        StopWash.Visible = False
                        
                        If UCase(.Splits(1).Columns(coll)) = "W" Then
                            toStart.FormatString = " Start"
                            toMove.FormatString = "Pindah Stall"
                            toStart.Visible = True
                            toMove.Visible = True
                            gantiMech.Visible = True
                            toFinish.Visible = False
                        ElseIf UCase(.Splits(1).Columns(coll)) = "P" Then
                            If Not sgnPre.Visible Then
                                toStart.FormatString = "Pause"
                                toMove.FormatString = "Pindah Stall"
                                toStart.Visible = True
                            Else
                                toFinish.Left = 1200
                                toFinish.Visible = True
                            End If
                        ElseIf UCase(.Splits(1).Columns(coll)) = "S" Then
                            toStart.FormatString = " Start"
                            toMove.FormatString = "Pindah Stall"
                            toStart.Visible = True
                            toMove.Visible = True
                            gantiMech.Visible = True
                        ElseIf UCase(.Splits(1).Columns(coll)) = "B" Then
                            toStart.Visible = False
                            toMove.FormatString = "Pindah Stall"
                            toMove.Visible = True
                            toFinish.Visible = False
                        ElseIf UCase(.Splits(1).Columns(coll)) = "H" Then
                            StartInsp.Visible = True
                            StartInsp.Left = toStart.Left
                            StartInsp.Top = toStart.Top
                            toMove.FormatString = "Pindah Stall"
                            toMove.Visible = True
                        ElseIf UCase(.Splits(1).Columns(coll)) = "Q" Then
                            toFinish.Left = 1200
                            toFinish.Visible = True
                        ElseIf UCase(.Splits(1).Columns(coll)) = "G" Then
                            toStart.FormatString = " Start"
                            toStart.Visible = True
                            StartWash.Visible = True
                            StartWash.Left = 1150
                            StartWash.Top = toStart.Top
                            toMove.FormatString = "Pindah Stall"
                            toMove.Visible = True
                        ElseIf UCase(.Splits(1).Columns(coll)) = "A" Then
                            toFinish.Left = 1200
                            toFinish.Visible = True
                        ElseIf UCase(.Splits(1).Columns(coll)) = "I" Then
                            toStart.Visible = False
                            toMove.Visible = False
                            toFinish.Visible = False
                        ElseIf UCase(.Splits(1).Columns(coll)) = "F" Then
                            toStart.FormatString = " Start"
                            toMove.FormatString = "Pindah Stall"
                            toStart.Visible = True
                            toMove.Visible = True
                        End If
                        Call panggilHelpX(jmS, gridStall.Splits(0).Columns(0))
                        frmeStall.Visible = True
                    End If
                End If
            
WSTATE:
                If UCase(gridStall.Splits(1).Columns(coll)) = "?" Then
                ' CARI JOB PENDING
                Dim rsSearch3 As ADODB.Recordset
                Dim hlp As String
                Dim hlp2 As String
                
                Set rsU = objProgress.takeDescPending(Trim(stl), jmS, Format(isnow, "mm/dd/yyyy"))
                If Not rsU.EOF Then
                    Do While Not rsU.EOF
                        hlp2 = ""
                        If rsPending.State Then rsPending.Close
                        If rsPending.State = adStateClosed Then createPending
                        If rsPending.State <> adStateClosed Then
                            For i = 1 To 8
                                hlp = "REP" & i
                                If Trim(rsU(hlp)) <> "" Then
                                    With rsPending
                                        .AddNew
                                        !NoWo = IIf(Trim(hlp2) = Trim(rsU("NOWO")), "", Trim(rsU("NOWO")))
                                        !noPol = IIf(Trim(hlp2) = Trim(rsU("NOWO")), "", Trim(rsU("NOPOL")))
                                        !cust = IIf(Trim(hlp2) = Trim(rsU("NOWO")), "", Trim(rsU("CUSTOMER")))
                                        !stall = IIf(Trim(hlp2) = Trim(rsU("NOWO")), "", Trim(rsU("NOSTALL")))
                                        !jamMulai = IIf(Trim(hlp2) = Trim(rsU("NOWO")), "", Trim(rsU("JAMMULAI")))
                                        !MECH = IIf(Trim(hlp2) = Trim(rsU("NOWO")), "", Trim(rsU("MECH")))
                                        !REP = Trim(rsU(hlp))
                                        .Update
                                    End With
                                End If
                                If i = 1 Then
                                    hlp2 = Trim(rsU("NOWO"))
                                End If
                            Next
                        End If
                        
                        Set rsSearch3 = objProgress.takeDescAdd(Trim(rsU("NOWO")))
                        If Not rsSearch3.EOF Then
                            Do While Not rsSearch3
                                If rsPending.State = adStateClosed Then createPending
                                If rsPending.State <> adStateClosed Then
                                    With rsPending
                                        .AddNew
                                        !NoWo = Trim(rsU("NO_WO"))
                                        !noPol = Trim(rsU("NOPOL"))
                                        !cust = Trim(rsU("CUSTOMER"))
                                        !stall = Trim(rsU("NOSTALL"))
                                        !jamMulai = Trim(rsU("JAMMULAI"))
                                        !MECH = Trim(rsU("MECH"))
                                        !REP = Trim(rsSearch("REP"))
                                        .Update
                                    End With
                                End If
                                rsSearch3.MoveNext
                            Loop
                        End If
                        rsSearch3.Close
                        
                        rsU.MoveNext
                    Loop
                    Set pendingGrid.DataSource = rsPending
                    Call headColumnsPending
                    frmePending.Visible = True
                Else
                    Set rsSearch3 = objProgress.takeDescBook(Trim(noOrd))
                    If Not rsSearch3.EOF Then
                        Do While Not rsSearch3.EOF
                            If rsPending.State = adStateClosed Then createPending
                            If rsPending.State <> adStateClosed Then
                                With rsPending
                                    .AddNew
                                    !NoWo = Trim(rsSearch3("NOFAK"))
                                    !noPol = Trim(rsSearch3("NO_POLISI"))
                                    !cust = Trim(rsSearch3("CUSTOMER"))
                                    !stall = Trim(rsSearch3("NOSTALL"))
                                    !jamMulai = Trim(rsSearch3("JAMMULAI"))
                                    !MECH = Trim(rsSearch3("MECH"))
                                    !REP = Trim(rsSearch3("REP"))
                                    .Update
                                End With
                            End If
                            rsSearch3.MoveNext
                        Loop
                    End If
                End If
                rsU.Close
                End If
                
            ElseIf UCase(.Splits(1).Columns(coll)) = "" And gridStall.Splits(0).Columns(0) <> "" Then
                If Me.Tag <> "" Then
                    toStart.Visible = False
                    toFinish.Visible = False
                    gantiMech.Visible = False
                    StartInsp.Visible = False
                    StopInsp.Visible = False
                    StartWash.Visible = False
                    StopWash.Visible = False

                    toMove.FormatString = "Pakai Stall"
                    toMove.Visible = True
                    jamskr = Round(njamAwl / 60) + IIf(coll <= batasAwal, Val((coll - 1) / batasAwal), Val((batasAwal / batasAwal) + (coll - batasAwal - 1) / 6))
                    menitskr = IIf(coll <= batasAwal, ((6 - batasAwal) + ((coll - 1) Mod batasAwal)) * 10, (((coll - batasAwal) - 1) Mod 6) * 10)
                    Text(4) = Format(jamskr, "0#") & ":" & Format(menitskr, "0#")
                    If Format(jamskr, "0#") & ":" & Format(menitskr, "0#") < Format(isnow, "HH:NN") Then
                        MsgBox "Tidak boleh lebih kecil dari jam sekarang. " & vbCr & "Jam Sekarang : " & Format(isnow, "HH:NN")
                    Else
                        If Trim(Text(3)) <> 0 Then
                            X = InStr(Format(Text(3), "0#.#0"), ",")
                            If X > 0 Then
                                Jam = jamskr + CDbl(Left(Format(Text(3), "0#.#0"), X - 1))
                                menit = menitskr + (Mid(Format(Text(3), "0#.#0"), X + 1, (Len(Format(Text(3), "0#.#0")) - X)) / 10 * 6)
                                If menit >= 60 Then
                                    Jam = Jam + Val(menit / 60)
                                    menit = menit Mod 60
                                End If
                            Else
                                Jam = jamskr + Text(3)
                                menit = menitskr
                            End If
                            If Jam >= 24 Then
                                Jam = 23
                                menit = 59
                            End If
                        End If
                        Text(5) = Format(Jam, "0#") & ":" & Format(menit, "0#")
                        frmeStall.Visible = True
                    End If
                Else
                    frmeStall.Visible = False
                End If
            End If
        End If
    End With
End Sub

Private Sub panggilHelpX(jamS As String, baris As String)
Dim rsHlp As ADODB.Recordset
Dim rsJ As ADODB.Recordset
Dim noOrd As String

    If UCase(gridStall.Splits(1).Columns(coll)) = "X" Then
        Set rsHlp = objProgress.searchHelpProgres(Trim(baris), jamS)
        If Not rsHlp.EOF Then
            If Left(hitungWaktu(rsHlp("RATE"), rsHlp("JAMMULAI"), rsHlp("TANGGAL")), 5) >= jamS Then
                Txt = Trim(rsHlp("KETERANGAN"))
                noOrd = Left(Mid(Txt, InStr(Txt, "|") + 1, Len(Txt) - InStr(Txt, "|")), InStr(Mid(Txt, InStr(Txt, "|") + 1, Len(Txt) - InStr(Txt, "|")), "|") - 1)
                
                ' CARI DI JOBDTA
                Dim rsUt As ADODB.Recordset
                Set rsUt = objProgress.takeDetailJob(noOrd, 2)
                If Not rsUt.EOF Then
                    Text(0) = rsUt("NO_CHASSIS")
                    Text(1) = rsUt("NO_POLISI")
                    Text(2) = IIf(Trim(rsHlp("MECH")) <> "", Trim(rsHlp("MECH")), Trim(rsUt("MECH")))
                    Text(3) = CDbl(rsUt("ACTRATE1")) + CDbl(rsUt("ACTRATE2")) + CDbl(rsUt("ACTRATE3")) + CDbl(rsUt("ACTRATE4")) + CDbl(rsUt("ACTRATE5")) + CDbl(rsUt("ACTRATE6")) + CDbl(rsUt("ACTRATE7")) + CDbl(rsUt("ACTRATE8"))
                    Text(4) = rsHlp("JAMMULAI")
                    Text(5) = Left(hitungWaktu(rsHlp("RATE"), rsHlp("JAMMULAI"), rsHlp("TANGGAL")), 5)
                    Text(10) = rsUt("MDL")
                    Text(11) = rsUt("FMDL")
                    Text(12) = rsUt("COLOR")
                    Label10.Caption = "No. WO"
                    Text(13) = noOrd
                    PRC = IIf(rsUt("FS_RC") <> "", True, False)
                    
                    If rsJob.State <> adStateClosed Then
                        rsJob.Close
                        Set rsJ = Nothing
                    End If
                    For i = 1 To 7
                        tmp2 = "KLPKERJA" & i
                        tmp = "REP" & i
                        tmp1 = "RTE" & i
                        If rsUt(tmp) <> "" Then
                            Set rsJ = objProgress.takeDeffJob(Trim(rsUt(tmp2)), Trim(rsUt(tmp)), Trim(rsUt("MDL")))
                            If Not rsJ.EOF Then
                                Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
                                With rsJob
                                    .AddNew
                                    !Group = rsJ("JOB")
                                    !Job = rsJ("FS_ITEM")
                                    !ket = rsJ("KETERANGAN")
                                    !Rate = rsUt(tmp1)
                                    .Update
                                End With
                            End If
                        End If
                    Next
                End If
                rsUt.Close
                
                ' CARI DI ADDITIONAL JOB
                Set rsUt = objO.FetchData("SELECT * FROM JOBADD2 WHERE NO_WO = '" & noOrd & "'")
                If Not rsUt.EOF Then
                    Do While Not rsUt.EOF
                        With rsJob
                            .AddNew
                            !Group = Trim(rsUt("KLPKERJA"))
                            !Job = Trim(rsUt("REP"))
                            !ket = ""
                            !Rate = CDbl(Trim(rsUt("RTE")))
                            .Update
                        End With
                        Text(3) = Text(3) + CDbl(rsUt("ACTRTE"))
                        rsUt.MoveNext
                    Loop
                End If
                rsUt.Close
                
                ' CARI DI SUBLET
                Set rsUt = objO.FetchData("SELECT * FROM JOBSUBLET WHERE NO_WO = '" & noOrd & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "'")
                If Not rsUt.EOF Then
                    Do While Not rsUt.EOF
                        With rsJob
                            .AddNew
                            !Group = "SUB"
                            !Job = Trim(rsUt("NAMABARANG"))
                            !ket = Trim(rsUt("KODESATUAN"))
                            !Rate = 0
                            .Update
                        End With
                        rsUt.MoveNext
                    Loop
                End If
                rsUt.Close
                
                Set gridJobOrder.DataSource = rsJob
                Call headColumns
            End If
        End If
        rsHlp.Close
    End If
End Sub

Private Sub lblStatus_Change()
    If Trim(lblStatus.Caption) = "DITUNGGU" Then
        lblStatus.ForeColor = vbRed
    Else
        lblStatus.ForeColor = vbBlue
    End If
End Sub

Private Sub searchClose_Click()
    framSearch.Visible = False
End Sub

Private Sub StartInsp_Click()

    If Not cekValidasi(Trim(Text(13))) Then
        Call objProgress.startInspection(Trim(Text(13)), Trim(Text(2)), stl, "", Trim(Text(0)), nTgl, IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), CDbl(Text(3)), Trim(Text(5)))
        frmeStall.Visible = False
        
        ' REFRESH GRID
        Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
        gridStall.Refresh
    Else
        MsgBox "Sudah divalidasi SA, tidak dapat diproses lagi"
        frmeStall.Visible = False
    End If
End Sub

Private Sub StartWash_Click()
    
    If Not cekValidasi(Trim(Text(13))) Then
        Call objProgress.startWashing(Trim(Text(13)), Trim(Text(2)), stl, "WASHING", Trim(Text(0)), nTgl, IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), CDbl(Text(3)), Trim(Text(5)))
        frmeStall.Visible = False
        
        ' REFRESH GRID
        Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
        gridStall.Refresh
    Else
        MsgBox "Sudah divalidasi SA, tidak dapat diproses lagi"
        frmeStall.Visible = False
    End If
End Sub

Private Sub StopInsp_Click()

    Call objProgress.stopInspection(Trim(Text(13)), Trim(Text(2)), stl, "SELESAI FINAL INSPECTION", Trim(Text(0)), nTgl, IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), CDbl(Text(3)), Trim(Text(5)))
    frmeStall.Visible = False
    
    ' REFRESH GRID
    Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
    gridStall.Refresh
End Sub

Private Sub tblOther_Click()
    gTag = "STALL"
    frmHistory.Text(0) = Text(0)
    frmHistory.Text(1) = cariNamaCustomer(Text(0))
    frmHistory.Text(2) = Replace(Replace(Text(1), ">", ""), " ", "")
    frmHistory.Show (vbModal)
End Sub

Private Function cariNamaCustomer(noRangka As String) As String
    cariNamaCustomer = objO.FetchData_String("SELECT CUSTOMER FROM CUSTDAT WHERE NO_CHASSIS = '" & Trim(noRangka) & "'")
End Function

Private Sub tblTutup_Click()
    frameHistory.Visible = False
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    Select Case Index
    End Select
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index) = gMainKeyCari
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
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 14
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ORDER"
                'LIHAT LAGI
                gBantuan2 = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
                gBantuan3 = Format(isnow, "DD/MM/YYYY")
                gBantuan4 = ""
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
Dim rsCari As ADODB.Recordset

    Select Case Index
        Case 9
            If Text(Index) <> "" And Len(Trim(Text(Index))) = 7 Then
                cmdPause.SetFocus
            End If
        Case 14
            Dim rsHlp As ADODB.Recordset
            Dim sisaRate As Integer
            Dim totalRate As Integer
            If Text(Index) <> "" Then
                If CDbl(Text(3)) > 0 Then
                    'LIHAT LAGI
                    Set rsCari = objO.FetchData("SELECT isnull(MENITKERJA,0) AS MENITKERJA, ISNULL(TERPAKAI,0) AS TERPAKAI, LEVMECH FROM mekanik A LEFT JOIN T_MEKANIK B ON A.MECH = B.MECH AND B.HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' AND B.TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' WHERE A.mech = '" & Text(Index) & "'")
                    If rsCari.EOF Then
                        gMainKeyCari = Text(Index)
                        gSearchOption = "MEKANIK"
                        gBantuan = "ORDER"
                        'LIHAT LAGI
                        gBantuan2 = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
                        gBantuan3 = Format(isnow, "DD/MM/YYYY")
                        gBantuan4 = ""
                        frmCari.Show (vbModal)
                    Else
                        Dim nSelisih As Integer
                        Set rsHlp = objProgress.countDifference(Trim(Text(13)))
                        If Not rsHlp.EOF Then
                            If Not rsHlp("SELISIH") = Null Then
                                nSelisih = CDbl(rsHlp("SELISIH"))
                            Else
                                nSelisih = 0
                            End If
                            totalRate = CDbl(Text(3)) * 60
                            sisaRate = totalRate - CDbl(nSelisih)
                        End If
                         
                        Text(15) = rsCari("MENITKERJA") - rsCari("TERPAKAI")
                    End If
                End If
            End If
    End Select
End Sub

Private Sub Timer1_Timer()
    If Not batalLoad Then
        Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
        gridStall.Refresh
    End If
End Sub

Private Sub toClose_Click()
    frmeStall.Visible = False
End Sub

Private Sub toDown_Click()
    frmeStall.Top = frmeStall.Top + 2000
End Sub

Private Sub toFinish_Click()
    
    vPass = IIf(opPause(0).value, opPause(0).Caption, IIf(opPause(1).value, opPause(1).Caption, IIf(opPause(2).value, opPause(2).Caption, opPause(3).Caption)))
    
    If chEM = Unchecked Then
        Call objProgress.finishProses(Trim(Text(13)), Trim(Text(2)), stl, Trim(vPass) & "|" & Trim(Text(9)), Trim(Text(0)), nTgl, IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), CDbl(Text(3)), PRC, Trim(Text(5))) ', Format(isnow, "HH:NN"), CDbl(Text(3)), PRC) '
    Else
        Call objProgress.finishProses2(Trim(Text(13)), stl, Trim(vPass) & "|" & Trim(Text(9)), Trim(Text(0)), nTgl, IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), CDbl(Text(3)), PRC, Trim(Text(5))) ', Format(isnow, "HH:NN"), CDbl(Text(3)), PRC) '
    End If
    frmeStall.Visible = False
    toFinish.Visible = False
    toStart.Visible = False
    toStart.FormatString = " Start"
    toMove.FormatString = " Pindah Stall"
    toMove.Visible = False
    
    ' REFRESH GRID
    Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
    gridStall.Refresh
End Sub

Private Sub toLeft_Click()
    frmeStall.Left = frmeStall.Left - 2000
End Sub

Private Sub toMove_Click()
    If Trim(toMove.FormatString) = "Pindah Stall" Then
        MsgBox "Pilih Stall dan Jam yang akan digunakan !"
        Me.Tag = "Move"
        gantiMech.Visible = False
    ElseIf toMove.FormatString = "Pakai Stall" Then
        
        If Left(Text(5), 2) < njamAk Then
            Call objProgress.moveProses(Trim(Text(13)), Trim(Text(2)), stl, "", Trim(Text(0)), nTgl, Trim(Text(4)), CDbl(Text(3)), Trim(Text(5)))
            Me.Tag = ""
            
            ' REFRESH GRID
            Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
            gridStall.Refresh
        Else
            MsgBox "Waktu Stall tidak Mencukupi !! "
            Exit Sub
        End If
    End If
    frmeStall.Visible = False
End Sub

Private Sub toRight_Click()
    frmeStall.Left = frmeStall.Left + 2000
End Sub

Private Sub toStart_Click()
Dim nStart1 As Boolean
    
    If toStart.FormatString = " Start" Then
        If Not cekValidasi(Trim(Text(13))) Then
            If chEM = Unchecked Then
                nStart1 = False
                If objProgress.searchTempProgress(Trim(Text(13))) Then
                    nStart1 = True
                End If
            
                Call objProgress.startProses(Trim(Text(13)), Trim(Text(2)), stl, "", Trim(Text(0)), nTgl, IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1), Trim(Text(4)), CDbl(Text(3)), nStart1, Trim(Text(5)))
                toStart.FormatString = "Pause"
                frmeStall.Visible = False
                toFinish.Left = 1200
                toFinish.Visible = True
                
                ' REFRESH GRID
                Call loadStall(Format(nTgl, "MM/DD/YYYY"), njamAwl, njamAk)
                gridStall.Refresh
            Else
                gantiMech_Click
            End If
        Else
            MsgBox "Sudah divalidasi SA, tidak dapat diproses lagi"
            frmeStall.Visible = False
        End If
    ElseIf toStart.FormatString = "Pause" Then
        Text(9).Enabled = True
        Text(9) = ""
        frmePause.Visible = True
        frmeMekanik.Visible = False
        frmeStall.Height = frmeStall.Height + frmePause.Height + 300
        toStart.FormatString = " Start"
        opPause(3).value = True
        toStart.Visible = False
        toFinish.Visible = False
    End If
End Sub

Private Sub toUp_Click()
    frmeStall.Top = frmeStall.Top - 2000
End Sub
