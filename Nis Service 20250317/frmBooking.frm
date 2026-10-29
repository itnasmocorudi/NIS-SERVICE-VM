VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MsComCtl.ocx"
Begin VB.Form frmBooking 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Booking"
   ClientHeight    =   12585
   ClientLeft      =   1965
   ClientTop       =   1815
   ClientWidth     =   17145
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   12585
   ScaleWidth      =   17145
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   147
      Top             =   435
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   142
      Top             =   435
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   13320
      TabIndex        =   143
      Top             =   120
      Visible         =   0   'False
      Width           =   7095
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame frameHasilDSS 
      BackColor       =   &H000080FF&
      Height          =   9855
      Left            =   480
      TabIndex        =   208
      Top             =   10605
      Visible         =   0   'False
      Width           =   12495
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF00FF&
         Caption         =   "BATAL"
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
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   217
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton cmdPilihDSS 
         BackColor       =   &H0000FFFF&
         Caption         =   "PILIH"
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
         Left            =   10320
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   150
         Width           =   975
      End
      Begin TrueOleDBGrid70.TDBGrid gridDSS 
         Height          =   3000
         Left            =   120
         TabIndex        =   209
         Top             =   600
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   5292
         _LayoutType     =   4
         _RowHeight      =   17
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
         Columns(4).NumberFormat=   "Short Time"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).DataField=   ""
         Columns(5).NumberFormat=   "Short Time"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=66,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=63,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=64,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=65,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
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
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&HF3F3C5&,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(75)  =   ":id=39,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=39,.fontname=MS Sans Serif"
         _StyleDefs(77)  =   "Named:id=40:OddRow"
         _StyleDefs(78)  =   ":id=40,.parent=33,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0"
         _StyleDefs(79)  =   ":id=40,.charset=0"
         _StyleDefs(80)  =   ":id=40,.fontname=MS Sans Serif"
         _StyleDefs(81)  =   "Named:id=41:RecordSelector"
         _StyleDefs(82)  =   ":id=41,.parent=34"
         _StyleDefs(83)  =   "Named:id=42:FilterBar"
         _StyleDefs(84)  =   ":id=42,.parent=33"
      End
      Begin VB.Label Label65 
         BackColor       =   &H000080FF&
         BackStyle       =   0  'Transparent
         Caption         =   "HASIL PENCARIAN DSS "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   210
         Top             =   120
         Width           =   8010
      End
   End
   Begin VB.Frame frameDSS 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7800
      TabIndex        =   202
      Top             =   6360
      Visible         =   0   'False
      Width           =   3855
      Begin VB.TextBox txtNRate 
         Height          =   285
         Left            =   360
         TabIndex        =   203
         Text            =   "Text2"
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSComctlLib.ProgressBar progDSS 
         Height          =   735
         Left            =   -15
         TabIndex        =   204
         Top             =   75
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   1296
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   315
         Index           =   122
         Left            =   2040
         TabIndex        =   205
         Top             =   720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99:99"
         PromptChar      =   ">"
      End
      Begin VB.CommandButton cmdBatalAdjustDMS 
         BackColor       =   &H000000FF&
         Caption         =   "BATAL"
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
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdSimpanAdjustDSS 
         BackColor       =   &H0000FF00&
         Caption         =   "SIMPAN"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label64 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "ADJUST TOTAL ACTUAL RATE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   207
         Top             =   0
         Width           =   4095
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Adjust Pengerjaan"
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
         Left            =   480
         TabIndex        =   206
         Top             =   795
         Width           =   1440
      End
   End
   Begin VB.Frame frmEstimasiSublet 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   7440
      TabIndex        =   197
      Top             =   7320
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdStalOk 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   960
         Width           =   1815
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   91
         Left            =   3000
         TabIndex        =   46
         Top             =   550
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "ESTIMASI PENGERJAAN SUBLET"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   480
         TabIndex        =   200
         Top             =   0
         Width           =   3915
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Stall"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1320
         TabIndex        =   199
         Top             =   30
         Width           =   2040
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1200
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1200
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         X1              =   3600
         X2              =   4680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000000FF&
         X1              =   3600
         X2              =   4680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "ESTIMASI PENGERJAAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   198
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   0
      Picture         =   "frmBooking.frx":0000
      ScaleHeight     =   10695
      ScaleWidth      =   495
      TabIndex        =   186
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton cmdBersih 
      BackColor       =   &H0000FF00&
      Caption         =   "Bersih"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   152
      Top             =   75
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Frame frmeWaktu 
      BackColor       =   &H00C0FFFF&
      Height          =   3375
      Left            =   12960
      TabIndex        =   97
      Top             =   3360
      Visible         =   0   'False
      Width           =   6135
      Begin TrueOleDBGrid70.TDBGrid GridWaktu1 
         Height          =   3855
         Left            =   8640
         TabIndex        =   98
         Top             =   1440
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6800
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
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
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
         AllowAddNew     =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
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
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
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
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
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
      Begin MSFlexGridLib.MSFlexGrid GridWaktu 
         Height          =   2625
         Left            =   120
         TabIndex        =   99
         Top             =   240
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   4630
         _Version        =   393216
         Cols            =   6
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hari H : Tanggal Customer Walk In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   120
         TabIndex        =   104
         Top             =   2880
         Width           =   4425
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   1
      Left            =   7515
      TabIndex        =   77
      Top             =   30
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   714
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   3120
      TabIndex        =   79
      Top             =   60
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   60
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "??"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   18
      Left            =   10200
      TabIndex        =   91
      Top             =   690
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
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
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   19
      Left            =   10200
      TabIndex        =   92
      Top             =   1020
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   56
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   57
      Left            =   2640
      TabIndex        =   2
      Top             =   60
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   16777215
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "CCC"
      PromptChar      =   ">"
   End
   Begin VB.CommandButton cmdWaktu 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Waktu Terpakai"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   360
      Width           =   1890
   End
   Begin VB.CommandButton cmdHO 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Check &HO"
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
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   95
      Top             =   120
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Data Pelanggan"
      Height          =   2895
      Left            =   360
      TabIndex        =   64
      Top             =   480
      Width           =   8565
      Begin VB.CommandButton cmdPlgn 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Pelanggan"
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
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   153
         Top             =   465
         Width           =   1050
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   915
         TabIndex        =   4
         Top             =   735
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   915
         TabIndex        =   5
         Top             =   1035
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   915
         TabIndex        =   6
         Top             =   1335
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   915
         TabIndex        =   8
         Top             =   1935
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   5430
         TabIndex        =   60
         Top             =   2520
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   7245
         TabIndex        =   61
         Top             =   2535
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   5430
         TabIndex        =   10
         Top             =   435
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   13
         Left            =   5430
         TabIndex        =   12
         Top             =   1320
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   5430
         TabIndex        =   63
         Top             =   1035
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   50
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   11
         Left            =   7695
         TabIndex        =   62
         Top             =   735
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   7560
         TabIndex        =   74
         Top             =   750
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   5430
         TabIndex        =   75
         Top             =   2235
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   915
         TabIndex        =   3
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   5430
         TabIndex        =   59
         Top             =   1935
         Width           =   3060
         _ExtentX        =   5398
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
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   915
         TabIndex        =   9
         Top             =   2235
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   915
         TabIndex        =   7
         Top             =   1635
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   41
         Left            =   5430
         TabIndex        =   100
         Top             =   1635
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   43
         Left            =   915
         TabIndex        =   58
         Top             =   2535
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   44
         Left            =   2640
         TabIndex        =   103
         Top             =   450
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   10
         Left            =   5430
         TabIndex        =   11
         Top             =   735
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   71
         Left            =   2865
         TabIndex        =   133
         Top             =   2520
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
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
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   8400
         TabIndex        =   141
         Top             =   1320
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   270
         Left            =   2500
         TabIndex        =   151
         Top             =   457
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Label lblModel 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   226
         Top             =   400
         Width           =   2175
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   6240
         X2              =   8400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   6240
         X2              =   8400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1920
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Customer && Kendaraan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   2280
         TabIndex        =   146
         Top             =   45
         Width           =   3675
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   8535
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOU"
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
         Left            =   450
         TabIndex        =   102
         Top             =   2580
         Width           =   345
      End
      Begin VB.Label Label22 
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
         Left            =   4440
         TabIndex        =   101
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   4695
         TabIndex        =   82
         Top             =   1950
         Width           =   585
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN No."
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
         Left            =   -60
         TabIndex        =   81
         Top             =   1650
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA No."
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
         Left            =   -60
         TabIndex        =   80
         Top             =   2250
         Width           =   945
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Lgn"
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
         Left            =   -165
         TabIndex        =   76
         Top             =   495
         Width           =   1065
      End
      Begin VB.Label Label19 
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
         Left            =   4800
         TabIndex        =   73
         Top             =   1380
         Width           =   480
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   4860
         TabIndex        =   72
         Top             =   780
         Width           =   420
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Th. Mobil"
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
         Left            =   4635
         TabIndex        =   71
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DDTE."
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
         Left            =   6720
         TabIndex        =   70
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOTAM"
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
         Left            =   4740
         TabIndex        =   69
         Top             =   2550
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GSM No."
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
         Left            =   -60
         TabIndex        =   68
         Top             =   1950
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
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
         Left            =   -60
         TabIndex        =   67
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Left            =   -60
         TabIndex        =   66
         Top             =   1050
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Left            =   -60
         TabIndex        =   65
         Top             =   750
         Width           =   945
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "History Service"
      Height          =   1935
      Left            =   9000
      TabIndex        =   83
      Top             =   1440
      Width           =   3735
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   22
         Left            =   1080
         TabIndex        =   86
         Top             =   990
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   21
         Left            =   1080
         TabIndex        =   88
         Top             =   705
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   30
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   20
         Left            =   1080
         TabIndex        =   84
         Top             =   435
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin VB.CommandButton cmdHistory 
         BackColor       =   &H00DCFCFB&
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
         Left            =   -60
         Picture         =   "frmBooking.frx":2049
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   1300
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00DCFCFB&
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
         Left            =   1800
         Picture         =   "frmBooking.frx":2806
         Style           =   1  'Graphical
         TabIndex        =   191
         Top             =   1300
         Width           =   1935
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On Line HO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   480
         TabIndex        =   193
         Top             =   1680
         Width           =   1050
      End
      Begin VB.Label Label58 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "On LOCAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   240
         Left            =   2400
         TabIndex        =   192
         Top             =   1680
         Width           =   930
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   3120
         X2              =   3600
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   3120
         X2              =   3600
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   600
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   600
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   1200
         TabIndex        =   145
         Top             =   45
         Width           =   1620
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   360
         TabIndex        =   89
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM Akhir"
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
         Left            =   360
         TabIndex        =   87
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Left            =   360
         TabIndex        =   85
         Top             =   480
         Width           =   570
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6240
      LargeChange     =   350
      Left            =   12480
      Max             =   5250
      SmallChange     =   350
      TabIndex        =   105
      Top             =   3480
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   6255
      Left            =   480
      ScaleHeight     =   6195
      ScaleWidth      =   12435
      TabIndex        =   106
      Top             =   3480
      Width           =   12495
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Input Booking"
         Height          =   11175
         Left            =   0
         TabIndex        =   107
         Top             =   -100
         Width           =   11955
         Begin VB.Frame frameDIDX 
            BackColor       =   &H80000004&
            Caption         =   "Frame13"
            Height          =   2895
            Left            =   5640
            TabIndex        =   232
            Top             =   0
            Width           =   6255
            Begin VB.TextBox Text2 
               Height          =   375
               Index           =   1
               Left            =   3840
               MaxLength       =   20
               TabIndex        =   263
               Top             =   1440
               Width           =   2295
            End
            Begin VB.TextBox Text2 
               Height          =   375
               Index           =   0
               Left            =   3840
               MaxLength       =   100
               TabIndex        =   262
               Top             =   720
               Width           =   2295
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   63
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   252
               Top             =   720
               Width           =   2650
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   64
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   251
               Top             =   1050
               Width           =   2650
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   65
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   250
               Top             =   1380
               Width           =   2650
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               IMEMode         =   3  'DISABLE
               Index           =   66
               Left            =   1080
               MaxLength       =   100
               TabIndex        =   249
               Top             =   1710
               Width           =   2650
            End
            Begin VB.TextBox txt 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   320
               Index           =   11
               Left            =   1080
               MaxLength       =   6
               TabIndex        =   236
               Top             =   360
               Width           =   1350
            End
            Begin VB.TextBox txt 
               Appearance      =   0  'Flat
               BackColor       =   &H0080FFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   795
               IMEMode         =   3  'DISABLE
               Index           =   8
               Left            =   1080
               MaxLength       =   150
               MultiLine       =   -1  'True
               TabIndex        =   235
               Top             =   2050
               Width           =   2745
            End
            Begin VB.Frame Frame13 
               BackColor       =   &H8000000D&
               BorderStyle     =   0  'None
               Caption         =   "Frame13"
               Height          =   350
               Left            =   0
               TabIndex        =   233
               Top             =   0
               Width           =   6255
               Begin VB.Label Label67 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ".:: Pickup Delivery ::."
                  BeginProperty Font 
                     Name            =   "Tahoma"
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
                  Left            =   120
                  TabIndex        =   266
                  Top             =   120
                  Width           =   2010
               End
               Begin VB.Label Label67 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   ".:: Decision Maker ::."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFFFFF&
                  Height          =   240
                  Index           =   0
                  Left            =   4200
                  TabIndex        =   234
                  Top             =   120
                  Width           =   2010
               End
            End
            Begin MSMask.MaskEdBox MaskEdBox17 
               Height          =   270
               Index           =   0
               Left            =   2400
               TabIndex        =   253
               Top             =   360
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   87
               Left            =   975
               TabIndex        =   254
               Top             =   0
               Width           =   3540
               _ExtentX        =   6244
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   50
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox MaskEdBox18 
               Height          =   270
               Left            =   3720
               TabIndex        =   258
               Top             =   720
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox19 
               Height          =   270
               Left            =   3720
               TabIndex        =   259
               Top             =   1050
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox20 
               Height          =   270
               Left            =   3720
               TabIndex        =   260
               Top             =   1380
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MaskEdBox17 
               Height          =   270
               Index           =   1
               Left            =   3720
               TabIndex        =   261
               Top             =   1710
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kompleks; No"
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
               Index           =   5
               Left            =   0
               TabIndex        =   265
               Top             =   2400
               Width           =   960
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Phone Number"
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
               Left            =   3840
               TabIndex        =   257
               Top             =   1200
               Width           =   1050
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Nama Decision"
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
               Left            =   3840
               TabIndex        =   256
               Top             =   480
               Width           =   1035
            End
            Begin VB.Label Label75 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nama"
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
               Left            =   0
               TabIndex        =   255
               Top             =   15
               Width           =   945
            End
            Begin VB.Label Label13 
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
               Index           =   9
               Left            =   1020
               TabIndex        =   248
               Top             =   750
               Width           =   60
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Provinsi"
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
               Index           =   13
               Left            =   120
               TabIndex        =   247
               Top             =   765
               Width           =   555
            End
            Begin VB.Label Label13 
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
               Left            =   1020
               TabIndex        =   246
               Top             =   1065
               Width           =   60
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kab/Kota"
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
               Index           =   14
               Left            =   120
               TabIndex        =   245
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label Label13 
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
               Left            =   1020
               TabIndex        =   244
               Top             =   1410
               Width           =   60
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kecamatan"
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
               Index           =   18
               Left            =   120
               TabIndex        =   243
               Top             =   1380
               Width           =   795
            End
            Begin VB.Label Label13 
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
               Index           =   12
               Left            =   1020
               TabIndex        =   242
               Top             =   1710
               Width           =   60
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kelurahan"
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
               Index           =   19
               Left            =   120
               TabIndex        =   241
               Top             =   1710
               Width           =   720
            End
            Begin VB.Label Label13 
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
               Index           =   13
               Left            =   1020
               TabIndex        =   240
               Top             =   2280
               Width           =   60
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jalan / "
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
               Index           =   21
               Left            =   0
               TabIndex        =   239
               Top             =   2115
               Width           =   525
            End
            Begin VB.Label Label13 
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
               Index           =   18
               Left            =   1020
               TabIndex        =   238
               Top             =   450
               Width           =   60
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kode Pos"
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
               Index           =   24
               Left            =   120
               TabIndex        =   237
               Top             =   405
               Width           =   660
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "                                                                                                                                "
            Height          =   2535
            Left            =   5640
            TabIndex        =   267
            Top             =   360
            Width           =   6015
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   46
               Left            =   3765
               TabIndex        =   268
               Top             =   240
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox MaskEdBox4 
               Height          =   270
               Left            =   3645
               TabIndex        =   269
               Top             =   360
               Visible         =   0   'False
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   47
               Left            =   3765
               TabIndex        =   270
               Top             =   600
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox MaskEdBox7 
               Height          =   270
               Left            =   3645
               TabIndex        =   271
               Top             =   720
               Visible         =   0   'False
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   48
               Left            =   3765
               TabIndex        =   272
               Top             =   960
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox MaskEdBox8 
               Height          =   270
               Left            =   3645
               TabIndex        =   273
               Top             =   1080
               Visible         =   0   'False
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   49
               Left            =   3765
               TabIndex        =   274
               Top             =   1320
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox MaskEdBox9 
               Height          =   270
               Left            =   3720
               TabIndex        =   275
               Top             =   1440
               Visible         =   0   'False
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   50
               Left            =   3765
               TabIndex        =   276
               Top             =   1680
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox MaskEdBox10 
               Height          =   270
               Left            =   3645
               TabIndex        =   277
               Top             =   1800
               Visible         =   0   'False
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   51
               Left            =   3765
               TabIndex        =   278
               Top             =   2040
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox MaskEdBox11 
               Height          =   270
               Left            =   3645
               TabIndex        =   279
               Top             =   2160
               Visible         =   0   'False
               Width           =   75
               _ExtentX        =   132
               _ExtentY        =   476
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   54
               Left            =   3840
               TabIndex        =   280
               Top             =   240
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   65
               Left            =   3840
               TabIndex        =   281
               Top             =   600
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   66
               Left            =   3840
               TabIndex        =   282
               Top             =   960
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   67
               Left            =   3840
               TabIndex        =   283
               Top             =   1320
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   68
               Left            =   3840
               TabIndex        =   284
               Top             =   1680
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   69
               Left            =   3840
               TabIndex        =   285
               Top             =   2040
               Visible         =   0   'False
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin VB.Label Label44 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Waktu"
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
               Left            =   840
               TabIndex        =   297
               Top             =   405
               Width           =   465
            End
            Begin VB.Label Label43 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frekuensi"
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
               Left            =   600
               TabIndex        =   296
               Top             =   765
               Width           =   690
            End
            Begin VB.Label Label42 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Lokasi"
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
               Left            =   840
               TabIndex        =   295
               Top             =   1125
               Width           =   435
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Lampu Peringatan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   -500
               TabIndex        =   294
               Top             =   1485
               Width           =   1890
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
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
               Left            =   720
               TabIndex        =   293
               Top             =   1920
               Width           =   555
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   " (Akhir-akhir ini, Seminggu lalu)"
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
               Left            =   1440
               TabIndex        =   292
               Top             =   405
               Width           =   2205
            End
            Begin VB.Label Label69 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "(Selalu, Kadang-kadang)"
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
               Left            =   1440
               TabIndex        =   291
               Top             =   765
               Width           =   1770
            End
            Begin VB.Label Label70 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "(Jalan datar, Jalan Tol, Tanjakan)"
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
               Left            =   1440
               TabIndex        =   290
               Top             =   1125
               Width           =   2430
            End
            Begin VB.Label Label71 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   " (Lampu menyala, Kedip sekali)"
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
               Left            =   1440
               TabIndex        =   289
               Top             =   1485
               Width           =   2205
            End
            Begin VB.Label Label72 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "(Starter, Pemanasan, Konstan)"
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
               Left            =   1440
               TabIndex        =   288
               Top             =   1920
               Width           =   2250
            End
            Begin VB.Label Label73 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "(Berhenti, Mesin dingin, Panas)"
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
               Left            =   1440
               TabIndex        =   287
               Top             =   2205
               Width           =   2235
            End
            Begin VB.Label Label57 
               BackColor       =   &H00C0FFFF&
               Caption         =   "-------------------------------------------------Kondisi ketika terjadi Keluhan ---------------------------------"
               Height          =   255
               Left            =   0
               TabIndex        =   286
               Top             =   120
               Width           =   5895
            End
         End
         Begin VB.CommandButton TblKeluar 
            BackColor       =   &H000000FF&
            Caption         =   "KELUAR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   10800
            MaskColor       =   &H00FF8080&
            Style           =   1  'Graphical
            TabIndex        =   220
            Top             =   10440
            Width           =   1095
         End
         Begin VB.CommandButton cmdcetak 
            BackColor       =   &H00FF8080&
            Caption         =   "CETAK"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9000
            MaskColor       =   &H00FF8080&
            Style           =   1  'Graphical
            TabIndex        =   219
            Top             =   10440
            Width           =   1095
         End
         Begin VB.CommandButton tblSimpan 
            BackColor       =   &H0080FF80&
            Caption         =   "SIMPAN"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7800
            MaskColor       =   &H0080FF80&
            Style           =   1  'Graphical
            TabIndex        =   218
            Top             =   10440
            Width           =   1095
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   255
            Left            =   5760
            TabIndex        =   189
            Top             =   120
            Width           =   6015
            Begin VB.TextBox txtKodeProgramApp 
               Enabled         =   0   'False
               Height          =   285
               Left            =   720
               TabIndex        =   228
               Top             =   0
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.Label JualNRM 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   75
               TabIndex        =   190
               Top             =   0
               Width           =   5865
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   495
            Left            =   7800
            TabIndex        =   187
            Top             =   10440
            Width           =   4095
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   2160
            ScaleHeight     =   855
            ScaleWidth      =   3090
            TabIndex        =   180
            Top             =   10440
            Width           =   3090
            Begin TDBNumber6Ctl.TDBNumber txtRp 
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   181
               Top             =   30
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   556
               Calculator      =   "frmBooking.frx":2FC3
               Caption         =   "frmBooking.frx":2FE3
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmBooking.frx":304F
               Keys            =   "frmBooking.frx":306D
               Spin            =   "frmBooking.frx":30B7
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   12640511
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   ","
               DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###,##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   -99999999999
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
            Begin TDBNumber6Ctl.TDBNumber txtRp 
               Height          =   315
               Index           =   1
               Left            =   1200
               TabIndex        =   182
               Top             =   360
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   556
               Calculator      =   "frmBooking.frx":30DF
               Caption         =   "frmBooking.frx":30FF
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmBooking.frx":316B
               Keys            =   "frmBooking.frx":3189
               Spin            =   "frmBooking.frx":31D3
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   12640511
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   ","
               DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###,##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   -99999999999
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
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Current AR"
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
               Left            =   0
               TabIndex        =   184
               Top             =   45
               Width           =   1050
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Platform AR"
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
               Left            =   0
               TabIndex        =   183
               Top             =   390
               Width           =   1155
            End
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   4215
            Left            =   120
            TabIndex        =   155
            Top             =   6120
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   7435
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   617
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Job Order && Sublet"
            TabPicture(0)   =   "frmBooking.frx":31FB
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame5"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdStall"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Part Order"
            TabPicture(1)   =   "frmBooking.frx":3217
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame10"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame10 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   3735
               Left            =   -74970
               TabIndex        =   167
               Top             =   360
               Width           =   11655
               Begin MSFlexGridLib.MSFlexGrid PartGrid 
                  Height          =   3105
                  Left            =   120
                  TabIndex        =   52
                  Top             =   600
                  Width           =   7665
                  _ExtentX        =   13520
                  _ExtentY        =   5477
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   35
                  Left            =   120
                  TabIndex        =   50
                  Top             =   240
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   503
                  _Version        =   393216
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
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   36
                  Left            =   3165
                  TabIndex        =   51
                  Top             =   240
                  Width           =   450
                  _ExtentX        =   794
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "99"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox MaskEdBox5 
                  Height          =   270
                  Left            =   3045
                  TabIndex        =   168
                  Top             =   240
                  Width           =   75
                  _ExtentX        =   132
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   70
                  Left            =   3240
                  TabIndex        =   169
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   690
                  _ExtentX        =   1217
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
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Order Part"
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
                  TabIndex        =   170
                  Top             =   0
                  Width           =   765
               End
            End
            Begin VB.CommandButton cmdStall 
               BackColor       =   &H00DCFCFB&
               Caption         =   "S&tall"
               Height          =   615
               Left            =   10680
               Style           =   1  'Graphical
               TabIndex        =   49
               Top             =   360
               Width           =   930
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   3735
               Left            =   30
               TabIndex        =   156
               Top             =   360
               Width           =   11655
               Begin VB.ComboBox cbJobType 
                  Height          =   315
                  ItemData        =   "frmBooking.frx":3233
                  Left            =   1680
                  List            =   "frmBooking.frx":3243
                  Style           =   2  'Dropdown List
                  TabIndex        =   300
                  Top             =   120
                  Width           =   1935
               End
               Begin VB.CommandButton cmdLihatProgram 
                  BackColor       =   &H0080FF80&
                  Caption         =   "Lihat Program"
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
                  Style           =   1  'Graphical
                  TabIndex        =   215
                  Top             =   120
                  Width           =   1455
               End
               Begin VB.CheckBox chSublet 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Part"
                  Height          =   255
                  Left            =   11040
                  TabIndex        =   44
                  Top             =   720
                  Width           =   615
               End
               Begin MSFlexGridLib.MSFlexGrid DGrid 
                  Height          =   2625
                  Left            =   120
                  TabIndex        =   38
                  Top             =   1080
                  Width           =   5505
                  _ExtentX        =   9710
                  _ExtentY        =   4630
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  ScrollBars      =   1
                  SelectionMode   =   1
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   34
                  Left            =   900
                  TabIndex        =   36
                  Top             =   720
                  Width           =   3135
                  _ExtentX        =   5530
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   50
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox MaskEdBox3 
                  Height          =   270
                  Left            =   4050
                  TabIndex        =   157
                  Top             =   720
                  Width           =   75
                  _ExtentX        =   132
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   32
                  Left            =   120
                  TabIndex        =   35
                  Top             =   720
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   3
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox MaskEdBox6 
                  Height          =   270
                  Left            =   750
                  TabIndex        =   158
                  Top             =   720
                  Width           =   75
                  _ExtentX        =   132
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   42
                  Left            =   3240
                  TabIndex        =   159
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
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
                  Mask            =   "99,99"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   64
                  Left            =   4200
                  TabIndex        =   37
                  Top             =   720
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   503
                  _Version        =   393216
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   3
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">CCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox MaskEdBox12 
                  Height          =   270
                  Left            =   4860
                  TabIndex        =   160
                  Top             =   720
                  Width           =   75
                  _ExtentX        =   132
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   33
                  Left            =   1920
                  TabIndex        =   164
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   2895
                  _ExtentX        =   5106
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
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSFlexGridLib.MSFlexGrid subletGrid 
                  Height          =   2625
                  Left            =   5640
                  TabIndex        =   45
                  Top             =   1080
                  Width           =   5985
                  _ExtentX        =   10557
                  _ExtentY        =   4630
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   79
                  Left            =   5640
                  TabIndex        =   40
                  Top             =   720
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   503
                  _Version        =   393216
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
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   80
                  Left            =   8640
                  TabIndex        =   41
                  Top             =   720
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "99"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox MaskEdBox15 
                  Height          =   270
                  Left            =   8550
                  TabIndex        =   165
                  Top             =   720
                  Width           =   75
                  _ExtentX        =   132
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   81
                  Left            =   9000
                  TabIndex        =   42
                  Top             =   720
                  Width           =   810
                  _ExtentX        =   1429
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
                  Mask            =   ">?????"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   82
                  Left            =   9840
                  TabIndex        =   43
                  Top             =   720
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
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
                  Format          =   "#,##0"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   315
                  Index           =   85
                  Left            =   3720
                  TabIndex        =   48
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  AutoTab         =   -1  'True
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">99:99"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox Text 
                  Height          =   285
                  Index           =   86
                  Left            =   5640
                  TabIndex        =   39
                  Top             =   240
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   503
                  _Version        =   393216
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
                  Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox MaskEdBox16 
                  Height          =   270
                  Left            =   8550
                  TabIndex        =   230
                  Top             =   240
                  Width           =   75
                  _ExtentX        =   132
                  _ExtentY        =   476
                  _Version        =   393216
                  Appearance      =   0
                  Enabled         =   0   'False
                  PromptChar      =   "_"
               End
               Begin VB.Label Label66 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Promo Sublet"
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
                  Left            =   5640
                  TabIndex        =   231
                  Top             =   0
                  Width           =   945
               End
               Begin VB.Label sgnDSS 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "NON DSS"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000080FF&
                  Height          =   375
                  Left            =   1800
                  TabIndex        =   216
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   3495
               End
               Begin VB.Label sgnStall 
                  Alignment       =   2  'Center
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "STALL"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H0000C000&
                  Height          =   345
                  Left            =   9000
                  TabIndex        =   211
                  Top             =   0
                  Width           =   900
               End
               Begin VB.Label Label60 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Estimasi Adjust"
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
                  Left            =   2160
                  TabIndex        =   195
                  Top             =   200
                  Visible         =   0   'False
                  Width           =   1440
               End
               Begin VB.Label Label56 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sublet Order"
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
                  Left            =   5640
                  TabIndex        =   166
                  Top             =   480
                  Width           =   915
               End
               Begin VB.Label Label52 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Mech"
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
                  Left            =   4200
                  TabIndex        =   163
                  Top             =   480
                  Width           =   375
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Job Order"
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
                  Left            =   900
                  TabIndex        =   162
                  Top             =   510
                  Width           =   720
               End
               Begin VB.Label Label27 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipe Job"
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
                  TabIndex        =   161
                  Top             =   510
                  Width           =   600
               End
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   5880
            Left            =   5760
            TabIndex        =   121
            Top             =   240
            Width           =   6015
            Begin VB.ComboBox cboPelanggan 
               Height          =   315
               ItemData        =   "frmBooking.frx":329E
               Left            =   840
               List            =   "frmBooking.frx":32A0
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   3360
               Width           =   1215
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   39
               Left            =   3495
               TabIndex        =   122
               Top             =   3900
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   40
               Left            =   3495
               TabIndex        =   123
               Top             =   4200
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   52
               Left            =   3495
               TabIndex        =   57
               Top             =   4860
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   53
               Left            =   3495
               TabIndex        =   124
               Top             =   5460
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   55
               Left            =   3495
               TabIndex        =   125
               Top             =   4500
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   270
               Index           =   75
               Left            =   2040
               TabIndex        =   56
               Top             =   3345
               Width           =   3885
               _ExtentX        =   6853
               _ExtentY        =   476
               _Version        =   393216
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   30
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   37
               Left            =   840
               TabIndex        =   53
               Top             =   2760
               Width           =   5100
               _ExtentX        =   8996
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   38
               Left            =   840
               TabIndex        =   54
               Top             =   3030
               Width           =   5100
               _ExtentX        =   8996
               _ExtentY        =   503
               _Version        =   393216
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
               Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   45
               Left            =   1080
               TabIndex        =   171
               Top             =   4320
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   59
               Left            =   1080
               TabIndex        =   172
               Top             =   3960
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
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
               Mask            =   ">99-99-9999"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   83
               Left            =   1080
               TabIndex        =   177
               Top             =   5160
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   84
               Left            =   1080
               TabIndex        =   178
               Top             =   4800
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
               Mask            =   ">99-99-9999"
               PromptChar      =   ">"
            End
            Begin VB.Label lblFOP 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   360
               TabIndex        =   229
               Top             =   5520
               Width           =   2055
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Janji Penyerahan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   0
               TabIndex        =   179
               Top             =   4920
               Width           =   960
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Estimasi Selesai  "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   240
               TabIndex        =   173
               Top             =   4080
               Width           =   720
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pesan Khusus"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   138
               Top             =   2640
               Width           =   615
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saran"
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
               Left            =   330
               TabIndex        =   137
               Top             =   3075
               Width           =   420
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Karakter"
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
               TabIndex        =   136
               Top             =   3405
               Width           =   615
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL"
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
               Left            =   2850
               TabIndex        =   131
               Top             =   5520
               Width           =   480
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SUBLET"
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
               Left            =   2790
               TabIndex        =   130
               Top             =   4920
               Width           =   540
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PART"
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
               Left            =   2940
               TabIndex        =   129
               Top             =   4260
               Width           =   390
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "________________________  +"
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
               Left            =   3525
               TabIndex        =   128
               Top             =   5175
               Width           =   2370
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "JASA"
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
               Left            =   2955
               TabIndex        =   127
               Top             =   3960
               Width           =   375
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "OLI"
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
               Left            =   3075
               TabIndex        =   126
               Top             =   4560
               Width           =   255
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4440
            Left            =   120
            TabIndex        =   117
            Top             =   1560
            Width           =   5640
            Begin VB.CommandButton Command3 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Info Keluhan"
               Height          =   465
               Left            =   0
               Style           =   1  'Graphical
               TabIndex        =   298
               Top             =   1800
               Width           =   975
            End
            Begin VB.CommandButton cmdFIR1 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Lengkapi FIR"
               Height          =   225
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   225
               Top             =   1320
               Width           =   1215
            End
            Begin VB.CommandButton cmdFIR2 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Lengkapi FIR"
               Height          =   225
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   224
               Top             =   1560
               Width           =   1215
            End
            Begin VB.CommandButton cmdFIR3 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Lengkapi FIR"
               Height          =   225
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   223
               Top             =   1820
               Width           =   1215
            End
            Begin VB.CommandButton cmdFIR4 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Lengkapi FIR"
               Height          =   225
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   222
               Top             =   2060
               Width           =   1215
            End
            Begin VB.CommandButton cmdFIR5 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Lengkapi FIR"
               Height          =   225
               Left            =   4320
               Style           =   1  'Graphical
               TabIndex        =   221
               Top             =   2310
               Width           =   1215
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00FFFFFF&
               Caption         =   "       Service       "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1080
               TabIndex        =   174
               Top             =   3960
               Width           =   2175
               Begin VB.OptionButton opSvc 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Ditunggu"
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
                  Left            =   120
                  TabIndex        =   176
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton opSvc 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Ditinggal"
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
                  Left            =   1080
                  TabIndex        =   175
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   4
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   25
               Top             =   2280
               Width           =   4455
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   3
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   24
               Top             =   2040
               Width           =   4455
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   2
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   23
               Top             =   1800
               Width           =   4455
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   22
               Top             =   1560
               Width           =   4455
            End
            Begin VB.TextBox Text1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   21
               Top             =   1320
               Width           =   4455
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   27
               Left            =   1095
               TabIndex        =   16
               Top             =   0
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   28
               Left            =   1095
               TabIndex        =   17
               Top             =   240
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   29
               Left            =   1095
               TabIndex        =   18
               Top             =   480
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   30
               Left            =   1095
               TabIndex        =   19
               Top             =   720
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   31
               Left            =   1095
               TabIndex        =   20
               Top             =   975
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   315
               Index           =   58
               Left            =   120
               TabIndex        =   132
               Top             =   840
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   72
               Left            =   5295
               TabIndex        =   26
               Top             =   1320
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   73
               Left            =   5295
               TabIndex        =   27
               Top             =   1560
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   74
               Left            =   5295
               TabIndex        =   28
               Top             =   1785
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   76
               Left            =   5295
               TabIndex        =   29
               Top             =   2010
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   240
               Index           =   77
               Left            =   5295
               TabIndex        =   30
               Top             =   2235
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
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
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   60
               Left            =   1095
               TabIndex        =   31
               Top             =   2760
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   61
               Left            =   1095
               TabIndex        =   32
               Top             =   3045
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   62
               Left            =   1095
               TabIndex        =   33
               Top             =   3360
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   63
               Left            =   1095
               TabIndex        =   34
               Top             =   3645
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   26
               Left            =   4440
               TabIndex        =   139
               Top             =   4080
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   8
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CCCCCCCC"
               PromptChar      =   ">"
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User"
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
               Left            =   3990
               TabIndex        =   140
               Top             =   4125
               Width           =   330
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Catatan"
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
               Left            =   360
               TabIndex        =   135
               Top             =   2805
               Width           =   585
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Keluhan Customer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   134
               Top             =   1320
               Width           =   945
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Permintaan Customer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   118
               Top             =   120
               Width           =   945
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            TabIndex        =   110
            Top             =   240
            Width           =   5625
            Begin VB.CheckBox chkKirim 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Kirim"
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
               Left            =   4680
               TabIndex        =   150
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox chSBNP 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SB NP"
               Height          =   255
               Left            =   2520
               TabIndex        =   149
               Top             =   940
               Width           =   855
            End
            Begin VB.CheckBox chEM 
               BackColor       =   &H00FFFFFF&
               Caption         =   "EM"
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
               Left            =   1800
               TabIndex        =   148
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox chkAmbil 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ambil"
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
               Left            =   3840
               TabIndex        =   144
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox chkThs 
               BackColor       =   &H00FFFFFF&
               Caption         =   "THS"
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
               Left            =   840
               TabIndex        =   120
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox chkRtn 
               BackColor       =   &H00FFFFFF&
               Caption         =   "RTJ"
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
               Left            =   120
               TabIndex        =   119
               Top             =   960
               Width           =   735
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   23
               Left            =   1260
               TabIndex        =   13
               Top             =   0
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               Mask            =   ">99-99-9999"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   24
               Left            =   1260
               TabIndex        =   14
               Top             =   300
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   25
               Left            =   1260
               TabIndex        =   15
               Top             =   600
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CC"
               PromptChar      =   ">"
            End
            Begin VB.Shape Shape3 
               Height          =   315
               Left            =   0
               Top             =   930
               Width           =   5535
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "(HH:MM)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2040
               TabIndex        =   116
               Top             =   322
               Width           =   1140
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "(DD-MM-YYYY)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2535
               TabIndex        =   115
               Top             =   0
               Width           =   1140
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "D [ Datang ] / T [Telepon] / S [Sms]"
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
               Left            =   1710
               TabIndex        =   114
               Top             =   645
               Width           =   2550
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Via"
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
               Left            =   1005
               TabIndex        =   113
               Top             =   645
               Width           =   210
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jam Booking"
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
               Left            =   345
               TabIndex        =   112
               Top             =   345
               Width           =   885
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal Booking"
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
               Left            =   60
               TabIndex        =   111
               Top             =   0
               Width           =   1170
            End
         End
         Begin VB.CommandButton cmdWarranty 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Warranty"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   10440
            Width           =   1890
         End
         Begin VB.CommandButton TblBatal 
            Caption         =   "&Batal"
            Height          =   390
            Left            =   9240
            TabIndex        =   108
            Top             =   11640
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label nAR2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 % "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   5280
            TabIndex        =   185
            Top             =   10800
            Width           =   840
         End
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   78
      Left            =   11280
      TabIndex        =   154
      Top             =   240
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
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
      Mask            =   ">CCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00FFFF80&
      Caption         =   "No &Polisi"
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
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   196
      Top             =   50
      Width           =   1170
   End
   Begin VB.TextBox txtKodeLgnLama 
      Height          =   285
      Left            =   1260
      TabIndex        =   227
      Top             =   615
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.TextBox txtNoBookingApp 
      Height          =   375
      Left            =   720
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   264
      Top             =   10080
      Width           =   6735
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   88
      Left            =   11280
      TabIndex        =   299
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
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
      Mask            =   ">CCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.Label lblKeteranganBawah 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   600
      TabIndex        =   201
      Top             =   9720
      Width           =   12375
   End
   Begin VB.Label lblEdit 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   194
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9855
      Left            =   0
      Picture         =   "frmBooking.frx":32A2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
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
      Left            =   11520
      TabIndex        =   94
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "(HH:MM)"
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
      Left            =   10965
      TabIndex        =   93
      Top             =   1065
      Width           =   630
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Order"
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
      Left            =   9030
      TabIndex        =   90
      Top             =   735
      Width           =   1020
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Booking"
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
      Left            =   6600
      TabIndex        =   78
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmBooking"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public baris As Integer
    Dim barispart As Integer
Dim barisSublet As Integer
'Dim Enter As Integer
Dim stkPart As Integer
Dim namaPart As String
Dim HrgPart As Double
Dim satPart As String
Dim oliPart As Boolean
Public hrgLabour As Double
Dim Jrate As Double
Dim fOther As Boolean
Dim nopol As String
Public nJRate As Double
Dim nWKT As String
Dim nMech As String
Dim bKondisi(6) As String
Private nX As Boolean
Dim rpt As Object

Const MaxBrs = 8
Const MaxBrsPart = 30
Dim simpan As Boolean
Dim gNoRangka As String
Dim rsTemp As ADODB.Recordset

Dim namabySA As String
Dim NewId As Double
Dim estimasiSublet As String
     
'BOOKING ONLINE
Public isBookingOnline      As Boolean
Public noPolBookingOnline   As String
Public IDBookingOnline      As Double
Public noTelpBookingOnline  As String
Public namaCustomerBookingOnline As String
Public kotaBookingOnline    As String
Public tanggalBookingOnline As String
Public jamBookingOnline     As String
Public keluhanOnline        As String
Public permintaanOnline     As String
Public mdlKendaraanOnline   As String
Public noBookingAppOnline   As String
Public kodePromoAppOnline   As String
Public nikSales             As String

Private isBookDiag          As Boolean
Private no_approval         As String

Private totalNilaiPart      As Double

'26-02-2015 RUDI job untuk mengambil job order
Private job                 As String
Private jobQuery            As String

Dim rsRecoStall             As New ADODB.Recordset
Dim tempWaktuAwalDSS        As String
Dim tempWaktuAkhirDSS       As String
Dim tempStall               As String
Dim nRate                   As Double
Private adRate              As String
Private isLoadBooking       As Boolean
''20171010 RETENTION
Public isRetentionBoard     As Boolean
Public tglAwalBook          As Date


    Dim isRES As Boolean
 
Private tagihanFOP  As String
Private kodeMDL     As String
Private daftarSBE   As String
Private rsFOP       As ADODB.Recordset

Public Property Let passingISBookingOnline(VALUES As Boolean)
    isBookingOnline = VALUES
    isLoadBooking = False
End Property

Public Property Let passingNoPolBookingOnline(VALUES As String)
    noPolBookingOnline = VALUES
End Property

Public Property Let passingIDBookingOnline(VALUES As Double)
    IDBookingOnline = VALUES
End Property

Public Property Let passingNoTelpBookingOnline(VALUES As String)
    noTelpBookingOnline = VALUES
End Property

Public Property Let passingNamaCustomerBookingOnline(VALUES As String)
    namaCustomerBookingOnline = VALUES
End Property

Public Property Let passingKotaBookingOnline(VALUES As String)
    kotaBookingOnline = VALUES
End Property

Public Property Let passingTanggalBookingOnline(VALUES As String)
    tanggalBookingOnline = VALUES
End Property

Public Property Let passingJamBookingOnline(VALUES As String)
    jamBookingOnline = VALUES
End Property

Public Property Let passingKeluhanOnline(VALUES As String)
    keluhanOnline = VALUES
End Property

Public Property Let passingPermintaanOnline(VALUES As String)
    permintaanOnline = VALUES
End Property

Public Property Let passingmdlKendaraanOnline(VALUES As String)
    mdlKendaraanOnline = VALUES
End Property

Public Property Let passingNoBookingAppOnline(VALUES As String)
    noBookingAppOnline = VALUES
End Property

Public Property Let passingKodePromosiAppOnline(VALUES As String)
    kodePromoAppOnline = VALUES
End Property

Public Property Let passingNisSALES(VALUES As String)
    nikSales = VALUES
End Property

Public Property Let passingX(VALUES As Boolean)
    nX = VALUES
End Property

Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
          Dim aa As String
          Dim xHasil As String
1        On Error GoTo GensSql_Error

2         xHasil = ""
3         If xUpdate Then
4             xHasil = "UPDATE " + xTabel + " SET "
5             i = 0
6             Do While gsSql1(i) <> ""
7                 If i <> 0 Then
8                     xHasil = xHasil + ","
9                 End If
                  
10                aa = UCase(gsSql1(i))
11                Select Case aa
                  Case "PPN", "DISC"
12                    xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
13                Case Else
14                    xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
15                End Select
16                i = i + 1
17            Loop
18        Else
19            xHasil = "INSERT INTO " + xTabel + " ("
20            i = 0
21            Do While gsSql1(i) <> ""
22                If i <> 0 Then
23                    xHasil = xHasil + ","
24                End If
25                xHasil = xHasil + gsSql1(i)
26                i = i + 1
27            Loop
28            xHasil = xHasil + ") VALUES ("
              
29            i = 0
30            Do While gsSql1(i) <> ""
31                If i <> 0 Then
32                    xHasil = xHasil + ","
33                End If
                  
34                aa = UCase(gsSql1(i))
35                Select Case aa
                  Case "PPN", "DISC"
36                    xHasil = xHasil + gsSql2(i)
37                Case Else
38                    xHasil = xHasil + "'" + gsSql2(i) + "'"
39                End Select
40                i = i + 1
41            Loop
42            xHasil = xHasil + ")"
43        End If
44        GensSql = xHasil

45       On Error GoTo 0
46       Exit Function

GensSql_Error:

47        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GensSql of Form frmBooking")
End Function

Private Sub hapusForm()
      ' *** FUNGSI UNTUK MEMBERSIHKAN TEXTBOX PADA FORM ***
1        On Error GoTo hapusForm_Error

2         Frame2.Top = -100
3         For i = 0 To 86
4             Select Case i
                  Case 0
5                     Text(i) = ""
6                     Text(i).Mask = "??"
7                 Case 56
8                     Text(i) = ""
9                     Text(i).Mask = "9999"
10                Case 57
11                    Text(i) = ""
12                    Text(i).Mask = "CCC"
13                Case 6, 7, 8
14                    Text(i) = ""
15                    Text(i).Mask = "9999-9999999999"
16                Case 9
17                    Text(i) = "0"
18                    Text(i).Mask = "9999"
19                Case 16, 17, 18, 20, 23, 59, 84
20                    Text(i) = ">>->>->>>>"
21                    Text(i).Mask = "99-99-9999"
22                Case 19, 24, 45, 58, 83, 85
23                    Text(i) = ">>:>>"
24                    Text(i).Mask = "99:99"
25                Case 36, 39, 40, 55, 52, 53, 70, 82
26                    Text(i) = 0
27                Case 42
28                    Text(i) = ">>,>>"
29                    Text(i).Mask = "99,99"
30                    Text(i).Enabled = False
31                Case Else
32                    Text(i) = ""
33            End Select
34        Next
          
35        For i = 0 To 4
36            Text1(i) = ""
37        Next
          
38        For i = 1 To 7
39            dGrid.TextMatrix(i, 1) = ""
40            dGrid.TextMatrix(i, 2) = ""
41            dGrid.TextMatrix(i, 3) = ""
42            dGrid.TextMatrix(i, 4) = ""
43            partGrid.TextMatrix(i, 1) = ""
44            partGrid.TextMatrix(i, 2) = ""
45            partGrid.TextMatrix(i, 3) = ""
46            partGrid.TextMatrix(i, 4) = ""
47            partGrid.TextMatrix(i, 5) = ""
48            partGrid.TextMatrix(i, 6) = ""
49            partGrid.TextMatrix(i, 7) = ""
50            partGrid.TextMatrix(i, 8) = ""
51        Next
          
52        For i = 1 To 30
53            subletGrid.TextMatrix(i, 0) = ""
54            subletGrid.TextMatrix(i, 1) = ""
55            subletGrid.TextMatrix(i, 2) = ""
56            subletGrid.TextMatrix(i, 3) = ""
57            subletGrid.TextMatrix(i, 4) = ""
58            subletGrid.TextMatrix(i, 5) = ""
59            subletGrid.TextMatrix(i, 6) = ""
60            If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
61                subletGrid.TextMatrix(i, 7) = ""
62                subletGrid.TextMatrix(i, 8) = ""
63            End If
64        Next
          
65        baris = 1
66        barispart = 1
67        barisSublet = 1
          
          
68        For i = 0 To 5
69            bKondisi(i) = ""
70        Next
          
71        nJRate = 0
72        cmdStall.Tag = "belum"
73        cmdWarranty.Enabled = False
74        cmdWarranty.BackColor = &HC0C0C0
75        cmdWaktu.Enabled = False
76        cmdHistory.Enabled = False
77        Command1.Enabled = False
78        chkRtn = Unchecked
79        chkThs = Unchecked
80        VScroll1.value = 0
81        isJualNRM = False
82        Text(54).Visible = False
83        For i = 65 To 69
84            Text(i).Visible = False
85        Next
86        chkAmbil = Unchecked
87        chkKirim = Unchecked
88        chkRtn = Unchecked
89        chkThs = Unchecked
          
90        Text(23).Enabled = True
91        Text(24).Enabled = True
          
          ' DEFAULT DOTAM & DDATE
92        Text(16) = "01-01-3000"
93        Text(17) = "01-01-3000"
94        gNoRangka = ""
95        Text(18) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
96        Text(19) = Format(isnow, "HH:NN")
97        If Log.isDSS = 1 Then
98            sgnDSS.Visible = True
99            sgnDSS.Caption = "NON DSS"
100       End If
101       If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then Call objO.UpdateDataACC("Update SOORD Set Status = '' Where Status = '" & gSA & "'")
          'Call isiNoBook
102       tglAwalBook = "01-01-1900"
          
103       If Log.NHS = 1 Then
104           chkThs.Caption = "NHS"
105       Else
106           chkThs.Caption = "THS"
107       End If
          
108       If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
109           Label66.Visible = True
110           Text(86).Visible = True
111           MaskEdBox16.Visible = True
112       Else
113           Label66.Visible = False
114           Text(86).Visible = False
115           MaskEdBox16.Visible = False
116       End If
          
117       If Log.DIDX = 1 Then
118           frameDIDX.Visible = True
119           Command3.Visible = True
120       Else
121           frameDIDX.Visible = False
122           Command3.Visible = False
123       End If
124       Text(88) = ""
125      On Error GoTo 0
126      Exit Sub

hapusForm_Error:

127       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmBooking")
End Sub

Private Sub isiNoBook()
      ' *** FUNGSI UNTUK MENGISI NOMOR URUT BOOKING, 6 DIGIT ***
          '- proses mengisi no booking berdasarkan t_profil
          'Set oRs2 = objO.FetchData("SELECT FN_NO_BOOK from t_profil WITH (NOLOCK)")
1        On Error GoTo isiNoBook_Error
         
      ''    Set oRs2 = objO.FetchData("SELECT FN_NO_BOOK from t_profil")

2        Set oRs2 = objHT.FetchDataBengkel("SELECT FN_NO_BOOK from t_profil")
          Dim no_urut As String
          
          '- pengisian no booking lama
          'Set oRs2 = objO.FetchData("SELECT nofak FROM Book WHERE nofak <> 'BOOKED' AND left(nofak,2) = '" & Right(Year(isnow), 2) & "' AND right(left(nofak,4),2) = '" & Format(Month(isnow), "0#") & "' ORDER BY nofak DESC")
      '    If Not oRs2.EOF Then
      '        Text(1) = Right(Year(isnow), 2) & Format(Month(isnow), "0#") & Format(CDbl(Right(oRs2!NoFak, 4)) + 1, "0###")
      '    Else
      '        Text(1) = Right(Year(isnow), 2) & Format(Month(isnow), "0#") & Format(1, "0###")
      '    End If
      '    oRs2.Close

3         no_urut = ""
4         If Not oRs2.EOF Then
5             If IsNull(oRs2.Fields(0)) Then
                  'jika ada yang null
6                 Text(1) = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(1, "0###")
7                 no_urut = Format(1, "0###")
8             Else
9                 If CDbl(Right(oRs2!FN_NO_BOOK, 4)) = 9999 Then
10                   erTranAll = True
11                   ErrMsg.judul = ""
12                   ErrMsg.message = "Silahkan reset no booking anda, proses simpan booking gagal!"
13                   ErrMsg.show
14                   Exit Sub
15                End If
16                Text(1) = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(CDbl(Right(oRs2!FN_NO_BOOK, 4)) + 1, "0###")
17                no_urut = Format(CDbl(Right(oRs2!FN_NO_BOOK, 4)) + 1, "0###")
18            End If
19        Else
20            Text(1) = Right(year(isnow), 2) & Format(Month(isnow), "0#") & Format(1, "0###")
21            no_urut = Format(1, "0###")
22        End If
          ''2016-02-24
23        Call objHT.InsertData_Bengkel("UPDATE t_profil SET FN_NO_BOOK = '" & no_urut & "' ")
          ''objO.InsertData ("UPDATE t_profil SET FN_NO_BOOK = '" & no_urut & "' ")
24        oRs2.Close

25       On Error GoTo 0
26       Exit Sub

isiNoBook_Error:
27        erTranAll = True
28        ErrMsg.judul = ""
29        ErrMsg.message = ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiNoBook of Form frmBooking")
30        ErrMsg.show
          
          ''MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiNoBook of Form frmBooking")
End Sub

Private Sub saveSublet()
1        On Error GoTo saveSublet_Error

2         Call objHT.DeleteData_Bengkel("DELETE BOOK_SUBLET WHERE NOFAK = '" & Trim(Text(1)) & "' AND NO_POLISI = '" & nopol & "' AND NO_CHASSIS = '" & Trim(Text(41)) & "'")
3         For i = 1 To barisSublet - 1
4             With subletGrid
5                 If .TextMatrix(i, 1) <> "" Then
6                     If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
7                         Call objHT.DeleteData_Bengkel("DELETE FROM BOOK_SUBLET WHERE NOFAK = '" & Trim(Text(1)) & "' AND NO_POLISI = '" & nopol & "' AND NO_CHASSIS = '" & Trim(Text(41)) & "' AND NAMABARANG = '" & .TextMatrix(i, 1) & "'")
8                         Call objHT.InsertData_Bengkel("INSERT INTO BOOK_SUBLET VALUES ('" & Trim(Text(1)) & "','" & Trim(Text(41)) & "','" & Trim(nopol) & "','" & .TextMatrix(i, 7) & "', '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 3) & "', '" & CDbl(.TextMatrix(i, 2)) & "', " & CDbl(.TextMatrix(i, 4)) & ",'" & IIf(.TextMatrix(i, 5) = "J", "J", "P") & "','" & .TextMatrix(i, 6) & "','" & .TextMatrix(i, 8) & "')")
9                     Else
10                        Call objHT.DeleteData_Bengkel("DELETE FROM BOOK_SUBLET WHERE NOFAK = '" & Trim(Text(1)) & "' AND NO_POLISI = '" & nopol & "' AND NO_CHASSIS = '" & Trim(Text(41)) & "' AND NAMABARANG = '" & .TextMatrix(i, 1) & "'")
11                        Call objHT.InsertData_Bengkel("INSERT INTO BOOK_SUBLET VALUES ('" & Trim(Text(1)) & "','" & Trim(Text(41)) & "','" & Trim(nopol) & "','', '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 3) & "', '" & CDbl(.TextMatrix(i, 2)) & "', " & CDbl(.TextMatrix(i, 4)) & ",'" & IIf(.TextMatrix(i, 5) = "J", "J", "P") & "','" & .TextMatrix(i, 6) & "')")
12                    End If
13                End If
14            End With
15        Next

16       On Error GoTo 0
17       Exit Sub

saveSublet_Error:
18        erTranAll = True
19        ErrMsg.judul = ""
20        ErrMsg.message = "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure saveSublet of Form frmBooking")
21        ErrMsg.show
          
      '    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure saveSublet of Form frmBooking")
End Sub

Private Sub savePart(noBok As String)
          Dim rsS As ADODB.Recordset
          Dim Flg As String
          Dim iNo As Double
          
          Dim rsSO As ADODB.Recordset
1        On Error GoTo savePart_Error

2         If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then
3             Set rsSO = objHT.FetchDataACC("Select KodeItem + NamaBySA + convert(varchar(10),Nomor) as Kode from SOORD WHERE NOSPK = '" & noBok & "' Order By Nomor DESC")
4         End If
          
          Dim rsSoH As ADODB.Recordset
          
5         Set rsSoH = objHT.FetchDataACC("Select NoSPK from SOHDR WHERE NOSPK = '" & noBok & "'")
          
          Dim nSyaratBayar As String
          Dim rsSb As ADODB.Recordset
          'Set rsSb = objO.FetchDataACC("SELECT TOPGRP as SyaratBayar FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & Trim(Text(2)) & "'")
6         Set rsSb = objHT.FetchDataACC("SELECT TOPGRP as SyaratBayar FROM ARMST WHERE KODELGN = '" & Trim(Text(2)) & "'")
7         If rsSb.RecordCount > 0 Then
8             nSyaratBayar = rsSb.Fields(0).value
9             Set rsSb = Nothing
10        Else
11            nSyaratBayar = "D0"
12        End If
          
          Dim nDelIndex As String
13        nDelIndex = GetDelIndex("SV")
          
14        Call objHT.DeleteData_Bengkel("DELETE BOOK_PART WHERE NOFAK = '" & noBok & "'")
          
          Dim nSat As String
15        nSat = ""

           Dim xNomor As String
16         xNomor = ""
           Dim rsSOno As ADODB.Recordset
           'Set rsSOno = objO.FetchDataACC("Select top 1 Nomor from SOORD WITH (NOLOCK) WHERE NOSO = '" & Trim(Text(64)) & "' Order By Nomor DESC")
17         Set rsSOno = objHT.FetchDataACC("Select top 1 Nomor from SOORD WHERE NOSO = '" & Trim(Text(64)) & "' Order By Nomor DESC")
           
18         If rsSOno.RecordCount > 0 Then
19             iNo = CDbl(rsSOno.Fields(0).value + 1)
20             Set rsSOno = Nothing
21         Else
22             iNo = 1
23         End If
24            For i = 1 To barispart - 1
25                If Trim(partGrid.TextMatrix(i, 1)) <> "" Or Trim(partGrid.TextMatrix(i, 7)) <> "" Then
26                    Set rsS = objHT.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & partGrid.TextMatrix(i, 1) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan ORDER BY A.kodeitem")
27                    If Not rsS.EOF Then
28                        If rsS("QTYSTOCK") < CDbl(partGrid.TextMatrix(i, 4)) Then
29                            Flg = "KURANG"
30                        Else
31                            Flg = "LENGKAP"
32                        End If
33                    Else
34                        Flg = "KURANG"
35                    End If
                      
                      'Call objHT.InsertData_Bengkel("INSERT INTO book_part VALUES ('" & noBok & "','" & Trim(partGrid.TextMatrix(i, 1)) & "','" & Trim(partGrid.TextMatrix(i, 2)) & "','" & IIf(rsS.RecordCount > 0, rsS("Kodesatuan"), "ITEM") & "','" & CDbl(partGrid.TextMatrix(i, 4)) & "','" & Flg & "','" & Format(Trim(Text(23)), "mm/dd/yyyy") & "','', '" & partGrid.TextMatrix(i, 7) & "'," & Val(partGrid.TextMatrix(i, 8)) & ")")
36                    Call objHT.InsertData_Bengkel("INSERT INTO book_part (NOFAK,KODEITEM,NAMABARANG,KODESATUAN,QTY,FLAG,ETA,KETERANGAN,NAMABYSA,nID) VALUES ('" & noBok & "','" & Trim(partGrid.TextMatrix(i, 1)) & "','" & Trim(partGrid.TextMatrix(i, 2)) & "','" & IIf(rsS.RecordCount > 0, rsS("Kodesatuan"), "ITEM") & "','" & CDbl(partGrid.TextMatrix(i, 4)) & "','" & Flg & "','" & Format(Trim(Text(23)), "mm/dd/yyyy") & "','', '" & partGrid.TextMatrix(i, 7) & "'," & iNo & ")")
      '                xNomor = xNomor & "," & iNo 'Trim(Val(partGrid.TextMatrix(i, 8)))
37                    If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then
38                        If rsSO.RecordCount > 0 Then
39                            rsSO.MoveFirst
40                            rsSO.Find "Kode = '" & Trim(partGrid.TextMatrix(i, 1)) + Trim(partGrid.TextMatrix(i, 7)) + Trim(partGrid.TextMatrix(i, 8)) & "'", , adSearchForward, 1
                              'rsSO.Find "Kode = '" & Trim(partGrid.TextMatrix(i, 1)) + Trim(partGrid.TextMatrix(i, 7)) + iNo & "'", , adSearchForward, 1
                              
                              'rsSO.Find "Kode = '" & Trim(PartGrid.TextMatrix(i, 8)) & "'", , adSearchForward, 1
41                            If rsSO.EOF = True Then
42                                xNomor = xNomor & "," & Trim(Val(partGrid.TextMatrix(i, 8)))
                                  ''2015-10-06 | APPRATE diganti bisa menit
                                  ''Call objHT.InsertData_ACC("Insert Into SOORD (NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) Values ('" & noBok & "', '" & partGrid.TextMatrix(i, 1) & "', '" & partGrid.TextMatrix(i, 2) & "', '" & partGrid.TextMatrix(i, 4) & "','1', '" & partGrid.TextMatrix(i, 3) & "', '" & noBok & "', '" & gSA & "', Getdate(),'" & Val(partGrid.TextMatrix(i, 8)) & "','" & partGrid.TextMatrix(i, 7) & "', '0', '" & TglNul2(Trim(Text(23))) + " " + Trim(Text(24)) & "','" & gSA & "','" & TglNul2(Trim(Text(23))) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), Trim(Text(24)))), "HH:MM:SS") & "')")
43                                Call objHT.InsertData_ACC("Insert Into SOORD (NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) Values ('" & noBok & "', '" & partGrid.TextMatrix(i, 1) & "', '" & partGrid.TextMatrix(i, 2) & "', '" & partGrid.TextMatrix(i, 4) & "','1', '" & partGrid.TextMatrix(i, 3) & "', '" & noBok & "', '" & gSA & "', Getdate(),'" & Val(partGrid.TextMatrix(i, 8)) & "','" & Trim(partGrid.TextMatrix(i, 7)) & "', '0', '" & TglNul2(Trim(Text(23))) + " " + Trim(Text(24)) & "','" & gSA & "','" & TglNul2(Trim(Text(23))) + " " + Format((DateAdd(Left(Log.AppRate, 1), Mid(Log.AppRate, 2, Len(Log.AppRate)), Trim(Text(24)))), "HH:MM:SS") & "')")
                                  
44                            Else
45                                xNomor = xNomor & "," & Trim(Val(partGrid.TextMatrix(i, 8)))
46                            End If
47                        Else
48                            xNomor = xNomor & "," & Trim(Val(partGrid.TextMatrix(i, 8)))
                              ''2015-10-06 | APPRATE diganti bisa menit
                              
                              ''Call objHT.InsertData_ACC("Insert Into SOORD " & _
                              ''    "(NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) " & _
                              ''    "Values ('" & noBok & "', '" & partGrid.TextMatrix(i, 1) & "', '" & partGrid.TextMatrix(i, 2) & "', '" & partGrid.TextMatrix(i, 4) & "','1', '" & partGrid.TextMatrix(i, 3) & "', '" & noBok & "', '" & gSA & "', Getdate(),'" & Val(partGrid.TextMatrix(i, 8)) & "','" & partGrid.TextMatrix(i, 7) & "', '0', '" & TglNul2(Trim(Text(23))) + " " + Trim(Text(24)) & "','" & gSA & "','" & TglNul2(Trim(Text(23))) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), Trim(Text(24)))), "HH:MM:SS") & "')")
49                            Call objHT.InsertData_ACC("Insert Into SOORD (NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) " & _
                                  "Values ('" & noBok & "', '" & partGrid.TextMatrix(i, 1) & "', '" & partGrid.TextMatrix(i, 2) & "', '" & partGrid.TextMatrix(i, 4) & "','1', '" & partGrid.TextMatrix(i, 3) & "', '" & noBok & "', '" & gSA & "', Getdate(),'" & Val(partGrid.TextMatrix(i, 8)) & "','" & partGrid.TextMatrix(i, 7) & "', '0', '" & TglNul2(Trim(Text(23))) + " " + Trim(Text(24)) & "','" & gSA & "','" & TglNul2(Trim(Text(23))) + " " + Format((DateAdd(Left(Log.AppRate, 1), Mid(Log.AppRate, 2, Len(Log.AppRate)), Trim(Text(24)))), "HH:MM:SS") & "')")
50                        End If
51                    End If
52                End If
53            Next
          'End If
          
54            If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then
                '  Set rsSoH = objO.FetchDataACC("Select NoSPK from SOHDR WHERE noso = '" & noBok & "'")
55                If barispart > 1 Then
                  
56                    If rsSoH.RecordCount > 0 Then
57                        Call objHT.UpdateData_ACC("Update SOHDR Set TglPerlu = '" & TglNul2(Trim(Text(23))) + " " + Trim(Text(24)) & "' Where NoSPK = '" & noBok & "'")
58                    Else
59                        If barispart > 1 Then
60                            Call objHT.InsertData_ACC("insert into SOHDR " & _
                              "(NOSO,TglSO, TypeTrn, TglPerlu, KodeLgn, NamaKirim, AlamatKirim1, AlamatKirim2,AlamatKirim3,SyaratBayar, KODESUMBER, DELINDEX, TglEntry, KodeUser, NoSPK, TUTUPORDER ) " & _
                              "values ('" & noBok & "', Getdate(), 'A','" & TglNul2(Trim(Text(23))) + " " + Trim(Text(24)) & "', '" & Trim(Text(2)) & "', '" & Trim(Text(3)) & "', '" & Trim(Text(4)) & "', '" & Trim(Text(5)) & "', '', '" & nSyaratBayar & "', 'SV', '" & nDelIndex & "', getdate(), '" & gSA & "', '" & noBok & "',0 )")
61                        End If
62                    End If
63                End If
64                If xNomor = "" Then
                  
65                Else
66                    Call objHT.UpdateData_ACC("Delete SOORD Where NOSPK = '" & noBok & "' AND IsNull(qtyreserve,0) = 0 and Nomor not in (" & Right(xNomor, Len(xNomor) - 1) & ")")
67                End If
                  
68            End If

69       On Error GoTo 0
70       Exit Sub

savePart_Error:
71        erTranAll = True
72        ErrMsg.judul = ""
73        ErrMsg.message = "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure savePart of Form frmBooking")
74        ErrMsg.show
      '    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure savePart of Form frmBooking")
End Sub

Private Sub TampilWaktu(Tggl As Date)
      Dim rsWaktu As ADODB.Recordset
      Dim nStall As String
      Dim i As Integer

1        On Error GoTo TampilWaktu_Error
          GridWaktu.Clear
2         GridWaktu.TextMatrix(0, 0) = "..."
3         GridWaktu.TextMatrix(0, 1) = "STALL"
          
          'FUNGSI MENGHASILKAN WAKTU TERSEDIA BERDASARKAN TGL TTN (HARI H)
4         If DateDiff("d", isnow, Tggl) >= 2 Then
5             Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl, "mm/dd/yyyy") & "'")
6             GridWaktu.TextMatrix(0, 2) = "H - 2"
7             GridWaktu.TextMatrix(0, 3) = "H - 1"
8             GridWaktu.TextMatrix(0, 4) = "  H  "
9             GridWaktu.TextMatrix(0, 5) = "H + 1"
10        ElseIf DateDiff("d", isnow, Tggl) = 1 Then
11            Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 1, "mm/dd/yyyy") & "'")
12            GridWaktu.TextMatrix(0, 2) = "H - 1"
13            GridWaktu.TextMatrix(0, 3) = "  H  "
14            GridWaktu.TextMatrix(0, 4) = "H + 1"
15            GridWaktu.TextMatrix(0, 5) = "H + 2"
16        ElseIf DateDiff("d", isnow, Tggl) <= 0 Then
17            Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 2, "mm/dd/yyyy") & "'")
18            GridWaktu.TextMatrix(0, 2) = "  H  "
19            GridWaktu.TextMatrix(0, 3) = "H + 1"
20            GridWaktu.TextMatrix(0, 4) = "H + 2"
21            GridWaktu.TextMatrix(0, 5) = "H + 3"
22        End If
          
          'FUNSGI UNTUK MENAMPILKAN WAKTU TERSEDIA H-2 SAMPAI H+1
23        Set rsWaktu = objO.FetchData("SELECT * FROM STALLWKF WITH (NOLOCK) where NoStall not in ('STALL FINL','STALL STOP','stall SUBL') ORDER BY LEFT(NOSTALL,1),CASE LEFT(NOSTALL,1) WHEN 'S' THEN CAST(SUBSTRING(NOSTALL, 7, LEN(NOSTALL)-6) AS INTEGER) ELSE CAST(SUBSTRING(NOSTALL, 5, LEN(NOSTALL)-4) AS INTEGER) END, [H-2], [H-1], H, [H+1]")
24        If Not rsWaktu.EOF Then
25            i = 1
26            GridWaktu.Rows = rsWaktu.RecordCount + 1
              
27            GridWaktu.ColWidth(0) = 200
28            GridWaktu.ColWidth(1) = 700
29            GridWaktu.ColWidth(2) = 1200
30            GridWaktu.ColWidth(3) = 1200
31            GridWaktu.ColWidth(4) = 1200
32            GridWaktu.ColWidth(5) = 1200
33            nStall = ""
              
34            Do While Not rsWaktu.EOF
35                If Trim(rsWaktu("NOSTALL")) = nStall Then
36                    GridWaktu.TextMatrix(i, 1) = ""
37                    GridWaktu.TextMatrix(i, 2) = Nul(Trim(rsWaktu("H-2")), "")
38                    GridWaktu.TextMatrix(i, 3) = Nul(Trim(rsWaktu("H-1")), "")
39                    GridWaktu.TextMatrix(i, 4) = Nul(Trim(rsWaktu("H")), "")
40                    GridWaktu.TextMatrix(i, 5) = Nul(Trim(rsWaktu("H+1")))
41                Else
42                    GridWaktu.TextMatrix(i, 0) = "*"
43                    GridWaktu.TextMatrix(i, 1) = Trim(rsWaktu("NOSTALL"))
44                    GridWaktu.TextMatrix(i, 2) = Nul(Trim(rsWaktu("H-2")), "")
45                    GridWaktu.TextMatrix(i, 3) = Nul(Trim(rsWaktu("H-1")), "")
46                    GridWaktu.TextMatrix(i, 4) = Nul(Trim(rsWaktu("H")), "")
47                    GridWaktu.TextMatrix(i, 5) = Nul(Trim(rsWaktu("H+1")))
48                    nStall = Trim(rsWaktu("NOSTALL"))
49                End If
50                i = i + 1
51                rsWaktu.MoveNext
52            Loop
53        End If

54       On Error GoTo 0
55       Exit Sub

TampilWaktu_Error:

56        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TampilWaktu of Form frmBooking")
End Sub

Public Sub panggilWarrantyBok(noRangka As String, modL As String)
      Dim rsRec As ADODB.Recordset
      Dim terlibat As Boolean

1        On Error GoTo panggilWarrantyBok_Error

          ' ================================================================================================================================
          ' 01 Mar 2017 | Thania | buka kuncian pencarian warranty berdasarkan model hanya berdasarkan noka (update ketinggalan periode yg lalu)
          ' ================================================================================================================================
      '          Set rsTemp = objO.FetchData("SELECT * FROM ho_warranty A WITH (NOLOCK) INNER JOIN ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
      '                     & "WHERE MDL = '" & modL & "' AND no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
      '                     & "ORDER BY A.id_warranty,mdl,no_Rangka1")
2         If Log.RPLHistori = 0 Then
3             Set rsTemp = objO.FetchData("SELECT * FROM ho_warranty A WITH (NOLOCK) INNER JOIN ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
                     & "WHERE no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
                     & "ORDER BY A.id_warranty,mdl,no_Rangka1")
4         Else
5             Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..ho_warranty A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
                     & "WHERE no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
                     & "ORDER BY A.id_warranty,mdl,no_Rangka1")
6         End If
          ' ================================================================================================================================
7         If Not rsTemp.EOF Then
8             Do While Not rsTemp.EOF
9                 Set rsRec = objO.FetchData("SELECT no_chassis FROM History WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
10                If Not rsRec.EOF Then
11                    terlibat = False
12                    Exit Do
13                Else
14                    terlibat = True
      '                Exit Do
15                End If
16                rsRec.Close
                  
17                Set rsRec = objO.FetchData("SELECT NO_RANGKA FROM TEMPWAR WITH (NOLOCK) WHERE NO_RANGKA = '" & noRangka & "' AND ID_WARRANTY = '" & rsTemp("ID_WARRANTY") & "'")
18                If Not rsRec.EOF Then
19                    terlibat = False
20                    Exit Do
21                Else
22                    terlibat = True
      '                Exit Do
23                End If
24                rsRec.Close
          
                  ''2018-08-31 hanung sunu
25                If Log.RPLHistori = 0 Then
26                    If Log.LogHOConn = 1 Then
27                        If ConHO = True Then
28                            Set rsRec = objO.FetchDataHO("SELECT no_chassis FROM Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
29                            If Not rsRec.EOF Then
30                                terlibat = False
31                                Exit Do
32                            Else
33                                terlibat = True
              '                    Exit Do
34                            End If
35                        End If
36                    End If
37                Else
38                    Set rsRec = objO.FetchData("SELECT no_chassis FROM [" & dbReplikasiHO & "]..Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
39                    If Not rsRec.EOF Then
40                        terlibat = False
41                        Exit Do
42                    Else
43                        terlibat = True
      '                    Exit Do
44                    End If
45                End If
                  
46                rsTemp.MoveNext
47            Loop
              
48            If terlibat Then
49                MsgBox "UNIT TERLIBAT WARRANTY !!!", vbCritical, "UNIT SSC"
50                cmdWarranty.BackColor = &HFF&
51                cmdWarranty.Enabled = True
52            End If
53        Else
54            cmdWarranty.BackColor = &HC0C0C0
55            cmdWarranty.Enabled = False
56        End If

57       On Error GoTo 0
58       Exit Sub

panggilWarrantyBok_Error:

59        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilWarrantyBok of Form frmBooking")
End Sub

Private Sub panggilCustHO(nPol As String, rs As ADODB.Recordset)
          
1        On Error GoTo panggilCustHO_Error

2         If Not rs.EOF Then
3             Text(2) = rs("FS_KD_PLG")
4             Text(44) = rs("FS_KD_WIL")
5             Text(3) = rs("CUSTOMER")
6             Text(4) = rs("CUST_ADRES")
7             Text(5) = rs("KOTA")
8             Text(6) = rs("GSM")
9             Text(7) = rs("CDMA")
10            Text(8) = rs("TELPON")
11            Text(9) = rs("FS_TH_MOBIL")
12            Text(10) = rs("MDL")
13            Text(11) = rs("MODEL")
14                        Call CexChassis_Discon(Text(11))
15                        If isOnCarClass = False Then
16                            MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
17                        End If
18            Text(12) = rs("FMDL")
19            Text(13) = rs("COLOR")
20            Text(14) = rs("FS_KD_DEALER")
21            Text(15) = rs("FS_NM_DEALER")
22            Text(16) = TglDDMMYYYY(IIf(rs("DOTAM") = "", "01-01-3000", rs("DOTAM")))    'TglDDMMYYYY(rs("DOTAM"))
23            Text(17) = TglDDMMYYYY(IIf(rs("DDATE") = "", "01-01-3000", rs("DDATE")))
24            gNoRangka = rs("NO_CHASSIS")
25            Text(41) = gNoRangka
26            Text(26) = gSA
27            cmdHistory.Enabled = True
28            Command1.Enabled = True
29        End If

30       On Error GoTo 0
31       Exit Sub

panggilCustHO_Error:

32        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilCustHO of Form frmBooking")
End Sub

Private Sub c_Click()

End Sub

Private Sub cboPelanggan_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub chEM_Click()
    If chSBNP.value = vbChecked And chEM.value = vbChecked Then
        MsgBox "Anda telah memilih tipe SBNP !", vbInformation
        chEM.value = vbUnchecked
        Exit Sub
    End If
    Call hitungEstimasiSelesai
End Sub

Private Sub chSBNP_Click()
    If chEM.value = vbChecked And chSBNP.value = vbChecked Then
        MsgBox "Anda telah memilih tipe EM !", vbInformation
        chSBNP.value = vbUnchecked
        Exit Sub
    End If
    Call hitungEstimasiSelesai
End Sub

Private Sub chSublet_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Public Sub PindahStallEM()
1        On Error GoTo PindahStallEM_Error

2         If Log.isEmStall = 1 Then
3             If chEM = Checked Then
                  Dim rsAStl As ADODB.Recordset
                  Dim keterangan As String
                  Dim noPolEM As String
                  Dim jamJM As Integer
                  Dim menitJM As Integer
                  Dim jamSelesai As String
                  
4                 noPolEM = Trim(Text(0)) & "-" & Trim(Text(56)) & "-" & Trim(Text(57))
                  
5                 If noPolEM <> "" Then
6                     If Trim(Text(1)) <> "" Then
7                         keterangan = noPolEM & " |" & Trim(Text(1))
8                     Else
9                         keterangan = noPolEM & " |"
10                    End If
11                    Set rsAStl = objO.FetchData("SELECT * FROM stall WITH(NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "%' ORDER BY noStall")
12                    If Not rsAStl.EOF Then
13                        If rsAStl("ADJUST") <> "00:30" Then
                              'lakukan pergantian nostall tersebut dengan adjust EM
14                            jamJM = CInt(Left(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") - 1))
15                            menitJM = Mid(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") + 1, Len(Trim(rsAStl("jammulai"))) - InStr(Trim(rsAStl("jammulai")), ":"))
                                                      
16                            menitJM = menitJM + 30
                              
17                            If menitJM >= 60 Then
18                               jamJM = jamJM + Val(menitJM / 60)
19                               menitJM = menitJM Mod 60
20                            End If
21                            jamSelesai = Format(jamJM, "0#") & ":" & Format(menitJM, "0#")
22                            objO.InsertData ("UPDATE STALL SET JAMSELESAI = '" & jamSelesai & "',ADJUST = '00:30' WHERE tanggal = '" & Format(Trim(Text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "%'")
23                        End If
24                    Else
25                        Set rsAStl = objO.FetchData("SELECT * FROM tempstall WITH(NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "' ORDER BY noStall")
26                        If Not rsAStl.EOF Then
27                            If rsAStl("ADJUST") <> "00:30" Then
                                  'lakukan pergantian nostall tersebut dengan adjust EM
28                                jamJM = CInt(Left(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") - 1))
29                                menitJM = Mid(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") + 1, Len(Trim(rsAStl("jammulai"))) - InStr(Trim(rsAStl("jammulai")), ":"))
                                                          
30                                menitJM = menitJM + 30
                                  
31                                If menitJM >= 60 Then
32                                   jamJM = jamJM + Val(menitJM / 60)
33                                   menitJM = menitJM Mod 60
34                                End If
35                                jamSelesai = Format(jamJM, "0#") & ":" & Format(menitJM, "0#")
36                                objO.InsertData ("UPDATE TEMPSTALL SET JAMSELESAI = '" & jamSelesai & "',ADJUST = '00:30' WHERE tanggal = '" & Format(Trim(Text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "'")
37                            End If
38                        End If
39                    End If
40                End If
41            End If
42        End If

43       On Error GoTo 0
44       Exit Sub

PindahStallEM_Error:

45        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PindahStallEM of Form frmBooking")
End Sub

Public Sub PindahStallSBNP()
1        On Error GoTo PindahStallSBNP_Error

2         If Log.isEmStall = 1 Then
3             If chEM = Checked Then
                  Dim rsAStl As ADODB.Recordset
                  Dim keterangan As String
                  Dim noPolEM As String
                  Dim jamJM As Integer
                  Dim menitJM As Integer
                  Dim jamSelesai As String
                  
4                 noPolEM = Trim(Text(0)) & "-" & Trim(Text(56)) & "-" & Trim(Text(57))
                  
5                 If noPolEM <> "" Then
6                     If Trim(Text(1)) <> "" Then
7                         keterangan = noPolEM & " |" & Trim(Text(1))
8                     Else
9                         keterangan = noPolEM & " |"
10                    End If
11                    Set rsAStl = objO.FetchData("SELECT * FROM stall WITH(NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "%' ORDER BY noStall")
12                    If Not rsAStl.EOF Then
13                        If rsAStl("ADJUST") <> "00:30" Then
                              'lakukan pergantian nostall tersebut dengan adjust EM
14                            jamJM = CInt(Left(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") - 1))
15                            menitJM = Mid(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") + 1, Len(Trim(rsAStl("jammulai"))) - InStr(Trim(rsAStl("jammulai")), ":"))
                                                      
16                            menitJM = menitJM + 60
                              
17                            If menitJM >= 60 Then
18                               jamJM = jamJM + Val(menitJM / 60)
19                               menitJM = menitJM Mod 60
20                            End If
21                            jamSelesai = Format(jamJM, "0#") & ":" & Format(menitJM, "0#")
22                            objO.InsertData ("UPDATE STALL SET JAMSELESAI = '" & jamSelesai & "',ADJUST = '00:30' WHERE tanggal = '" & Format(Trim(Text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "%'")
23                        End If
24                    Else
25                        Set rsAStl = objO.FetchData("SELECT * FROM tempstall WITH(NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "' ORDER BY noStall")
26                        If Not rsAStl.EOF Then
27                            If rsAStl("ADJUST") <> "00:30" Then
                                  'lakukan pergantian nostall tersebut dengan adjust EM
28                                jamJM = CInt(Left(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") - 1))
29                                menitJM = Mid(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") + 1, Len(Trim(rsAStl("jammulai"))) - InStr(Trim(rsAStl("jammulai")), ":"))
                                                          
30                                menitJM = menitJM + 60
                                  
31                                If menitJM >= 60 Then
32                                   jamJM = jamJM + Val(menitJM / 60)
33                                   menitJM = menitJM Mod 60
34                                End If
35                                jamSelesai = Format(jamJM, "0#") & ":" & Format(menitJM, "0#")
36                                objO.InsertData ("UPDATE TEMPSTALL SET JAMSELESAI = '" & jamSelesai & "',ADJUST = '00:30' WHERE tanggal = '" & Format(Trim(Text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "'")
37                            End If
38                        End If
39                    End If
40                End If
41            End If
42        End If

43       On Error GoTo 0
44       Exit Sub

PindahStallSBNP_Error:

45        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PindahStallSBNP of Form frmBooking")
End Sub


Private Sub cmdBatalAdjustDMS_Click()
    frameDSS.Visible = False
End Sub

Private Sub cmdBersih_Click()
Dim nPol1 As String
Dim nPol2 As String
Dim nPol3 As String
    
    nPol1 = Trim(Text(0))
    nPol2 = Trim(Text(56))
    nPol3 = Trim(Text(57))
    Call hapusForm
    Text(0) = nPol1
    Text(56) = nPol2
    Text(57) = nPol3
    
    Text(3).SelLength = Len(Text(3))
    Text(3).SetFocus
End Sub

Private Sub cmdCetak_Click()
      Dim rsSearch As ADODB.Recordset

1        On Error GoTo cmdCetak_Click_Error

2         Set rsSearch = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & Trim(Text(1)) & "' AND DELIV <> 'BTL'")
3         If Not rsSearch.EOF Then
              'Call CetakForm
4             Call CetakForm1
               ''2016-06-02 "" cetak detail FIR
5              If Log.FIR = 1 Then
6                 Set rsTmp = objO.FetchData("SELECT IDKELUHAN FROM DETAILFIR WITH (NOLOCK) WHERE NOFAK = '" & Trim(Text(1)) & "'")
7                 If Not rsTmp.EOF Then
8                    rsTmp.MoveFirst
9                    Do While Not rsTmp.EOF
10                      Call cetakDetailFIR(Trim(rsTmp!idkeluhan))
11                      rsTmp.MoveNext
12                   Loop
13                End If
14             End If
15        Else
16            MsgBox "Booking belum disimpan", vbCritical, "SIMPAN BOOKING"
17            Exit Sub
18        End If

19       On Error GoTo 0
20       Exit Sub

cmdCetak_Click_Error:

21        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCetak_Click of Form frmBooking")
End Sub

Private Sub cmdFIR1_Click()
   If Trim(Text1(0)) = "" Then
      MsgBox "Lengkapi keluhan 1 untuk mengisi FIR 1", vbInformation
      Exit Sub
   Else
      frmDetailFIR.isBooking = True
      frmDetailFIR.NoBooking = Trim(Text(1))
      frmDetailFIR.noChassis = Trim(Text(41))
      frmDetailFIR.Text(14).Text = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) ''no polisi
      frmDetailFIR.Text(16).Text = Trim(Text(23)) ''tanggal
      frmDetailFIR.Text(10).Text = Trim(Text(10)) '' full model
      frmDetailFIR.Text(0).Text = Trim(Text(12))  '' model
      frmDetailFIR.Text1.Text = Trim(Text1(0)) '' keluhan
      frmDetailFIR.Label15.Caption = "1"
      frmDetailFIR.show vbModal
      
   End If
End Sub

Private Sub cmdFIR2_Click()
   If Trim(Text1(1)) = "" Then
      MsgBox "Lengkapi keluhan 2 untuk mengisi FIR 2", vbInformation
      Exit Sub
   Else
      frmDetailFIR.isBooking = True
      frmDetailFIR.NoBooking = Trim(Text(1))
      frmDetailFIR.noChassis = Trim(Text(41))
      frmDetailFIR.Text(14).Text = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) ''no polisi
      frmDetailFIR.Text(16).Text = Trim(Text(23)) ''tanggal
      frmDetailFIR.Text(10).Text = Trim(Text(10)) '' full model
      frmDetailFIR.Text(0).Text = Trim(Text(12))  '' model
      frmDetailFIR.Text1.Text = Trim(Text1(1)) '' keluhan
      frmDetailFIR.Label15.Caption = "2"
      frmDetailFIR.show vbModal
   End If
End Sub

Private Sub cmdFIR3_Click()
   If Trim(Text1(2)) = "" Then
      MsgBox "Lengkapi keluhan 3 untuk mengisi FIR 3", vbInformation
      Exit Sub
   Else
      frmDetailFIR.isBooking = True
      frmDetailFIR.NoBooking = Trim(Text(1))
      frmDetailFIR.noChassis = Trim(Text(41))
      frmDetailFIR.Text(14).Text = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) ''no polisi
      frmDetailFIR.Text(16).Text = Trim(Text(23)) ''tanggal
      frmDetailFIR.Text(10).Text = Trim(Text(10)) '' full model
      frmDetailFIR.Text(0).Text = Trim(Text(12))  '' model
      frmDetailFIR.Text1.Text = Trim(Text1(2)) '' keluhan
      frmDetailFIR.Label15.Caption = "3"
      frmDetailFIR.show vbModal
   End If
End Sub

Private Sub cmdFIR4_Click()
   If Trim(Text1(3)) = "" Then
      MsgBox "Lengkapi keluhan 4 untuk mengisi FIR 4", vbInformation
      Exit Sub
   Else
      frmDetailFIR.isBooking = True
      frmDetailFIR.NoBooking = Trim(Text(1))
      frmDetailFIR.noChassis = Trim(Text(41))
      frmDetailFIR.Text(14).Text = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) ''no polisi
      frmDetailFIR.Text(16).Text = Trim(Text(23)) ''tanggal
      frmDetailFIR.Text(10).Text = Trim(Text(10)) '' full model
      frmDetailFIR.Text(0).Text = Trim(Text(12))  '' model
      frmDetailFIR.Text1.Text = Trim(Text1(3)) '' keluhan
      frmDetailFIR.Label15.Caption = "4"
      frmDetailFIR.show vbModal
   End If
End Sub

Private Sub cmdFIR5_Click()
   If Trim(Text1(4)) = "" Then
      MsgBox "Lengkapi keluhan 5 untuk mengisi FIR 5", vbInformation
      Exit Sub
   Else
      frmDetailFIR.isBooking = True
      frmDetailFIR.NoBooking = Trim(Text(1))
      frmDetailFIR.noChassis = Trim(Text(41))
      frmDetailFIR.Text(14).Text = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) ''no polisi
      frmDetailFIR.Text(16).Text = Trim(Text(23)) ''tanggal
      frmDetailFIR.Text(10).Text = Trim(Text(10)) '' full model
      frmDetailFIR.Text(0).Text = Trim(Text(12))  '' model
      frmDetailFIR.Text1.Text = Trim(Text1(4)) '' keluhan
      frmDetailFIR.Label15.Caption = "5"
      frmDetailFIR.show vbModal
   End If
End Sub

Private Sub cmdHistory_Click()
 
    If gNoRangka = "" Then Exit Sub
    If Log.LogHOConn = 0 Then
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
    gTag = "ONLINE"
    dRangka = gNoRangka
    nData(0) = gNoRangka
    nData(1) = Text(3)
    nData(2) = nopol
'    Dim fH As New FrmHistory_new
'    ''20180928|DEV ''PENAMBAHAN DATA INSPECTION
'    fH.isOrder = False
'     fH.noBaris = baris
'    fH.modelKendaraan = Trim(Text(10))
'    fH.show (vbModal)
    Dim ft As New FormTransparan
    ft.xBaris = baris
    ft.xModelKendaraan = Trim(Text(10))
    ft.xIsForm = "HIST_BOOKING"
    ft.show (vbModal)
End Sub

Private Sub cmdLihatProgram_Click()
    gMainKeyCari = Text(Index)
    gSearchOption = "PROGRAM"
    gBantuan = "AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
    frmCari.show (vbModal)
End Sub

Private Sub cmdPilihDSS_Click()
1        On Error GoTo cmdPilihDSS_Click_Error

2         If gridDSS.Columns(1).value <> "Stall Manual" Then
3       If MsgBox("Apakah yakin memilih " & gridDSS.Columns(1).value & " ?", vbYesNo) = vbYes Then
4           erTranAll = False
5           onTR = True
            
6           Call objHT.CloseTransactionAll_Err
7           Call objHT.OpenTransactionBengkel
8           On Error GoTo errorTrans
            'mencari jam selesai (jammulai + adjust)
            Dim jamSelesai As String
            Dim total As Double
            Dim PilihMekanik As String
            Dim waktuMulai As String
9           PilihMekanik = UCase(gridDSS.Splits(0).Columns(3))
            
10          If Len(PilihMekanik) > 3 Then
            'jika mekanik lebih dari 1
11              PilihMekanik = Left(PilihMekanik, InStr(PilihMekanik, " ,"))
12          End If
            
13          total = cariJumlahMenit(gridDSS.Columns(4).value) + cariJumlahMenit(Trim(Text(122)))
14          Text(64) = PilihMekanik '
            
15          jamSelesai = Format(Val(total / 60), "0#") & ":" & Format((total Mod 60), "0#")
16          waktuMulai = gridDSS.Columns(4).value
17            If Log.newJpcb = 1 Then
18                Set rsTemp = objO.FetchData("SELECT WAKTUMULAI,WAKTUSELESAI,WAKTUMULAIISTIRAHAT,WAKTUSELESAIISTIRAHAT FROM T_STALL WITH(NOLOCK) WHERE hari = '" & IIf(Weekday(Trim(Text(23))) = 1, 6, Weekday(Trim(Text(23))) - 2) & WeekdayName(IIf(Weekday(Trim(Text(23))) = 1, 8, Weekday(Trim(Text(23)))) - 1) & "'")
19                If rsTemp.EOF = False Then
20                    rsTemp.MoveFirst
21                    If cariJumlahMenit(waktuMulai) <= cariJumlahMenit(rsTemp!waktuMulaiIstirahat) Or cariJumlahMenit(waktuMulai) <= cariJumlahMenit(rsTemp!waktuSelesaiIstirahat) Then
                          'If (cariJumlahMenit(rsTemp!waktuSelesaiIstirahat) >= cariJumlahMenit(jamSelesai) And cariJumlahMenit(Trim(rsTemp!waktuMulaiIstirahat)) >= cariJumlahMenit(waktuMulai)) Or cariJumlahMenit(jamSelesai) >= cariJumlahMenit(rsTemp!waktuSelesaiIstirahat) Then
22                        If (cariJumlahMenit(Trim(rsTemp!waktuMulaiIstirahat)) <= cariJumlahMenit(jamSelesai)) Or cariJumlahMenit(jamSelesai) >= cariJumlahMenit(rsTemp!waktuSelesaiIstirahat) Then
23                            total = cariJumlahMenit(gridDSS.Columns(4).value) + cariJumlahMenit(Trim(Text(122))) + (cariJumlahMenit(rsTemp!waktuSelesaiIstirahat) - cariJumlahMenit(rsTemp!waktuMulaiIstirahat))
24                        End If
25                    End If
                      
26                End If
27                jamSelesai = Format(Val(total / 60), "0#") & ":" & Format((total Mod 60), "0#")
28            End If
29          Call saveStallDSS(gridDSS.Columns(1).value, Format(Trim(Text(23)), "dd/mm/yyyy"), Trim(gridDSS.Columns(4).value), Trim(jamSelesai), Trim(txtNRate.Text), "", Trim(Text(64)), Replace(Replace(nopol, ">", ""), " ", "") & " |" & Trim(Text(1)) & "| " & Trim(Text(3)), Trim(Text(122)))
            '            jamAwal = Trim(Text(3))
30          If erTranAll = False Then
31              Call objHT.CloseTransactionBengkel
32              onTR = False
33              sgnStall.Visible = True
34              sgnDSS.Caption = "DSS"
35              sgnStall.Caption = Trim(gridDSS.Columns(1).value)
36          Else
37              Call objHT.CloseTransactionBengkel_err
38              onTR = False
39              MsgBox "Stall tidak berhasil dipilih !", vbInformation + vbOKOnly
            '    Exit Sub
40          End If
41          frameHasilDSS.Visible = False
42          Call hitungEstimasiSelesai
43          Exit Sub
errorTrans:
44          Call objHT.CloseTransactionBengkel_err
45          onTR = False
46          MsgBox "Stall tidak berhasil dipilih !", vbInformation + vbOKOnly
47          Exit Sub
48      End If
49        Else
50            frameHasilDSS.Visible = False
51            sgnDSS.Caption = "NON DSS"
52                    With frmStall
                        
53                      .Text(7) = Text(23)
54                      .Text(14) = Trim(Text(122)) ''Trim(Text(85))
55                      .Text(4).Mask = ""
56                      .Text(4) = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) ' noPol
57                      .Text(5) = Text(2)
58                      .Text(6) = Text(3)
59                      .Text(12) = Text(1)
60                      .Text(9) = IIf(chEM = Checked, 0.5, IIf(chSBNP = Checked, 1, nRate))
61                      .Text(8) = nMech
62                      .chkThs.value = chkThs.value
63                      .chEM.value = IIf(chEM = Checked, chEM.value, IIf(chSBNP = Checked, chSBNP.value, 0))
64                      .frmStal.Visible = False
65                      .Text(13).Enabled = True
66                      If isBookByMrs = False Then
67                          NotChild frmStall.hWnd
68                      Else
69                          TOChild frmBooking.hWnd, frmStall.hWnd
                            '.Show
70                      End If
71                      .WindowState = vbMaximized
                        
72                  End With
73                  frmStall.show
74            Exit Sub
75        End If
        
76      On Error GoTo 0
77      Exit Sub

cmdPilihDSS_Click_Error:
78        frameHasilDSS.Visible = False
79        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdPilihDSS_Click of Form frmBooking")

80       On Error GoTo 0
81       Exit Sub
End Sub

Private Sub cmdPilih_Click()

End Sub

Private Sub cmdSimpanAdjustDSS_Click()
1        On Error GoTo cmdSimpanAdjustDSS_Click_Error
          ''2016-05-17 || Nilai adjust dibatasi minimal 00:10
2           If cariJumlahMenit(Trim(Text(122))) < 10 Then
3              MsgBox "Adjust Minimal 10 menit / 00:10 ", vbInformation
4              Exit Sub
5           End If
            
            ''2016-05-17 || format > 59 menit dan format > 23 jam
6           If CInt(Left(Trim(Text(122)), InStr(1, Trim(Text(122)), ":") - 1)) > 23 Then
7              MsgBox "Format adjust salah, jam tidak boleh melebihi 23 jam ", vbInformation
8              Exit Sub
9           ElseIf CInt(Right(Trim(Text(122)), InStr(1, Trim(Text(122)), ":") - 1)) > 59 Then
10             MsgBox "Format adjust salah, format dalam jam : menit", vbInformation
11             Exit Sub
12          End If
            
13        If Log.isDSS = 1 Then
14            Call pilihStallDSS
15            frameDSS.Visible = False
16        Else 'If Log.dmsNonDss = 1 Then
17            frameDSS.Visible = False
18            With frmStall
19                .Text(13) = Trim(Text(122))
20                .Text(7) = Text(23)
21                .Text(14) = Trim(Text(122)) ''Trim(Text(85))
22                .Text(4).Mask = ""
23                .Text(4) = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) ' noPol
24                .Text(5) = Text(2)
25                .Text(6) = Text(3)
26                .Text(12) = Text(1)
27                .Text(9) = IIf(chEM = Checked, 0.5, IIf(chSBNP = Checked, 1, nRate))
28                .Text(8) = nMech
29                .chkThs.value = chkThs.value
30                .chEM.value = IIf(chEM = Checked, chEM.value, IIf(chSBNP = Checked, chSBNP.value, 0))
31                .frmStal.Visible = False
32                .Text(13).Enabled = True
33                If isBookByMrs = False Then
34                    NotChild frmStall.hWnd
35                Else
36                    TOChild frmBooking.hWnd, frmStall.hWnd
                      '.Show
37                End If
38                .WindowState = vbMaximized
                  
39            End With
40            frmStall.show
41        End If

42       On Error GoTo 0
43       Exit Sub

cmdSimpanAdjustDSS_Click_Error:

44        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpanAdjustDSS_Click of Form frmBooking")
End Sub

Private Sub cmdStalOk_Click()
          
1        On Error GoTo cmdStalOk_Click_Error

2         estimasiSublet = ""
          Dim jJ As Integer
          Dim jM As Integer
          Dim total As Integer
          
3         jJ = CInt(Left(Trim(Text(91)), InStr(Trim(Text(91)), ":") - 1))
4         jM = Mid(Trim(Text(91)), InStr(Trim(Text(91)), ":") + 1, Len(Trim(Text(91))) - InStr(Trim(Text(91)), ":"))
          
5         total = (jJ * 60) + jM
6         If total = 0 Then
7             MsgBox "Estimasi pengerjaan sublet harus diisi", vbCritical
8         ElseIf total > 240 Then
9             MsgBox "Estimasi pengerjaan sublet maksimal 4 Jam", vbCritical
10        Else
11            estimasiSublet = Trim(Text(91))
12            frmEstimasiSublet.Visible = False
              
13            If MaxBrs >= barisSublet Then
14            With subletGrid
15                If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
16                    If NulSID(Text(86)) <> "" Then
                          ''ambil data promo
17                        sSQL = "SELECT D.KDSUPP,E.SUPPLIER,A.KDPROMO,A.NAMAPROMO,A.KETERANGAN,A.HARGAJUAL,D.HARGABELI,A.KDSUBLET,D.NAMASUBLET,D.KETERANGAN AS KETSUBLET,D.KODESATUAN"
18                        sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLETPROMO A WITH (NOLOCK)"
19                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLET D WITH (NOLOCK) ON A.KDSUBLET = D.KDSUBLET"
20                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMOMDL B WITH (NOLOCK) ON A.KDPROMO = B.KDPROMO"
21                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMODEALER C WITH (NOLOCK) ON A.KDPROMO = C.KDPROMO"
22                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUPPLIER E WITH (NOLOCK) ON E.KDSUPP = D.KDSUPP "
23                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL F WITH (NOLOCK) ON D.KDSUBLET = F.KDSUBLET"
24                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER G WITH (NOLOCK) ON D.KDSUBLET = G.KDSUBLET"
25                        sSQL = sSQL & " LEFT JOIN T_SUPPLIER H WITH (NOLOCK) ON H.KDSUPP = E.KDSUPP"
26                        sSQL = sSQL & " WHERE ISNULL(H.KODEAP,'') <> '' AND A.FS_LOG_DEL = '' AND A.KDPROMO = '" & NulSID(Text(86)) & "'"
27                    Else
                          ''ambil data sublet
28                        sSQL = "SELECT DISTINCT '' as KDPROMO,D.KDSUBLET,D.NAMASUBLET,D.KETERANGAN,D.HARGAJUAL,D.HARGABELI,E.KDSUPP,E.SUPPLIER,D.KODESATUAN"
29                        sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLET D "
30                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL B WITH (NOLOCK) ON D.KDSUBLET = B.KDSUBLET"
31                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER C WITH (NOLOCK) ON D.KDSUBLET = C.KDSUBLET"
32                        sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUPPLIER E WITH (NOLOCK) ON D.KDSUPP = E.KDSUPP"
33                        sSQL = sSQL & " WHERE D.KDSUBLET = '" & NulSID(Text(79)) & "'"
34                        sSQL = sSQL & " AND D.FS_LOG_DEL = ''"
35                    End If
36                    Set rsDetailSublet = objO.FetchData(sSQL)
37                    If Not rsDetailSublet.EOF Then
38                        rsDetailSublet.MoveFirst
39                        .TextMatrix(barisSublet, 1) = NulSID(rsDetailSublet!namaSublet)             ' NAMA SUBLET
40                        .TextMatrix(barisSublet, 2) = CDbl(Text(80))                                ' QTY
41                        .TextMatrix(barisSublet, 3) = NulSID(rsDetailSublet!KODESATUAN)             ' SATUAN
42                        .TextMatrix(barisSublet, 4) = Format(CDbl(Text(82)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
43                        .TextMatrix(barisSublet, 5) = IIf(chSublet = Checked, 1, 0)
44                        .TextMatrix(barisSublet, 6) = estimasiSublet
45                        .TextMatrix(barisSublet, 7) = NulSID(rsDetailSublet!KDSUBLET)         ''kodeitem
46                        .TextMatrix(barisSublet, 8) = NulSID(rsDetailSublet!KDPROMO)         ''kodepromo
47                    Else
48                        MsgBox "Master sublet tidak terdaftar" & vbCrLf & "Pastikan Enter untuk kode sublet / promo", vbInformation
49                        Exit Sub
50                    End If
51                Else
52                    .TextMatrix(barisSublet, 1) = Trim(Text(79))                                ' NAMA SUBLET
53                    .TextMatrix(barisSublet, 2) = CDbl(Text(80))                                ' QTY
54                    .TextMatrix(barisSublet, 3) = Trim(Text(81))                                ' SATUAN
55                    .TextMatrix(barisSublet, 4) = Format(CDbl(Text(82)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
56                    .TextMatrix(barisSublet, 5) = IIf(chSublet = Checked, 1, 0)
57                    .TextMatrix(barisSublet, 6) = estimasiSublet
58                End If
59            End With
60            barisSublet = barisSublet + 1
61            Text(52) = Format(CDbl(Text(52)) + (CDbl(Text(80)) * CDbl(Text(82))), "#0,#")
62            Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
              
63            Text(79) = ""
64            Text(80) = 0
65            Text(81) = ""
66            Text(82) = 0
67            chSublet = Unchecked
68            Call hitungEstimasiSelesai
69            Text(79).SetFocus
70        End If
71        End If

72       On Error GoTo 0
73       Exit Sub

cmdStalOk_Click_Error:

74        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdStalOk_Click of Form frmBooking")
          
End Sub

Private Sub Command1_Click()
1        On Error GoTo Command1_Click_Error

2     If Replace(Replace(Text(15), ">", ""), " ", "") = "" Then Exit Sub
3         gTag = "LOCAL"
4         dRangka = gNoRangka
5         nData(0) = gNoRangka
6         nData(1) = Text(3)
7         nData(2) = nopol
          
         ' fH.Show (vbModal)
8         If Log.dmsNonDss = "1" Then
              Dim fHDND As New FrmHistory_newDMSNonDSS
9             fHDND.isOrder = False
10            fHDND.noBaris = baris
11            fHDND.modelKendaraan = Trim(Text(10))
12            fHDND.show (vbModal)
13        Else
              Dim fH As New FrmHistory_new
14            fH.show (vbModal)
15        End If

16       On Error GoTo 0
17       Exit Sub

Command1_Click_Error:

18        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Command1_Click of Form frmBooking")
End Sub

Private Sub cmdHO_Click()
Dim rsHlp As ADODB.Recordset
    If nopol <> "" Then
        If Log.LogHOConn = 1 Then
            If ConHO = True Then
                Set rsHlp = objO.FetchDataHO("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_POLISI = '" & nopol & "'")
                If Not rsHlp.EOF Then
                    Call panggilCustHO(nopol, rsHlp)
                    Call panggilHistoryHO(rsHlp("NO_CHASSIS"))
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdNo_Click()
    If cmdNo.Caption = "No &Polisi" Then
        Call hapusForm
        cmdNo.Caption = "No &Rekom"
    ElseIf cmdNo.Caption = "No &Rekom" Then
        Call hapusForm
        cmdNo.Caption = "No &Polisi"
    End If
End Sub

Private Sub cmdNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub cmdPlgn_Click()
'    With frmtPlgn
'        .Text(57) = Text(2)
'        .Text(1) = Text(3)
'        .Text(2) = Left(Text(4), 40)
'        .Text(3) = Text(5)
'        .Text(4) = Text(8) + "," + Trim(Text(6)) + "," + Trim(Text(7))
'        gBantuan2 = "kosong"
'        frmtPlgn.Show vbModal, Me
'    End With
'    Dim Nama() As String
'
'    With frmSIDMasterPelanggan
'        .setKodeLgnLama = Nul(txtKodeLgnLama.Text)
'
'        Nama() = Split(Nul(Text(3)), " ")
'        For i = 0 To UBound(Nama)
'            If Nama(i) <> "" And i = 0 Then
'                .txt(4) = Nul(Nama(i))
'            ElseIf Nama(i) <> "" And i = 1 Then
'                .txt(5) = Nul(Nama(i))
'            Else
'                .txt(6) = .txt(6) & Nul(Nama(i))
'            End If
'        Next
'
'        .txt(13) = Nul(Text(8))
'        .txt(14) = Nul(Text(6))
'        .txt(15) = Nul(Text(7))
'
'        .txt(8) = Nul(Text(4))
'
'        .Show vbModal
'    End With
    Call panggilCari
    If gMainKeyCari <> "" Then
    Text(2) = gMainKeyCari
    End If
End Sub

Private Sub panggilCari()
    gMainKeyCari = Nul(Text(41))
     
    With frmSIDSearchCust
        .Text(1) = gMainKeyCari
        .Text(4) = Nul(Text(2))
        gSearchOption = "frmSIDCustomer"
        .Left = Me.asxPowerBanner1.Width '+ 100
        .Top = Me.Top + 450
        .show vbModal
    End With
End Sub

Private Sub cmdSettingPrinter_Click()
    rpt.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdStall_Click()
          
          Dim pesan As String
          Dim boleh As Boolean
1         If chkThs.value = vbChecked Then
2             MsgBox "Unit THS tidak perlu menggunakan stall ", vbExclamation
3             Exit Sub
4         End If
5         nopol = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57))
6         If nMech = "" Then
7             nMech = Trim(Text(64))
8         End If
9         boleh = True
10        nRate = 0
11        job = ""
12        jobQuery = ""
13        For i = 1 To baris - 1
14            If Trim(dGrid.TextMatrix(i, 1)) <> "" Then
                  
                  'If Log.RPLHistori = 0 Then
15                    Set rsTemp = objO.FetchData("SELECT RATE FROM actRAte WITH (NOLOCK) WHERE MDL = '" & Text(10) & "' AND FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
                  'Else
                  '    Set rsTemp = objO.FetchData("SELECT RATE FROM [" & dbReplikasiHO & "]..actRAte WITH (NOLOCK) WHERE MDL = '" & text(10) & "' AND FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
                  'End If
16                If Not rsTemp.EOF Then
17                    nRate = nRate + CDbl(rsTemp("RATE")) 'DGrid.TextMatrix(i, 2)
18                    job = job & "'" & Trim(dGrid.TextMatrix(i, 1)) & "',"
19                    jobQuery = jobQuery & " B.FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND "
20                Else
21                    boleh = False
22                    pesan = pesan + vbCr & "Actual Rate untuk " & Trim(dGrid.TextMatrix(i, 1)) & " belum ada "
23                    Exit For
24                End If
25            End If
26        Next
27        If NulSID(jobQuery) <> "" Then
28            jobQuery = Left(jobQuery, Len(jobQuery) - 4)
29        End If
          ' HITUNG NJRATE
30        If nJRate = 0 Then
31            For i = 1 To baris - 1
32                If Trim(dGrid.TextMatrix(i, 1)) <> "" Then
33                    nJRate = nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(Text(10)) & "'")
34                End If
35            Next
36        End If
          
37        If baris = 1 Then
38            pesan = pesan + "Job Order. "
39            boleh = False
40        Else
41            If nRate = 0 Then
42                pesan = pesan + "Nilai Actual Rate 0. "
43                boleh = False
44            End If
45        End If
46        If Trim(nopol) = "" Then
47            nopol = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57))
48            If Trim(nopol) = "" Then
49                pesan = pesan + "No Polisi. "
50                boleh = False
51            End If
52        End If
53        If Trim(Text(3)) = "" Then
54            pesan = pesan + "Nama Customer. "
55            boleh = False
56        End If
57        If Trim(Text(23)) = ">>->>->>>>" Then
58            pesan = pesan + "Tanggal Booking. "
59            boleh = False
60        End If
          
61        If boleh Then
62            If chEM = Checked Then
63                Call PindahStallEM
64                Text(85) = "00:30"
65            End If
66            If chSBNP = Checked Then
67                Call PindahStallSBNP
68                Text(85) = "01:00"
69            End If
              ' ISI TEMPSTALL
              
70            Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                           & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%') AND " _
                           & "tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%' ORDER BY noStall, Tanggal, JamMulai")
              
71            gStall = "booking"
72            gTanggalStall = Trim(Text(23))
73            job = Left(job, Len(job) - 1)
74            frameDSS.Visible = False
              
75            If Log.isEmStall = 1 Then
              '========---- TAMPILKAN ESTIMASI SERVICE ----------========
76                frameDSS.Visible = True
77                progDSS.Visible = False
78                cmdSimpanAdjustDSS.Enabled = True
79                cmdBatalAdjustDMS.Enabled = True
80                txtNRate.Text = nRate
81                adRate = "00:00"
82                If Trim(nRate) <> 0 Then
      '                If InStr(nRate, ",") > 0 Then
      '                    Text(122) = Format(Left(nRate, InStr(nRate, ",") - 1), "0#") & ":" & Format(Val((CDbl(nRate) - Left(nRate, InStr(nRate, ",") - 1)) * 60), "0#")
      '                Else
      '                    Text(122) = Format(nRate, "0#") & ":00"
      '                End If
83                    If chEM.value = vbChecked Then
84                        Text(122) = "00:30"
85                    ElseIf chSBNP = vbChecked Then
86                        Text(122) = "01:00"
87                    Else
88                        If InStr(nRate, ",") > 0 Then
89                            Text(122) = Format(Left(nRate, InStr(nRate, ",") - 1), "0#") & ":" & Format(Val((CDbl(nRate) - Left(nRate, InStr(nRate, ",") - 1)) * 60), "0#")
90                        Else
91                            Text(122) = Format(nRate, "0#") & ":00"
92                        End If
93                    End If
94                    adRate = Trim(Text(122))
95                End If

96            Else
97                With frmStallNonEM
                      '.Text(13).Enabled = True
98                    .Text(7) = Text(23)
99                    .Text(4).Mask = ""
100                   .Text(4) = nopol
101                   .Text(5) = Text(2)
102                   .Text(6) = Text(3)
103                   .Text(12) = Text(1)
104                   .Text(9) = IIf(chEM = Checked, 0.5, IIf(chSBNP = Checked, 1, nRate))
105                   .Text(8) = nMech
106                   .chkThs.value = chkThs.value
107                   .chEM.value = IIf(chEM = Checked, chEM.value, IIf(chSBNP = Checked, chSBNP.value, 0))
108                   .frmStal.Visible = False
109                   If isBookByMrs = False Then
110                       NotChild frmStallNonEM.hWnd
111                   Else
112                       TOChild frmBooking.hWnd, frmStallNonEM.hWnd
                          '.Show
113                   End If
114                   .WindowState = vbMaximized
                      'frmStall.Show
115               End With
116           End If
117       Else
118           MsgBox "Lengkapi data Booking !!" & vbCr & pesan, vbCritical, "Isi Field Kosong"
119           Exit Sub
120       End If
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub cmdWaktu_Click()
    If frmeWaktu.Visible Then
        frmeWaktu.Visible = False
    Else
        If Text(23) <> "" And Text(23) <> ">>->>->>>>" Then
            frmeWaktu.Visible = True
            frmeWaktu.Top = cmdWaktu.Top + cmdWaktu.Height + 100
            frmeWaktu.Left = cmdWaktu.Left - (frmeWaktu.Width - cmdWaktu.Width)
            If Log.NHS = 1 And chkThs.value = 1 Then
                Call TampilWaktuNHS(Trim(Text(23)))
            Else
                Call TampilWaktu(Trim(Text(23)))
            End If
        Else
            MsgBox "Tanggal booking harus diisi"
            Text(23).SelLength = Len(Text(23))
            Text(23).SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdWarranty_Click()
    If Text(23) = ">>->>->>>>" And Text(24) = ">>:>>" Then
        MsgBox "Isi tanggal & jam Booking !"
        Text(23).SetFocus
        Exit Sub
    Else
        gTag = "BOOKING"
        With frmWarranty
            .Text(0) = gNoRangka
            .Text(1) = Text(3)
            .Text(2) = nopol
            .Brs.Caption = baris
            .BrsPart.Caption = barispart
            .Text(3) = Trim(Text(23))
            .Text(4) = Trim(Text(24))
            .show (vbModal)
        End With
    End If
End Sub


Private Sub Command2_Click()
    frameHasilDSS.Visible = False
End Sub

Private Sub Command3_Click()
    If Log.DIDX = 1 Then
        If Frame9.Left = 100 Then
            Frame9.Left = 5640
        Else
            Frame9.Left = 100
        End If
        
        
    End If
End Sub

Private Sub DGrid_DblClick()
1         If baris > 1 And dGrid.TextMatrix(dGrid.row, 1) <> "" Then
2             a = dGrid.row
              
3             If Trim(Text(1)) <> "" Then
                  Dim nCexPP As String
                  Dim rsPP As ADODB.Recordset
                  'Set rsPP = objO.FetchData("Select Item_Part from HO_PartPack Where MDL = '" & Text(10) & "' and FS_Item = '" & dGrid.TextMatrix(a, 1) & "' and JOB = '" & dGrid.TextMatrix(a, 4) & "'")
                  
4                 If Log.RPLHistori = 0 Then
5                     Set rsPP = objO.FetchData("Select Nama_Part from HO_PartPack WITH (NOLOCK) Where MDL = '" & Text(10) & "' and FS_Item = '" & dGrid.TextMatrix(a, 1) & "' and JOB = '" & dGrid.TextMatrix(a, 4) & "'")
6                 Else
7                     Set rsPP = objO.FetchData("Select Nama_Part from [" & dbReplikasiHO & "]..HO_PartPack WITH (NOLOCK) Where MDL = '" & Text(10) & "' and FS_Item = '" & dGrid.TextMatrix(a, 1) & "' and JOB = '" & dGrid.TextMatrix(a, 4) & "'")
8                 End If
9                 If rsPP.RecordCount > 0 Then
10                    nCexPP = "Select NOSO FROM SOORD WITH (NOLOCK) Where NoSPK = '" & Trim(Text(1)) & "' and ("
11                    rsPP.MoveFirst
12                    Do Until rsPP.EOF
                          'nCexPP = nCexPP & " KodeItem like '%" & rsPP.Fields(0).value & "' OR "
13                        nCexPP = nCexPP & " NamaBySA = '" & rsPP.Fields(0).value & "' OR "
14                        rsPP.MoveNext
15                    Loop
16                    nCexPP = Left(nCexPP, Len(nCexPP) - 3) & ")"
17                    nCexPP = nCexPP & " and ((IsNull(QtyKirimBL,0) + IsNull(QtyKirimBI,0) + IsNull(QtyKeluar,0)) <> 0 Or IsNull(QtyReserve,0) <> 0 ) AND ([TO] <> 0 Or BO not in ('','CANCEL') )"
                      Dim rsPSO As ADODB.Recordset
18                    Set rsPSO = objO.FetchDataACC(nCexPP)
19                    If rsPSO.RecordCount > 0 Then
20                        MsgBox "Paket Part sudah di proses oleh partman !," & Chr(13) & "Job Order tidak dapat di hapus" & Chr(13) & Chr(13) & "Lakukan pembatalan proses part oleh Partman, untuk menghapus Job Order"
21                        Exit Sub
22                    End If
23                End If

24            End If
              
              'AMBIL
25            With dGrid
26                Text(34) = .TextMatrix(a, 1)
27                Jrate = .TextMatrix(a, 2)
28                Text(32) = .TextMatrix(a, 4)
      '            Text(50) = .TextMatrix(a, 5)
29                If .TextMatrix(a, 3) <> "" Then
30                    If .TextMatrix(a, 3) Then
31                        fOther = True
32                        Text(42).Visible = True
33                        Text(42).Mask = ""
34                        Text(42) = Jrate
35                    Else
36                        fOther = False
37                        Text(42).Visible = False
38                    End If
39                Else
40                    fOther = False
41                    Text(42).Visible = False
42                End If
43            End With
44            myChassis = Trim(Text(41))
45            If Log.freepart = 1 And tagihanFOP <> "" Then
46                Call cekFopPart(Text(32), Text(34), Text(10), 2, 1, barispart, frmBooking)
47            Else
48                Call cekPartPack(Text(32), Text(34), Text(10), 2, 1, barispart, frmBooking)
49            End If
              


50            If Log.freepart = 1 Then Call hapusSubletFOP(NulSID(Text(10)), Text(34))
              
              ''Call cekPartPack(DGrid.TextMatrix(a, 4), DGrid.TextMatrix(a, 1), Text(10), 2, 1, barispart, frmBooking)
              'hrgLabour = hitungLabour(DGrid.TextMatrix(baris - 1, 2), Text(11))
51            hrgLabour = hitungLabour(dGrid.TextMatrix(a, 2), Text(11))
52            If isJualNRM = True Then
53                Me.JualNRM.Caption = nJualNRM
54            Else
55                Me.JualNRM.Caption = ""
56            End If
57            nJRate = nJRate - Jrate
58            If Text(58) <> ">>:>>" And Text(23) <> ">>->>->>>>" Then
59                nWKT = hitungWaktu(CDbl(nJRate), Text(58), Text(23))
60                Text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
61                Text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
62                Text(84) = Trim(Text(59))
63                Text(83) = Trim(Text(45))
                  
64                Call hitungEstimasiSelesai
65            End If
              
              'HAPUS
66            For i = a To baris
67                With dGrid
68                    .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
69                    .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
70                    .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
71                    .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
72                    .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
73                End With
74            Next
              
75            Text(39) = Format(CDbl(Text(39) - (CDbl(hrgLabour))), "#0,#")
76            Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
77            baris = baris - 1
              
78            jobQuery = ""
79            sgnDSS.Caption = ""
80            If baris > 1 Then
81              For i = 1 To baris - 1
82                  If Trim(dGrid.TextMatrix(i, 1)) <> "" Then
                        
                        'If Log.RPLHistori = 0 Then
83                        Set rsTemp = objO.FetchData("SELECT RATE FROM actRAte WITH (NOLOCK) WHERE MDL = '" & Text(10) & "' AND FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
                        'Else
                        '  Set rsTemp = objO.FetchData("SELECT RATE FROM [" & dbReplikasiHO & "]..actRAte WITH (NOLOCK) WHERE MDL = '" & text(10) & "' AND FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
                        'End If
84                      If Not rsTemp.EOF Then
85                          jobQuery = jobQuery & " B.FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND "
86                      End If
87                  End If
88              Next
                
89              If NulSID(jobQuery) <> "" Then
90                jobQuery = Left(jobQuery, Len(jobQuery) - 4)
91              End If
                
92              If Log.isDSS = 1 And NulSID(jobQuery) <> "" Then
                  Dim strSql As String
                  
                  
93                strSql = "SELECT B.TIPEBP, C.SKILLMIN, E.LEVEL AS LvlTech, C.STALLMIN, D.level AS LvlStall, C.LiftNeed " & _
                  "FROM  HO_Flate B " & _
                  "LEFT JOIN MASTERDSS C ON B.TIPEBP = C.KODEDSS " & _
                  "LEFT JOIN LevelStall D ON C.StallMin = D.KodeStall " & _
                  "LEFT JOIN T_MECHLEV E ON C.SkillMin = E.LEVELMECH " & _
                  "WHERE B.MDL = '" & Trim(Text(10)) & "' AND (" & jobQuery & ") AND ISNULL(C.SkillMin,'') <> '' AND dss = 1 " & _
                  "order by LvlStall desc ,LvlTech desc "
                  '-------------------------------------------------
                  
      '            Set rsTemp = objO.FetchData(strSql)
      '
      '            If rsTemp.EOF Then
      '              sgnDSS.Caption = "NON DSS"
      '            Else
      '              sgnDSS.Caption = "DSS"
      '            End If
94              End If
                
95           End If
96        End If
End Sub

Private Sub dGrid_GotFocus()
          Dim bolehTambah As Boolean

1        On Error GoTo dGrid_GotFocus_Error

2         If Log.isDSS = 1 Then
3             If Trim(Text(34)) = "" And Text(23) = ">>->>->>>>" And Text(24) = ">>->>->>>>" Then
4                 Exit Sub
5             End If
6         Else
7             If Trim(Text(34)) = "" And Text(23) = ">>->>->>>>" And Text(24) = ">>->>->>>>" And Text(64) = "" Then
8                 Exit Sub
9             End If
10        End If
          
11            For i = 0 To baris
12                If dGrid.TextMatrix(i, 1) = Trim(Text(34)) Then
13                    bolehTambah = False
14                    Text(34) = ""
15                    Exit Sub
16                Else
17                    bolehTambah = True
18                End If
19            Next
              
20            nMech = Trim(Text(64))
21            If Trim(Text(32)) = "SBE" Or Trim(Text(32)) = "SBI" Then
                      '' banding last sbe service dengan yg baru masuk
                      Dim rsSearch As ADODB.Recordset
                      Dim lastService As Double
22                    If Log.RPLHistori = 0 Then
23                        Set rsSearch = objO.FetchData("SELECT FS_NILAI FROM T_REGULAR WITH (NOLOCK) WHERE FS_KETERANGAN = '" & Mid(Text(34), InStr(Text(34), "KALA ") + 5, InStr(Text(34), "KM") - (InStr(Text(34), "KALA ") + 5)) & "'")
24                    Else
25                        Set rsSearch = objO.FetchData("SELECT FS_NILAI FROM [" & dbReplikasiHO & "]..T_REGULAR WITH (NOLOCK) WHERE FS_KETERANGAN = '" & Mid(Text(34), InStr(Text(34), "KALA ") + 5, InStr(Text(34), "KM") - (InStr(Text(34), "KALA ") + 5)) & "'")
26                    End If
27                    If Not rsSearch.EOF Then
                          
28                        If Log.RPLHistori = 0 Then
29                            lastService = objO.FetchData_Double("SELECT TOP 1 FS_NILAI FROM HISTORY A WITH (NOLOCK) INNER JOIN T_REGULAR B WITH (NOLOCK) ON A.FS_RC = B.FS_RC WHERE NO_CHASSIS = '" & Trim(Text(41)) & "' AND a.FS_RC <> '' AND FS_LOG_DEL = '' ORDER BY DATE DESC")
30                        Else
31                            lastService = objO.FetchData_Double("SELECT TOP 1 FS_NILAI FROM HISTORY A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..T_REGULAR B WITH (NOLOCK) ON A.FS_RC = B.FS_RC WHERE NO_CHASSIS = '" & Trim(Text(41)) & "' AND a.FS_RC <> '' AND FS_LOG_DEL = '' ORDER BY DATE DESC")
32                        End If
33                       If CDbl(lastService) >= CDbl(rsSearch("FS_NILAI")) Then
34                           MsgBox "Service berkala yang dipilih tidak boleh lebih kecil atau sama dengan History sebelumnya !!", vbInformation, "Salah SBE"
35                           bolehTambah = False
36                           Text(34).SetFocus
37                           Exit Sub
38                       Else
39                           bolehTambah = True
40                       End If
41                   End If
42               End If
43            If bolehTambah Then
                  
                  ' JIKA PEKERJAAN INPUT SENDIRI
44                If fOther Then
45                    Jrate = CDbl(Replace(Replace(Text(42), " ", ""), ">", ""))
46                End If
                  
47                If MaxBrs >= baris Then
48                    With dGrid
49                        .TextMatrix(baris, 1) = Text(34)
50                        .TextMatrix(baris, 2) = Jrate
51                        .TextMatrix(baris, 3) = fOther
52                        .TextMatrix(baris, 4) = Text(32)
53                    End With
54                    baris = baris + 1
                      
                      ' ISI PART PACK
55                    If Log.freepart = 1 And tagihanFOP <> "" Then
56                        Call cekFopPart(Text(32), Text(34), Text(10), 1, 1, barispart, frmBooking)
57                    Else
58                        Call cekPartPack(Text(32), Text(34), Text(10), 1, 1, barispart, frmBooking)
59                    End If
                      
60                    If Log.freepart = 1 Then Call cekSubletFOP(NulSID(Text(10)), Text(34))

61                    myChassis = Trim(Text(41))
62                    Text(39) = Format(CDbl(Text(39)) + Val(hitungLabour(CDbl(Jrate), Trim(Text(11)))), "#0,0")
                      
                     ' If Log.RPLHistori = 0 Then
63                        nJRate = nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(dGrid.TextMatrix(baris - 1, 1)) & "' AND MDL = '" & Trim(Text(10)) & "'")
                     ' Else
                      '    nJRate = nJRate + objO.FetchData_Double("select Rate from [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(dGrid.TextMatrix(baris - 1, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                      'End If
64                    If Text(58) <> ">>:>>" And Text(23) <> ">>->>->>>>" Then
65                        nWKT = hitungWaktu(CDbl(nJRate), Text(58), Text(23))
66                        Text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
67                        Text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
68                        Text(84) = Trim(Text(59))
69                        Text(83) = Trim(Text(45))
70                        Call hitungEstimasiSelesai
71                    End If
72                    If isJualNRM = True Then
73                        Me.JualNRM.Caption = nJualNRM
74                    Else
75                        Me.JualNRM.Caption = ""
76                    End If
77                    Text(34) = ""
78                    Text(32) = ""
79                    Text(42) = ">>,>>"
80                    Text(42).Visible = False
81                    Text(32).SetFocus
82                    hrgLabour = Val(hitungLabour(CDbl(Jrate), Trim(Text(11)))) / IIf(CDbl(Jrate) = 0, 1, CDbl(Jrate))
83                    If isJualNRM = True Then
84                        Me.JualNRM.Caption = nJualNRM
85                    Else
86                        Me.JualNRM.Caption = ""
87                    End If
88                    Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
89                    jobQuery = ""
90                    For i = 1 To baris - 1
91                        If Trim(dGrid.TextMatrix(i, 1)) <> "" Then
                              
                             ' If Log.RPLHistori = 0 Then
92                                Set rsTemp = objO.FetchData("SELECT RATE FROM actRAte WITH (NOLOCK) WHERE MDL = '" & Text(10) & "' AND FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
                             ' Else
                             '     Set rsTemp = objO.FetchData("SELECT RATE FROM [" & dbReplikasiHO & "]..actRAte WITH (NOLOCK) WHERE MDL = '" & text(10) & "' AND FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
                             ' End If
93                            If Not rsTemp.EOF Then
94                                jobQuery = jobQuery & " B.FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND "
95                            End If
96                        End If
97                    Next
                      
98                    If NulSID(jobQuery) <> "" Then
99                        jobQuery = Left(jobQuery, Len(jobQuery) - 4)
100                   End If
101                   If Log.isDSS = 1 And NulSID(jobQuery) <> "" Then
                          Dim strSql As String
                          
102                       strSql = "SELECT B.TIPEBP, C.SKILLMIN, E.LEVEL AS LvlTech, C.STALLMIN, D.level AS LvlStall, C.LiftNeed " & _
                            "FROM  HO_Flate B " & _
                            "LEFT JOIN MASTERDSS C ON B.TIPEBP = C.KODEDSS " & _
                            "LEFT JOIN LevelStall D ON C.StallMin = D.KodeStall " & _
                            "LEFT JOIN T_MECHLEV E ON C.SkillMin = E.LEVELMECH " & _
                            "WHERE B.MDL = '" & Trim(Text(10)) & "' AND (" & jobQuery & ") AND ISNULL(C.SkillMin,'') <> '' AND dss = 1 " & _
                            "order by LvlStall desc ,LvlTech desc "
                    '-------------------------------------------------
          
      '                    Set rsTemp = objO.FetchData(strSql)
      '                    sgnDSS.Caption = ""
      '                    If rsTemp.EOF Then
      '                        sgnDSS.Caption = "NON DSS"
      '                    Else
      '                        sgnDSS.Caption = "DSS"
      '                    End If
103                   End If
104               Else
105                   MsgBox "Maksimal " & str(MaxBrs) & " !"
106                   Text(34).SetFocus
107               End If
108           End If

109      On Error GoTo 0
110      Exit Sub

dGrid_GotFocus_Error:

111       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure dGrid_GotFocus of Form frmBooking")
End Sub

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call DGrid_DblClick
    End If
    If KeyCode = vbKeyReturn Then
        Call DGrid_DblClick
    End If
End Sub

Private Sub Form_Activate()
      Dim nDpn As String
      Dim nTgh As String
      Dim nBlg As String
1     nDpn = ""
2     nTgh = ""
3     nBlg = ""

4         If isBookingOnline And isLoadBooking = False Then
5             Call pecahNoPol(noPolBookingOnline, nDpn, nTgh, nBlg)
              
6             Text(6) = noTelpBookingOnline
7             Text(3) = namaCustomerBookingOnline
8             Text(5) = kotaBookingOnline
              
9             Text(23) = TglDDMMYYYY(CDate(tanggalBookingOnline)) 'Format(tanggalBookingOnline, "")
10            Text(24) = Left(jamBookingOnline, 5)
              
11            Text(27) = Left(permintaanOnline, 40)
12            If Len(permintaanOnline) > 40 Then
13                Text(28) = Mid(permintaanOnline, 41, 40)
14                If Len(permintaanOnline) > 80 Then
15                    Text(29) = Mid(permintaanOnline, 81, 40)
16                End If
17            End If
              
18            Text1(0) = Left(keluhanOnline, 200)
19            lblModel.Caption = Trim(mdlKendaraanOnline)
              
20            If NulSID(noBookingAppOnline) <> "" Then
21                lblKeteranganBawah.Caption = "BOOKING ONLINE APPS, NOPOLISI = " & noPolBookingOnline & " ATAS NAMA " & namaCustomerBookingOnline
                  ''cek jika dari didx
22                If Log.DiDxP1 = 1 Then
23                    sSQL = "SELECT a.nobookingapps, a.service_category + '-' + a.job_type as PermintaanCustomer,"
24                    sSQL = sSQL & " a.vehicle_problem keluhan_Customer, 'Accesories : ' + b.DaftarProduk Accesories from [" & dbReplikasiHO & "]..bookingonline_didx_toSend a Inner join"
25                    sSQL = sSQL & " (      SELECT"
26                    sSQL = sSQL & "     P1.booking_ID,"
27                    sSQL = sSQL & "     STUFF(("
28                    sSQL = sSQL & "         SELECT ', ' + P2.part_name"
29                    sSQL = sSQL & "         FROM [" & dbReplikasiHO & "]..bookingonline_part_didx P2"
30                    sSQL = sSQL & "         where P1.booking_ID = P2.booking_ID"
31                    sSQL = sSQL & "         FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '') AS DaftarProduk"
32                    sSQL = sSQL & "        FROM [" & dbReplikasiHO & "]..bookingonline_part_didx P1"
33                    sSQL = sSQL & "        GROUP BY P1.booking_ID) b on a.booking_ID = b.booking_ID"
34                    sSQL = sSQL & " where a.nobookingapps = '" & NulSID(noBookingAppOnline) & "'"
35                    Set rsTemp = objO.FetchData(sSQL)
36                    If Not rsTemp.EOF Then
37                        rsTemp.MoveFirst
38                        Text(27) = Left(Replace(NulSID(rsTemp!PermintaanCustomer), "-", " "), 40)
39                        Text(28) = Left(NulSID(rsTemp!keluhan_Customer), 40)
40                        Text(29) = Left(NulSID(rsTemp!Accesories), 40)
41                    End If
42                End If
43            Else
44                lblKeteranganBawah.Caption = "BOOKING ONLINE, NOPOLISI = " & noPolBookingOnline & " ATAS NAMA " & namaCustomerBookingOnline
45            End If
46            isLoadBooking = True
              
47            Text(0) = nDpn
48            Text(56) = nTgh
49            Text(57) = nBlg
              
              '' =================================================
              '' 2018-04-23 penambahan nobookingapp dan kodepromo
              '' =================================================
50            txtNoBookingApp.Text = noBookingAppOnline
51            txtKodeProgramApp.Text = kodePromoAppOnline
              
52            Sendkeys ("{TAB}")
53        End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub Form_Load()
          
          ''2016-05-24 """ detail fir
1         If Log.FIR = 1 Then
2           cmdFIR1.Visible = True
3           cmdFIR2.Visible = True
4           cmdFIR3.Visible = True
5           cmdFIR4.Visible = True
6           cmdFIR5.Visible = True
7         Else
8           cmdFIR1.Visible = False
9           cmdFIR2.Visible = False
10          cmdFIR3.Visible = False
11          cmdFIR4.Visible = False
12          cmdFIR5.Visible = False
13        End If
14        isBookDiag = False
15        bookAktif = True
16        baris = 1
17        dGrid.Rows = MaxBrs
18        dGrid.cols = 5
19        dGrid.ColWidth(0) = 0
20        dGrid.ColWidth(1) = 3500
21        dGrid.ColWidth(2) = 800
22        dGrid.ColWidth(3) = 0
23        dGrid.ColWidth(4) = 0
24        dGrid.ColAlignment(1) = flexAlignLeftCenter
25        dGrid.TextMatrix(0, 1) = "Order Booking"
          
26        barispart = 1
27        partGrid.Rows = MaxBrsPart
28        partGrid.cols = 9
29        partGrid.ColWidth(0) = 0
30        partGrid.ColWidth(1) = 1500
31        partGrid.ColWidth(2) = 2500
32        partGrid.ColWidth(3) = 0
33        partGrid.ColWidth(4) = 800
34        partGrid.ColWidth(5) = 0
35        partGrid.ColWidth(6) = 0
36        partGrid.ColWidth(7) = 2500
37        partGrid.ColWidth(8) = 0
38        partGrid.ColAlignment(1) = flexAlignLeftCenter
39        partGrid.TextMatrix(0, 1) = "Kode Part"
40        partGrid.TextMatrix(0, 2) = "Nama Part"
41        partGrid.TextMatrix(0, 4) = "Qty"
42        partGrid.TextMatrix(0, 7) = "Nama By PIC"
          
          ' ATUR GRID SUBLET
43        barisSublet = 1
44        subletGrid.Rows = 31
45        If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
46            subletGrid.cols = 8
47        Else
48            subletGrid.cols = 6
49        End If
50        subletGrid.Clear
51        If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
52            subletGrid.FormatString = " | Sublet | Qty | Satuan | Harga | Part | Est Waktu | Kode Item | Kode Promo"
53        Else
54            subletGrid.FormatString = " | Sublet | Qty | Satuan | Harga | Part | Est Waktu"
55        End If
56        subletGrid.ColWidth(0) = 0
57        subletGrid.ColWidth(1) = 2500
58        subletGrid.ColWidth(2) = 600
59        subletGrid.ColWidth(3) = 1000
60        subletGrid.ColWidth(4) = 1200
61        subletGrid.ColWidth(5) = 0
62        subletGrid.ColWidth(6) = 0
63        If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
64            subletGrid.ColWidth(7) = 0
65            subletGrid.ColWidth(8) = 0
66            Text(81).BackColor = &HC0E0FF
67            Text(82).BackColor = &HC0E0FF
68        End If
69        subletGrid.TextMatrix(0, 1) = "Order Sublet"
          
70        Call hapusForm
71        If Layar.Tag = "estBook" Then
72            baris = gBantuan
73            barispart = gBantuan2
74            barisSublet = gBantuan4
75            nopol = gBantuan3
76        End If
77        Layar.Tag = ""
          
78        Text(0).Mask = "??"
79        Text(56).Mask = "9999"
80        Text(57).Mask = "CCC"
81        Text(26) = gSA
          
          ' ISI COMBO KARAKTER
82        cboPelanggan.Clear
83        Set rsTemp = objO.FetchData("SELECT * FROM T_Karakter WITH (NOLOCK) ORDER BY kode")
84        If Not rsTemp.EOF Then
85            Do While Not rsTemp.EOF
86                cboPelanggan.AddItem rsTemp("KODE")
87                rsTemp.MoveNext
88            Loop
89        End If
90        cboPelanggan.ListIndex = 0
91        If cmdNo.Caption = "No &Polisi" Then cmdBersih.Visible = False
          
92        simpan = True
93        nX = False
          
94        frameHasilDSS.Visible = False
95        frameHasilDSS.Top = 54000
          
          'If Log.isEmStall = 0 Or Log.isDSS = 1 Then
96            Label60.Visible = False
97            Text(85).Visible = False
         'Else
            '  Label60.Visible = True
             ' Text(85).Visible = True
         ' End If
98        NewId = (objO.FetchData_DoubleACC("select top 1 Nomor from SOORD WITH (NOLOCK) Where NOSO = '' Order By Nomor DESC")) + 1
99        If Log.isDSS = 1 Then
100           Text(64).TabIndex = 38
101           Label52.Visible = False
102           Text(64).Visible = False
103           MaskEdBox12.Visible = False
104           sgnDSS.Visible = True
              
105       End If
          
106       If Log.RPLHistori = 1 Then
107           Command1.Visible = False
108           Label58.Visible = False
109           cmdHistory.Left = 1050
110           Label59.Visible = False
111       End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bookAktif = False
End Sub

Private Sub gridDSS_AfterColEdit(ByVal ColIndex As Integer)
1     On Error GoTo vb_error
2         If IsNull(gridDSS.Columns(5).value) Then
3             Exit Sub
4         End If
          'If (hitungTotalMenit(tempWaktuAkhirDSS) > hitungTotalMenit(gridDSS.Columns(5).value)) And tempStall = gridDSS.Columns(1).value Then
5          If (hitungTotalMenit(tempWaktuAkhirDSS) < hitungTotalMenit(gridDSS.Columns(5).value)) Then
6             gridDSS.Columns(4).value = tempWaktuAwalDSS
7             gridDSS.Columns(5).value = tempWaktuAkhirDSS
8             Exit Sub
9         End If
          'If (hitungTotalMenit(gridDSS.Columns(4).value) < hitungTotalMenit(tempWaktuAwalDSS)) And tempStall = gridDSS.Columns(1).value Then
10        If (hitungTotalMenit(gridDSS.Columns(4).value) < hitungTotalMenit(tempWaktuAwalDSS)) Then
11            gridDSS.Columns(4).value = tempWaktuAwalDSS
12            gridDSS.Columns(5).value = tempWaktuAkhirDSS
13            Exit Sub
14        End If
         ' If tempStall = gridDSS.Columns(1).value And (hitungTotalMenit(tempWaktuAkhirDSS) <= hitungTotalMenit(gridDSS.Columns(5).value)) And (hitungTotalMenit(gridDSS.Columns(4).value) >= hitungTotalMenit(tempWaktuAwalDSS)) Then
      '    If (hitungTotalMenit(tempWaktuAkhirDSS) <= hitungTotalMenit(gridDSS.Columns(5).value)) And (hitungTotalMenit(gridDSS.Columns(4).value) >= hitungTotalMenit(tempWaktuAwalDSS)) Then
              'hitung waktu adjust
              
15            If hitungTotalMenit(Text(122)) > ((hitungTotalMenit(gridDSS.Columns(5).value)) - (hitungTotalMenit(gridDSS.Columns(4).value))) Then
16                MsgBox "Jam mulai dan selesai tidak sesuai adjust waktu " & Trim(Text(122)), vbInformation
                              
17                gridDSS.Columns(4).value = tempWaktuAwalDSS
18                gridDSS.Columns(5).value = tempWaktuAkhirDSS
19            End If
        '  Else
       '       gridDSS.Columns(4).value = tempWaktuAwalDSS
       '       gridDSS.Columns(5).value = tempWaktuAkhirDSS
      '   End If
20        Exit Sub
vb_error:
21        MsgBox "Terjadi kesalahan saat edit waktu", vbInformation
22        gridDSS.Columns(4).value = tempWaktuAwalDSS
23        gridDSS.Columns(5).value = tempWaktuAkhirDSS
End Sub

Private Sub gridDSS_AfterColUpdate(ByVal ColIndex As Integer)
1     On Error GoTo vb_error
          'If (hitungTotalMenit(tempWaktuAkhirDSS) > hitungTotalMenit(gridDSS.Columns(5).value)) And tempStall = gridDSS.Columns(1).value Then
2          If (hitungTotalMenit(tempWaktuAkhirDSS) < hitungTotalMenit(gridDSS.Columns(5).value)) Then
3             gridDSS.Columns(4).value = tempWaktuAwalDSS
4             gridDSS.Columns(5).value = tempWaktuAkhirDSS
5             Exit Sub
6         End If
          'If (hitungTotalMenit(gridDSS.Columns(4).value) < hitungTotalMenit(tempWaktuAwalDSS)) And tempStall = gridDSS.Columns(1).value Then
7         If (hitungTotalMenit(gridDSS.Columns(4).value) < hitungTotalMenit(tempWaktuAwalDSS)) Then
8             gridDSS.Columns(4).value = tempWaktuAwalDSS
9             gridDSS.Columns(5).value = tempWaktuAkhirDSS
10            Exit Sub
11        End If
         ' If tempStall = gridDSS.Columns(1).value And (hitungTotalMenit(tempWaktuAkhirDSS) <= hitungTotalMenit(gridDSS.Columns(5).value)) And (hitungTotalMenit(gridDSS.Columns(4).value) >= hitungTotalMenit(tempWaktuAwalDSS)) Then
         ' If (hitungTotalMenit(tempWaktuAkhirDSS) <= hitungTotalMenit(gridDSS.Columns(5).value)) And (hitungTotalMenit(gridDSS.Columns(4).value) >= hitungTotalMenit(tempWaktuAwalDSS)) Then
              'hitung waktu adjust
              
12            If hitungTotalMenit(Text(122)) > ((hitungTotalMenit(gridDSS.Columns(5).value)) - (hitungTotalMenit(gridDSS.Columns(4).value))) Then
13                MsgBox "Jam mulai dan selesai tidak sesuai adjust waktu " & Trim(Text(122)), vbInformation
                              
14                gridDSS.Columns(4).value = tempWaktuAwalDSS
15                gridDSS.Columns(5).value = tempWaktuAkhirDSS
16            End If
         ' Else
          '    gridDSS.Columns(4).value = tempWaktuAwalDSS
          '    gridDSS.Columns(5).value = tempWaktuAkhirDSS
         ' End If
17        Exit Sub
vb_error:
18        MsgBox "Terjadi kesalahan saat edit waktu", vbInformation
19        gridDSS.Columns(4).value = tempWaktuAwalDSS
20        gridDSS.Columns(5).value = tempWaktuAkhirDSS
End Sub

Private Sub gridDSS_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    On Error GoTo vb_error
    tempStall = gridDSS.Columns(1).value
    tempWaktuAwalDSS = gridDSS.Columns(4).value
    tempWaktuAkhirDSS = gridDSS.Columns(5).value
    Exit Sub
vb_error:
    MsgBox "Terjadi kesalahan saat edit waktu", vbInformation
    gridDSS.Columns(4).value = tempWaktuAwalDSS
    gridDSS.Columns(5).value = tempWaktuAkhirDSS
End Sub

Private Sub gridDSS_LostFocus()
1         On Error GoTo vb_error
2         If tempStall <> gridDSS.Columns(1).value Then
3             Exit Sub
4         End If
          'If (hitungTotalMenit(tempWaktuAkhirDSS) > hitungTotalMenit(gridDSS.Columns(5).value)) And tempStall = gridDSS.Columns(1).value Then
5         If (hitungTotalMenit(tempWaktuAkhirDSS) < hitungTotalMenit(gridDSS.Columns(5).value)) Then
6             gridDSS.Columns(4).value = tempWaktuAwalDSS
7             gridDSS.Columns(5).value = tempWaktuAkhirDSS
8             Exit Sub
9         End If
10        If (hitungTotalMenit(gridDSS.Columns(4).value) < hitungTotalMenit(tempWaktuAwalDSS)) And tempStall = gridDSS.Columns(1).value Then
          'If (hitungTotalMenit(gridDSS.Columns(4).value) < hitungTotalMenit(tempWaktuAwalDSS)) Then
11            gridDSS.Columns(4).value = tempWaktuAwalDSS
12            gridDSS.Columns(5).value = tempWaktuAkhirDSS
13            Exit Sub
14        End If
        '  If tempStall = gridDSS.Columns(1).value And (hitungTotalMenit(tempWaktuAkhirDSS) <= hitungTotalMenit(gridDSS.Columns(5).value)) And (hitungTotalMenit(gridDSS.Columns(4).value) >= hitungTotalMenit(tempWaktuAwalDSS)) Then
         ' If (hitungTotalMenit(tempWaktuAkhirDSS) <= hitungTotalMenit(gridDSS.Columns(5).value)) And (hitungTotalMenit(gridDSS.Columns(4).value) >= hitungTotalMenit(tempWaktuAwalDSS)) Then
              'hitung waktu adjust

15            If hitungTotalMenit(Text(122)) > ((hitungTotalMenit(gridDSS.Columns(5).value)) - (hitungTotalMenit(gridDSS.Columns(4).value))) Then
16                MsgBox "Jam mulai dan selesai tidak sesuai adjust waktu " & Trim(Text(122)), vbInformation

17                gridDSS.Columns(4).value = tempWaktuAwalDSS
18                gridDSS.Columns(5).value = tempWaktuAkhirDSS
19            End If
       '   Else
          '    gridDSS.Columns(4).value = tempWaktuAwalDSS
          '    gridDSS.Columns(5).value = tempWaktuAkhirDSS
         'End If
20        Exit Sub
vb_error:
21        MsgBox "Terjadi kesalahan saat edit waktu", vbInformation
22        gridDSS.Columns(4).value = tempWaktuAwalDSS
23        gridDSS.Columns(5).value = tempWaktuAkhirDSS
End Sub

Private Sub opSvc_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub partGrid_DblClick()
    If barispart > 1 And (partGrid.TextMatrix(partGrid.row, 1) <> "" Or partGrid.TextMatrix(partGrid.row, 7) <> "") Then

    If Trim(Text(1)) <> "" Then
        Dim rsSO As ADODB.Recordset
        Set rsSO = objO.FetchDataACC("Select NOSPK from SOORD WITH (NOLOCK) Where NOSPK = '" & Trim(Text(1)) & "' and KodeItem = '" & (partGrid.TextMatrix(partGrid.row, 1)) & "' and isNull(NamaBYSA,'') = '" & (partGrid.TextMatrix(partGrid.row, 7)) & "' and Nomor = " & partGrid.TextMatrix(partGrid.row, 8) & "  AND ((IsNull(qtyreserve,0)) <> 0 ) AND ([TO] <> 0 Or BO not in ('','CANCEL'))")
        If rsSO.RecordCount > 0 Then
            MsgBox "Spare Part : " & (partGrid.TextMatrix(partGrid.row, 1)) & " " & (partGrid.TextMatrix(partGrid.row, 7)) & Chr(13) & "Sudah disupply !, tidak dapat dihapus/dirubah !", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
    
        a = partGrid.row
        'AMBIL
        With partGrid
            If .TextMatrix(partGrid.row, 1) <> "" Then
                Text(35) = .TextMatrix(a, 1)
                namaPart = .TextMatrix(a, 2)
                satPart = .TextMatrix(a, 3)
                Text(36) = CDbl(.TextMatrix(a, 4))
                HrgPart = CDbl(.TextMatrix(a, 5))
                oliPart = NulA(.TextMatrix(a, 6))
                namabySA = .TextMatrix(a, 7)
                
                If oliPart Then
                    Text(55) = Format(CDbl(Text(55)) - hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5)), "#0,#")
                Else
                    Text(40) = Format(CDbl(Text(40)) - hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5)), "#0,#")
                End If
        
            Else
                Text(35) = .TextMatrix(a, 7)
                namaPart = ""
                satPart = .TextMatrix(a, 3)
                Text(36) = CDbl(.TextMatrix(a, 4))
                HrgPart = 0
                oliPart = .TextMatrix(a, 6)
                namabySA = .TextMatrix(a, 7)
                
                Text(55) = Format(CDbl(Text(55)) - hitungPart(partGrid.TextMatrix(a, 4), 0), "#0,#")

            End If
        End With
        
        'If oliPart Then
        '    Text(55) = Format(CDbl(Text(55)) - hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5)), "#0,#")
        'Else
        '    Text(40) = Format(CDbl(Text(40)) - hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5)), "#0,#")
        'End If
        Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
        
        'HAPUS
        For i = a To barispart
            With partGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
                .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)
                .TextMatrix(i, 6) = .TextMatrix(i + 1, 6)
                .TextMatrix(i, 7) = .TextMatrix(i + 1, 7)
                .TextMatrix(i, 8) = .TextMatrix(i + 1, 8)
            End With
        Next

        barispart = barispart - 1
        Text(35).SetFocus
    End If
End Sub

Private Sub PartGrid_GotFocus()
Dim bolehTambah As Boolean

    If Text(36) = "" Then
        Text(36) = 0
    End If
    If Trim(Text(35)) <> "" And CDbl(Text(36)) > 0 Then
        For i = 1 To barispart
            If partGrid.TextMatrix(i, 1) = Trim(Text(35)) Or partGrid.TextMatrix(i, 7) = Trim(Text(35)) Then
                If MsgBox("Part sudah pernah diinput, lanjutkan ?", vbQuestion + vbYesNo, "Part Booking") = vbYes Then
                    bolehTambah = True
                Else
                    Text(35) = ""
                    Exit Sub
                End If
            Else
                bolehTambah = True
                If partGrid.TextMatrix(i, 8) = "" Then
                    NewId = NewId + 1
                Else
                    NewId = partGrid.TextMatrix(i, 8)
                End If
            End If
        Next
        
        If bolehTambah Then
            If MaxBrsPart >= barispart Then
                With partGrid
                    If namaPart <> "" Then
                    .TextMatrix(barispart, 1) = Text(35)
                    .TextMatrix(barispart, 2) = namaPart
                    .TextMatrix(barispart, 3) = satPart
                    .TextMatrix(barispart, 4) = Text(36)
                    .TextMatrix(barispart, 5) = HrgPart
                    .TextMatrix(barispart, 6) = oliPart
                    .TextMatrix(barispart, 7) = ""
                    .TextMatrix(barispart, 8) = NewId
                    Else
                    .TextMatrix(barispart, 1) = ""
                    .TextMatrix(barispart, 2) = ""
                    .TextMatrix(barispart, 3) = satPart
                    .TextMatrix(barispart, 4) = Text(36)
                    .TextMatrix(barispart, 5) = HrgPart
                    .TextMatrix(barispart, 6) = oliPart
                    .TextMatrix(barispart, 7) = Text(35)
                    .TextMatrix(barispart, 8) = NewId
                    End If
                End With
                barispart = barispart + 1
                If oliPart Then
                    Text(55) = Format(CDbl(Text(55)) + hitungPart(partGrid.TextMatrix(barispart - 1, 4), partGrid.TextMatrix(barispart - 1, 5)), "#0,0")
                Else
                    Text(40) = Format(CDbl(Text(40)) + hitungPart(partGrid.TextMatrix(barispart - 1, 4), partGrid.TextMatrix(barispart - 1, 5)), "#0,0")
                End If
                Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
                
                Text(35) = ""
                Text(36) = 0
                namaPart = ""
                Text(35).SetFocus
                NewId = NewId + 1
            Else
                MsgBox "Maksimal Part " & str(MaxBrsPart) & " !"
                Text(35).SetFocus
            End If
        End If
    End If
End Sub

Private Sub PartGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call partGrid_DblClick
    End If
    If KeyCode = vbKeyReturn Then
        Call partGrid_DblClick
    End If
End Sub

Private Sub tblBatal_Click()
    If MsgBox("Yakin Membatalkan booking ?", vbYesNo, "Konfirmasi") = vbYes Then
        Call objO.UpdateData("UPDATE BOOK set DELIV='BTL',PAID='D' where NO_POLISI ='" & nopol & "' and DELIV='BOK'")
        Call hapusForm
        Text(0) = ""
        Text(56) = ""
        Text(57) = ""
    End If
End Sub

Private Sub TblKeluar_Click()
    Call cekTempStal
    If cmdWarranty.Visible Then
        Call cekWarranty(Trim(Text(41)), Trim(Text(71)))
    End If
    If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then Call objO.UpdateDataACC("Update SOORD Set Status = '' Where Status = '" & gSA & "'")
    Unload Me
End Sub

Private Sub cekWarranty(nRgk As String, kdWar As String)
          Dim rsHlp As ADODB.Recordset
1        On Error GoTo cekWarranty_Error

2         Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM BOOK WITH (NOLOCK) WHERE NO_CHASSIS = '" & nRgk & "' AND KODETIPE = '" & kdWar & "'")
3         If rsHlp.EOF Then
4            Call objO.DeleteData("DELETE TEMPWAR WHERE NO_RANGKA = '" & nRgk & "' AND ID_WARRANTY = '" & kdWar & "'")
5         End If

6        On Error GoTo 0
7        Exit Sub

cekWarranty_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekWarranty of Form frmBooking")
End Sub

Private Sub cekTempStal()
1        On Error GoTo cekTempStal_Error

2         If Trim(nopol) = "" Then
3             Text(0) = ""
4             Text(56) = ""
5             Text(57) = ""
6         ElseIf nopol <> "--" And Tanggal <> ">>->>->>>>" Then
7             If Text(23) <> ">>->>->>>>" Then
8                 Set rsTemp = objO.FetchData("SELECT tanggal FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%'")
9                 If Not rsTemp.EOF Then
10                    Call objO.DeleteData("DELETE tempStall WHERE tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '" & Trim(nopol) & " |%'")
11                End If
12            End If
13        End If

14       On Error GoTo 0
15       Exit Sub

cekTempStal_Error:

16        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekTempStal of Form frmBooking")
End Sub

Private Sub TblKeluar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub tblSimpan_Click()
    Dim boleh As Boolean
    Dim pesan As String
    Dim rsSearch As ADODB.Recordset
    Dim nobookingdidx As String
    Dim eventID As String
    Dim pekerjaan As String
    Dim rsSingleID  As ADODB.Recordset
    Dim oneAccount As String
    
    
1        On Error GoTo tblSimpan_Click_Error

2         boleh = True
    
   ' On Error GoTo vbError
    ' MEMERIKSA FIELD MANDATORY YANG KOSONG, DENGAN FLAG
3         If Trim(nopol) <> "" Then
4       If CekTgl(Text(23)) = False Then
5           pesan = pesan + "Tgl Booking. "
6           boleh = False
7       End If
8       If Text(24) = ">>:>>" Then
9           pesan = pesan + "Jam Booking. "
10          boleh = False
11      End If
12      If Trim(Text(10)) = "" Then
13          pesan = pesan + "Model. "
14          boleh = False
15      Else
16          If NulSID(Text(10)) <> "" Then
17               If Log.RPLHistori = 0 Then
18                  Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(10) & "'")
19              Else
20                  Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & Text(10) & "'")
21              End If
22              If rsSearch.BOF Then
23                  MsgBox "Kelas kendaraan tidak sesuai master Model. Perbaiki data Model kendaraan di master Customer", vbInformation
24                  boleh = False
25              End If
26          End If
27      End If
28      If Trim(Text(3)) = "" Then
29          pesan = pesan + "Nama. "
30          boleh = False
31      End If
32      If Trim(Text(4)) = "" Then
33          pesan = pesan + "Alamat. "
34          boleh = False
35      End If
36      If Trim(Text(8)) = "" And Trim(Text(6)) = "" And Trim(Text(7)) = "" Then
37          pesan = pesan + "Telepon. "
38          boleh = False
39      End If
40      If CekTgl(Text(23)) And chkThs.value = vbUnchecked Then
41          If baris > 1 Then
42              If cmdStall.Tag = "" Or cmdStall.Tag = "belum" Then
43                  Set rsTemp = objO.FetchData("SELECT tanggal FROM tempstall WITH (NOLOCK) WHERE tanggal = '" & Format(Text(23), "mm/dd/yyyy") & "' AND keterangan like '%" & Trim(nopol) & " |%'")
44              Else
45                  Set rsTemp = objO.FetchData("SELECT tanggal FROM stall WITH (NOLOCK) WHERE tanggal = '" & Format(Text(23), "mm/dd/yyyy") & "' AND keterangan like '%" & Trim(nopol) & " |%'")
46              End If
47              If rsTemp.EOF Then
48                  pesan = pesan + "Stall. "
49                  boleh = False
50                  cmdStall.Tag = "belum"
51                  cmdStall.BackColor = &HFFFF&
52              Else
53                  cmdStall.Tag = ""
54                  cmdStall.BackColor = &H8000000F
55              End If
56          Else
                ' non job order only sublet
57              If barisSublet > 1 Then
58              Else
59                  boleh = False
60                  pesan = pesan + "Job / Sublet Order. "
61                  cmdStall.Tag = "belum"
62                  cmdStall.BackColor = &HFFFF&
63              End If
64          End If
65      End If
66      If Text(27) = "" And Text(28) = "" And Text(29) = "" And Text(30) = "" And Text(31) = "" Then
67          boleh = False
68          pesan = pesan + "Permintaan Customer. "
69      End If
70      If Text(64) <> "" Then
'
'            If Text(1) = "" Then
'                nTgl = Text(23)
'                gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
'                Set rsTemp = objO.FetchData("select A.MECH from mekanik A WITH (NOLOCK) LEFT JOIN T_MEKANIK B WITH (NOLOCK) ON A.MECH = B.MECH AND hari = '" & gBantuan2 & "' AND TANGGAL = '" & gBantuan3 & "' LEFT JOIN T_STALL C WITH (NOLOCK) ON C.HARI = '" & gBantuan2 & "' WHERE DATEDIFF(n,WAKTUMULAI,WAKTUSELESAI) - ISNULL(ALOKASI,0) > 0 AND A.MECH = '" & Text(64) & "'")
'                If rsTemp.EOF Then
'                    boleh = False
'                    pesan = pesan + "Waktu Mekanik tidak mencukupi"
'                End If
'            End If
71      End If
72      If Log.dmsNonDss = 1 Then
73          If Trim(Text(1)) <> "" Then
74              totalNilaiPart = getTotalPart(Trim(Text(1)), False)
75          Else
76              totalNilaiPart = CDbl(Trim(Text(40))) + CDbl(Trim(Text(55)))
77          End If
            
78          If totalNilaiPart > Log.plafonDP Then
                'lakukan pengisian DP
79              MsgBox "Nilai part membutuhkan DP minimal Rp." & Format(((Log.persentaseDP / 100) * totalNilaiPart), "0,#"), vbInformation, "Informasi DP Part"
                'boleh = False
80          End If
            
81      End If
        '20/12/2012  Cust Non Axtif
82      If Trim(Text(2).Text) <> "" Then
            Dim rsNon As ADODB.Recordset
83          Set rsNon = objO.FetchData("SELECT * FROM [SINGLEIDORI]..BRIDGING WITH (NOLOCK) WHERE IDCUST = '" & Text(2) & "' AND KodeDealer = '" & Right(gDb, 2) & "'")
84          If Not rsNon.EOF Then
85              txtKodeLgnLama.Text = NulSID(rsNon!kodeLgn)
86          Else
87              txtKodeLgnLama.Text = NulSID(Text(2))
88          End If
89          Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & txtKodeLgnLama.Text & "' and nonaktifflag = 0")
90          If rsNon.RecordCount = 0 Then
91              boleh = False
92              pesan = "Kode Langgan : " & txtKodeLgnLama.Text & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Rubah Kode Langgan" & vbCrLf
93          End If
94      End If
        
        
95      If NulSID(Text(2)) <> "" Then
96          Set rsSingleID = objO.FetchData("SELECT * FROM [SINGLEIDORI]..newcustomer WITH (NOLOCK) WHERE IDCUST = '" & Text(2) & "'")
97          If rsSingleID.EOF Then
98              boleh = False
99              pesan = "Kode Langgan : " & NulSID(Text(2)) & ", Tidak terdaftar dalam single ID !" & Chr(13) & "Silahkan proses dimaster Customer" & vbCrLf

100         End If
101     Else
102         boleh = False
103         pesan = "Kode Langgan : " & NulSID(Text(2)) & ", Tidak terdaftar dalam single ID !" & Chr(13) & "Silahkan proses dimaster Customer" & vbCrLf

104     End If
        
        ''20181023 | REVISI | DITAMBAHKAN VALIDASI UNTUK MENGECEK JOBORDER DENGAN MODEL
        ''-----------------------------------------------------------------------------------------
        Dim rsAct As ADODB.Recordset
        Dim pesanJobOrder As String
105     pesanJobOrder = ""
106     For i = 1 To 7
107         If dGrid.TextMatrix(i, 1) <> "" Then
108             Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(Text(10)) & "'")
109             If rsAct.EOF Then
110                boleh = False
                   'pesan = pesan & "Job order " & NulSID(DGrid.TextMatrix(i, 1)) & " tidak terdaftar dalam model " & NulSID(text(10)) & vbCrLf
111                pesanJobOrder = pesanJobOrder & " - " & NulSID(dGrid.TextMatrix(i, 1)) & vbCrLf
                   
112             End If
113         End If
114     Next
115     If pesanJobOrder <> "" Then
116         MsgBox "Job Order :" & vbCrLf & pesanJobOrder & "Tidak terdaftar pada model : " & NulSID(Text(10)), vbInformation
            
117     End If
        '' ----------------------------------------------------------------------------------------
        
        '' 20191003 || req mb thania | booking
        '' --------------------------------------------------------------
        '' 20191017 request mb thania validasi ditutup
        '' --------------------------------------------------------------
'        If Trim(Text(1)) <> "" Then
'            If tglAwalBook = "01-01-1900" Then
'                Dim rsBooK As ADODB.Recordset
'                Set rsBooK = objO.FetchData("SELECT Date FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & NulSID(Text(1)) & "'")
'                If Not rsBooK.EOF Then
'                    tglAwalBook = Format(rsBooK!Date, "dd/mm/yyyy")
'                End If
'            End If
'
'            If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(tglAwalBook, "DD/MM/YYYY")) < 0 Then
'                boleh = False
'                pesan = pesan & "Tanggal Booking awal sudah lewat tanggal."
'            ''jika tanggal book < tanggal sekarang
'            ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(23), "DD/MM/YYYY")) < 0 Then
'                boleh = False
'                pesan = pesan & "Tanggal Booking tidak boleh lebih kecil dengan hari ini"
'                ''jika tanggal book < tanggal book awal dan tanggal book = hari ini tidak boleh
'            ElseIf DateDiff("d", Format(Text(23), "dd/mm/yyyy"), Format(tglAwalBook, "DD/MM/YYYY")) < 0 And DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(23), "DD/MM/YYYY")) = 0 Then
'                boleh = False
'                pesan = pesan & "Tanggal Booking tidak boleh sama dengan hari ini"
'            ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(23), "DD/MM/YYYY")) > 30 Then
'                boleh = False
'                pesan = pesan & "Tanggal Booking tidak boleh lebih besar dari 1 bulan"
'            End If
'
'
'        Else
'            If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(23), "DD/MM/YYYY")) <= 0 And Trim(Text(1)) = "" Then
'                Text(23) = TglDDMMYYYY(Format(DateAdd("D", 1, isnow), "dd/mm/yyyy")) ''text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
'                boleh = False
'                pesan = pesan & "Tanggal Booking tidak boleh lebih kecil atau sama dengan hari ini"
'                'Exit Sub
'            ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(23), "DD/MM/YYYY")) > 30 Then
'                Text(23) = TglDDMMYYYY(Format(DateAdd("D", 1, isnow), "dd/mm/yyyy")) ''text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
'                boleh = False
'                pesan = pesan & "Tanggal Booking tidak boleh lebih besar dari 1 bulan"
'                'Exit Sub
'            End If
'        End If

        
118     If Get_ConnectionProfil_Bengkel99 = False Then
119        MsgBox "Koneksi HO terputus , proses tidak dapat dilanjutkan.", vbInformation
120        Exit Sub
121     End If
122     nobookingdidx = ""
123     If Log.DiDxP1 = 1 Then
124         nobookingdidx = objO.FetchDataHO_String("SELECT booking_ID from didx..bookingonline_didx WHERE nobookingapps = '" & noBookingAppOnline & "'")
125     End If
126     isRES = True
127     If boleh = True Then
128         nMech = Trim(Text(64))
129         Call saveLogMenu(gSA, ComptName, gVersi, "frmBooking", "tblSimpan", 0)
            
130         On Error GoTo vbError
131         Call objHT.CloseTransactionAll_Err
132         Call objHT.OpenTransactionAll
133         Call objHT.OpenTransactionHO
134         erTranAll = False
            
            'Call SaveDataCust
            
            
135         If erTranAll = False Then Call SaveDataBooking(Trim(nopol))
136         If erTranAll = False Then Call savePart(Text(1))
137         If barisSublet > 1 Then
138             If erTranAll = False Then Call saveSublet
139         End If
140         If erTranAll = False Then Call saveStall
            ''20240119DEV
141         If Log.DIDX = 1 Then
142             If erTranAll = False Then
143                 Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM BookDX WHERE nofak = '" & Trim(Text(1)) & "'")
144                 If Not rsTemp.EOF Then
145                     sSQL = "UPDATE [BookDX] SET DscMName='" & NulSID(Text2(0)) & "',DscMHP='" & NulSID(Text2(1)) & "',DelivAlamat='" & NulSID(txt(8)) & "',DelivKecamatan='" & NulSID(txt(65)) & "',DelivKelurahan='" & NulSID(txt(66)) & "',DelivKota='" & NulSID(txt(64)) & "',DelivPropinsi='" & NulSID(txt(63)) & "',DelivKodePos = '" & NulSID(txt(11)) & "',EstbyWAC = '" & NulSID(Text(88)) & "' WHERE  nofak = '" & Trim(Text(1)) & "'"
146                 Else
147                     sSQL = "INSERT INTO [BookDX] ([NOFAK],[DscMName],[DscMHP],[DelivAlamat],[DelivKecamatan],[DelivKelurahan],[DelivKota],[DelivPropinsi],[DelivKodePos],EstbyWAC)"
148                     sSQL = sSQL & " VALUES ('" & Trim(Text(1)) & "','" & NulSID(Text2(0)) & "','" & NulSID(Text2(1)) & "','" & NulSID(txt(8)) & "','" & NulSID(txt(65)) & "','" & NulSID(txt(66)) & "','" & NulSID(txt(64)) & "','" & NulSID(txt(63)) & "','" & NulSID(txt(11)) & "','" & NulSID(Text(88)) & "') "
149                 End If
150                 Call objHT.InsertData_Bengkel(sSQL)
151             End If
152         End If
153         If chkRtn Then
154             If erTranAll = False Then Call rtjFU(Text(41))
155         End If
            'Call CetakForm
            '------------------------------------------------------
            '- 2015-05-08 | kUDIs | PENAMBAHAN LOG DSS            |
            '------------------------------------------------------
156         If Log.isDSS = 1 Then
                Dim jamMulai As String
                Dim jamSelesai As String
                Dim noStall As String
157             noStall = ""
158             jamMulai = "00:00"
159             jamSelesai = "00:00"
160             Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM STALL WHERE TANGGAL = '" & Format(Text(23), "MM/DD/YYYY") & "' and keterangan like '%" & Trim(Text(1)) & "%'")
161             If Not rsTemp.EOF Then
162                 rsTemp.MoveFirst
163                 jamMulai = Trim(rsTemp!jamMulai)
164                 jamSelesai = Trim(rsTemp!jamSelesai)
165                 noStall = UCase(Trim(rsTemp!noStall))
166             End If
167             Call objHT.InsertData_Bengkel("INSERT INTO LOGDSS VALUES ('" & Trim(Text(1)) & "','" & Trim(sgnDSS.Caption) & "','" & adRate & "','" & Trim(Text(122)) & "',GETDATE(),'" & jamMulai & "','" & jamSelesai & "','" & noStall & "') ")
168         End If
            
169         If Log.FIR = 1 Then
               ''update nobooking didetailfir
170            Set rsTemp = objHT.FetchDataBengkel("SELECT NOFAK FROM DETAILFIR WITH (NOLOCK) WHERE NO_POLISI = '" & Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) & "' AND TANGGAL = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND NOFAK = ''")
171            If Not rsTemp.EOF Then
172               Call objHT.InsertData_Bengkel("UPDATE DETAILFIR SET NOFAK = '" & Trim(Text(1)) & "' WHERE NO_POLISI = '" & Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) & "' AND TANGGAL = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND NOFAK = ''")
173            End If
174         End If
            
175         If isRetentionBoard = True Then
176             gBantuan = Trim(Text(1))
177             gBantuan2 = Trim(nopol)
                
178         End If
179         If Log.DiDxP1 = 1 Then
                
180             If nobookingdidx <> "" Then
181                 eventID = objO.FetchDataHO_String("select event_id from didx..bookingonline_didx_toSend where booking_ID = '" & nobookingdidx & "'")
                    
182                 Call objHT.InsertData_Ho("INSERT INTO didx..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & noBookingAppOnline & "','" & Right(gDb, 2) & "." & Trim(Text(1)) & "','','" & Right(gDb2, 2) & "','CONFIRMED','0','" & eventID & "',getdate())")
                    
183                 Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM STALL WHERE TANGGAL = '" & Format(Text(23), "MM/DD/YYYY") & "' and keterangan like '%" & Trim(Text(1)) & "%'")
184                 If Not rsTemp.EOF Then
185                     rsTemp.MoveFirst
186                     jamMulai = Trim(rsTemp!jamMulai)
187                     jamSelesai = Trim(rsTemp!jamSelesai)
188                     noStall = UCase(Trim(rsTemp!noStall))
189                 End If
                    
190                 Call objHT.InsertData_Ho("DELETE FROM didx..bookingonline_part_didx_tosend WHERE booking_ID = '" & nobookingdidx & "' ")
191                 For i = 1 To barispart - 1
192                     If Trim(partGrid.TextMatrix(i, 1)) <> "" Or Trim(partGrid.TextMatrix(i, 7)) <> "" Then
193                         Call objHT.InsertData_Ho("INSERT INTO didx..bookingonline_part_didx_tosend (booking_ID,part_number,part_name) VALUES ('" & nobookingdidx & "','" & Trim(partGrid.TextMatrix(i, 1)) & "','" & Trim(partGrid.TextMatrix(i, 2)) & "')")
194                     End If
                        
195                 Next
                    
196                 Call objHT.InsertData_Ho("DELETE FROM didx..bookingonline_job_didx_tosend WHERE booking_ID = '" & nobookingdidx & "' ")
                    
197                 For i = 1 To 7
198                     If dGrid.TextMatrix(i, 1) <> "" Then
199                         pekerjaan = pekerjaan & dGrid.TextMatrix(i, 1) & ";"
200                         Call objHT.InsertData_Ho("INSERT INTO didx..bookingonline_job_didx_tosend (booking_ID,job_ID,job_name) VALUES ('" & nobookingdidx & "','" & i & "','" & Trim(dGrid.TextMatrix(i, 1)) & "')")
201                     End If
202                 Next
203                 If isRES Then
204                     Call objHT.InsertData_Ho("UPDATE didx..bookingonline_didx_toSend SET [booking_status] = 'CONFIRMED', slot_requested_datetime_start = '" & Format(Text(23), "YYYY-MM-DD") & " " & jamMulai & "',slot_requested_datetime_end = '" & Format(Text(23), "YYYY-MM-DD") & " " & jamSelesai & "',job_description = '" & pekerjaan & "',stall_or_squad_number = '" & noStall & "',dealer_customer_ID = '" & NulSID(Text(2)) & "' WHERE booking_ID = '" & nobookingdidx & "' ")
                    
205                 Else
206                     Call objHT.InsertData_Ho("UPDATE didx..bookingonline_didx_toSend SET [booking_status] = 'CONFIRMED', slot_datetime_start = '" & Format(Text(23), "YYYY-MM-DD") & " " & jamMulai & "',slot_datetime_end = '" & Format(Text(23), "YYYY-MM-DD") & " " & jamSelesai & "',job_description = '" & pekerjaan & "',stall_or_squad_number = '" & noStall & "',dealer_customer_ID = '" & NulSID(Text(2)) & "' WHERE booking_ID = '" & nobookingdidx & "' ")
                    
207                 End If
208             Else
209                 oneAccount = objO.FetchDataHO_String("SELECT one_account_ID FROM didx..bookingonline_didx_toSend WHERE phone_number = '" & Text(6) & "'")
210                 Call objHT.InsertData_Ho("INSERT INTO didx..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,[status],proses,event_no,tglentry) VALUES ('BOOKGRP','" & Right(gDb, 2) & "." & Trim(Text(1)) & "','" & Right(gDb, 2) & "." & Trim(Text(1)) & "','','" & Right(gDb2, 2) & "','CONFIRMED','0','',getdate())")
                    
211                 Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM STALL WHERE TANGGAL = '" & Format(Text(23), "MM/DD/YYYY") & "' and keterangan like '%" & Trim(Text(1)) & "%'")
212                 If Not rsTemp.EOF Then
213                     rsTemp.MoveFirst
214                     jamMulai = Trim(rsTemp!jamMulai)
215                     jamSelesai = Trim(rsTemp!jamSelesai)
216                     noStall = UCase(Trim(rsTemp!noStall))
217                 End If
                    
218                 Call objHT.InsertData_Ho("DELETE FROM didx..bookingonline_part_didx_tosend WHERE booking_ID = '" & Right(gDb, 2) & "." & Trim(Text(1)) & "' ")
219                 For i = 1 To barispart - 1
220                     If Trim(partGrid.TextMatrix(i, 1)) <> "" Or Trim(partGrid.TextMatrix(i, 7)) <> "" Then
221                         Call objHT.InsertData_Ho("INSERT INTO didx..bookingonline_part_didx_tosend (booking_ID,part_number,part_name) VALUES ('" & Right(gDb, 2) & "." & Trim(Text(1)) & "','" & Trim(partGrid.TextMatrix(i, 1)) & "','" & Trim(partGrid.TextMatrix(i, 2)) & "')")
222                     End If
                        
223                 Next
                    
224                 Call objHT.InsertData_Ho("DELETE FROM didx..bookingonline_job_didx_tosend WHERE booking_ID = '" & Right(gDb, 2) & "." & Trim(Text(1)) & "' ")
                    
225                 For i = 1 To 7
226                     If dGrid.TextMatrix(i, 1) <> "" Then
227                         pekerjaan = pekerjaan & dGrid.TextMatrix(i, 1) & ";"
228                         Call objHT.InsertData_Ho("INSERT INTO didx..bookingonline_job_didx_tosend (booking_ID,job_ID,job_name) VALUES ('" & Right(gDb, 2) & "." & Trim(Text(1)) & "','" & i & "','" & Trim(dGrid.TextMatrix(i, 1)) & "')")
229                     End If
230                 Next
231                 If isRES Then
232                     Call objHT.InsertData_Ho("UPDATE didx..bookingonline_didx_toSend SET [booking_status] = 'CONFIRMED', phone_number = '" & Text(6) & "',slot_requested_datetime_start = '" & Format(Text(23), "YYYY-MM-DD") & " " & jamMulai & "',slot_requested_datetime_end = '" & Format(Text(23), "YYYY-MM-DD") & " " & jamSelesai & "',job_description = '" & pekerjaan & "',stall_or_squad_number = '" & noStall & "' WHERE booking_ID = '" & Right(gDb, 2) & "." & Trim(Text(1)) & "' ")
                    
233                 Else
234                     sSQL = "INSERT INTO didx..[bookingonline_didx_toSend] ([event_id] ,[timestamp],[one_account_ID],[first_name],[last_name],[gender],[phone_number],[email],[VIN]"
235                     sSQL = sSQL & " ,[model],[police_number],[actual_mileage],[booking_number],[booking_ID],[booking_created_datetime],[slot_start_datetime],[service_category],[service_type],[outlet_ID]"
236                     sSQL = sSQL & " ,[outlet_name],[vehicle_problem],[cancellation_reason],[mobile_service_address],[subdistrict],[district],[city],[province],[postal_code],[booking_status],"
237                     sSQL = sSQL & " [nobookingapps],[primary_user],[variant],[created_datetime],[slot_datetime_start],[slot_datetime_end],[service_location],[job_type],[slot_requested_datetime_start],"
238                     sSQL = sSQL & " [slot_requested_datetime_end],[job_description],[stall_or_squad_number],dealer_customer_ID) VALUES"
239                     sSQL = sSQL & " ('',getdate(),'" & oneAccount & "','" & NulSID(rsSingleID!namacust1) & "','" & NulSID(rsSingleID!namacust2) & "','MALE','" & NulSID(Text(6)) & "','" & IIf(NulSID(rsSingleID!email) <> "", NulSID(rsSingleID!email), "jogn.doe@gmail.com") & "','" & NulSID(Text(41)) & "','" & NulSID(Text(12)) & "'"
240                     sSQL = sSQL & " ,'" & Replace(Trim(nopol), "-", "") & "','0','" & Right(gDb, 2) & "." & Trim(Text(1)) & "','" & Right(gDb, 2) & "." & Trim(Text(1)) & "',getdate(),'" & Format(Text(23), "YYYY-MM-DD") & " " & jamSelesai & "','" & pekerjaan & "','SBE','" & Log.kdDealerDidx & "','" & NulSID(Log.LogNamaNas) & "','" & pekerjaan & "','','" & NulSID(rsSingleID!Alamat) & "'"
241                     sSQL = sSQL & " ,'" & NulSID(rsSingleID!kelurahan) & "','" & NulSID(rsSingleID!kecamatan) & "','" & NulSID(rsSingleID!Kota) & "','" & NulSID(rsSingleID!Propinsi) & "','" & NulSID(rsSingleID!KodePos) & "','CONFIRMED','" & Right(gDb, 2) & "." & Trim(Text(1)) & "','MASTER','" & NulSID(Text(12)) & "',getdate(),'" & Format(Text(23), "YYYY-MM-DD") & " " & jamMulai & "','" & Format(Text(23), "YYYY-MM-DD") & " " & jamSelesai & "'"
242                     sSQL = sSQL & " ,'WORKSHOP','" & cbJobType.Text & "','" & Format(Text(23), "YYYY-MM-DD") & " " & jamMulai & "','" & Format(Text(23), "YYYY-MM-DD") & " " & jamSelesai & "','Servis Rutin','" & noStall & "','" & NulSID(Text(2)) & "')"
243                     Call objHT.InsertData_Ho(sSQL)
                    
244                 End If
245             End If
246         End If
247         If isBookingOnline Then
248             Call objHT.InsertData_Bengkel("UPDATE BOOKINGONLINE SET FPROSES = 1,NOFAK = '" & Trim(Text(1)) & "' WHERE ID = '" & IDBookingOnline & "'")
249         End If
            
250         If Log.RPLHistori = 1 Then
                '' 20190816DEV1
251             If Log.hansaRenkei = 1 Then
                    Dim rsHansaRenkei   As ADODB.Recordset
                    Dim kodeSales       As String
                    Dim noHP            As String
                    Dim tglNotif        As Date
                    
252                 If NulSID(Text(6)) <> "" Then
253                     noHP = NulSID(Text(6))
254                 ElseIf NulSID(Text(7)) <> "" Then
255                     noHP = NulSID(Text(7))
256                 Else
257                     noHP = NulSID(Text(8))
258                 End If
                    
                    '' tanggal notif 3 hari sebelum hari booking
                    '' jika 3 hari sebelum booking kurang dari hari ini tanggal notif hari ini
                    '' --------------------------------------------------------------------------------
                    
259                 If DateDiff("D", Format(isnow, "YYYY-MM-DD"), DateAdd("D", -3, Text(23))) < 0 Then
260                     tglNotif = Format(isnow, "YYYY-MM-DD HH:MM")
261                 Else
262                     tglNotif = Format(DateAdd("D", -3, Text(23)), "YYYY-MM-DD") & " " & Text(24)
263                 End If
                    
264                 Set rsHansaRenkei = objHT.FetchDataBengkel("SELECT * FROM [" & dbReplikasiHO & "]..NotifHansaRenkei WITH (NOLOCK) WHERE KODEDEALER = '" & Right(gDb, 2) & "' AND NOWO = '" & NulSID(Text(1)) & "' ")
265                 If rsHansaRenkei.EOF Then
266                     If NulSID(Text(41)) <> "" Then
                            ''nik sales diambil dari HO yang dikirim kodesales dan kodecabang
267                         kodeSales = objHT.FetchDataBengkel_String("SELECT Kode_Sales FROM [SPK" & Right(gDb, 2) & "]..SPK WITH (NOLOCK) WHERE NoRangka = '" & NulSID(Text(41)) & "' AND SPK_Status NOT IN (2,3) ")
268                     End If
269                     sSQL = "INSERT INTO [" & dbReplikasiHO & "]..NotifHansaRenkei (KodeSumber,Tipe,Nowo,NoRangka,KodeDealer,KodePesan,KodeSales,NIK,TglNotif,Proses,UserEntry,TglEntry,keterangan)"
270                     sSQL = sSQL & " VALUES ('GRP','BOOK','" & NulSID(Text(1)) & "','" & NulSID(Text(41)) & "','" & Right(gDb, 2) & "','BOOKING : " & NulSID(Text(1)) & "','" & kodeSales & "','','" & Format(tglNotif, "YYYY-MM-DD HH:MM") & "',0,'" & gSA & "',GETDATE(),'" & NulSID(nopol) & "#" & NulSID(Text(12)) & "#" & NulSID(Text(3)) & "#" & Format(Text(23), "yyyy-mm-dd") & " " & NulSID(Text(24)) & "##" & noHP & "')"
271                     Call objHT.InsertData_Bengkel(sSQL)
272                 End If
273             End If
274         End If
            
275         If erTranAll = False Then
276             Call objHT.CloseTransactionAll
277             Call objHT.CloseTransactionHO
278             Call saveLogMenu(gSA, ComptName, gVersi, "frmBooking", "tblSimpan", 1)
279             If gPreviewCetak = True Then Text(0).SetFocus
280             nJRate = 0
                'pengecekan apakah dari diag
281             If isBookDiag Then
                    'jika booking dari diagnose
                    'lakukan proses pengupdatetan no booking ke dalam diagdta dan diagapproval
282                 objO.InsertData ("UPDATE DIAGDTA SET NOFAK = '" & Trim(Text(1)) & "' WHERE NO_DIAG = '" & no_approval & "'")
283                 objO.InsertData ("UPDATE DIAGAPPROVAL SET NOFAK = '" & Trim(Text(1)) & "' WHERE NO_DIAG = '" & no_approval & "'")
284             End If
285             Call CetakForm1
                ''2016-06-02 "" cetak detail FIR
286               If Log.FIR = 1 Then
287                  Set rsTmp = objO.FetchData("SELECT IDKELUHAN FROM DETAILFIR WITH (NOLOCK) WHERE NOFAK = '" & Trim(Text(1)) & "'")
288                  If Not rsTmp.EOF Then
289                     rsTmp.MoveFirst
290                     Do While Not rsTmp.EOF
291                        Call cetakDetailFIR(Trim(rsTmp!idkeluhan))
292                        rsTmp.MoveNext
293                     Loop
294                  End If
295               End If
296         Else
297             Call objHT.CloseTransactionAll_Err
298             Call objHT.CloseTransactionHO_err
299             Call saveLogMenu(gSA, ComptName, gVersi, "frmBooking", "tblSimpan", 2)
300             MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
301         End If
302         simpan = True
303     Else
304         MsgBox "Masukkan " + pesan + "!"
305     End If
306       Else
307     MsgBox "Masukkan No. Polisi !", vbOKOnly, "Perhatian"
308       End If
309       Exit Sub
    
vbError:
310       If erTranAll = True Then
311     err.Raise err.Description
312     Call objHT.CloseTransactionAll_Err
313     Call saveLogMenu(gSA, ComptName, gVersi, "frmBooking", "tblSimpan", 2)
314     MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
315       End If
316       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmBooking")
317       Exit Sub

318      On Error GoTo 0
319      Exit Sub

tblSimpan_Click_Error:

320       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmBooking")
End Sub

Private Sub CetakForm1()
1        On Error GoTo CetakForm1_Error

2     Set rpt = New RptBooking2
          
3         If MsgBox("Mencetak Booking ?", vbYesNo, "Konfirmasi") = vbYes Then
              Dim sHlp As Object
              Dim sHlp1 As Object
              Dim sHlp2 As Object
              Dim sHlp3 As Object
              Dim sHlp4 As Object
              Dim rRs As New ADODB.Recordset
              Dim crxSubReportPart As CRAXDRT.Report
              Dim crxSubReportKeluhan As CRAXDRT.Report
              Dim crxSubReportJob As CRAXDRT.Report
              Dim crxSubReportCat As CRAXDRT.Report
              Dim crxSubReportReq As CRAXDDRT.Report
                  
4                 Set rRs = objO.FetchData("SELECT CUSTOMER, CUST_ADRES, KOTA, TELPON, NO_POLISI, NO_ENGINE, VDS, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, " _
                                        & "FS_KD_DEALER, FS_NM_DEALER, DDATE, CODE, DOTAM, NPWP, PPN, DISC, DISCT, DATE, FRONT, TIME, TIME_D, MECH, REQ1, REQ2, REQ3, REQ4, " _
                                        & "REQ5, REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, TOT_RATE, PAID, NOFAK, " _
                                        & "REMARK1, SUGGEST, ESTDATE, ESTTIME, DELIV, FS_TH_MOBIL, FS_LOG_ADD, FS_LOG_DEL, FD_TGL_DEL, FS_RC, FS_NM_HOBI, DATE_BOOK, " _
                                        & "FS_KD_PLG, FS_KD_WIL, GSM, CDMA, FLAGSVC, TOTPART, TOTLABOUR, TOTSUBLET, TOTOLI, REKOM, COND1, COND2, COND3, COND4, COND5, " _
                                        & "COND6, THS, FN_RETURN, KARPLG, ALASAN, CATATAN1, CATATAN2, CATATAN3, CATATAN4, KODETIPE, CMPL1, CMPL2, CMPL3, CMPL4, CMPL5, AMBIL, DATESERAH, TIMESERAH FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & Trim(Text(1)) & "'")
5                 rpt.database.SetDataSource rRs
                  
6                 sSQL = "select NOFAK, KODEITEM, NAMABARANG, QTY, NAMABYSA FROM BOOK_PART WITH (NOLOCK) WHERE NOFAK = '" & Trim(Text(1)) & "'"
7                 Set sHlp = objO.FetchData(sSQL)
8                 Set crxSubReportPart = rpt.Subreport1.OpenSubreport
9                 crxSubReportPart.database.SetDataSource sHlp
                  
10                sSQL = "select KELUHAN FROM repoBook('" & Trim(Text(1)) & "') WHERE KELUHAN <> '' ORDER BY IDB"
11                Set sHlp1 = objO.FetchData(sSQL)
12                Set crxSubReportKeluhan = rpt.Subreport4.OpenSubreport
13                crxSubReportKeluhan.database.SetDataSource sHlp1
                  
14                sSQL = "select JOBORDER FROM repoBook('" & Trim(Text(1)) & "') WHERE JOBORDER <> '' UNION SELECT NAMABARANG FROM BOOK_SUBLET WHERE NOFAK = '" & Trim(Text(1)) & "'"
15                Set sHlp2 = objO.FetchData(sSQL)
16                Set crxSubReportJob = rpt.Subreport5.OpenSubreport
17                crxSubReportJob.database.SetDataSource sHlp2
                  
18                sSQL = "select CATATAN FROM repoBook('" & Trim(Text(1)) & "') WHERE CATATAN <> '' ORDER BY IDB"
19                Set sHlp3 = objO.FetchData(sSQL)
20                Set crxSubReportCat = rpt.Subreport6.OpenSubreport
21                crxSubReportCat.database.SetDataSource sHlp3
                  
22                sSQL = "select REQUEST FROM repoBook('" & Trim(Text(1)) & "') WHERE REQUEST <> '' ORDER BY IDB"
23                Set sHlp4 = objO.FetchData(sSQL)
24                Set crxSubReportReq = rpt.Subreport2.OpenSubreport
25                crxSubReportReq.database.SetDataSource sHlp4
                  
26                rpt.txtPT.SetText Trim(gPerush)
27                rpt.txtPT2.SetText Trim(gPerush)
                  
                  Dim rsSearch As ADODB.Recordset
28                Set rsSearch = objO.FetchData("SELECT NOSTALL FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & Trim(nopol) & " |%' AND FLAG = ''")
29                If Not rsSearch.EOF Then
30                    rpt.txtStall.SetText Trim(rsSearch("NOSTALL"))
31                End If
                   
                  
32                With CRViewer1
33                    .ReportSource = rpt
34                    .ViewReport
35                    .Zoom 100
36                    .Visible = True
37                    .Top = 0
38                    .Left = asxPowerBanner1.Left + asxPowerBanner1.Width
39                    .Height = asxPowerBanner1.Height + 200
40                    .Width = Me.Width - .Left
41                End With
42                cmdTutup.Visible = True
43                cmdTutup.Left = CRViewer1.Left + 20
44                cmdTutup.Top = 435
                  
45                cmdSettingPrinter.Top = cmdTutup.Top
46                cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
47                cmdSettingPrinter.Visible = True
48                Screen.MousePointer = vbDefault
49        End If

50       On Error GoTo 0
51       Exit Sub

CetakForm1_Error:

52        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure CetakForm1 of Form frmBooking")
End Sub

Private Sub tblSimpan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub Text_Change(Index As Integer)
    If Index = 52 Or Index = 58 Or Index = 70 Or Index = 46 Or Index = 47 Or Index = 48 Or Index = 49 Or Index = 50 Or Index = 51 Or Index = 57 Or Index = 1 Then
        If Text(0) <> "" Then
            Select Case Index
                Case 52
                    If Text(Index) = "" Then
                        Text(Index) = 0
                    ElseIf CDbl(Text(Index)) >= 0 Then
                        Text(Index) = Format(Text(Index), "#0,#")
                        Text(Index).SelStart = Len(Text(Index))
                        Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
                    End If
                Case 58
                    If Text(58) <> ">>:>>" And Text(23) <> ">>->>->>>>" Then
                        nWKT = hitungWaktu(CDbl(nJRate), Text(58), Text(23))
                        Text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
                        Text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
                        Text(84) = Trim(Text(59))
                        Text(83) = Trim(Text(45))
                        Call hitungEstimasiSelesai
                    End If
                Case 70
                    baris = CDbl(gBantuan)
                    barispart = CDbl(gBantuan2)
                    
                Case 46
                    Me.Text(54).Visible = IIf(Me.Text(46) = "LAINNYA", True, False)
                Case 47
                    Me.Text(65).Visible = IIf(Me.Text(47) = "LAINNYA", True, False)
                Case 48
                    Me.Text(66).Visible = IIf(Me.Text(48) = "LAINNYA", True, False)
                Case 49
                    Me.Text(67).Visible = IIf(Me.Text(49) = "LAINNYA", True, False)
                Case 50
                    Me.Text(68).Visible = IIf(Me.Text(50) = "LAINNYA", True, False)
                Case 51
                    Me.Text(69).Visible = IIf(Me.Text(51) = "LAINNYA", True, False)
                Case 57
                    If nX Then
                        Call text_Validate(Index, False)
                    End If
                Case 1
                    If Trim(Text(1)) <> "" Then
                        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then
                            Dim rsPM As ADODB.Recordset
                            Set rsPM = objO.FetchDataACC("select Top 1 Status from SOORD WITH (NOLOCK) Where NoSPK = '" & Trim(Text(1)) & "'")
                            If rsPM.RecordCount > 0 Then
                                If rsPM.Fields(0).value <> "" And rsPM.Fields(0).value <> gSA Then
                                    MsgBox "Sparepart sedang diproses oleh User (" & rsPM.Fields(0).value & ") !", vbInformation + vbOKOnly, "Reserve Sparepart"
                                    Call cmdBersih_Click
                                    Exit Sub
                                Else
                                    Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NoSPK = '" & Trim(Text(1)) & "'")
                                End If
                            Else
                                Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NoSPK = '" & Trim(Text(1)) & "'")
                            End If
                        End If
                    Else
                        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then Call objO.UpdateDataACC("Update SOORD Set Status = '' Where Status ='" & gSA & "'")
                    End If
            End Select
        End If
        
        'lakukan pengecekan booking
        
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index).Mask = ""
            Text(Index) = gMainKeyCari
            If gBantuan <> "" And Index = 2 Then
                Text(41) = gBantuan
            End If
            If Index = 0 Then
                Dim nDpn1 As String
                Dim nTgh1 As String
                Dim nBlk1 As String
                nDpn1 = ""
                nTgh1 = ""
                nBlk1 = ""
                Call pecahNoPol(gMainKeyCari, nDpn1, nTgh1, nBlk1)
                Text(0) = Trim(nDpn1)
                Text(56) = Trim(nTgh1)
                Text(57) = Trim(nBlk1)
            End If
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    Select Case Index
        Case 23, 46
            VScroll1.value = 0
        Case 32
            VScroll1.value = 4900
        Case 60
            VScroll1.value = 1000
        Case 63, 37
            VScroll1.value = 2000
    End Select
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub SaveDataBooking(NoPolisi As String)
          
          ' *** MENYIMPAN DATA BOOKING ***
1        On Error GoTo SaveDataBooking_Error

2         If Text(1) = "" Then
3             Call isiNoBook
4             If erTranAll = True Then Exit Sub
5         End If
          
6         Call ClearsSql
7         gsSql1(0) = "NO_POLISI"
8         gsSql2(0) = nopol
9         gsSql1(1) = "NOFAK"
10        gsSql2(1) = Text(1)
11        gsSql1(2) = "CUSTOMER"
12        gsSql2(2) = Text(3)
13        gsSql1(3) = "CUST_ADRES"
14        gsSql2(3) = Text(4)
15        gsSql1(4) = "KOTA"
16        gsSql2(4) = Text(5)
17        gsSql1(5) = "GSM"
18        gsSql2(5) = Text(6)
19        gsSql1(6) = "CDMA"
20        gsSql2(6) = Text(7)
21        gsSql1(7) = "TELPON"
22        gsSql2(7) = Text(8)
23        gsSql1(8) = "FS_TH_MOBIL"
24        gsSql2(8) = Text(9)
25        gsSql1(9) = "MDL"
26        gsSql2(9) = Text(10)
27        gsSql1(10) = "MODEL"
28        gsSql2(10) = Text(11)
29        gsSql1(11) = "FMDL"
30        gsSql2(11) = Text(12)
31        gsSql1(12) = "COLOR"
32        gsSql2(12) = Text(13)
33        gsSql1(13) = "FS_KD_DEALER"
34        gsSql2(13) = Text(14)
35        gsSql1(14) = "FS_NM_DEALER"
36        gsSql2(14) = Text(15)
37        gsSql1(15) = "DOTAM"
38        gsSql2(15) = Replace(Text(16), ">", "")
39        gsSql1(16) = "DDATE"
40        gsSql2(16) = Replace(Text(17), ">", "")
          
41        gsSql1(17) = "[DATE]"
42        gsSql2(17) = Format(TglYYYYMMDD(Text(23)), "mm/dd/yyyy")
43        gsSql1(18) = "TIME"
44        gsSql2(18) = Text(24)
45        gsSql1(19) = "PAID"
46        gsSql2(19) = Text(25)
          
47        gsSql1(20) = "REQ1"
48        gsSql2(20) = Text(27)
49        gsSql1(21) = "REQ2"
50        gsSql2(21) = Text(28)
51        gsSql1(22) = "REQ3"
52        gsSql2(22) = Text(29)
53        gsSql1(23) = "REQ4"
54        gsSql2(23) = Text(30)
55        gsSql1(24) = "REQ5"
56        gsSql2(24) = Text(31)

          Dim hlp As String
          Dim hlp2 As String
          Dim nTotRt As Double
          Dim nTotAct As Integer
          Dim rsLama As ADODB.Recordset
          Dim rsAct As ADODB.Recordset
          
          ' HAPUS JAM MEKANIK LAMA
         ''2016-02-24
      '    Set rsLama = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & Trim(Text(1)) & "'")
57        Set rsLama = objHT.FetchDataBengkel("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & Trim(Text(1)) & "'")
58        If Not rsLama.EOF Then
59            For i = 1 To 7
60                hlp = "REP" & i
61                If rsLama(hlp) <> "" Then
                        ''2016-02-24
                      ''Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'")
                      
                      'If Log.RPLHistori = 0 Then
62                        Set rsAct = objHT.FetchDataBengkel("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'")
                     ' Else
                      '    Set rsAct = objHT.FetchDataBengkel("SELECT ISNULL(RATE,0)*60 AS RATE FROM [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'")
                      'End If
63                    If Not rsAct.EOF Then
64                        nTotAct = nTotAct + CDbl(rsAct("RATE"))
65                    End If
66                    rsAct.Close
67                End If
68            Next
69            Call createMekanik
              'Call decresRestTime(rsLama("MECH"), CDbl(nTotAct), 1, Format(rsLama("DATE"), "MM/DD/YYYY"))
70        End If
71        rsLama.Close
          
72        nTotRt = 0
73        nTotAct = 0
74        For i = 1 To 7
75            hlp = "REP" & i
76            hlp2 = "RTE" & i
77            gsSql1(24 + i) = hlp
78            gsSql2(24 + i) = dGrid.TextMatrix(i, 1)
79            gsSql1(31 + i) = hlp2
80            gsSql2(31 + i) = Replace(IIf(dGrid.TextMatrix(i, 2) <> "", dGrid.TextMatrix(i, 2), 0), ",", ".")
81            nTotRt = nTotRt + IIf(dGrid.TextMatrix(i, 2) <> "", dGrid.TextMatrix(i, 2), 0)
82            If dGrid.TextMatrix(i, 1) <> "" Then
               ''2016-02-24
      '         Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(Text(10)) & "'")
               
               'If Log.RPLHistori = 0 Then
83                Set rsAct = objHT.FetchDataBengkel("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(Text(10)) & "'")
               'Else
               '   Set rsAct = objHT.FetchDataBengkel("SELECT ISNULL(RATE,0)*60 AS RATE FROM [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
              ' End If
84                If Not rsAct.EOF Then
85                   nTotAct = nTotAct + CDbl(rsAct("RATE"))
86                End If
87                rsAct.Close
88            End If
89        Next
          'Call countRestTime(Trim(text(64)), CDbl(nTotAct), 1, Format(text(23), "MM/DD/YYYY"))
          
90        If Text(18) = ">>->>->>>>" Then
91            Text(18) = TglDDMMYYYY(Format(isnow, "DD/MM/YYYY"))
92            Text(19) = Format(isnow, "HH:NN")
93        End If
          
94        gsSql1(39) = "DATE_BOOK"
95        gsSql2(39) = Format(TglYYYYMMDD(Text(18)), "mm/dd/yyyy")
96        gsSql1(40) = "TIME_D"
97        gsSql2(40) = Text(19)
          
98        gsSql1(41) = "REMARK1"
99        gsSql2(41) = Text(37)
100       gsSql1(42) = "SUGGEST"
101       gsSql2(42) = Text(38)
102       gsSql1(43) = "DELIV"
103       gsSql2(43) = "BOK"
          
104       gsSql1(44) = "FS_LOG_ADD"
105       gsSql2(44) = gSA
          
106       gsSql1(45) = "FLAGSVC"
107       If opSvc(0).value Then
108           gsSql2(45) = "TUNGGU"
109       ElseIf opSvc(1).value Then
110           gsSql2(45) = "TINGGAL"
111       End If
          
112       gsSql1(46) = "TOTPART"
113       gsSql2(46) = CDbl(Text(40))
114       gsSql1(47) = "TOTLABOUR"
115       gsSql2(47) = CDbl(Text(39))
116       gsSql1(48) = "TOTSUBLET"
117       gsSql2(48) = CDbl(Text(52))
118       gsSql1(49) = "TOTOLI"
119       gsSql2(49) = CDbl(Text(55))
120       gsSql1(50) = "FS_KD_PLG"
121       gsSql2(50) = NulSID(txtKodeLgnLama.Text) 'text(2)
122       gsSql1(51) = "NO_CHASSIS"
123       gsSql2(51) = gNoRangka
          
124       gsSql1(52) = "ESTTIME"
125       gsSql2(52) = Text(45)
126       gsSql1(53) = "TOT_RATE"
127       gsSql2(53) = Replace(CDbl(nTotRt), ",", ".")
          
128       gsSql1(54) = "REKOM"
129       gsSql2(54) = IIf(cmdNo.Caption = "No &Rekom", "Y", "N")
          
130       For i = 1 To 6
131           hlp = "COND" & i
132           gsSql1(54 + i) = hlp
133           gsSql2(54 + i) = IIf(Text(45 + i) = "LAINNYA", "L| " & Trim(bKondisi(i - 1)), Text(45 + i))
134       Next
          
135       gsSql1(61) = "THS"
136       gsSql2(61) = IIf(chkThs = Checked, "1", "0")
137       gsSql1(62) = "FN_RETURN"
138       gsSql2(62) = IIf(chkRtn = Checked, "1", "0")
139       gsSql1(63) = "CATATAN1"
140       gsSql2(63) = Text(60)
141       gsSql1(64) = "CATATAN2"
142       gsSql2(64) = Text(61)
143       gsSql1(65) = "CATATAN3"
144       gsSql2(65) = Text(62)
145       gsSql1(66) = "CATATAN4"
146       gsSql2(66) = Text(63)
147       gsSql1(67) = "ESTDATE"
148       gsSql2(67) = Format(Text(59), "MM/DD/YYYY")
149       gsSql1(68) = "MECH"
150       gsSql2(68) = nMech
151       gsSql1(69) = "KODETIPE"
          '' ===============================================================
          '' 2018-04-23 kodetipe diisi dengan nobookingapp dan kode program
          '' ===============================================================
152       If NulSID(txtNoBookingApp.Text) <> "" Or NulSID(txtKodeProgramApp.Text) <> "" Then
153           gsSql2(69) = NulSID(txtNoBookingApp.Text) & "#|#" & IIf(NulSID(txtKodeProgramApp.Text) = "", "kosong", NulSID(txtKodeProgramApp.Text))
154       Else
155           gsSql2(69) = Trim(Text(71))
156       End If
          
157       gsSql1(70) = "AMBIL"
158       gsSql2(70) = IIf(chkAmbil = Checked, "1", "0")
159       gsSql1(71) = "CODE"
160       gsSql2(71) = IIf(chEM = Checked, "1", "0")
161       gsSql1(72) = "KD_K"
162       gsSql2(72) = IIf(chSBNP = Checked, "1", "0")
163       gsSql1(73) = "KIRIM"
164       gsSql2(73) = IIf(chkKirim = Checked, "1", "0")
165       gsSql1(74) = "FS_KD_WIL"
166       gsSql2(74) = Trim(Text(44))
167       gsSql1(75) = "DATESERAH"
168       gsSql2(75) = Format(Text(84), "MM/DD/YYYY")
169       gsSql1(76) = "TIMESERAH"
170       gsSql2(76) = Trim(Text(83))
          
171       For i = 1 To 5
172           gsSql1(76 + i) = "CMPL" & i
              'gsSql2(76 + i) = Trim(text(71 + i))
173           gsSql2(76 + i) = Trim(Text1(-1 + i))
174       Next
          
175       If Trim(Text(1)) = "" Then
176           gsSql1(82) = "FRONT"
177           gsSql2(82) = Text(26)
178       End If
          
          ''2016-02-24
      '    Set oRs = objO.FetchData("Select * from BOOK WITH (NOLOCK) where nofak = '" & Text(1) & "'")
179       Set oRs = objHT.FetchDataBengkel("Select * from BOOK WITH (NOLOCK) where nofak = '" & Text(1) & "'")
180       If Not oRs.EOF Then
181           If oRs("DELIV") <> "" Then
182               gsSql1(43) = "DELIV"
183               gsSql2(43) = Trim(oRs("DELIV"))
184           End If
              
185           If DateDiff("d", Format(oRs("DATE"), "MM/DD/YYYY"), Format(TglYYYYMMDD(Text(23)), "mm/dd/yyyy")) <> 0 Then
                  'If oRs("TIME") <> Trim(text(24)) Then
186                   If Left(oRs("DELIV"), 3) = "RES" Then
187                       gsSql1(43) = "DELIV"
188                       gsSql2(43) = "RES" & CDbl(Right(Trim(oRs("DELIV")), 1)) + 1
189                   Else
190                       gsSql1(43) = "DELIV"
191                       gsSql2(43) = "RES1"
192                   End If
                  'End If
193               Call objHT.DeleteData_Bengkel("DELETE STALL WHERE KETERANGAN LIKE '%|" & Trim(Text(1)) & "|%' AND TANGGAL = '" & Format(oRs("DATE"), "MM/DD/YYYY") & "'")
194           Else
195               If oRs("TIME") <> Trim(Text(24)) Then
196                   If Left(oRs("DELIV"), 3) = "RES" Then
197                       gsSql1(43) = "DELIV"
198                       gsSql2(43) = "RES" & CDbl(Right(Trim(oRs("DELIV")), 1)) + 1
199                   Else
200                       gsSql1(43) = "DELIV"
201                       gsSql2(43) = "RES1"
202                   End If
203               End If
204               Call objHT.DeleteData_Bengkel("DELETE STALL WHERE KETERANGAN LIKE '%|" & Trim(Text(1)) & "|%' AND TANGGAL = '" & Format(oRs("DATE"), "MM/DD/YYYY") & "'")
205           End If
206           sSQL = GensSql("BOOK", True) + " where NOFAK = '" & Text(1) & "'"
207           Call updateAntrian(nopol, Text(12), Text(3), Format(Text(23), "DD/MM/YYYY"), Text(24), 2, Trim(Text(1)))
208       Else
209           isRES = False
210           sSQL = GensSql("BOOK", False)
211           nopol = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57))
212           Call updateAntrian(nopol, Text(12), Text(3), Format(Text(23), "DD/MM/YYYY"), Text(24), 1, Trim(Text(1)))
213       End If
214       Call objHT.UpdateData_Bengkel(sSQL)
215       Call objHT.UpdateData_Bengkel("UPDATE BOOK SET FS_RC ='' where NO_POLISI ='" & NoPolisi & "' and DELIV='BOK'")
          
216       If Log.hansaRenkei = 1 Then
217           If NulSID(nikSales) <> "" Then
218               Call objHT.UpdateData_Bengkel("UPDATE BOOK SET nikSales = '" & NulSID(nikSales) & "' where NOFAK = '" & Text(1) & "'")
219           End If
220       End If
221       If Log.RPLHistori = 0 Then
222           Set oRs = objO.FetchData("SELECT * FROM T_REGULAR WITH (NOLOCK)")
223       Else
224           Set oRs = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_REGULAR WITH (NOLOCK)")
225       End If
226       If Not oRs.BOF Then
227           Do While Not oRs.EOF
228               sSQL = "UPDATE BOOK SET FS_RC ='" & Trim(oRs("FS_RC")) & "' where NO_POLISI ='" & NoPolisi & "' and DELIV='BOK'" _
                       & " AND REP1+ REP2+ REP3+ REP4+ REP5+ REP6+ REP7 like '% " & Trim(oRs("FS_KETERANGAN")) & "[^0]%'"
229               Call objHT.UpdateData_Bengkel(sSQL)
230               oRs.MoveNext
231           Loop
232       End If
233       If oRs.State <> adStateClosed Then oRs.Close

234      On Error GoTo 0
235      Exit Sub

SaveDataBooking_Error:
236       erTranAll = True
237       ErrMsg.judul = ""
238       ErrMsg.message = "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure SaveDataBooking of Form frmBooking")
239       ErrMsg.show
          
      '    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure SaveDataBooking of Form frmBooking")
End Sub

Private Sub PanggilDataBooking(NoPolisi As String, sDeliv As String)
          Dim j As Integer
          Dim rsHlp As ADODB.Recordset
          Dim rsHlp2 As ADODB.Recordset
          
1        On Error GoTo PanggilDataBooking_Error

2         sSQL = "Select * from BOOK WITH (NOLOCK) where NO_POLISI = '" & NoPolisi & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE' ORDER BY DELIV, [date] desc"
3         Set rsHlp2 = objO.FetchData(sSQL)
4         If Not rsHlp2.BOF Then
5             Set rsHlp = objO.FetchData("SELECT no_chassis FROM jobDta WITH (NOLOCK) where no_chassis = '" & rsHlp2("NO_CHASSIS") & "' AND DELIV <> 'BTL' AND RPR_COND <> 'K' AND TPSVC <> 'W' ORDER BY no_WO desc")
6             If Not rsHlp.EOF Then
7                 If MsgBox("Transaksi Lama belum closing !!!" & vbCr & "Tetap akan melanjutkan ?", vbYesNo, "Closing") = vbYes Then
8                     GoTo lanjut
9                 Else
10                    GoTo LEWAT
11                End If
12            Else
13                GoTo lanjut
14            End If
lanjut:
                     ' j = ""
                      
                      
15                    nikSales = NulSID(rsHlp2!nikSales)
16                    tglAwalBook = Format(NulSID(rsHlp2!Date), "dd/mm/yyyy")
17                    txtKodeLgnLama.Text = Nul(rsHlp2("FS_KD_PLG"))
18                    Set rsHlp = objO.FetchData("select IDCust from SingleIDOri..Bridging with (nolock) where KodeLgn = '" & Nul(rsHlp2("FS_KD_PLG")) & "' and KodeDealer = '" & Right(gDb, 2) & "'")
19                    If Not rsHlp.EOF Then
20                        Text(2) = Nul(rsHlp("IDCust"))
21                    Else
22                        Text(2) = Nul(rsHlp2("FS_KD_PLG"))
23                    End If
                      
24                    Text(44) = Nul(rsHlp2("FS_KD_WIL"))
25                    Text(3) = Nul(rsHlp2("CUSTOMER"))
26                    Text(4) = Nul(rsHlp2("CUST_ADRES"))
27                    Text(5) = Nul(rsHlp2("KOTA"))
28                    Text(6) = Nul(rsHlp2("GSM"))
29                    Text(7) = Nul(rsHlp2("CDMA"))
30                    Text(8) = Nul(rsHlp2("TELPON"))
31                    Text(9) = Nul(rsHlp2("FS_TH_MOBIL"))
32                    Text(10) = rsHlp2("MDL")
33                    Text(11) = rsHlp2("MODEL")
34                        Call CexChassis_Discon(Text(11))
35                        If isOnCarClass = False Then
36                            MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
37                        End If
38                    Text(12) = rsHlp2("FMDL")
39                    Text(13) = rsHlp2("COLOR")
40                    Text(14) = rsHlp2("FS_KD_DEALER")
41                    Text(15) = rsHlp2("FS_NM_DEALER")
42                    Text(16).Mask = ""
43                    Text(16) = TglDDMMYYYY(IIf(rsHlp2("DOTAM") = "", "01/01/1900", rsHlp2("DOTAM")))
44                    Text(17).Mask = ""
45                    Text(17) = TglDDMMYYYY(IIf(rsHlp2("DDATE") = "", "01/01/1900", rsHlp2("DDATE")))
46                    Text(41) = rsHlp2("NO_CHASSIS")
47                    gNoRangka = rsHlp2("NO_CHASSIS")
48                    cmdStall.Tag = ""
49                    cmdStall.BackColor = &H8000000F
                      
50                    If Trim(rsHlp2("DELIV")) = "BOK" Or Left(Trim(rsHlp2("DELIV")), 3) = "RES" Then
51                        Text(1) = rsHlp2("NOFAK")
52                        Text(18) = TglDDMMYYYY(rsHlp2("DATE_BOOK"))
53                        Text(19) = Format(rsHlp2("TIME_D"), "HH:NN")
54                        Text(23) = TglDDMMYYYY(rsHlp2("DATE"))
55                        Text(24) = Format(rsHlp2("TIME"), "HH:NN")
                          
56                        Text(23).Enabled = True
57                        Text(24).Enabled = True
58                        If Trim(rsHlp2("DELIV")) = "RES3" Then
59                            MsgBox "Sudah di Reschedule 3x, tidak dapat di reschedule kembali." & vbCr & "Batalkan Booking, buat booking baru ."
60                            Text(23).Enabled = False
61                            Text(24).Enabled = False
62                        End If
                          
                          ''20190930 | req mb thania | 2.  Untuk booking, ketika edit minta ditambahkan pengecekan, apakah tgl booking (janji kedatangan) sama dengan tgl server, jika sama, maka user hanya bisa mengedit jam booking. Ini untuk keperluan jika ternyata customer minta reschedule jam lain krn macet / dalam perjalanan.
                          '' --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                          '' 20191003 | req mb thania | tanggal dibuka kembali karena klo reschedule beda tanggal
                          ''------------------------------------------------------------------------------------------------------------
      '                    If Format(rsHlp2("DATE"), "dd/mm/yyyy") = Format(isnow, "dd/mm/yyyy") Then
      '                        Text(23).Enabled = False
      '                    End If
                          '' --------------------------------------------------------------------------------
                          
63                        Text(25) = rsHlp2("PAID")
64                        Text(26) = gSA
65                        Text(27) = rsHlp2("REQ1")
66                        Text(28) = rsHlp2("REQ2")
67                        Text(29) = rsHlp2("REQ3")
68                        Text(30) = rsHlp2("REQ4")
69                        Text(31) = rsHlp2("REQ5")
                          
                          ' ISI KELUHAN
                          Dim hlp As String
70                        For i = 1 To 5
71                            hlp = "CMPL" & i
72                            Text1(-1 + i) = rsHlp2(hlp)
73                        Next
                          
74                        Text(39) = NulA(Format(rsHlp2("TOTLABOUR"), "#0,#"))
75                        Text(40) = NulA(Format(rsHlp2("TOTPART"), "#0,#"))
76                        Text(52) = NulA(Format(rsHlp2("TOTSUBLET"), "#0,#"))
77                        Text(55) = NulA(Format(rsHlp2("TOTOLI"), "#0,#"))
78                        Text(53) = Format(CDbl(Text(39) + CDbl(Text(40)) + CDbl(Text(52))) + CDbl(Text(55)), "#0,#")
                          
79                        Text(37) = Trim(rsHlp2("REMARK1"))
80                        Text(38) = Trim(rsHlp2("SUGGEST"))
81                        Text(75) = Trim(rsHlp2("KARPLG"))
82                        Text(60) = Trim(rsHlp2("CATATAN1"))
83                        Text(61) = Trim(rsHlp2("CATATAN2"))
84                        Text(62) = Trim(rsHlp2("CATATAN3"))
85                        Text(63) = Trim(rsHlp2("CATATAN4"))
                          
                          '' ===============================================
                          '' 2018-04-23 load nobooking app dan kodetipe app
                          '' ===============================================
86                        If NulSID(rsHlp2("KODETIPE")) <> "" Then
                              '' isi nobookingapp dan kodeprogram
87                            If InStr(1, NulSID(rsHlp2("KODETIPE")), "#|#") > 0 Then
88                                txtNoBookingApp.Text = Left(NulSID(rsHlp2("KODETIPE")), InStr(1, NulSID(rsHlp2("KODETIPE")), "#|#") - 1)
89                                noBookingAppOnline = txtNoBookingApp.Text
90                                If NulSID(Mid(NulSID(rsHlp2("KODETIPE")), InStr(1, NulSID(rsHlp2("KODETIPE")), "#|#") + 3, Len(NulSID(rsHlp2("KODETIPE"))))) <> "KOSONG" Then
91                                    txtKodeProgramApp.Text = NulSID(Mid(NulSID(rsHlp2("KODETIPE")), InStr(1, NulSID(rsHlp2("KODETIPE")), "#|#") + 3, Len(NulSID(rsHlp2("KODETIPE")))))
92                                End If
93                            End If
94                        End If
                          ' TAMPIL CONDITION
                          Dim kCond As String
                          Dim kBant As String
                          Dim rsTmp As ADODB.Recordset
95                        Set rsTmp = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & rsHlp2("NOFAK") & "' ")
96                        If Not rsTmp.EOF Then
97                            For i = 1 To 6
98                                kBant = ""
99                                kCond = ""
100                               hlp = "COND" & i
101                               If rsTmp(hlp) <> "" Then
102                                   If Left(rsTmp(hlp), 3) = "L| " Then
103                                       kBant = Mid(rsTmp(hlp), 4, Len(rsTmp(hlp)) - 3)
104                                       kCond = "LAINNYA"
105                                   Else
106                                       kCond = Trim(rsTmp(hlp))
107                                       kBant = ""
108                                   End If
109                               End If
110                               Text(45 + i) = kCond
111                               bKondisi(i - 1) = kBant
112                           Next
113                       End If
114                       Text(54).Visible = False
115                       rsTmp.Close
                              
116                       chkRtn = IIf(rsHlp2("FN_RETURN") = "1", Checked, Unchecked)
117                       chkThs = IIf(rsHlp2("THS") = "1", Checked, Unchecked)
118                       chkAmbil = IIf(rsHlp2("AMBIL") = "1", Checked, Unchecked)
119                       chkKirim = IIf(rsHlp2("KIRIM") = "1", Checked, Unchecked)
120                       chEM = IIf(rsHlp2("CODE") = "1", Checked, Unchecked)
121                       chSBNP = IIf(rsHlp2("KD_K") = "1", Checked, Unchecked)
                          
122                       dGrid.TextMatrix(1, 1) = rsHlp2("REP1")
123                       dGrid.TextMatrix(2, 1) = rsHlp2("REP2")
124                       dGrid.TextMatrix(3, 1) = rsHlp2("REP3")
125                       dGrid.TextMatrix(4, 1) = rsHlp2("REP4")
126                       dGrid.TextMatrix(5, 1) = rsHlp2("REP5")
127                       dGrid.TextMatrix(6, 1) = rsHlp2("REP6")
128                       dGrid.TextMatrix(7, 1) = rsHlp2("REP7")
                          
129                       dGrid.TextMatrix(1, 2) = IIf(dGrid.TextMatrix(1, 1) <> "", rsHlp2("RTE1"), "")
130                       dGrid.TextMatrix(2, 2) = IIf(dGrid.TextMatrix(2, 1) <> "", rsHlp2("RTE2"), "")
131                       dGrid.TextMatrix(3, 2) = IIf(dGrid.TextMatrix(3, 1) <> "", rsHlp2("RTE3"), "")
132                       dGrid.TextMatrix(4, 2) = IIf(dGrid.TextMatrix(4, 1) <> "", rsHlp2("RTE4"), "")
133                       dGrid.TextMatrix(5, 2) = IIf(dGrid.TextMatrix(5, 1) <> "", rsHlp2("RTE5"), "")
134                       dGrid.TextMatrix(6, 2) = IIf(dGrid.TextMatrix(6, 1) <> "", rsHlp2("RTE6"), "")
135                       dGrid.TextMatrix(7, 2) = IIf(dGrid.TextMatrix(7, 1) <> "", rsHlp2("RTE7"), "")
                          
136                       baris = 1
137                       nJRate = 0
138                       For j = 1 To 7
139                           If Trim(dGrid.TextMatrix(j, 1)) <> "" Then
140                               Set rsTemp = objO.FetchData("SELECT rate,job FROM HO_flate WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(dGrid.TextMatrix(j, 1)) & "' AND MDL = '" & rsHlp2("MDL") & "' and FS_LOG_DEL = ''")
141                               If Not rsTemp.EOF Then
142                                   dGrid.TextMatrix(j, 2) = rsTemp("RATE")
143                                   dGrid.TextMatrix(j, 3) = ""
144                                   dGrid.TextMatrix(j, 4) = rsTemp("JOB")
145                               Else
146                                   dGrid.TextMatrix(j, 3) = True
147                               End If
                                  
                                 ' If Log.RPLHistori = 0 Then
148                                   nJRate = nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(dGrid.TextMatrix(j, 1)) & "' AND MDL = '" & Trim(Text(10)) & "'")
                                 ' Else
                                  '    nJRate = nJRate + objO.FetchData_Double("select Rate from [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(dGrid.TextMatrix(j, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                                  'End If
149                               baris = baris + 1
150                           End If
151                       Next
                          
152                       Set rsTemp = objO.FetchData("SELECT * FROM BOOK_SUBLET WITH (NOLOCK) WHERE NOFAK = '" & rsHlp2("NOFAK") & "'")
153                       If Not rsTemp.EOF Then
154                           i = 1
155                           barisSublet = 1
156                           Do While Not rsTemp.EOF
157                               subletGrid.TextMatrix(i, 1) = rsTemp("NAMABARANG")
158                               subletGrid.TextMatrix(i, 2) = CDbl(rsTemp("QTY"))
159                               subletGrid.TextMatrix(i, 3) = rsTemp("KODESATUAN")
160                               subletGrid.TextMatrix(i, 4) = Format(CDbl(rsTemp("HARGA")), "#0,#")
                                  ''2016-02-24
161                               subletGrid.TextMatrix(i, 5) = IIf(rsTemp("FLAG") = "P", 1, 0)
                                  
                                  ''subletGrid.TextMatrix(i, 5) = rsTemp("FLAG")
162                               subletGrid.TextMatrix(i, 6) = rsTemp("ESTWAKTU")
163                               i = i + 1
164                               rsTemp.MoveNext
165                           Loop
166                       End If
167                       If rsTemp.State <> adStateClosed Then rsTemp.Close
                          'CEX 280514
168                        If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then
169                        Set rsTemp = objO.FetchDataACC("select NOSPK as NOFAK, KodeItem, Keterangan as namaBarang, SatuanNS as KodeSatuan, Qty, '' as FLAG, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID, '' as KDOLI from SOORD WITH (NOLOCK) Where NoSPK = '" & rsHlp2("NOFAK") & "' and (isnull(BO,'null') = 'null' or isnull(bo,'') <> 'CANCEL')")
170                        Else
171                        Set rsTemp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM book_part A WITH (NOLOCK) LEFT JOIN T_OLI B WITH (NOLOCK) ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & rsHlp2("NOFAK") & "'")
172                        End If
                         ' Set rsTemp = objO.FetchDataACC("select NOSPK as NOFAK, KodeItem, Keterangan as namaBarang, SatuanNS as KodeSatuan, Qty, '' as FLAG, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID, '' as KDOLI from SOORD Where NoSPK = '" & rsHlp2("NOFAK") & "' and isnull(bo,'') <> 'CANCEL'")
                         ' If rsTemp.RecordCount = 0 Then
                         '     Set rsTemp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM book_part A LEFT JOIN T_OLI B ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & rsHlp2("NOFAK") & "'")
                         ' End If
173                       If Not rsTemp.EOF Then
174                           i = 1
175                           barispart = 1
176                           Do While Not rsTemp.EOF
177                               partGrid.TextMatrix(i, 1) = rsTemp("KODEITEM")
178                               partGrid.TextMatrix(i, 2) = rsTemp("NAMABARANG")
179                               partGrid.TextMatrix(i, 3) = rsTemp("KODESATUAN")
180                               partGrid.TextMatrix(i, 4) = rsTemp("QTY")
181                               partGrid.TextMatrix(i, 6) = IIf(rsTemp("KDOLI") <> "", True, False)
182                               partGrid.TextMatrix(i, 7) = rsTemp("NAMABYSA")
183                               partGrid.TextMatrix(i, 8) = Val(NulA(rsTemp("nID")))
                                  Dim rshrgPart As ADODB.Recordset
184                                   Set rshrgPart = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM LIKE '%" & Trim(rsTemp("KODEITEM")) & "%'")
185                                   If Not rshrgPart.EOF Then
186                                       partGrid.TextMatrix(i, 5) = rshrgPart("HARGAJUALA")
187                                   End If
188                                   rshrgPart.Close
189                                   Set rshrgPart = Nothing
190                               barispart = barispart + 1
191                               i = i + 1
192                               rsTemp.MoveNext
193                           Loop
194                           NewId = (objO.FetchData_DoubleACC("select top 1 Nomor from SOORD WITH (NOLOCK) Where NOSPK = '" & rsHlp2("NOFAK") & "' Order By Nomor DESC")) + 1
195                       End If
                          
                          ' ISI STALL
                          'HAPUS TEMPSTALL
196                       Call objO.InsertData("DELETE FROM TEMPSTALL WHERE tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%'")
                          
197                       Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                                       & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%') AND " _
                                       & "tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%' ORDER BY noStall, Tanggal, JamMulai")
198                       cmdStall.Tag = ""
199                       Set rsHlp = objO.FetchData("SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%'")
200                       If Not rsHlp.EOF Then
201                           Text(58) = Trim(rsHlp("JAMMULAI"))
202                           nMech = Trim(rsHlp("MECH"))
203                       Else
204                           Text(58) = Trim(Text(24))
205                           nMech = Trim(rsHlp2("MECH"))
206                       End If
207                       rsHlp.Close
208                       Text(64) = nMech
                          
209                       If Text(58) <> ">>:>>" And Text(23) <> ">>->>->>>>" Then
210                           nWKT = hitungWaktu(CDbl(nJRate), Text(58), Text(23))
211                           Text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
212                           Text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
213                           Text(84) = TglDDMMYYYY(rsHlp2("DATESERAH"))
214                           Text(83) = Trim(rsHlp2("TIMESERAH"))
215                           Call hitungEstimasiSelesai
216                       End If
217                   Else
218                       If PanggilDataCust(NoPolisi) Then
219                           cmdStall.Tag = "belum"
220                       End If
221                   End If
222                   lblEdit.Visible = True
223       If Text(1) = "" Then Call cmdBersih_Click
                  'End If
LEWAT:
          
224       End If
          
225       If Log.RPLHistori = 0 Then
226           sSQL = "select * from MODEL WITH (NOLOCK) where MDL ='" & Text(10) & "'"
227       Else
228           sSQL = "select * from [" & dbReplikasiHO & "]..MODEL WITH (NOLOCK) where MDL ='" & Text(10) & "'"
229       End If
230       Set rsHlp2 = objO.FetchData(sSQL)
231       If rsHlp2.BOF Then
232           gGroupModel = ""
233       Else
234           gGroupModel = rsHlp2("FS_GRP_MDL")
235       End If
236       If rsHlp2.State <> adStateClosed Then rsHlp2.Close
237       If Log.DIDX = 1 Then
238            sSQL = "select * from BookDX WITH (NOLOCK) where NOFAK ='" & Text(1) & "'"
239           Set rsHlp2 = objO.FetchData(sSQL)
240           If Not rsHlp2.EOF Then
241               txt(11).Text = NulSID(rsHlp2!DelivKodePos)
242               txt(63).Text = NulSID(rsHlp2!DelivPropinsi)
243               txt(64).Text = NulSID(rsHlp2!DelivKota)
244               txt(65).Text = NulSID(rsHlp2!DelivKecamatan)
245               txt(66).Text = NulSID(rsHlp2!DelivKelurahan)
246               txt(8).Text = NulSID(rsHlp2!DelivKodePos)
247               Text2(0).Text = NulSID(rsHlp2!DscMName)
248               Text2(1).Text = NulSID(rsHlp2!DscMHP)
249               Text(88).Text = NulSID(rsHlp2!EstbyWAC)
250           End If
251       End If
252      On Error GoTo 0
253      Exit Sub

PanggilDataBooking_Error:

254       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PanggilDataBooking of Form frmBooking")
End Sub

Private Sub PanggilDataBooking2(NoPolisi As String, NoFak As String, sDeliv As String)
          Dim j As Integer
          Dim rsHlp As ADODB.Recordset
          Dim rsHlp2 As ADODB.Recordset
          
1        On Error GoTo PanggilDataBooking2_Error

2         sSQL = "Select * from BOOK WITH (NOLOCK) where NO_POLISI = '" & NoPolisi & "' AND nofak = '" & NoFak & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE' ORDER BY DELIV, [date] desc"
3         Set rsHlp2 = objO.FetchData(sSQL)
4         If Not rsHlp2.BOF Then
5       Set rsHlp = objO.FetchData("SELECT no_chassis FROM jobDta WITH (NOLOCK) where no_chassis = '" & rsHlp2("NO_CHASSIS") & "' AND DELIV <> 'BTL' AND RPR_COND <> 'K' AND TPSVC <> 'W' ORDER BY no_WO desc")
6       If Not rsHlp.EOF Then
7           If MsgBox("Transaksi Lama belum closing !!!" & vbCr & "Tetap akan melanjutkan ?", vbYesNo, "Closing") = vbYes Then
8               GoTo lanjut
9           Else
10              GoTo LEWAT
11          End If
12      Else
13          GoTo lanjut
14      End If
lanjut:
15              nikSales = NulSID(rsHlp2!nikSales)
16              tglAwalBook = Format(NulSID(rsHlp2!Date), "dd/mm/yyyy")
                ''Text(2) = Nul(rsHlp2("FS_KD_PLG"))
17              txtKodeLgnLama.Text = Nul(rsHlp2("FS_KD_PLG"))
18              Set rsHlp = objO.FetchData("select IDCust from SingleIDOri..Bridging with (nolock) where KodeLgn = '" & Nul(rsHlp2("FS_KD_PLG")) & "' and KodeDealer = '" & Right(gDb, 2) & "'")
19              If Not rsHlp.EOF Then
20                Text(2) = Nul(rsHlp("IDCust"))
21              Else
22                Text(2) = Nul(rsHlp2("FS_KD_PLG"))
23              End If
24              Text(44) = Nul(rsHlp2("FS_KD_WIL"))
25              Text(3) = Nul(rsHlp2("CUSTOMER"))
26              Text(4) = Nul(rsHlp2("CUST_ADRES"))
27              Text(5) = Nul(rsHlp2("KOTA"))
28              Text(6) = Nul(rsHlp2("GSM"))
29              Text(7) = Nul(rsHlp2("CDMA"))
30              Text(8) = Nul(rsHlp2("TELPON"))
31              Text(9) = Nul(rsHlp2("FS_TH_MOBIL"))
32              Text(10) = rsHlp2("MDL")
33              Text(11) = rsHlp2("MODEL")
34                  Call CexChassis_Discon(Text(11))
35                  If isOnCarClass = False Then
36                      MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
37                  End If
38              Text(12) = rsHlp2("FMDL")
39              Text(13) = rsHlp2("COLOR")
40              Text(14) = rsHlp2("FS_KD_DEALER")
41              Text(15) = rsHlp2("FS_NM_DEALER")
42              Text(16).Mask = ""
43              Text(16) = TglDDMMYYYY(IIf(rsHlp2("DOTAM") = "", "01/01/1900", rsHlp2("DOTAM")))
44              Text(17).Mask = ""
45              Text(17) = TglDDMMYYYY(IIf(rsHlp2("DDATE") = "", "01/01/1900", rsHlp2("DDATE")))
46              Text(41) = rsHlp2("NO_CHASSIS")
47              gNoRangka = rsHlp2("NO_CHASSIS")
48              cmdStall.Tag = ""
49              cmdStall.BackColor = &H8000000F
                
50              If Trim(rsHlp2("DELIV")) = "BOK" Or Left(Trim(rsHlp2("DELIV")), 3) = "RES" Then
51                  Text(1) = rsHlp2("NOFAK")
52                  Text(18) = TglDDMMYYYY(rsHlp2("DATE_BOOK"))
53                  Text(19) = Format(rsHlp2("TIME_D"), "HH:NN")
54                  Text(23) = TglDDMMYYYY(rsHlp2("DATE"))
55                  Text(24) = Format(rsHlp2("TIME"), "HH:NN")
                    
56                  Text(23).Enabled = True
57                  Text(24).Enabled = True
58                  If Trim(rsHlp2("DELIV")) = "RES3" Then
59                      MsgBox "Sudah di Reschedule 3x, tidak dapat di reschedule kembali." & vbCr & "Batalkan Booking, buat booking baru ."
60                      Text(23).Enabled = False
61                      Text(24).Enabled = False
62                  End If
                      
                      ''20190930 | req mb thania | 2.  Untuk booking, ketika edit minta ditambahkan pengecekan, apakah tgl booking (janji kedatangan) sama dengan tgl server, jika sama, maka user hanya bisa mengedit jam booking. Ini untuk keperluan jika ternyata customer minta reschedule jam lain krn macet / dalam perjalanan.
                      '' --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                      '' 20191003 | req mb thania | tanggal dibuka kembali karena klo reschedule beda tanggal
                      ''------------------------------------------------------------------------------------------------------------
      '                If Format(rsHlp2("DATE"), "dd/mm/yyyy") = Format(isnow, "dd/mm/yyyy") Then
      '                    Text(23).Enabled = False
      '                End If
                      '' --------------------------------------------------------------------------------
                      
63                  Text(25) = rsHlp2("PAID")
64                  Text(26) = gSA
65                  Text(27) = rsHlp2("REQ1")
66                  Text(28) = rsHlp2("REQ2")
67                  Text(29) = rsHlp2("REQ3")
68                  Text(30) = rsHlp2("REQ4")
69                  Text(31) = rsHlp2("REQ5")
                    
                    ' ISI KELUHAN
                    Dim hlp As String
70                  For i = 1 To 5
71                      hlp = "CMPL" & i
72                      Text1(-1 + i) = rsHlp2(hlp)
73                  Next
                    
74                  Text(39) = NulA(Format(rsHlp2("TOTLABOUR"), "#0,#"))
75                  Text(40) = NulA(Format(rsHlp2("TOTPART"), "#0,#"))
76                  Text(52) = NulA(Format(rsHlp2("TOTSUBLET"), "#0,#"))
77                  Text(55) = NulA(Format(rsHlp2("TOTOLI"), "#0,#"))
78                  Text(53) = Format(CDbl(Text(39) + CDbl(Text(40)) + CDbl(Text(52))) + CDbl(Text(55)), "#0,#")
                    
79                  Text(37) = Trim(rsHlp2("REMARK1"))
80                  Text(38) = Trim(rsHlp2("SUGGEST"))
81                  Text(75) = Trim(rsHlp2("KARPLG"))
82                  Text(60) = Trim(rsHlp2("CATATAN1"))
83                  Text(61) = Trim(rsHlp2("CATATAN2"))
84                  Text(62) = Trim(rsHlp2("CATATAN3"))
85                  Text(63) = Trim(rsHlp2("CATATAN4"))
                    
                    '' ===============================================
                    '' 2018-04-23 load nobooking app dan kodetipe app
                    '' ===============================================
86                  If NulSID(rsHlp2("KODETIPE")) <> "" Then
                        '' isi nobookingapp dan kodeprogram
87                      If InStr(1, NulSID(rsHlp2("KODETIPE")), "#|#") > 0 Then
88                          txtNoBookingApp.Text = Left(NulSID(rsHlp2("KODETIPE")), InStr(1, NulSID(rsHlp2("KODETIPE")), "#|#") - 1)
89                          noBookingAppOnline = txtNoBookingApp.Text
90                          If NulSID(Mid(NulSID(rsHlp2("KODETIPE")), InStr(1, NulSID(rsHlp2("KODETIPE")), "#|#") + 3, Len(NulSID(rsHlp2("KODETIPE"))))) <> "KOSONG" Then
91                              txtKodeProgramApp.Text = NulSID(Mid(NulSID(rsHlp2("KODETIPE")), InStr(1, NulSID(rsHlp2("KODETIPE")), "#|#") + 3, Len(NulSID(rsHlp2("KODETIPE")))))
92                          End If
93                      End If
94                  End If
                    
                    ' TAMPIL CONDITION
                    Dim kCond As String
                    Dim kBant As String
                    Dim rsTmp As ADODB.Recordset
95                  Set rsTmp = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & rsHlp2("NOFAK") & "' ")
96                  If Not rsTmp.EOF Then
97                      For i = 1 To 6
98                          kBant = ""
99                          kCond = ""
100                         hlp = "COND" & i
101                         If rsTmp(hlp) <> "" Then
102                             If Left(rsTmp(hlp), 3) = "L| " Then
103                                 kBant = Mid(rsTmp(hlp), 4, Len(rsTmp(hlp)) - 3)
104                                 kCond = "LAINNYA"
105                             Else
106                                 kCond = Trim(rsTmp(hlp))
107                                 kBant = ""
108                             End If
109                         End If
110                         Text(45 + i) = kCond
111                         bKondisi(i - 1) = kBant
112                     Next
113                 End If
114                 Text(54).Visible = False
115                 rsTmp.Close
                        
116                 chkRtn = IIf(rsHlp2("FN_RETURN") = "1", Checked, Unchecked)
117                 chkThs = IIf(rsHlp2("THS") = "1", Checked, Unchecked)
118                 chkAmbil = IIf(rsHlp2("AMBIL") = "1", Checked, Unchecked)
119                 chkKirim = IIf(rsHlp2("KIRIM") = "1", Checked, Unchecked)
                    
                    
120                 dGrid.TextMatrix(1, 1) = rsHlp2("REP1")
121                 dGrid.TextMatrix(2, 1) = rsHlp2("REP2")
122                 dGrid.TextMatrix(3, 1) = rsHlp2("REP3")
123                 dGrid.TextMatrix(4, 1) = rsHlp2("REP4")
124                 dGrid.TextMatrix(5, 1) = rsHlp2("REP5")
125                 dGrid.TextMatrix(6, 1) = rsHlp2("REP6")
126                 dGrid.TextMatrix(7, 1) = rsHlp2("REP7")
                    
127                 dGrid.TextMatrix(1, 2) = IIf(dGrid.TextMatrix(1, 1) <> "", rsHlp2("RTE1"), "")
128                 dGrid.TextMatrix(2, 2) = IIf(dGrid.TextMatrix(2, 1) <> "", rsHlp2("RTE2"), "")
129                 dGrid.TextMatrix(3, 2) = IIf(dGrid.TextMatrix(3, 1) <> "", rsHlp2("RTE3"), "")
130                 dGrid.TextMatrix(4, 2) = IIf(dGrid.TextMatrix(4, 1) <> "", rsHlp2("RTE4"), "")
131                 dGrid.TextMatrix(5, 2) = IIf(dGrid.TextMatrix(5, 1) <> "", rsHlp2("RTE5"), "")
132                 dGrid.TextMatrix(6, 2) = IIf(dGrid.TextMatrix(6, 1) <> "", rsHlp2("RTE6"), "")
133                 dGrid.TextMatrix(7, 2) = IIf(dGrid.TextMatrix(7, 1) <> "", rsHlp2("RTE7"), "")
                    
134                 baris = 1
135                 nJRate = 0
136                 jobQuery = ""
137                 For j = 1 To 7
138                     If Trim(dGrid.TextMatrix(j, 1)) <> "" Then
139                         Set rsTemp = objO.FetchData("SELECT rate,job FROM HO_flate WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(dGrid.TextMatrix(j, 1)) & "' AND MDL = '" & rsHlp2("MDL") & "' and FS_LOG_DEL = ''")
140                         If Not rsTemp.EOF Then
141                             dGrid.TextMatrix(j, 2) = rsTemp("RATE")
142                             dGrid.TextMatrix(j, 3) = ""
143                             dGrid.TextMatrix(j, 4) = rsTemp("JOB")
144                             jobQuery = jobQuery & " B.FS_ITEM = '" & Trim(dGrid.TextMatrix(j, 1)) & "' AND "
145                         Else
146                             dGrid.TextMatrix(j, 3) = True
147                         End If
                            
                           ' If Log.RPLHistori = 0 Then
148                           nJRate = nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(dGrid.TextMatrix(j, 1)) & "' AND MDL = '" & Trim(Text(10)) & "'")
                           ' Else
                           '   nJRate = nJRate + objO.FetchData_Double("select Rate from [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(dGrid.TextMatrix(j, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                           ' End If
149                         baris = baris + 1
150                     End If
151                 Next
                    
152                 If NulSID(jobQuery) <> "" Then
153                   jobQuery = Left(jobQuery, Len(jobQuery) - 4)
154                 End If
155                   If Log.isDSS = 1 And NulSID(jobQuery) <> "" Then
                              Dim strSql As String
                              
156                           strSql = "SELECT B.TIPEBP, C.SKILLMIN, E.LEVEL AS LvlTech, C.STALLMIN, D.level AS LvlStall, C.LiftNeed " & _
                                "FROM  HO_Flate B " & _
                                "LEFT JOIN MASTERDSS C ON B.TIPEBP = C.KODEDSS " & _
                                "LEFT JOIN LevelStall D ON C.StallMin = D.KodeStall " & _
                                "LEFT JOIN T_MECHLEV E ON C.SkillMin = E.LEVELMECH " & _
                                "WHERE B.MDL = '" & Trim(Text(10)) & "' AND (" & jobQuery & ") AND ISNULL(C.SkillMin,'') <> '' AND dss = 1 " & _
                                "order by LvlStall desc ,LvlTech desc "
                        '-------------------------------------------------
              
157                           Set rsTemp = objO.FetchData(strSql)
158                           sgnDSS.Caption = ""
159                           If rsTemp.EOF Then
160                               sgnDSS.Caption = "NON DSS"
161                           Else
162                               sgnDSS.Caption = "DSS"
163                           End If
164                   End If
                      
165                 Set rsTemp = objO.FetchData("SELECT * FROM BOOK_SUBLET WITH (NOLOCK) WHERE NOFAK = '" & rsHlp2("NOFAK") & "'")
166                 If Not rsTemp.EOF Then
167                     i = 1
168                     barisSublet = 1
169                     Do While Not rsTemp.EOF
170                         subletGrid.TextMatrix(i, 1) = rsTemp("NAMABARANG")
171                         subletGrid.TextMatrix(i, 2) = CDbl(rsTemp("QTY"))
172                         subletGrid.TextMatrix(i, 3) = rsTemp("KODESATUAN")
173                         subletGrid.TextMatrix(i, 4) = Format(CDbl(rsTemp("HARGA")), "#0,#")
                              ''2016-09-22 || ambil data flag diconvert ke flag P atau J
174                         subletGrid.TextMatrix(i, 5) = IIf(rsTemp("FLAG") = "P", 1, 0) ''rsTemp("FLAG")
175                         subletGrid.TextMatrix(i, 6) = rsTemp("ESTWAKTU")
176                         i = i + 1
177                         barisSublet = barisSublet + 1
178                         rsTemp.MoveNext
179                     Loop
180                 End If
181                 If rsTemp.State <> adStateClosed Then rsTemp.Close
                    'CEX 280514
182                  If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then
                          ''Set rsHrg = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & NulSID(Trim(oRs2("ITEM_PART"))) & "' and B.Jenis = 'OLI'")
                          ''Set rsTemp = objO.FetchDataACC("select NOSPK as NOFAK, KodeItem, Keterangan as namaBarang, SatuanNS as KodeSatuan, Qty, '' as FLAG, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID, '' as KDOLI from SOORD WITH (NOLOCK) Where NoSPK = '" & rsHlp2("NOFAK") & "' and (isnull(BO,'null') = 'null' or isnull(bo,'') <> 'CANCEL')")
183                       Set rsTemp = objO.FetchDataACC("select a.NOSPK as NOFAK, a.KodeItem, a.Keterangan as namaBarang, a.SatuanNS as KodeSatuan, a.Qty, '' as FLAG, a.TglPerlu as ETA, '' as Keterangan, a.NamaBYSA, a.Nomor as nID, case when isnull(B.Jenis,'') = 'OLI' then '1' else '' end as KDOLI from SOORD a WITH (NOLOCK) left join inmst c WITH (NOLOCK) on a.kodeitem = c.kodeitem left Join ProfilACC..SCC B WITH (NOLOCK) On c.Scc_code = B.SCC_Code Where a.NoSPK = '" & rsHlp2("NOFAK") & "' and (isnull(a.BO,'null') = 'null' or isnull(a.bo,'') <> 'CANCEL')")
184                  Else
185                  Set rsTemp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM book_part A WITH (NOLOCK) LEFT JOIN T_OLI B WITH (NOLOCK) ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & rsHlp2("NOFAK") & "'")
186                  End If
                   ' Set rsTemp = objO.FetchDataACC("select NOSPK as NOFAK, KodeItem, Keterangan as namaBarang, SatuanNS as KodeSatuan, Qty, '' as FLAG, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID, '' as KDOLI from SOORD Where NoSPK = '" & rsHlp2("NOFAK") & "' and isnull(bo,'') <> 'CANCEL'")
                   ' If rsTemp.RecordCount = 0 Then
                   '     Set rsTemp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM book_part A LEFT JOIN T_OLI B ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & rsHlp2("NOFAK") & "'")
                   ' End If
187                 If Not rsTemp.EOF Then
188                     i = 1
189                     barispart = 1
                        
190                 Text(40) = 0 'NulA(Format(rsHlp2("TOTPART"), "#0,#"))
191                 Text(55) = 0 'NulA(Format(rsHlp2("TOTOLI"), "#0,#"))
192                     Do While Not rsTemp.EOF
193                         partGrid.TextMatrix(i, 1) = rsTemp("KODEITEM")
194                         partGrid.TextMatrix(i, 2) = rsTemp("NAMABARANG")
195                         partGrid.TextMatrix(i, 3) = rsTemp("KODESATUAN")
196                         partGrid.TextMatrix(i, 4) = rsTemp("QTY")
197                         partGrid.TextMatrix(i, 6) = IIf(rsTemp("KDOLI") <> "", True, False)
198                         partGrid.TextMatrix(i, 7) = rsTemp("NAMABYSA")
199                         partGrid.TextMatrix(i, 8) = Val(NulA(rsTemp("nID")))
                            Dim rshrgPart As ADODB.Recordset
200                             Set rshrgPart = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM LIKE '%" & Trim(rsTemp("KODEITEM")) & "%'")
201                             If Not rshrgPart.EOF Then
202                                 partGrid.TextMatrix(i, 5) = rshrgPart("HARGAJUALA")
203                                 If rsTemp("KDOLI") <> "" Then
204                                   Text(55) = Text(55) + (rshrgPart("HARGAJUALA") * rsTemp("QTY"))
205                                 Else
206                                   Text(40) = Text(40) + (rshrgPart("HARGAJUALA") * rsTemp("QTY"))
207                                 End If
208                             End If
209                             rshrgPart.Close
210                             Set rshrgPart = Nothing
211                         barispart = barispart + 1
212                         i = i + 1
213                         rsTemp.MoveNext
214                     Loop
215                     Text(40) = NulA(Format(Text(40), "#0,#"))
216                     Text(55) = NulA(Format(Text(55), "#0,#"))
217                     Text(53) = Format(CDbl(Text(39) + CDbl(Text(40)) + CDbl(Text(52))) + CDbl(Text(55)), "#0,#")
                    
218                     NewId = (objO.FetchData_DoubleACC("select top 1 Nomor from SOORD WITH (NOLOCK) Where NOSPK = '" & rsHlp2("NOFAK") & "' Order By Nomor DESC")) + 1
219                 End If
                    
                    ' ISI STALL
220                 Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                                 & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%') AND " _
                                 & "tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%' ORDER BY noStall, Tanggal, JamMulai")
221                 cmdStall.Tag = ""
222                 Set rsHlp = objO.FetchData("SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(Text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & nopol & " |%'")
223                 If Not rsHlp.EOF Then
224                     Text(58) = Trim(rsHlp("JAMMULAI"))
225                     nMech = Trim(rsHlp("MECH"))
226                 Else
227                     Text(58) = Trim(Text(24))
228                     nMech = Trim(rsHlp2("MECH"))
229                 End If
230                 rsHlp.Close
231                 Text(64) = nMech
                    
232                 chEM = IIf(rsHlp2("CODE") = "1", Checked, Unchecked)
233                 chSBNP = IIf(rsHlp2("KD_K") = "1", Checked, Unchecked)
                    
234                 If Text(58) <> ">>:>>" And Text(23) <> ">>->>->>>>" Then
235                     nWKT = hitungWaktu(CDbl(nJRate), Text(58), Text(23))
236                     Text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
237                     Text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
238                     Text(84) = TglDDMMYYYY(rsHlp2("DATESERAH"))
239                     Text(83) = Trim(rsHlp2("TIMESERAH"))
240                     Call hitungEstimasiSelesai
241                 End If
242             Else
243                 If PanggilDataCust(NoPolisi) Then
244                     cmdStall.Tag = "belum"
245                 End If
246             End If
247             lblEdit.Visible = True
248             cmdHistory.Enabled = True
249             Command1.Enabled = True
250       If Text(1) = "" Then Call cmdBersih_Click
            'End If
LEWAT:
          
251       End If
          
252       If Log.RPLHistori = 0 Then
253           sSQL = "select * from MODEL WITH (NOLOCK) where MDL ='" & Text(10) & "'"
254       Else
255           sSQL = "select * from [" & dbReplikasiHO & "]..MODEL WITH (NOLOCK) where MDL ='" & Text(10) & "'"
256       End If
257       Set rsHlp2 = objO.FetchData(sSQL)
258       If rsHlp2.BOF Then
259         gGroupModel = ""
260       Else
261         gGroupModel = rsHlp2("FS_GRP_MDL")
262       End If
263       If rsHlp2.State <> adStateClosed Then rsHlp2.Close
          
264       If rsHlp2.State <> adStateClosed Then rsHlp2.Close
265       If Log.DIDX = 1 Then
266            sSQL = "select * from BookDX WITH (NOLOCK) where NOFAK ='" & Text(1) & "'"
267           Set rsHlp2 = objO.FetchData(sSQL)
268           If Not rsHlp2.EOF Then
269               txt(11).Text = NulSID(rsHlp2!DelivKodePos)
270               txt(63).Text = NulSID(rsHlp2!DelivPropinsi)
271               txt(64).Text = NulSID(rsHlp2!DelivKota)
272               txt(65).Text = NulSID(rsHlp2!DelivKecamatan)
273               txt(66).Text = NulSID(rsHlp2!DelivKelurahan)
274               txt(8).Text = NulSID(rsHlp2!DelivKodePos)
275               Text2(0).Text = NulSID(rsHlp2!DscMName)
276               Text2(1).Text = NulSID(rsHlp2!DscMHP)
277           End If
278       End If
279      On Error GoTo 0
280      Exit Sub

PanggilDataBooking2_Error:

281       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PanggilDataBooking2 of Form frmBooking")
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
1          Call CekTombol(KeyCode)
2         If KeyCode = vbKeyF12 Then
3             Select Case Index
                  Case 0
4                     gMainKeyCari = Text(Index)
5                     gSearchOption = "BOOK"
6                     frmCari.show (vbModal)
7                 Case 2
8                     Call panggilCari
                      
9                 Case 10
10                    If nopol = "" Then
11                        MsgBox "Isi No Polisi kendaraan !!!"
12                        Text(0).SelLength = Len(Text(0))
13                        Text(0).SetFocus
14                        Exit Sub
15                    Else
16                        gMainKeyCari = Text(Index)
17                        gSearchOption = "MDL"
18                        frmCari.show (vbModal)
19                    End If
20                Case 13
21                    If Text(10) <> "" Then
22                        gMainKeyCari = Text(Index)
23                        gSearchOption = "WARNA"
24                        gBantuan = Trim(Text(10))
25                        frmCariSPK.show (vbModal)
26                    Else
27                        MsgBox "Isi model kendaraan !!!"
28                        Text(10).SelLength = 10
29                        Text(10).SetFocus
30                        Exit Sub
31                    End If
32                Case 32
33                    If Text(10) = "" Then
34                        MsgBox "Isi model kendaraan !!!"
35                        Text(10).SelLength = 10
36                        Text(10).SetFocus
37                        Exit Sub
38                    Else
39                        gMainKeyCari = Text(Index)
40                        gSearchOption = "JENIS_KERJA"
41                        frmCari.show (vbModal)
42                    End If
43                Case 34
44                    If Text(32) = "" Then
45                        MsgBox "Isi Tipe Job Order !!!"
46                        Text(32).SelLength = 3
47                        Text(32).SetFocus
48                        Exit Sub
49                    Else
50                        If Text(24) = ">>:>>" Or Text(23) = ">>->>->>>>" Then
51                            MsgBox "Tanggal dan Jam Booking !!!"
52                            Text(23).SelLength = Len(Text(23))
53                            Text(23).SetFocus
54                            Exit Sub
55                        Else
56                            gMainKeyCari = Text(Index)
57                            gSearchOption = "JOBLIST2"
58                            gBantuan = Text(10)
59                            gBantuan2 = Text(32)
60                            frmCari.show (vbModal)
61                        End If
62                    End If
63                Case 35
64                    If Text(10) = "" Then
65                        MsgBox "Isi model kendaraan !!!"
66                        Text(10).SelLength = 10
67                        Text(10).SetFocus
68                        Exit Sub
69                    Else
70                        gMainKeyCari = Text(Index)
71                        gSearchOption = "PART"
72                        frmCariAcess.show (vbModal)
73                    End If
74                Case 46, 47, 48, 49, 50, 51
75                    gMainKeyCari = Text(Index)
76                    gSearchOption = "KONDISI"
77                    gBantuan = Index - 45
78                    frmCari.show (vbModal)
79                Case 56, 57
80                    Text(0).SetFocus
81                    Exit Sub
82                Case 64
83                    If Text(23) <> ">>->>->>>>" Then
84                        gMainKeyCari = Text(Index)
85                        gSearchOption = "MEKANIK"
86                        If DateDiff("d", Format(Text(23), "MM/DD/YYYY"), Format(isnow, "MM/DD/YYYY")) = 0 Then
87                            gBantuan = "ORDER"
88                            gBantuan2 = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
89                            gBantuan3 = Format(isnow, "DD/MM/YYYY")
90                            gBantuan4 = ""
91                        Else
92                            gBantuan = "BOOKING"
                              Dim nTgl As String
93                            nTgl = Format(Text(23), "MM/DD/YYYY")
94                            gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
95                            gBantuan3 = Format(nTgl, "DD/MM/YYYY")
96                            gBantuan4 = ""
97                        End If
98                        frmCari.show (vbModal)
99                    End If
100               Case 79
101                   If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
102                       gMainKeyCari = Text(Index)
103                       gSearchOption = "KODESUBLETSENTRAL"
104                       gBantuan = Format(NulSID(Text(23)), "yyyy-mm-dd")
105                       gBantuan2 = NulSID(Text(10))
106                       frmCari.show (vbModal)
107                   Else
108                       gMainKeyCari = Text(Index)
109                       gSearchOption = "SUBLET"
110                       frmCari.show (vbModal)
111                   End If
112               Case 86
113                   If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
114                       gMainKeyCari = Text(Index)
115                       gSearchOption = "PROMOSUBLETSENTRAL"
116                       gBantuan = Format(NulSID(Text(23)), "yyyy-mm-dd")
117                       gBantuan2 = NulSID(Text(10))
118                       frmCari.show (vbModal)
119                   End If
120           End Select
121       End If
122       If KeyCode = vbKeyEscape Then
123           If Not simpan Then
124               If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
125                   Exit Sub
126               End If
127           End If
128           Call hapusForm
129           cmdNo.Caption = "No &Polisi"
130           Text(0).SetFocus
131       End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
    KeyAscii = UpCase(KeyAscii)
    If Index = 27 And KeyAscii = 13 Then
        Dim nCase As Boolean
        nCase = False
        Dim xSB As Integer
        For xSB = 5 To 500 Step 5
            If Val(Text(27).Text) = xSB Then
                nCase = True
                Exit For
            End If
        Next xSB
        If nCase = True Then Text(27).Text = "SERVIS BERKALA " & Trim(Text(27)) & ".000 KM"
    End If
    Select Case Index
        Case 52, 23
            KeyAscii = CekNumerik(KeyAscii)
            simpan = False
    End Select
End Sub

Private Sub PanggilDataCustBook(nPol As String)
      Dim rsHlp As ADODB.Recordset
1        On Error GoTo PanggilDataCustBook_Error

2         Set rsHlp = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NO_POLISI = '" & nPol & "' ORDER BY [DATE] DESC")
3         If Not rsHlp.EOF Then
4             Text(3) = rsHlp("CUSTOMER")
5             Text(4) = rsHlp("CUST_ADRES")
6             Text(5) = rsHlp("KOTA")
7             Text(6) = rsHlp("GSM")
8             Text(7) = rsHlp("CDMA")
9             Text(8) = rsHlp("TELPON")
10            Text(9) = rsHlp("FS_TH_MOBIL")
11            Text(10) = rsHlp("MDL")
12            Text(11) = rsHlp("MODEL")
13                        Call CexChassis_Discon(Text(11))
14                        If isOnCarClass = False Then
15                            MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
16                        End If
17            Text(12) = rsHlp("FMDL")
18            Text(13) = rsHlp("COLOR")
19        End If

20       On Error GoTo 0
21       Exit Sub

PanggilDataCustBook_Error:

22        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PanggilDataCustBook of Form frmBooking")
End Sub

Private Sub PanggilDataCustDiagnose(nopol As String, noPrediag As String)
          Dim rsDiagapproval As New ADODB.Recordset
1        On Error GoTo PanggilDataCustDiagnose_Error

2         Set rsDiagapproval = objO.FetchData("select da.*,d.no_polisi,d.* from diagdta d WITH (NOLOCK) inner join diagapproval da WITH (NOLOCK) on da.no_diag = d.no_diag where da.no_wo = '' and isnull(da.alasan_tdksetuju,'') = '' and  d.no_polisi = '" & nopol & "' and da.NO_DIAG = '" & noPrediag & "'")
          
3         If Not rsDiagapproval.EOF Then
              'jika ada data
4             no_approval = Trim(rsDiagapproval.Fields(0))
              Dim rsDiagdta As New ADODB.Recordset
              Dim rsDiagPart As New ADODB.Recordset
              Dim rsDiagDtl As New ADODB.Recordset
              
5             Set rsDiagdta = objO.FetchData("select * from diagdta WITH (NOLOCK) where no_polisi = '" & Trim(nopol) & "' and no_diag = '" & no_approval & "' order by no_diag desc")
6             Set rsDiagPart = objO.FetchData("select * from diagpart WITH (NOLOCK) where no_diag = '" & Trim(rsDiagapproval.Fields(0)) & "'")
7             Set rsDiagDtl = objO.FetchData("select cmpl1,cmpl2,cmpl3,cmpl4,cmpl5 from diagdtl WITH (NOLOCK) where no_diag = '" & Trim(rsDiagapproval.Fields(0)) & "'")
              
8             If Not rsDiagdta.EOF Then
                  
9                 isBookDiag = True
                  
                  'jika ada data berdasarkan no polisi
10                Text(2) = Nul(rsDiagdta("FS_KD_PLG"))
11                Text(44) = Nul(rsDiagdta("FS_KD_WIL"))
12                Text(3) = Nul(rsDiagdta("CUSTOMER"))
13                Text(4) = Nul(rsDiagdta("CUST_ADRES"))
14                Text(5) = Nul(rsDiagdta("KOTA"))
                  'text(6) = Nul(rsDiagdta("GSM"))
                  'text(7) = Nul(rsDiagdta("CDMA"))
                  'text(8) = Nul(rsDiagdta("TELPON"))
15                Text(9) = Nul(rsDiagdta("FS_TH_MOBIL"))
16                Text(10) = rsDiagdta("MDL")
17                Text(11) = rsDiagdta("MODEL")
18                    Call CexChassis_Discon(Text(11))
19                    If isOnCarClass = False Then
20                        MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
21                    End If
22                Text(12) = rsDiagdta("FMDL")
23                Text(13) = rsDiagdta("COLOR")
24                Text(14) = rsDiagdta("FS_KD_DEALER")
25                Text(15) = rsDiagdta("FS_NM_DEALER")
26                Text(16).Mask = ""
27                Text(16) = TglDDMMYYYY(IIf(rsDiagdta("DOTAM") = "", "01/01/1900", rsDiagdta("DOTAM")))
28                Text(17).Mask = ""
29                Text(17) = TglDDMMYYYY(IIf(rsDiagdta("DDATE") = "", "01/01/1900", rsDiagdta("DDATE")))
30                Text(41) = rsDiagdta("NO_CHASSIS")
31                gNoRangka = rsDiagdta("NO_CHASSIS")
32                cmdStall.Tag = ""
33                cmdStall.BackColor = &H8000000F
                  
34                Text(1) = rsDiagdta("NOFAK")
                  'text(18) = TglDDMMYYYY(rsDiagdta("DATE_BOOK"))
                  'text(19) = Format(rsDiagdta("TIME_D"), "HH:NN")
35                Text(23) = TglDDMMYYYY(rsDiagdta("DATE"))
36                Text(24) = Format(rsDiagdta("TIME"), "HH:NN")
                  
37                Text(23).Enabled = True
38                Text(24).Enabled = True
      '                        If Trim(rsHlp2("DELIV")) = "RES3" Then
      '                            MsgBox "Sudah di Reschedule 3x, tidak dapat di reschedule kembali." & vbCr & "Batalkan Booking, buat booking baru ."
      '                            text(23).Enabled = False
      '                            text(24).Enabled = False
      '                        End If

39                Text(26) = gSA
40                Text(27) = rsDiagdta("REQ1")
41                Text(28) = rsDiagdta("REQ2")
42                Text(29) = rsDiagdta("REQ3")
43                Text(30) = rsDiagdta("REQ4")
44                Text(31) = rsDiagdta("REQ5")
                  
                  ' ISI KELUHAN
                  Dim hlp As String
45                If Not rsDiagDtl.EOF Then
46                    For i = 1 To 5
47                        hlp = "CMPL" & i
48                        Text1(-1 + i) = rsDiagDtl(hlp)
49                    Next
50                End If
51            End If
              
52            If Not rsDiagPart.EOF Then
53                i = 1
54                barispart = 1
55                Do While Not rsDiagPart.EOF
56                    partGrid.TextMatrix(i, 1) = rsDiagPart.Fields(4) '("KODEITEM")
57                    partGrid.TextMatrix(i, 2) = rsDiagPart.Fields(5) '("NAMAITEM")
58                    partGrid.TextMatrix(i, 3) = rsDiagPart.Fields(6) '("SAT_ITEM")
59                    partGrid.TextMatrix(i, 4) = rsDiagPart.Fields(7) '("QTY")
                      'PartGrid.TextMatrix(i, 6) = IIf(rsDiagPart.Fields() <> "", True, False) '("KDOLI") <> "", True, False)
60                    partGrid.TextMatrix(i, 7) = "" 'rsDiagPart("NAMABYSA")
61                    partGrid.TextMatrix(i, 8) = barispart 'Val(NulA(rsDiagPart("nID")))
                      Dim rshrgPart As ADODB.Recordset
62                    If Trim(rsDiagPart.Fields(4)) <> "" Then
63                        Set rshrgPart = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM LIKE '%" & Trim(rsDiagPart.Fields(4)) & "%'") '("KODEITEM")) & "%'")
64                    Else
65                        Set rshrgPart = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM LIKE '%zzzzzzzzzzzzzzzzzzzzzzzzzzzzz%'") '("KODEITEM")) & "%'")
66                    End If
                      
67                    If Not rshrgPart.EOF Then
68                        partGrid.TextMatrix(i, 5) = rshrgPart("HARGAJUALA")
69                    Else
70                        partGrid.TextMatrix(i, 5) = 0
71                    End If
72                    rshrgPart.Close
73                    Set rshrgPart = Nothing
                      
                      'If oliPart Then
                      '    Text(55) = Format(CDbl(Text(55)) + hitungPart(partGrid.TextMatrix(barispart - 1, 4), partGrid.TextMatrix(barispart - 1, 5)), "#0,0")
                      'Else
74                        Text(40) = Format(CDbl(Text(40)) + hitungPart(partGrid.TextMatrix(i, 4), partGrid.TextMatrix(i, 5)), "#,0")
                      'End If
75                    Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#,0#")
76                    barispart = barispart + 1
77                    i = i + 1
78                    rsDiagPart.MoveNext
79                Loop
80                NewId = barispart '(objO.FetchData_DoubleACC("select top 1 Nomor from SOORD Where NOSPK = '" & rsHlp2("NOFAK") & "' Order By Nomor DESC")) + 1
81            End If
82        End If

83       On Error GoTo 0
84       Exit Sub

PanggilDataCustDiagnose_Error:

85        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PanggilDataCustDiagnose of Form frmBooking")
End Sub


Private Function PanggilDataCust(nPol As String) As Boolean
          Dim rsHlp3 As ADODB.Recordset
1        On Error GoTo PanggilDataCust_Error

2         On Error Resume Next
3         Set rsHlp3 = objO.FetchData("Select * from CUSTDAT WITH (NOLOCK) where NO_POLISI = '" & Trim(Replace(Replace(nPol, ">", ""), " ", "")) & "'")
4         If Not rsHlp3.EOF Then
5             rsHlp3.MoveLast
6             If rsTemp.State Then rsTemp.Close
              
7             Set rsTemp = objO.FetchData("SELECT * FROM jobdta WITH (NOLOCK) WHERE no_polisi = '" & Trim(Replace(Replace(nPol, ">", ""), " ", "")) & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL' AND TPSVC <> 'W' ORDER BY no_WO desc")
8             If rsTemp.EOF Then
balik:
9                 Text(2) = rsHlp3("FS_KD_PLG")
10                txtKodeLgnLama.Text = Nul(rsHlp3("FS_KD_PLG"))
11                Text(44) = rsHlp3("FS_KD_WIL")
12                Text(3) = rsHlp3("CUSTOMER")
13                Text(4) = rsHlp3("CUST_ADRES")
14                Text(5) = rsHlp3("KOTA")
15                Text(6) = rsHlp3("GSM")
16                Text(7) = rsHlp3("CDMA")
17                Text(8) = rsHlp3("TELPON")
18                Text(9) = rsHlp3("FS_TH_MOBIL")
19                Text(10) = rsHlp3("MDL")
20                Text(11) = rsHlp3("MODEL")
21                        Call CexChassis_Discon(Text(11))
22                        If isOnCarClass = False Then
23                            MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
24                        End If
25                Text(12) = rsHlp3("FMDL")
26                Text(13) = rsHlp3("COLOR")
27                Text(14) = rsHlp3("FS_KD_DEALER")
28                Text(15) = rsHlp3("FS_NM_DEALER")
29                Text(16) = TglDDMMYYYY(rsHlp3("DOTAM"))
30                Text(17) = TglDDMMYYYY(rsHlp3("DDATE"))
31                gNoRangka = rsHlp3("NO_CHASSIS")
32                Text(41) = gNoRangka
33                Text(26) = gSA
34                cmdHistory.Enabled = True
35                Command1.Enabled = True
36            Else
37                If cmdNo.Caption = "No &Rekom" Then
38                    Text(3).SetFocus
39                Else
                      'proses ada data no polisi dari prediagnos
                      '- lakukan pengecekan no polisi di diagapproval
                      
40                        If rsTemp("RPR_COND") <> "K" Then
41                            If Trim(cekMOU(rsHlp3("MDL"), rsHlp3("NO_CHASSIS"))) = "" Then
42                                If cekWoInvoice(Trim(rsHlp3("NO_CHASSIS"))) Then
43                                    MsgBox "WO terakhir harus ditutup !", vbCritical, "CLOSING"
44                                    Call hapusForm
45                                    Text(57).SetFocus
46                                Else
47                                    If Not cekLastIsSBIDealer(Trim(rsHlp3("NO_CHASSIS"))) Then
48                                        If Not cekLastIsPDS(Trim(rsHlp3("NO_CHASSIS"))) Then
49                                            MsgBox "WO terakhir harus ditutup !", vbCritical, "CLOSING"
50                                            Call hapusForm
51                                            Text(57).SetFocus
52                                        Else
53                                            If Not cekSBIDealerInvoice(Trim(rsHlp3("NO_CHASSIS"))) Then
54                                                MsgBox "WO terakhir harus ditutup !", vbCritical, "CLOSING"
55                                                Call hapusForm
56                                                Text(57).SetFocus
57                                            Else
58                                                GoTo balik
59                                            End If
60                                        End If
61                                    Else
62                                        GoTo balik
63                                    End If
64                                End If
65                            Else
66                                GoTo balik
67                            End If
68                    End If
69                End If
70            End If
      '        If cekDataSingleID("unitcustomer", "where norangka = '" & Nul(Text(41)) & "'") = True Then
      '            ''jika diunit customer tampilkan idcust dan kodelgn lama
      '            ''cek data bridging atas kodelgn lama sudah ada belum
      '            If cekDataSingleID("bridging", "where kodelgn = '" & txtKodeLgnLama.Text & "' and kodedealer = '" & Right(gDb, 2) & "'") = True Then
      '                ''jika di bridging ada maka di ambil data idcust
      '                Set rsTemp = objO.FetchData("SELECT idcust FROM [SingleIDOri]..bridging WITH (NOLOCK) WHERE kodelgn = '" & txtKodeLgnLama.Text & "' and kodedealer = '" & Right(gDb, 2) & "'")
      '                If Not rsTemp.EOF Then
      '                    Text(1) = Nul(rsTemp!IDCust)
      '                Else
      '                    Text(1) = ""
      '                    MsgBox "Gagal mengambil data ID Customer", vbInformation
      '                    Exit Function
      '                End If
      '            Else
      '                Call cmdKodePlg_Click
      '            End If
      '        Else
      '            ''jika unit tidak ada diarahkan ke form customer single id
      '            Call cmdKodePlg_Click
      '        End If
71            PanggilDataCust = True
72        Else
73            cmdHistory.Enabled = False
74            Command1.Enabled = False
75            PanggilDataCust = False
76        End If
77        If rsHlp3.State <> adStateClosed Then rsHlp3.Close
          
78        If Log.RPLHistori = 0 Then
79            gGroupModel = objO.FetchData_String("select FS_GRP_MDL from MODEL WITH (NOLOCK) where MDL ='" & Text(10) & "'")
80        Else
81            gGroupModel = objO.FetchData_String("select FS_GRP_MDL from [" & dbReplikasiHO & "]..MODEL WITH (NOLOCK) where MDL ='" & Text(10) & "'")
82        End If
83       On Error GoTo 0
84       Exit Function

PanggilDataCust_Error:

85        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure PanggilDataCust of Form frmBooking")
End Function

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 9
            If Text(Index) = 0 Then
                Text(Index).SelLength = Len(Text(Index))
                Text(Index).SetFocus
                Exit Sub
            End If
        Case 23
            If Text(Index) <> "" Or Text(Index) <> ">>->>->>>>" Then
                cmdWaktu.Enabled = True
            Else
                cmdWaktu.Enabled = False
            End If
    End Select
End Sub

Private Sub isCurrentAR()
1        On Error GoTo isCurrentAR_Error

2         If Trim(Text(2).Text) = "" Then Exit Sub
          Dim aBatasKr As Double
          Dim aCurrent As Double
3         aBatasKr = objO.Fetch_BatasKredit_Ar(Trim(Replace(Text(2).Text, ">", "")))

4         txtRp(1) = aBatasKr
5         aCurrent = objO.Fetch_Current_Ar(Trim(Replace(Text(2).Text, ">", "")))
6         txtRp(0) = aCurrent
7         If aCurrent <> 0 And aBatasKr <> 0 Then
8             nPersenAR = Round((aCurrent / aBatasKr) * 100, 0)
9         Else
10            nPersenAR = 0
11        End If
12        nAR2.Caption = nPersenAR & "%"
13        Call isOverdueAR

14       On Error GoTo 0
15       Exit Sub

isCurrentAR_Error:

16        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isCurrentAR of Form frmBooking")
End Sub

Private Sub isOverdueAR()
          Dim nOAR As ADODB.Recordset
1        On Error GoTo isOverdueAR_Error

2         Set nOAR = objO.FetchOverdueAR(Trim(Replace(Text(2).Text, ">", "")))
3         If nOAR.RecordCount > 0 Then
4             Set nOAR = objO.FetchOverdueAR_Data(Trim(Replace(Text(2).Text, ">", "")))
5             If nOAR.RecordCount > 0 Then
6                 nAR2.Caption = nAR2.Caption & "  ( AR OVERDUE )"
7             End If
8         End If

9        On Error GoTo 0
10       Exit Sub

isOverdueAR_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isOverdueAR of Form frmBooking")
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
          Dim boleh As Boolean
          Dim birthDay As String
          Dim rsSearch As ADODB.Recordset
          
1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 57
3                 nopol = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57))
4                 Text(78) = Trim(nopol)
5                 If Trim(nopol) <> "" Then
6                     isBookDiag = False
7                     no_approval = ""
8                     boleh = True
                      
9                     If simpan = False Then
10                        If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
11                            boleh = False
12                        End If
13                    End If
14                    If boleh Then
15                        sgnStall.Caption = ""
16                        If Layar.Tag <> "estBook" Then
17                            lblEdit.Visible = False
18                            If noPolBookingOnline = "" Then
19                            Call hapusForm
20                            End If
21                            If frmBooking.Tag = "FU" Then
22                                chkRtn = Checked
23                            End If
                              Dim nDpn1 As String
                              Dim nTgh1 As String
                              Dim nBlk1 As String
24                            nDpn1 = ""
25                            nTgh1 = ""
26                            nBlk1 = ""
27                            Call pecahNoPol(nopol, nDpn1, nTgh1, nBlk1)
28                            simpan = True
29                            Text(0) = Trim(nDpn1)
30                            Text(56) = Trim(nTgh1)
31                            Text(57) = Trim(nBlk1)
                              'lakukan pengecekan dari diagnose yang oke
                              Dim rsDiagapproval As New ADODB.Recordset
                              Dim pilihNoPrediagnose As String
                              Dim kumpulanPrediagnose As String
                              Dim jumlahPrediagnose As Integer
                              
32                            Set rsDiagapproval = objO.FetchData("select da.*,d.no_polisi,d.* from diagdta d WITH (NOLOCK) inner join diagapproval da WITH (NOLOCK) on da.no_diag = d.no_diag where da.no_wo = '' and isnull(da.alasan_tdksetuju,'') = '' and  d.no_polisi = '" & nopol & "' and d.VALIDTWC = 'OKE' and isnull(da.NOFAK,'') = ''")
                              
33                            If Not rsDiagapproval.EOF Then
                                  'jika ada data
34                                jumlahPrediagnose = 1
35                                rsDiagapproval.MoveFirst
36                                Do While Not rsDiagapproval.EOF
37                                    kumpulanPrediagnose = kumpulanPrediagnose & jumlahPrediagnose & ". " & rsDiagapproval.Fields(0) & vbCrLf
38                                    jumlahPrediagnose = jumlahPrediagnose + 1
39                                    rsDiagapproval.MoveNext
40                                Loop
                                  
41                                If MsgBox("Sudah ada Prediagnosis dengan No Prediagnosis :" & vbCrLf & kumpulanPrediagnose & " Untuk No Polisi tersebut. " & vbCrLf & " Apakah ingin membuat booking atas prediagnosis tersebut?", vbYesNo, "Konfirmasi") = vbYes Then
ulang:
42                                    pilihNoPrediagnose = UCase(InputBox("Masukan Prediagnose yang ingin dibuat Booking " & vbCrLf & kumpulanPrediagnose & " : ", "Pilih Prediagnose "))
43                                    If pilihNoPrediagnose = "" Then
44                                        If MsgBox("Apakah ingin membatalkan Booking dari Prediagnose?", vbYesNo, "Konfirmasi") = vbYes Then
45                                            Exit Sub
46                                        Else
47                                            GoTo ulang
48                                        End If
49                                    Else
50                                        rsDiagapproval.filter = "NO_DIAG = '" & pilihNoPrediagnose & "'"
51                                        If rsDiagapproval.EOF Then
52                                            MsgBox "Tidak ada No Prediagnose untuk No Polisi tersebut!", vbOKOnly, ""
53                                            GoTo ulang
54                                        Else
55                                            Call PanggilDataCustDiagnose(nopol, Trim(pilihNoPrediagnose))
56                                            Exit Sub
57                                        End If
58                                    End If
59                                Else
                                      'jika tidak membuat booking dari prediagnose
                                      
                                      'LAKUKAN PEMANGGILAN DATA BOOKING OPEN
                                      'update tanggal 26-08-2014
60                                    Set rsSearch = objO.FetchData("Select DELIV,nofak from BOOK WITH (NOLOCK) where NO_POLISI ='" & Trim(nopol) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
                                      Dim pilihNoBooking As String
                                      Dim kumpulanBooking As String
                                      Dim jumlahBooking As Integer
61                                    If Not rsSearch.EOF Then
62                                        rsSearch.MoveFirst
63                                        Do While Not rsSearch.EOF
64                                            jumlahBooking = jumlahBooking + 1
65                                            kumpulanBooking = kumpulanBooking & jumlahBooking & ". " & rsSearch.Fields(1) & vbCrLf
66                                            rsSearch.MoveNext
67                                        Loop
68                                        If MsgBox("Sudah ada Booking dengan No Booking :" & vbCrLf & kumpulanBooking & " Untuk No Polisi tersebut. " & vbCrLf & " Apakah ingin mengubah booking tersebut?", vbYesNo, "Konfirmasi") = vbYes Then
balikLagi:
69                                            pilihNoBooking = UCase(InputBox("Masukan No Booking yang ingin diubah " & vbCrLf & kumpulanBooking & " : ", "Pilih Booking "))
70                                            If pilihNoBooking = "" Then
71                                                If MsgBox("Apakah ingin membatalkan mengubah data Booking?", vbYesNo, "Konfirmasi") = vbYes Then
72                                                    Exit Sub
73                                                Else
74                                                    GoTo balikLagi
75                                                End If
76                                            Else
77                                                rsSearch.filter = "nofak = '" & pilihNoBooking & "'"
78                                                If rsSearch.EOF Then
79                                                    MsgBox "Tidak ada No Booking untuk No Polisi tersebut!", vbOKOnly, ""
80                                                    GoTo balikLagi
81                                                Else
82                                                    Call PanggilDataBooking2(nopol, Trim(pilihNoBooking), rsSearch("DELIV"))
83                                                    If Log.NHS = 1 Then
84                                                        If NulSID(gNoRangka) <> "" Then
85                                                            Call panggilWarranty(gNoRangka, "")
86                                                        End If
87                                                    End If
88                                                    If Text(23) <> "" Or Text(23) <> ">>->>->>>>" Then
89                                                        cmdWaktu.Enabled = True
90                                                    End If
91                                                    Exit Sub
92                                                End If
93                                            End If
94                                        Else
                                          'jika tidak mengubah data booking
95                                            Call hapusForm
96                                            Text(0).SetFocus
97                                            Exit Sub
98                                        End If
99                                    End If
100                                   Call PanggilDataCust(nopol)
101                               End If
102                           Else
                                  'ambil data non diagnose
                                  'update tanggal 26-08-2014
103                               Set rsSearch = objO.FetchData("Select DELIV,nofak from BOOK WITH (NOLOCK) where NO_POLISI ='" & Trim(nopol) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
                                      Dim pilihNoBooking2 As String
                                      Dim kumpulanBooking2 As String
                                      Dim jumlahBooking2 As Integer
104                                   If Not rsSearch.EOF Then
105                                       rsSearch.MoveFirst
106                                       Do While Not rsSearch.EOF
107                                           jumlahBooking2 = jumlahBooking2 + 1
108                                           kumpulanBooking2 = kumpulanBooking2 & jumlahBooking2 & ". " & rsSearch.Fields(1) & vbCrLf
109                                           rsSearch.MoveNext
110                                       Loop
111                                       If MsgBox("Sudah ada Booking dengan No Booking :" & vbCrLf & kumpulanBooking2 & " Untuk No Polisi tersebut. " & vbCrLf & " Apakah ingin mengubah booking tersebut?", vbYesNo, "Konfirmasi") = vbYes Then
balikLagiUbahBook:
112                                           pilihNoBooking2 = UCase(InputBox("Masukan No Booking yang ingin diubah " & vbCrLf & kumpulanBooking2 & " : ", "Pilih Booking "))
113                                           If pilihNoBooking2 = "" Then
114                                               If MsgBox("Apakah ingin membatalkan mengubah data Booking?", vbYesNo, "Konfirmasi") = vbYes Then
115                                                   Exit Sub
116                                               Else
117                                                   GoTo balikLagiUbahBook
118                                               End If
119                                           Else
120                                               rsSearch.filter = "nofak = '" & pilihNoBooking2 & "'"
121                                               If rsSearch.EOF Then
122                                                   MsgBox "Tidak ada No Booking untuk No Polisi tersebut!", vbOKOnly, ""
123                                                   GoTo balikLagiUbahBook
124                                               Else
125                                                   Call PanggilDataBooking2(nopol, Trim(pilihNoBooking2), rsSearch("DELIV"))
126                                                   If Log.NHS = 1 Then
127                                                       If NulSID(gNoRangka) <> "" Then
128                                                           Call panggilWarranty(gNoRangka, "")
129                                                       End If
130                                                   End If
                                                      
131                                                   If Text(23) <> "" Or Text(23) <> ">>->>->>>>" Then
132                                                       cmdWaktu.Enabled = True
133                                                   End If
134                                                   Exit Sub
135                                               End If
136                                           End If
137                                       Else
                                          'jika tidak mengubah data booking
138                                           Call hapusForm
139                                           Text(0).SetFocus
140                                           Exit Sub
141                                       End If
142                                   End If
143                                   Call PanggilDataCust(nopol)
144                               End If
                                  
      '                            Set rsSearch = objO.FetchData("Select DELIV from BOOK where NO_POLISI ='" & Trim(noPol) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
      '                            If rsSearch.BOF Then
      '                                baris = 1
      '                                If Not PanggilDataCust(noPol) Then
      '                                '    Call PanggilDataCustBook(noPol)
      '                                End If
      '                            Else
      '                                Call PanggilDataBooking(Trim(noPol), rsSearch("DELIV"))
      '                            End If
                              'End If
      '                        Set rsSearch = objO.FetchData("Select DELIV from BOOK where NO_POLISI ='" & Trim(noPol) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
      '                        If rsSearch.BOF Then
      '                            baris = 1
      '                            If Not PanggilDataCust(noPol) Then
      '                            '    Call PanggilDataCustBook(noPol)
      '                            End If
      '                        Else
      '                            Call PanggilDataBooking(Trim(noPol), rsSearch("DELIV"))
      '                        End If
145                           If cekPiutang(Trim(Text(2)), Trim(Text(10)), gNoRangka) Then
146                           End If
147                           birthDay = ""
148                           birthDay = birthDayReminder(Trim(Text(41)))
149                           If birthDay <> "" Then
150                               If Abs(DateDiff("d", Format(birthDay, "mm/dd/yyyy"), isnow)) <= 5 Then
151                                   MsgBox "Selamat Ulang tahun Bp./Ibu./Sdr." & cariNamaBirth(Trim(Text(15))) & vbCr & "Pada Tanggal " & birthDay
152                               End If
153                           End If
154                           Text(43) = cekMOU(Trim(Text(10)), gNoRangka)
155                           Text(26) = gSA
156                           simpan = False
157                       End If
                          
158                       If Log.freepart = 1 And NulSID(NulSID(Text(41))) <> "" Then
159                            tagihanFOP = ""
160                            kodeMDL = objO.FetchData_String("select Mdl from SingleIDOri..UnitCustomer where norangka = '" & NulSID(NulSID(Text(41))) & "'")
161                            If CekUnitTerlibatFOP(NulSID(Text(41)), kodeMDL, tagihanFOP) Then
162                                lblFOP.Caption = "Free Part + Oli : " & tagihanFOP
163                               sSQL = "select distinct FS_ITEM from [" & dbReplikasiHO & "]..HO_FOP_PROGRAM where mdl = '" & NulSID(kodeMDL) & "' and NIK = '" & Mid(NulSID(NulSID(gNoRangka)), 10, 1) & "'"
164                               Set rsFOP = objO.FetchData(sSQL)
165                               daftarSBE = ""
166                               If Not rsFOP.EOF Then
167                                  rsFOP.MoveFirst
168                                  Do While Not rsFOP.EOF
169                                      daftarSBE = daftarSBE & " - " & NulSID(rsFOP!fs_item) & vbCrLf
170                                      rsFOP.MoveNext
171                                  Loop
172                               End If
173                               MsgBox "Unit customer berikut terlibat FREE OIL PART " & tagihanFOP & " untuk SB kelipatan " & vbCrLf & daftarSBE, vbInformation, "Unit terlibat Free Service + Part"
                                  
174                            Else
175                                lblFOP.Caption = ""
176                            End If
                          
177                       End If
178                       If Log.NHS = 1 Then
179                           If NulSID(gNoRangka) <> "" Then
180                               Call panggilWarranty(gNoRangka, "")
181                           End If
182                       End If
183                   End If
184                   If cmdNo.Caption = "No &Polisi" Then
185                       cmdBersih.Visible = False
186                       If Trim(gNoRangka) <> "" Then
187                           Call panggilHistoryBok(gNoRangka, Trim(nopol))
188                           If Log.NHS = 0 Then
189                               Call panggilWarranty(gNoRangka, "")
                                  'Call panggilWarrantyBok(gNoRangka, Trim(text(10)))
190                           End If
191                       End If
192                   ElseIf cmdNo.Caption = "No &Rekom" Then
193                       cmdBersih.Visible = True
194                   End If
195                   If NulSID(Text(10)) <> "" Then
196                        If Log.RPLHistori = 0 Then
197                           Set rsSearch = objO.FetchData("Select * from MODEL WITH(NOLOCK) where MDL ='" & Text(10) & "'")
198                       Else
199                           Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH(NOLOCK) where MDL ='" & Text(10) & "'")
200                       End If
201                       If rsSearch.BOF Then
202                           MsgBox "Kelas kendaraan tidak sesuai master Model. Perbaiki data Model kendaraan di master Customer", vbInformation
lanjutKeMasterCustomer:
203                           With frmSIDCustomer
204                               If NulSID(gNoRangka) <> "" Then
205                                   If Left(gNoRangka, 3) <> "" Then
206                                   .Text(50) = Left(gNoRangka, 3)
207                                   End If
208                                   .Text(1) = Mid(gNoRangka, 4, 6)
209                                   .Text(0) = Mid(gNoRangka, 10, IIf(Len(gNoRangka) > 0, Len(gNoRangka) - 9, 0))
210                               End If
211                               frmSIDCustomer.show vbModal
212                           End With
213                       Else
214                           If NulSID(rsSearch!model) <> NulSID(Text(11)) Then
215                               MsgBox "Kelas kendaraan tidak sesuai master Model. Perbaiki data Model kendaraan di master Customer", vbInformation
216                               GoTo lanjutKeMasterCustomer
217                           End If
218                       End If
219                   End If
220               Else
221                   Call hapusForm
222               End If
223           Case 2
224               If Text(Index) <> "" Then
225                       Call isCurrentAR
                          'Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WHERE kodeLgn = '" & Text(Index) & "'")
                          '20/12/2012 Cust Non Axtif
      '                    Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & Text(Index) & "' and nonaktifflag = 0")
      '                    If rsSearch.EOF Then
      '                        gMainKeyCari = Text(Index)
      '                        gSearchOption = "KODEPLG"
      '                        frmCariAcess.Show (vbModal)
      '                    Else
      '                        Text(44) = Nul(rsSearch("KODEWIL"))
      '                        Text(3) = Nul(rsSearch("NAMALGN"))
      '                        Text(4) = Nul(rsSearch("ALAMAT1"))
      '                        Text(5) = Nul(rsSearch("ALAMAT2"))
      '
      '                        nDpn1 = ""
      '                        nTgh1 = ""
      '                        nBlk1 = ""
      '                        Call pecahNoTelp(Trim(Nul(rsSearch("TELEPON"))), nDpn1, nTgh1, nBlk1, ",")
      '                        Text(8) = Trim(nDpn1)
      '                        Text(6) = Trim(nTgh1)
      '                        Text(7) = Trim(nBlk1)
      '                    End If
226                   sSQL = "select b.*,a.kodelgn"
227                   sSQL = sSQL & " from [SingleIDOri]..Bridging a"
228                   sSQL = sSQL & " inner join [SingleIDOri]..NewCustomer b on a.IDCust = b.IDCust"
229                   sSQL = sSQL & " where a.KodeDealer = '" & Right(gDb, 2) & "' and a.idcust = '" & Nul(Text(Index)) & "'"
230                   Set rsSearch = objO.FetchData(sSQL)
231                   If Not rsSearch.EOF Then
232                       Text(44) = Left(Nul(rsSearch("idcust")), 3)
                          'text(3) = Nul(rsSearch("NamaCust1")) & " " & Nul(rsSearch("NamaCust2")) & " " & Nul(rsSearch("NamaCust3"))
                          'RUD.rev20180301 || 1.  Untuk title customer selain tipe Perorangan, minta dimunculkan dalam Cetakkan Invoice & Faktur Pajak
                          
233                       If UCase(Nul(rsSearch!TypeCustomer)) = "PERORANGAN" Then
234                           Text(3) = Left(Nul(Nul(rsSearch("NamaCust1")) & " " & Nul(rsSearch("NamaCust2")) & " " & Nul(rsSearch("NamaCust3"))), 50)
235                       Else
236                           Text(3) = Left(Nul(Nul(rsSearch!Title) & " " & Nul(rsSearch("NamaCust1")) & " " & Nul(rsSearch("NamaCust2")) & " " & Nul(rsSearch("NamaCust3"))), 50)
                              
237                       End If
238                       Text(4) = Nul(rsSearch("Alamat"))
239                       Text(5) = Nul(rsSearch("Kota"))

240                       Text(8) = Nul(rsSearch!IdTelp) & Nul(rsSearch!telp)
241                       Text(6) = Nul(rsSearch!hp1)
242                       Text(7) = Nul(rsSearch!hp2)
243                   Else
244                       Call panggilCari
245                   End If
                      
246               End If
247           Case 3
248               If Text(Index) <> "" Then
      'balik:
      '                    Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WITH (NOLOCK) WHERE NAMALgn = '" & Text(Index) & "'")
      '                    If rsSearch.EOF Then
      '                        Text(2) = ""
      '                        Text(44) = ""
      '                    Else
      '                        Dim rsKode As ADODB.Recordset
      'lanjut:
      '                        Dim telp As String
      '                        Set rsKode = objO.FetchDataACC("SELECT KODEWIL,ALAMAT1,ALAMAT2,TELEPON FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & Trim(Text(2)) & "'")
      '                        If Not rsKode.EOF Then
      '                            Text(44) = rsKode("KODEWIL")
      '                            Text(4) = rsKode("ALAMAT1")
      '                            Text(5) = rsKode("ALAMAT2")
      '                            telp = rsKode("TELEPON")
      '                        Else
      '                            Text(44) = rsSearch("KODEWIL")
      '                            Text(4) = rsSearch("ALAMAT1")
      '                            Text(5) = rsSearch("ALAMAT2")
      '                            telp = rsSearch("TELEPON")
      '                        End If
      '                        If rsKode.State <> adStateClosed Then rsKode.Close
      '
      '                        nDpn1 = ""
      '                        nTgh1 = ""
      '                        nBlk1 = ""
      '                        Call pecahNoTelp(Trim(telp), nDpn1, nTgh1, nBlk1, ",")
      '                        Text(8) = Trim(nDpn1)
      '                        Text(6) = Trim(nTgh1)
      '                        Text(7) = Trim(nBlk1)
      '                    End If

                      ''2018-01-05
                      '' PERMINTAAN CABANG KARENA BEBERAPA KALI MUNCUL PENCARIAN
                      ''------------------------------------------------------------------
      '                If Nul(text(2)) <> "" Then
      '                    sSQL = "select b.*,a.kodelgn"
      '                    sSQL = sSQL & " from [SingleIDOri]..Bridging a"
      '                    sSQL = sSQL & " inner join [SingleIDOri]..NewCustomer b on a.IDCust = b.IDCust"
      '                    sSQL = sSQL & " where a.KodeDealer = '" & Right(gDb, 2) & "' and a.idcust = '" & Nul(text(2)) & "'"
      '                    Set rsSearch = objO.FetchData(sSQL)
      '                Else
      '                    sSQL = "select b.*,a.kodelgn"
      '                    sSQL = sSQL & " from [SingleIDOri]..Bridging a"
      '                    sSQL = sSQL & " inner join [SingleIDOri]..NewCustomer b on a.IDCust = b.IDCust"
      '                    sSQL = sSQL & " where a.KodeDealer = '" & Right(gDb, 2) & "' and rtrim(ltrim(NamaCust1 + ' ' + NamaCust2 + ' ' + NamaCust3)) = '" & Nul(text(Index)) & "'"
      '                    Set rsSearch = objO.FetchData(sSQL)
      '                End If
      '                If Not rsSearch.EOF Then
      '                    text(44) = Left(Nul(rsSearch("idcust")), 3)
      '                    text(3) = Nul(rsSearch("NamaCust1")) & " " & Nul(rsSearch("NamaCust2")) & " " & Nul(rsSearch("NamaCust3"))
      '                    text(4) = Nul(rsSearch("Alamat"))
      '                    text(5) = Nul(rsSearch("Kota"))
      '
      '                    text(8) = Nul(rsSearch!IdTelp) & Nul(rsSearch!telp)
      '                    text(6) = Nul(rsSearch!HP1)
      '                    text(7) = Nul(rsSearch!HP2)
      '                Else
      '                    Call panggilCari
      '                End If
249               End If
250           Case 9
251               If Text(Index) = "" Then
252                   Text(Index) = 0
253                   Text(Index).SetFocus
254               ElseIf CDbl(Text(Index)) < (year(isnow) - 50) Or CDbl(Text(Index)) > year(isnow) Then
255                   Text(Index) = 0
256                   Text(Index).SelLength = 4
257                   Text(Index).SetFocus
258                   Exit Sub
259               End If
260           Case 10
261               If Text(Index) <> "" Then
                      
262                   If Log.RPLHistori = 0 Then
263                       Set rsSearch = objO.FetchData("Select * from MODEL WITH (NOLOCK) where MDL ='" & Text(Index) & "'")
264                   Else
265                       Set rsSearch = objO.FetchData("Select * from [" & dbReplikasiHO & "]..MODEL WITH (NOLOCK) where MDL ='" & Text(Index) & "'")
266                   End If
267                   If rsSearch.BOF Then
268                       Cancel = True
269                       gMainKeyCari = Text(Index)
270                       gSearchOption = "MDL"
271                       frmCari.show (vbModal)
272                       gGroupModel = ""
273                   Else
274                       Text(10) = rsSearch("MDL")
275                       Text(12) = rsSearch("Desc")
276                       Text(11) = rsSearch("MODEL")
277                       Call CexChassis_Discon(Text(11))
278                       If isOnCarClass = False Then
279                           MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
280                       End If
281                       gGroupModel = rsSearch("FS_GRP_MDL")
282                   End If
283               Else
284                   Text(12) = ""
285                   Text(11) = ""
286                   gGroupModel = ""
287               End If
288           Case 23
289               If CekTgl(Text(Index)) Then
                       '' 20191017 request mb thania validasi diganti dengan yang lama
                       '' -------------------------------------------------------------
      '                If NulSID(Text(1)) <> "" Then
      '                    '' jika tanggal book < tanggal book awal dan tanggal book = tanggal sekarang
      '                    If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) = 0 And DateDiff("d", Format(Text(Index), "dd/mm/yyyy"), Format(tglAwalBook, "DD/MM/YYYY")) > 0 Then
      '                        Text(Index) = TglDDMMYYYY(Format(tglAwalBook, "dd/mm/yyyy")) ''text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
      '                        MsgBox "Tanggal Booking tidak boleh sama dengan hari ini"
      '
      '                    ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) <= 0 Then
      '                        Text(Index) = TglDDMMYYYY(Format(tglAwalBook, "dd/mm/yyyy")) ''text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
      '                        MsgBox "Tanggal Booking tidak boleh lebih kecil atau sama dengan hari ini"
      '                        'Exit Sub
      '                    ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) > 30 Then
      '                        Text(Index) = TglDDMMYYYY(Format(tglAwalBook, "dd/mm/yyyy")) ''text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
      '                        MsgBox "Tanggal Booking tidak boleh lebih besar dari 1 bulan"
      '                        'Exit Sub
      '                    End If
      '                Else
      '                    If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) <= 0 Then
      '                        Text(Index) = TglDDMMYYYY(Format(DateAdd("D", 1, isnow), "dd/mm/yyyy")) ''text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
      '                        MsgBox "Tanggal Booking tidak boleh lebih kecil atau sama dengan hari ini"
      '                        'Exit Sub
      '                    ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) > 30 Then
      '                        Text(Index) = TglDDMMYYYY(Format(DateAdd("D", 1, isnow), "dd/mm/yyyy")) ''text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
      '                        MsgBox "Tanggal Booking tidak boleh lebih besar dari 1 bulan"
      '                        'Exit Sub
      '                    End If
      '                End If
290                   If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) < 0 Then
291                       Text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
292                       MsgBox "Tanggal Booking tidak boleh lebih kecil dari hari ini"
                          'Exit Sub
293                   ElseIf Log.tcm = 1 Then
294                         If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) > Log.tcmMaxHari Then ' ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) > 30 Then
295                             Text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
296                             MsgBox "Tanggal Booking tidak boleh lebih besar dari " & Log.tcmMaxHari & " hari" 'MsgBox "Tanggal Booking tidak boleh lebih besar dari 1 bulan"
                         
297                         End If
298                   ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) > gMaxHariBooking Then ' ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(Index), "DD/MM/YYYY")) > 30 Then
299                       Text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
300                       MsgBox "Tanggal Booking tidak boleh lebih besar dari " & gMaxHariBooking & " hari" 'MsgBox "Tanggal Booking tidak boleh lebih besar dari 1 bulan"
                          'Exit Sub
301                   End If
302               ElseIf Not CekTgl(Text(Index)) Then
303                   If Text(Index) <> ">>->>->>>>" Then
304                       MsgBox "Format tanggal salah"
305                   End If
                      ''Text(Index) = TglDDMMYYYY(Format(DateAdd("D", 1, isnow), "dd/mm/yyyy"))
306                   Text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
                      'Exit Sub
307               End If
308               If Text(58) <> ">>:>>" And Text(23) <> ">>->>->>>>" Then
309                   nWKT = hitungWaktu(CDbl(nJRate), Text(58), Text(23))
310                   Text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
311                   Text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
312               End If
313           Case 24
314               If Not CekJam(Text(Index)) Then
315                   Text(Index) = Format(isnow, "HH:NN")
316               Else
317                   If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(Text(23), "DD/MM/YYYY")) = 0 Then
318                       If DateDiff("s", Format(isnow, "HH:NN"), Format(Text(Index), "HH:NN")) < 0 Then
319                           Text(Index) = Format(isnow, "HH:NN")
320                       End If
321                   End If
322               End If
323               Text(58) = Trim(Text(24))
324               If Text(58) <> ">>:>>" And Text(23) <> ">>->>->>>>" And Not CekJam(Text(58)) Then
325                   nWKT = hitungWaktu(CDbl(nJRate), Text(58), Text(23))
326                   Text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
327                   Text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
328               End If
                  
329               If Trim(Text(1)) <> "" Then
330                   noPolis = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57))
331                   Set rsTemp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & noPolis & " |" & Trim(Text(1)) & "%'")
332                   If Not rsTemp.EOF Then
333                       If DateDiff("d", Format(rsTemp("TANGGAL"), "MM/DD/YYYY"), Format(Trim(Text(23)), "MM/DD/YYYY")) = 0 Then
334                           If Trim(rsTemp("JAMMULAI")) < Trim(Text(24)) Then
335                               Call objO.DeleteData("DELETE TEMPSTALL WHERE KETERANGAN LIKE '" & noPolis & " |" & Trim(Text(1)) & "|%'")
336                               MsgBox "Atur Ulang Posisi Stall"
337                               cmdStall.Tag = "belum"
338                           End If
339                       Else
340                           MsgBox "Atur Ulang Posisi Stall"
341                           cmdStall.Tag = "belum"
342                       End If
343                   End If
344                   If rsTemp.State <> adStateClosed Then rsTemp.Close
345               End If
346           Case 25
347               If Text(Index) = "" Then
348                   Text(Index) = "T"
349               ElseIf (Trim(Text(Index)) <> "T" And Trim(Text(Index)) <> "D" And Trim(Text(Index)) <> "S") Then
350                   Text(Index) = "T"
351               End If
352           Case 32
353               If Text(Index) <> "" Then
354                   If Text(10) = "" Then
355                       MsgBox "Isi model kendaraan !!!"
356                       Sendkeys ("+{TAB}")
357                       Text(10).SelLength = Len(Text(10))
358                       Text(10).SetFocus
359                       Exit Sub
360                   Else
361                       Set rsSearch = objO.FetchData("Select DISTINCT JOB, REMARK from HO_FLATE WITH (NOLOCK) where job = '" & Text(Index) & "' AND FS_LOG_DEL = ''")
362                       If rsSearch.BOF Then
363                           Cancel = True
364                           gMainKeyCari = Text(Index)
365                           gSearchOption = "JENIS_KERJA"
366                           frmCari.show (vbModal)
367                       Else
368                           Text(33) = rsSearch("REMARK")
369                       End If
370                   End If
371               End If
372           Case 34
373               If Text(Index) <> "" Then
374                   If Text(32) = "" Then
375                       MsgBox "Isi Tipe Job Order !!!"
376                       Sendkeys ("+{TAB}")
377                       Text(32).SelLength = 3
378                       Text(32).SetFocus
379                       Exit Sub
380                   Else
381                       If Text(23) <> ">>->>->>>>" Then
382                           Set rsSearch = objO.FetchData("select DISTINCT FS_ITEM, RATE from HO_FLATE WITH (NOLOCK) where job = '" & Text(32) & "' and fs_item = '" & Text(Index) & "' AND mdl = '" & Text(10) & "' and FS_LOG_DEL = ''")
383                           If rsSearch.BOF Then
384                               Cancel = True
385                               gMainKeyCari = Text(Index)
386                               gSearchOption = "JOBLIST2"
387                               gBantuan = Text(10)
388                               gBantuan2 = Text(32)
389                               frmCari.show (vbModal)
390                               If fOther Then
391                                   Text(42).SelLength = Len(Text(42))
392                               End If
393                           Else
394                               If Trim(rsSearch("FS_ITEM")) = "OTHER" Then
395                                   fOther = True
396                                   Text(34) = ""
397                                   Text(34).Mask = ">CCCCCCCCCCCCCCCCCCCCCCCCC"
398                                   Text(42).Visible = True
399                                   Text(42).Enabled = True
400                                   Text(34).SelLength = Len(Text(34))
401                                   Text(34).SetFocus
402                               Else
403                                   fOther = False
404                                   Text(42).Visible = False
405                                   Jrate = CDbl(rsSearch("Rate"))
406                               End If
407                           End If
408                       Else
409                           MsgBox "Isi Tanggal & Jam Booking", vbCritical, "KELENGKAPAN"
410                           Sendkeys ("+{TAB}")
411                           Text(23).SelLength = Len(Text(23))
412                           Text(23).SetFocus
413                           Exit Sub
414                       End If
415                   End If
416               End If
417           Case 35
                  'PART
418               If Text(Index) <> "" Then
419                       Set rsSearch = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
420                       If rsSearch.BOF Then
421                           If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then
422                               stkPart = 0
423                               namaPart = ""
424                               HrgPart = 0
425                               satPart = "ITEM"
426                               oliPart = False
427                               namabySA = Text(Index)
428                               If rsTemp.State <> adStateClosed Then rsTemp.Close
429                           Else
430                               Cancel = True
431                               gMainKeyCari = Text(Index)
432                               Text(Index) = ""
433                               gSearchOption = "PART"
434                               frmCariAcess.show (vbModal)
435                           End If
436                       Else
437                           stkPart = CDbl(rsSearch("QTYSTOCK"))
438                           namaPart = rsSearch("NAMABARANG")
439                           HrgPart = Val(rsSearch("HARGAJUALA"))
440                           satPart = rsSearch("KODESATUAN")
441                           Set rsTemp = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "' AND TIPE = 'O'")
442                           If Not rsTemp.EOF Then
443                               oliPart = True
444                           Else
445                               oliPart = False
446                           End If
447                           namabySA = ""
448                           If rsTemp.State <> adStateClosed Then rsTemp.Close
449                       End If
                      
450               End If
              
451           Case 36
452               If Text(35) <> "" Then
453                   If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(18)), "DD/MM/YYYY") Then
454                       If Trim(Text(Index)) = "" Then
455                               Text(Index) = 0
456                           ElseIf CDbl(Text(Index)) = 0 Then
457                               Sendkeys ("+{TAB}")
458                               Exit Sub
459                           Else
              ''''                    If CDbl(text(Index)) > CDbl(stkPart) Then
              ''''                        temp = MsgBox("Melebihi Stok yang ada !!!" & vbCr & "Tetap akan memproses ?", vbYesNo, "PESAN MELEBIHI STOCK")
              ''''                        If temp = vbYes Then
              ''''                            ' NOTE KE BAGIAN SPAREPART UNTUK PESAN
              ''''                        Else
              ''''                            text(Index) = stkPart
              ''''                        End If
              ''''                    End If
460                           End If
                          
461                       Else
462                           For i = 1 To barispart
463                               If UCase(partGrid.TextMatrix(i, 1)) = UCase(Trim(Text(35))) Or UCase(partGrid.TextMatrix(i, 7)) = UCase(Trim(Text(35))) Then
464                                   Text(Index) = 0
465                                   Exit Sub
466                               End If
467                           Next
468                       End If
          '
                      
469               End If
470           Case 42
471               If Text(Index) = "" Then
472                   Text(Index) = 0
473               Else
474                   If Text(32) = "" Then
475                       MsgBox "Isi Tipe Job Order !!!"
476                       Text(32).SelLength = 3
477                       Text(32).SetFocus
478                       Exit Sub
479                   End If
480               End If
481           Case 46, 47, 48, 49, 50, 51
482               If Text(Index) <> "" Then
483                   gBantuan = Index - 45
484                   Select Case gBantuan
                          Case 1
485                           Set rsSearch = objO.FetchData("SELECT waktu FROM T_KONDISI WITH (NOLOCK) WHERE waktu = '" & Text(Index) & "' ORDER BY id")
486                       Case 2
487                           Set rsSearch = objO.FetchData("SELECT frekuensi FROM T_KONDISI WITH (NOLOCK) WHERE frekuensi = '" & Text(Index) & "' ORDER BY id")
488                       Case 3
489                           Set rsSearch = objO.FetchData("SELECT tempat FROM T_KONDISI WITH (NOLOCK) WHERE tempat = '" & Text(Index) & "' ORDER BY id")
490                       Case 4
491                           Set rsSearch = objO.FetchData("SELECT warn FROM T_KONDISI WITH (NOLOCK) WHERE warn = '" & Text(Index) & "' ORDER BY id")
492                       Case 5
493                           Set rsSearch = objO.FetchData("SELECT mulai FROM T_KONDISI WITH (NOLOCK) WHERE mulai = '" & Text(Index) & "' ORDER BY id")
494                       Case 6
495                           Set rsSearch = objO.FetchData("SELECT selesai FROM T_KONDISI WITH (NOLOCK) WHERE selesai = '" & Text(Index) & "' ORDER BY id")
496                   End Select
497                   If rsSearch.EOF Then
498                       Text(Index) = ""
499                   Else
                          Dim nT As Integer
500                       Select Case gBantuan
                              Case 1
501                               Text(54) = ""
502                               nT = 54
503                           Case 2
504                               Text(65) = ""
505                               nT = 65
506                           Case 3
507                               Text(66) = ""
508                               nT = 66
509                           Case 4
510                               Text(67) = ""
511                               nT = 67
512                           Case 5
513                               Text(68) = ""
514                               nT = 68
515                           Case 6
516                               Text(69) = ""
517                               nT = 69
518                       End Select
519                       If Text(Index) = "LAINNYA" Then
520                           Select Case gBantuan
                                  Case 1
521                                   Text(54) = bKondisi(0)
522                               Case 2
523                                   Text(65) = bKondisi(1)
524                               Case 3
525                                   Text(66) = bKondisi(2)
526                               Case 4
527                                   Text(67) = bKondisi(3)
528                               Case 5
529                                   Text(68) = bKondisi(4)
530                               Case 6
531                                   Text(69) = bKondisi(5)
532                           End Select
533                           Text(nT).Visible = True
534                           Text(nT).SetFocus
535                       Else
536                           Text(nT) = ""
537                           Me.Text(nT).Visible = IIf(Me.Text(nT).Visible, False, False)
538                       End If
539                   End If
540               End If
541           Case 54
542               If Text(Index) <> "" Then
543                   bKondisi(0) = Text(Index)
544               End If
545           Case 65
546               If Text(Index) <> "" Then
547                   bKondisi(1) = Text(Index)
548               End If
549           Case 66
550               If Text(Index) <> "" Then
551                   bKondisi(2) = Text(Index)
552               End If
553           Case 67
554               If Text(Index) <> "" Then
555                   bKondisi(3) = Text(Index)
556               End If
557           Case 68
558               If Text(Index) <> "" Then
559                   bKondisi(4) = Text(Index)
560               End If
561           Case 69
562               If Text(Index) <> "" Then
563                   bKondisi(5) = Text(Index)
564               End If
565           Case 64
566               If Text(Index) <> "" Then
567                   Set rsSearch = objO.FetchData("SELECT * FROM MEKANIK WITH (NOLOCK) WHERE MECH = '" & Text(Index) & "'")
568                   If rsSearch.EOF Then
569                       gMainKeyCari = Text(Index)
570                       gSearchOption = "MEKANIK"
571                       gTanggal = isnow
572                       If DateDiff("d", Format(Text(23), "MM/DD/YYYY"), Format(gTanggal, "MM/DD/YYYY")) = 0 Then
573                           gBantuan = "ORDER"
574                           gBantuan2 = IIf(Weekday(gTanggal) = 1, 6, Weekday(gTanggal) - 2) & WeekdayName(IIf(Weekday(gTanggal) = 1, 8, Weekday(gTanggal)) - 1)
575                           gBantuan3 = Format(gTanggal, "DD/MM/YYYY")
576                           gBantuan4 = ""
577                       Else
578                           gBantuan = "BOOKING"
                              Dim nTgl As String
579                           nTgl = Format(Text(23), "MM/DD/YYYY")
580                           gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
581                           gBantuan3 = Format(nTgl, "DD/MM/YYYY")
582                           gBantuan4 = ""
583                       End If
584                       frmCari.show (vbModal)
585                   End If
586               End If
587           Case 79
588               If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
589                   If Trim(Text(Index)) <> "" And NulSID(Text(86)) = "" Then
590                       sSQL = "SELECT DISTINCT D.KDSUBLET,D.NAMASUBLET,D.KETERANGAN,D.HARGAJUAL,D.HARGABELI,E.KDSUPP,E.SUPPLIER,D.KODESATUAN"
591                       sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLET D "
592                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL B WITH (NOLOCK) ON D.KDSUBLET = B.KDSUBLET"
593                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER C WITH (NOLOCK) ON D.KDSUBLET = C.KDSUBLET"
594                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUPPLIER E WITH (NOLOCK) ON D.KDSUPP = E.KDSUPP"
595                       sSQL = sSQL & " WHERE ('" & Format(NulSID(Text(23)), "yyyy-mm-dd") & "' BETWEEN D.AWALPER anD D.AKHIRPER + '23:59:59' or D.AWALPER = D.AKHIRPER)"
596                       sSQL = sSQL & " AND D.FS_LOG_DEL = '' AND (D.ALLMODEL = 1 OR B.MDL = '" & NulSID(Text(10)) & "') AND (D.ALLCABANG = 1 OR C.FS_KD_DEALER = '" & gKdCbg & "')"
                          '' 20201022DEV
597                       sSQL = sSQL & " AND (SUBSTRING(D.KDSUBLET,1,4) = 'NR00' OR SUBSTRING(D.KDSUBLET,1,4) = '" & Replace(gKdCbg, " ", "") & "')"
598                       sSQL = sSQL & " AND D.KDSUBLET = '" & NulSID(Text(Index)) & "'"
599                       Set rsTemp = objO.FetchData(sSQL)
600                       If rsTemp.EOF Then
601                           MsgBox "Kode Pekerjaan sublet tidak ditemukan.", vbInformation
602                           Text(79) = ""
603                           Text(81) = ""
604                           Text(82) = ""
605                       Else
606                           Text(79).BackColor = &H80000005
607                           Text(79) = NulSID(rsTemp!KDSUBLET)
608                           Text(82) = Round(NulSID(rsTemp!hargaJual) * profilACC.nilaitotal / profilACC.nilaidpp, 0) ''NulSID(rsTemp!hargaJual)'' PRICE
609                           Text(81) = NulSID(rsTemp!KODESATUAN)
610                       End If
611                   End If
612               End If
613           Case 86
614               If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
615                   If Trim(Text(Index)) <> "" Then
616                       sSQL = "SELECT D.KDSUPP,E.SUPPLIER,A.KDPROMO,A.NAMAPROMO,A.KETERANGAN,A.HARGAJUAL,D.HARGABELI,A.KDSUBLET,D.NAMASUBLET,D.KETERANGAN AS KETSUBLET,D.KODESATUAN"
617                       sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLETPROMO A WITH (NOLOCK)"
618                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLET D WITH (NOLOCK) ON A.KDSUBLET = D.KDSUBLET"
619                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMOMDL B WITH (NOLOCK) ON A.KDPROMO = B.KDPROMO"
620                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMODEALER C WITH (NOLOCK) ON A.KDPROMO = C.KDPROMO"
621                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUPPLIER E WITH (NOLOCK) ON E.KDSUPP = D.KDSUPP "
622                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL F WITH (NOLOCK) ON D.KDSUBLET = F.KDSUBLET"
623                       sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER G WITH (NOLOCK) ON D.KDSUBLET = G.KDSUBLET"
624                       sSQL = sSQL & " LEFT JOIN T_SUPPLIER H WITH (NOLOCK) ON H.KDSUPP = E.KDSUPP"
625                       sSQL = sSQL & " WHERE ('" & Format(NulSID(Text(23)), "yyyy-mm-dd") & "' BETWEEN A.AWALPROMO anD A.AKHIRPROMO + '23:59:59' or A.AWALPROMO = A.AKHIRPROMO)"
626                       sSQL = sSQL & " AND ISNULL(H.KODEAP,'') <> '' AND A.FS_LOG_DEL = '' AND (A.ALLMODEL = 1 OR B.MDL = '" & NulSID(Text(1)) & "') AND (A.ALLCABANG = 1 OR C.FS_KD_DEALER = '" & gKdCbg & "')"
627                       sSQL = sSQL & " AND ('" & Format(NulSID(Text(23)), "yyyy-mm-dd") & "' BETWEEN D.AWALPER anD D.AKHIRPER + '23:59:59' or D.AWALPER = D.AKHIRPER) "
                           '' 20201022DEV
628                       sSQL = sSQL & " AND (SUBSTRING(A.KDSUBLET,1,4) = 'NR00' OR SUBSTRING(A.KDSUBLET,1,4) = '" & Replace(gKdCbg, " ", "") & "')"
629                       sSQL = sSQL & " AND D.FS_LOG_DEL = '' AND (D.ALLMODEL = 1 OR F.MDL = '" & NulSID(Text(10)) & "') AND (D.ALLCABANG = 1 OR G.FS_KD_DEALER = '" & gKdCbg & "') and A.KDPROMO = '" & NulSID(Text(Index)) & "'"
630                       Set rsTemp = objO.FetchData(sSQL)
631                       If Not rsTemp.EOF Then
632                           Text(79).BackColor = &HC0E0FF
633                           Text(79) = NulSID(rsTemp!KDSUBLET)
                              ''20200506 REVISI HARGA JUAL DAN BELI EXCLUDE PPN
                              ''------------------------------------------------
634                           Text(82) = Round(NulSID(rsTemp!hargaJual) * profilACC.nilaitotal / profilACC.nilaidpp, 0) ''NulSID(rsTemp!hargaJual)
635                           Text(81) = NulSID(rsTemp!KODESATUAN)
636                       Else
637                           gMainKeyCari = Text(Index)
638                           gSearchOption = "PROMOSUBLETSENTRAL"
639                           gBantuan = Format(NulSID(Text(23)), "yyyy-mm-dd")
640                           gBantuan2 = NulSID(Text(10))
641                           frmCari.show (vbModal)
642                           Text(Index).SetFocus
643                       End If
644                   Else
645                       If Text(21) = "" Then
646                           Text(24).BackColor = &H80000005
647                           Text(24) = ""
648                       End If
649                   End If
650               End If
651       End Select

652      On Error GoTo 0
653      Exit Sub

text_Validate_Error:

654       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmBooking")
           
End Sub


Public Sub panggilHistoryBok(noRgk As String, noPols As String)
1        On Error GoTo panggilHistoryBok_Error

2         On Error Resume Next
          Dim rsHist As ADODB.Recordset
          
              ' AMBIL HISTORY DARI HO
          Dim rsHlp  As ADODB.Recordset
          Dim rsHlp2 As ADODB.Recordset
          Dim dBaseHO As String
          
3         If Trim(nopol) = "" Then
4             nopol = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57))
5         End If
6         gNoRangka = noRgk

7         If rsHist.State Then rsHist.Close
8         If noRgk = "" Then
9             sSQL = "SELECT * FROM History WITH (NOLOCK) WHERE no_polisi = '" & Trim(noPols) & "' AND FS_LOG_DEL = '' ORDER BY [date] DESC, NO_WO DESC"
10        Else
11            sSQL = "SELECT * FROM History WITH (NOLOCK) WHERE no_chassis = '" & Trim(noRgk) & "' AND FS_LOG_DEL = '' ORDER BY [date] DESC, NO_WO DESC"
12        End If
13        Set rsHist = objO.FetchData(sSQL)
14        If Not rsHist.EOF Then
15            Text(20) = Format(Trim(rsHist("DATE")), "dd-mm-yyyy")
16            Text(21) = Nul(Trim(rsHist("FS_NM_DEALER")))
17            Text(22) = Format(CDbl(rsHist("KM")), "#,#")
18            gNoRangka = Trim(rsHist("NO_CHASSIS"))
19            If DateDiff("d", Format(rsHist("INVDATE"), "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) > 15 Then
20                chkRtn.Enabled = False
21            Else
22                chkRtn.Enabled = True
23            End If
24            cmdHistory.Enabled = True
25            Command1.Enabled = True
26        Else
27            chkRtn.Enabled = False
28            Text(22) = 0
              'cmdHistory.Enabled = False
29        End If
30        cmdHistory.Enabled = True
31        Command1.Enabled = True

32       On Error GoTo 0
33       Exit Sub

panggilHistoryBok_Error:

34        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilHistoryBok of Form frmBooking")
End Sub

Private Sub panggilHistoryHO(noRgk As String)
      Dim rsHlp As ADODB.Recordset

                  
                      
1        On Error GoTo panggilHistoryHO_Error

2         Set rsRec = objO.FetchDataHO("SELECT no_chassis FROM Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
3         If Not rsRec.EOF Then
4             terlibat = False
5         Else
6             terlibat = True
7         End If


8         If Log.LogHOConn = 1 Then
9             If ConHO = True Then
10            dBaseHO = HODb
               
11            Set rsHlp = objO.FetchDataHO("SELECT * FROM HISTORYnew WITH (NOLOCK) WHERE NO_CHASSIS = '" & noRgk & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC, NO_WO DESC")
12            If Not rsHlp.EOF Then
13                Do While Not rsHlp.EOF
14                    If rsHlp("NO_CHASSIS") <> "" Then
                           
15                        Set rsHlp2 = objO.FetchData("SELECT * FROM HISTORY WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "' AND KM= '" & rsHlp("KM") & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC")
16                        If rsHlp2.EOF Then
17                            Call objO.InsertData("INSERT INTO HISTORY VALUES ('" & rsHlp("CUSTOMER") & "', '" & rsHlp("CUST_ADRES") & "', '" & rsHlp("KOTA") & "','" & rsHlp("NO_POLISI") & "', '" & rsHlp("NO_ENGINE") & "','" & rsHlp("NO_CHASSIS") & "','" & rsHlp("FRONT") & "','" & rsHlp("MECH") & "','" & rsHlp("KM") & "','" & rsHlp("DATE") & "'," & rsHlp("PART_G") & ", " _
                                           & "" & rsHlp("PART_NG") & "," & rsHlp("OIL_G") & "," & rsHlp("OIL_NG") & "," & rsHlp("MTRL") & "," & rsHlp("LABOR") & "," & rsHlp("BODY") & "," & rsHlp("SBLT") & "," & rsHlp("MET") & "," & rsHlp("TOTAL") & ",'" & rsHlp("JOB") & "','" & rsHlp("REP1") & "','" & rsHlp("REP2") & "','" & rsHlp("REP3") & "','" & rsHlp("REP4") & "'," _
                                           & "'" & rsHlp("REP5") & "','" & rsHlp("REP6") & "','" & rsHlp("REP7") & "','" & rsHlp("MEKANIK1") & "','" & rsHlp("MEKANIK2") & "','" & rsHlp("MEKANIK3") & "'," _
                                           & "'" & rsHlp("MEKANIK4") & "','" & rsHlp("MEKANIK5") & "','" & rsHlp("MEKANIK6") & "','" & rsHlp("MEKANIK7") & "','" & rsHlp("MEKANIK8") & "','" & rsHlp("KLPKERJA1") & "','" & rsHlp("KLPKERJA2") & "','" & rsHlp("KLPKERJA3") & "','" & rsHlp("KLPKERJA4") & "','" & rsHlp("KLPKERJA5") & "','" & rsHlp("KLPKERJA6") & "'," _
                                           & "'" & rsHlp("KLPKERJA7") & "','" & rsHlp("KLPKERJA8") & "','" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "','" & rsHlp("NOFAK") & "','" & rsHlp("STATUS") & "','" & rsHlp("PAID") & "','" & rsHlp("DISC") & "','" & rsHlp("DISCT") & "','" & rsHlp("REMARK1") & "','" & rsHlp("FN_RETURN") & "','" & rsHlp("FS_RC") & "'," _
                                           & "'" & rsHlp("FS_KD_DEALER") & "','" & rsHlp("FS_NM_DEALER") & "','" & rsHlp("TIPEWO") & "','" & rsHlp("KODETIPE") & "', '" & rsHlp("SBE") & "','" & rsHlp("FS_KD_PLG") & "','" & rsHlp("FS_KD_WIL") & "','" & rsHlp("THS") & "','" & rsHlp("MOU") & "','" & rsHlp("PDS") & "','" & rsHlp("FS_LOG_DEL") & "','" & rsHlp("FD_TGL_DEL") & "', " _
                                           & "'" & rsHlp("ALASANDEL") & "','" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "')")
18                        End If
19                    End If
20                    rsHlp.MoveNext
21                Loop
22            End If
23            End If
               
24        End If

25       On Error GoTo 0
26       Exit Sub

panggilHistoryHO_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilHistoryHO of Form frmBooking")
End Sub

Private Sub saveStall()
          Dim rsExist As ADODB.Recordset
          Dim rsSTemp As ADODB.Recordset
          Dim nKet As String
          Dim nTD As Double
          Dim doup As Integer
          Dim filterStall As String
          
1        On Error GoTo saveStall_Error

2         If lblEdit.Visible = True Then
3             filterStall = nopol & " |" & Text(1) & "%'"
4         Else
5             filterStall = nopol & " |%'"
6         End If
          
          
7         Set rsSTemp = objHT.FetchDataBengkel("SELECT noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, keterangan, CHARINDEX('||',KETERANGAN) AS TNDA,adjust FROM TempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
8         If rsSTemp.RecordCount > 0 Then
9             nKet = rsSTemp.Fields(7).value
10            nKet = Left(nKet, InStr(nKet, "|")) & Text(1) & Right(nKet, Len(nKet) - InStr(nKet, "| ") + 1)
11            nTD = CDbl(rsSTemp("TNDA"))
12        End If
13        doup = 0
14        If Not rsSTemp.EOF Then
15            If nTD > 0 Then
16                Set rsSTemp = objHT.FetchDataBengkel("SELECT noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, '" & nKet & "' as keterangan ,FLAG,adjust FROM TempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
17                doup = 1
18            Else
19                Set rsSTemp = objHT.FetchDataBengkel("SELECT noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, keterangan,FLAG,adjust FROM TempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
20            End If
21        Else
22            Set rsSTemp = objHT.FetchDataBengkel("SELECT noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, keterangan,FLAG,adjust FROM TempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
23        End If
24        On Error GoTo errStall
25        If Not rsSTemp.EOF Then
26            If Log.helpJpcb = 1 Then
                  Dim rsHlpJPCB As ADODB.Recordset
                  Dim rsLogJPCB As ADODB.Recordset
                  
27                Set rsLogJPCB = objHT.FetchDataBengkel("SELECT * FROM PROFILBENGKEL.DBO.LOGJPCB")
28                Set rsHlpJPCB = objHT.FetchDataBengkel("SELECT TOP 1 * FROM HELPJPCB WHERE KETERANGAN = '" & nKet & "' ")
29                If rsHlpJPCB.EOF Then
30                     If rsLogJPCB.RecordCount > 0 Then
31                        rsLogJPCB.MoveFirst
32                        Do While Not rsLogJPCB.EOF
                             'lakukan pengisian data help jpcb berdasarkan mech
33                           Call objHT.InsertData_BengkelOnly("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('" & Trim(Text(1)) & "','T','S','" & UCase(Trim(rsSTemp("nostall"))) & "','" & Format(rsSTemp("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsSTemp("JAMMULAI")) & "','" & Trim(rsSTemp("JAMSELESAI")) & "'," & Replace(rsSTemp("rate"), ",", ".") & ",'" & Trim(rsSTemp("mech")) & "','','" & Trim(nKet) & "','" & Trim(rsSTemp("adjust")) & "','CREATE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
34                           rsLogJPCB.MoveNext
35                        Loop
36                     End If
37                Else
38                    rsHlpJPCB.MoveFirst
39                    If rsHlpJPCB("STATUS") = "F" Then
40                        If rsLogJPCB.RecordCount > 0 Then
41                            rsLogJPCB.MoveFirst
42                            Do While Not rsLogJPCB.EOF
43                               Call objHT.InsertData_BengkelOnly("INSERT INTO [helpJPCB] ([NoWO],[Status] ,[Tipe],[NoStall],[Tanggal],[JamMulai] ,[JamSelesai],[Rate],[Mech],[Flag],[Keterangan],[adjust],[TipeUpdate],[Komp],[TglEntry],[userid]) VALUES ('" & Trim(Text(1)) & "','T','S','" & UCase(Trim(rsSTemp("nostall"))) & "','" & Format(rsSTemp("tanggal"), "MM/DD/YYYY") & "','" & Trim(rsSTemp("JAMMULAI")) & "','" & Trim(rsSTemp("JAMSELESAI")) & "'," & Replace(rsSTemp("rate"), ",", ".") & ",'" & Trim(rsSTemp("mech")) & "','','" & Trim(nKet) & "','" & Trim(rsSTemp("adjust")) & "','CREATE','" & Trim(rsLogJPCB("KOMP")) & "',GETDATE(),'" & rsLogJPCB("userid") & "')")
44                               rsLogJPCB.MoveNext
45                            Loop
46                        End If
47                    End If
48                End If
49            End If
50            Call objHT.DeleteData_Bengkel("DELETE FROM stall WHERE tanggal = '" & Format(Text(23), "mm/dd/yyyy") & "' AND keterangan like '" & nopol & " |" & Text(1) & "%'")
51                Call objHT.InsertData_BengkelOnly("INSERT INTO stall (noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, keterangan, FLAG,adjust)  VALUES " _
                             & "('" & rsSTemp("NOSTALL") & "','" & Format(rsSTemp("TANGGAL"), "MM/DD/YYYY") & "', '" & rsSTemp("JAMMULAI") & "', " _
                             & "'" & rsSTemp("JAMSELESAI") & "', '" & Replace(rsSTemp("RATE"), ",", ".") & "', '" & rsSTemp("GROUP") & "', '" & rsSTemp("MECH") & "', '" & nKet & "', '','" & rsSTemp("adjust") & "')")
52                Call objHT.DeleteData_Bengkel("DELETE tempStall WHERE keterangan like '" & filterStall & "")
53            If doup = 1 Then Call objHT.UpdateData_Bengkel("UPDATE TEMPSTALL SET KETERANGAN = Left(KETERANGAN, CHARINDEX('|',KETERANGAN)) + '" & Text(1) & "' + Right(KETERANGAN, Len(KETERANGAN) - CHARINDEX('||',KETERANGAN))  WHERE tanggal = '" & Format(Text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
54        End If
55        Exit Sub
errStall:
56        err.Raise err.Number, err.Source, err.Description

57       On Error GoTo 0
58       Exit Sub

saveStall_Error:
59        erTranAll = True
60        ErrMsg.judul = ""
61        ErrMsg.message = "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure saveStall of Form frmBooking")
62        ErrMsg.show
          
      '    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure saveStall of Form frmBooking")
End Sub
                                
Private Sub rtjFU(nChs As String)
          ' UPDATE STATUS BOOK DI DTLFU
1        On Error GoTo rtjFU_Error

2         Call objHT.UpdateData_Bengkel("UPDATE DTLFU SET BOOK = '1' WHERE NO_CHASSIS = '" & Trim(nChs) & "' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "'")

3        On Error GoTo 0
4        Exit Sub

rtjFU_Error:
5         erTranAll = True
6         ErrMsg.judul = ""
7         ErrMsg.message = "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure rtjFU of Form frmBooking")
8         ErrMsg.show
          
      ''    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure rtjFU of Form frmBooking")
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii <= 57 And KeyAscii >= 48 Then
            KeyAscii = KeyAscii
        ElseIf KeyAscii = 46 Or KeyAscii = 8 Then
            KeyAscii = KeyAscii
        Else
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If txt(Index).BackColor = fColor Then
        txt(Index).Locked = True
        Sendkeys ("{TAB}")
    Else
        txt(Index).Locked = False
        If Index = 9 Then
            If txt(8).Text = "" Then txt(8).SetFocus
        End If
    End If
    If gMainKeyCari <> "" Then
        txt(Index) = UCase(NulSID(gMainKeyCari))
        gMainKeyCari = ""
        DiTab = True
        Sendkeys ("{TAB}")
        'txt(Index).SetFocus
        'SendKeys ("{TAB}")
    End If
    
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 63, 64, 65, 66, 11
                
                If gMainKeyCari = "" Then gMainKeyCari = txt(Index)
                gSearchOption = "TOPOLOGY"
                If frmCari.Visible Then
                    frmCari.Hide
                End If
                frmCari.show (vbModal)
             
        End Select
        onFind = False
    ElseIf KeyCode = 13 Then
        If Index = 11 Then
            If NulSID(txt(Index)) <> "" Then
                Dim rsTopology As ADODB.Recordset
                Set rsTopology = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TOPOLOGYDAERAH WITH (NOLOCK) WHERE KODEPOS = '" & NulSID(txt(11)) & "'")
                If Not rsTopology.EOF Then
                    rsTopology.MoveFirst
                    txt(63) = NulSID(rsTopology!Propinsi)
                    txt(64) = NulSID(rsTopology!Kota) 'nulsid(DGrid.Columns(1))
                    txt(65) = NulSID(rsTopology!kecamatan) 'nulsid(DGrid.Columns(2))
                    txt(66) = NulSID(rsTopology!kelurahan) 'nulsid(DGrid.Columns(3))
                    txt(11) = NulSID(rsTopology!KodePos)
                Else
                    MsgBox "Data kodepos tidak terdaftar.", vbInformation
                    txt(63) = ""
                    txt(64) = "" 'nulsid(DGrid.Columns(1))
                    txt(65) = "" 'nulsid(DGrid.Columns(2))
                    txt(66) = "" 'nulsid(DGrid.Columns(3))
                    txt(11) = "" 'nulsid(DGrid.Columns(4))
                End If
            End If
        Else
            Call txt_Validate(Index, False)
        End If
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
        
    Dim rsCari As ADODB.Recordset
    
    If Index = 63 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
                
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 64 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(63)) & "' and KOTA = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 65 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(63)) & "' and KOTA = '" & NulSID(txt(64)) & "' and KECAMATAN = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 66 Then
        If NulSID(txt(Index)) <> "" Then
            
            Set rsCari = objO.FetchData("SELECT * FROM [SINGLEIDORI]..TopologyDaerah WITH (NOLOCK) WHERE PROPINSI = '" & NulSID(txt(63)) & "' and KOTA = '" & NulSID(txt(64)) & "' and KECAMATAN = '" & NulSID(txt(65)) & "' and KELURAHAN = '" & NulSID(txt(Index)) & "'")
            If rsCari.EOF Then
                gMainKeyCari = txt(Index)
                txt(Index) = ""
                Call txt_KeyDown(Index, vbKeyF12, 0)
'                gMainKeyCari = txt(Index)
'                gSearchOption = "TOPOLOGY"
'                frmCari.Show (vbModal)
            End If
        End If
    ElseIf Index = 9 Or Index = 10 Then
        txt(Index) = Format(NulSID(txt(Index)), "0##")
    End If
End Sub

Private Sub VScroll1_Change()
    If VScroll1.value > 0 Then
        Frame2.Top = -VScroll1.value + 200
    ElseIf VScroll1.value = 0 Then
        Frame2.Top = 0
    End If
End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub VScroll1_Scroll()
    If VScroll1.value > 0 Then
        Frame2.Top = -VScroll1.value + 200
    ElseIf VScroll1.value = 0 Then
        Frame2.Top = 0
    End If
End Sub

Private Sub subletGrid_DblClick()
      Dim a As Integer

1         If barisSublet > 1 And subletGrid.TextMatrix(subletGrid.row, 1) <> "" Then
2             a = subletGrid.row
              
3             With subletGrid
4                 If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
5                     Text(86) = Trim(.TextMatrix(a, 8))
6                     Text(79) = Trim(.TextMatrix(a, 7))
7                     Text(80) = CDbl(.TextMatrix(a, 2))
8                     Text(81) = Trim(.TextMatrix(a, 3))
9                     Text(82) = CDbl(.TextMatrix(a, 4)) '* (100 / (100 + gUpSublet))
10                    chSublet = IIf(CDbl(.TextMatrix(a, 5)) = 0, Unchecked, Checked)
11                Else
12                    Text(79) = Trim(.TextMatrix(a, 1))
13                    Text(80) = CDbl(.TextMatrix(a, 2))
14                    Text(81) = Trim(.TextMatrix(a, 3))
15                    Text(82) = CDbl(.TextMatrix(a, 4)) '* (100 / (100 + gUpSublet))
16                    chSublet = IIf(CDbl(.TextMatrix(a, 5)) = 0, Unchecked, Checked)
17                End If
18            End With
              
19            For i = a To barisSublet
20                With subletGrid
21                    .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
22                    .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
23                    .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
24                    .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
25                    .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
26                    .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)
27                    .TextMatrix(i, 6) = .TextMatrix(i + 1, 6)
28                    If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
29                        .TextMatrix(i, 7) = .TextMatrix(i + 1, 7)
30                        .TextMatrix(i, 8) = .TextMatrix(i + 1, 8)
31                    End If
32                End With
33            Next
              
34            Text(52) = Format(CDbl(Text(52)) - (CDbl(Text(80)) * CDbl(Text(82))), "#0,#")
35            Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
                      
36            barisSublet = barisSublet - 1
37            Call hitungEstimasiSelesai
38            Text(79).SetFocus
39        End If
End Sub

Private Sub subletGrid_GotFocus()
          Dim bolehTambah As Boolean
          Dim rsDetailSublet  As ADODB.Recordset
1         If Trim(Text(79)) <> "" And Trim(Text(80)) <> "" And Trim(Text(81)) <> "" And Trim(Text(82)) <> "" Then
          
2             For i = 0 To baris
3                 If subletGrid.TextMatrix(i, 1) = Trim(Text(79)) Then
4                     bolehTambah = False
5                     Text(79) = ""
6                     Exit Sub
7                 Else
8                     bolehTambah = True
9                 End If
10            Next
11            If bolehTambah Then
12                If Log.isEmStall = 1 Then
13                    frmEstimasiSublet.Visible = True
14                    Text(91) = "00:00"
15                Else
16                    estimasiSublet = "00:00"
17                    If MaxBrs >= barisSublet Then
18                        With subletGrid
19                            If Log.NHS = 1 And Log.standarisasiSublet = 1 Then
20                                If NulSID(Text(86)) <> "" Then
                                      ''ambil data promo
21                                    sSQL = "SELECT D.KDSUPP,E.SUPPLIER,A.KDPROMO,A.NAMAPROMO,A.KETERANGAN,A.HARGAJUAL,D.HARGABELI,A.KDSUBLET,D.NAMASUBLET,D.KETERANGAN AS KETSUBLET,D.KODESATUAN"
22                                    sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLETPROMO A WITH (NOLOCK)"
23                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLET D WITH (NOLOCK) ON A.KDSUBLET = D.KDSUBLET"
24                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMOMDL B WITH (NOLOCK) ON A.KDPROMO = B.KDPROMO"
25                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETPROMODEALER C WITH (NOLOCK) ON A.KDPROMO = C.KDPROMO"
26                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUPPLIER E WITH (NOLOCK) ON E.KDSUPP = D.KDSUPP "
27                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL F WITH (NOLOCK) ON D.KDSUBLET = F.KDSUBLET"
28                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER G WITH (NOLOCK) ON D.KDSUBLET = G.KDSUBLET"
29                                    sSQL = sSQL & " LEFT JOIN T_SUPPLIER H WITH (NOLOCK) ON H.KDSUPP = E.KDSUPP"
30                                    sSQL = sSQL & " WHERE ISNULL(H.KODEAP,'') <> '' AND A.FS_LOG_DEL = '' AND A.KDPROMO = '" & NulSID(Text(86)) & "'"
31                                Else
                                      ''ambil data sublet
32                                    sSQL = "SELECT DISTINCT '' as KDPROMO,D.KDSUBLET,D.NAMASUBLET,D.KETERANGAN,D.HARGAJUAL,D.HARGABELI,E.KDSUPP,E.SUPPLIER,D.KODESATUAN"
33                                    sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..MSTSUBLET D "
34                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL B WITH (NOLOCK) ON D.KDSUBLET = B.KDSUBLET"
35                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER C WITH (NOLOCK) ON D.KDSUBLET = C.KDSUBLET"
36                                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..MSTSUPPLIER E WITH (NOLOCK) ON D.KDSUPP = E.KDSUPP"
37                                    sSQL = sSQL & " WHERE D.KDSUBLET = '" & NulSID(Text(79)) & "'"
38                                    sSQL = sSQL & " AND D.FS_LOG_DEL = ''"
39                                End If
40                                Set rsDetailSublet = objO.FetchData(sSQL)
41                                If Not rsDetailSublet.EOF Then
42                                    rsDetailSublet.MoveFirst
43                                    .TextMatrix(barisSublet, 1) = NulSID(rsDetailSublet!namaSublet)             ' NAMA SUBLET
44                                    .TextMatrix(barisSublet, 2) = CDbl(Text(80))                                ' QTY
45                                    .TextMatrix(barisSublet, 3) = NulSID(rsDetailSublet!KODESATUAN)             ' SATUAN
46                                    .TextMatrix(barisSublet, 4) = Format(CDbl(Text(82)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
47                                    .TextMatrix(barisSublet, 5) = IIf(chSublet = Checked, 1, 0)
48                                    .TextMatrix(barisSublet, 6) = estimasiSublet
49                                    .TextMatrix(barisSublet, 7) = NulSID(rsDetailSublet!KDSUBLET)         ''kodeitem
50                                    .TextMatrix(barisSublet, 8) = NulSID(rsDetailSublet!KDPROMO)         ''kodepromo
51                                Else
52                                    MsgBox "Master sublet tidak terdaftar" & vbCrLf & "Pastikan Enter untuk kode sublet / promo", vbInformation
53                                    Exit Sub
54                                End If
55                            Else
56                                .TextMatrix(barisSublet, 1) = Trim(Text(79))                                ' NAMA SUBLET
57                                .TextMatrix(barisSublet, 2) = CDbl(Text(80))                                ' QTY
58                                .TextMatrix(barisSublet, 3) = Trim(Text(81))                                ' SATUAN
59                                .TextMatrix(barisSublet, 4) = Format(CDbl(Text(82)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
60                                .TextMatrix(barisSublet, 5) = IIf(chSublet = Checked, 1, 0)
61                                .TextMatrix(barisSublet, 6) = estimasiSublet
62                            End If
63                        End With
64                        barisSublet = barisSublet + 1
65                        Text(52) = Format(CDbl(Text(52)) + (CDbl(Text(80)) * CDbl(Text(82))), "#0,#")
66                        Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
                          
67                        Text(79) = ""
68                        Text(80) = 0
69                        Text(81) = ""
70                        Text(82) = 0
71                        chSublet = Unchecked
72                        Call hitungEstimasiSelesai
73                        Text(79).SetFocus
74                    End If
                      '
75                End If
76            End If
77        End If
End Sub


Public Function GetDelIndex(ModilID As String) As String
          Dim acak As Double
1        On Error GoTo GetDelIndex_Error

2         Randomize Timer
3         acak = Int(Rnd * 99999)
4         GetDelIndex = ModilID & acak

5        On Error GoTo 0
6        Exit Function

GetDelIndex_Error:

7         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GetDelIndex of Form frmBooking")
End Function

Public Sub hitungEstimasiSelesai()
          Dim rsJob As ADODB.Recordset
          Dim rsSublet As ADODB.Recordset
          Dim totalJam As Integer
          Dim totalSublet As Integer
          Dim totalJob As Double
          Dim hasilJam As String
          'Dim nRate As Double
          Dim keterangan As String
          Dim jamTemSkrg As Integer
          Dim menitTemSkrg As Integer
          Dim jamMulai As String
          Dim jamTempSublet As Integer
          Dim menitTempSublet As Integer
          Dim totalTempSublet As Integer
1        On Error GoTo hitungEstimasiSelesai_Error

2         If Trim(Text(58)) <> ">>:>>" Then
3             jamMulai = Trim(Text(58))
4         Else
5             jamMulai = Trim(Text(24))
6         End If
          
7         totalSublet = 0
          
8         For i = 1 To barisSublet
9             With subletGrid
10                If .TextMatrix(i, 1) <> "" Then
11                    If .TextMatrix(i, 6) <> "" Then
12                    jamTempSublet = CInt(Left(Trim(.TextMatrix(i, 6)), InStr(.TextMatrix(i, 6), ":") - 1))
13                    menitTempSublet = Mid(Trim(.TextMatrix(i, 6)), InStr(Trim(.TextMatrix(i, 6)), ":") + 1, Len(Trim(.TextMatrix(i, 6))) - InStr(Trim(.TextMatrix(i, 6)), ":"))
14                    totalTempSublet = (jamTempSublet * 60) + menitTempSublet
15                    totalSublet = totalSublet + totalTempSublet
16                    End If
17                End If
18            End With
19        Next
              
20        totalJob = 0
21        nRate = 0
22        For i = 1 To baris - 1
23            If Trim(dGrid.TextMatrix(i, 1)) <> "" Then
                  
24               ' If Log.RPLHistori = 0 Then
25                    Set rsTemp = objO.FetchData("SELECT RATE FROM actRAte WITH (NOLOCK) WHERE MDL = '" & Text(10) & "' AND FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
26               ' Else
27                '    Set rsTemp = objO.FetchData("SELECT RATE FROM [" & dbReplikasiHO & "]..actRAte WITH (NOLOCK) WHERE MDL = '" & text(10) & "' AND FS_ITEM = '" & Trim(dGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
28                'End If
29                If Not rsTemp.EOF Then
30                    nRate = nRate + CDbl(rsTemp("RATE")) 'DGrid.TextMatrix(i, 2)
31                End If
32            End If
33        Next
          
         ' totalJob = 0
34        If Trim(Text(1)) <> "" Then
35            keterangan = Trim(Text(0)) & "-" & Trim(Text(56)) & "-" & Trim(Text(57)) & " |" & Trim(Text(1))
36        Else
37            keterangan = Trim(Text(0)) & "-" & Trim(Text(56)) & "-" & Trim(Text(57)) & " |"
38        End If
39        If chEM = Checked Then
40            totalJob = 0.5 * 60
41        ElseIf chSBNP = Checked Then
42            totalJob = 1 * 60
43        Else
              
44            Set rsJob = objO.FetchData("SELECT adjust,jammulai FROM tempstall WITH (NOLOCK) WHERE flag NOT IN ('k','x') and keterangan like '" & keterangan & "%'")
45            If Not rsJob.EOF Then
46                jamTemSkrg = CInt(Left(Trim(rsJob("adjust")), InStr(rsJob("adjust"), ":") - 1))
47                menitTemSkrg = Mid(Trim(rsJob("adjust")), InStr(Trim(rsJob("adjust")), ":") + 1, Len(Trim(rsJob("adjust"))) - InStr(Trim(rsJob("adjust")), ":"))
48                totalJob = (jamTemSkrg * 60) + menitTemSkrg
49                jamMulai = Trim(rsJob("jammulai"))
50            Else
51                Set rsJob = objO.FetchData("SELECT adjust,jammulai FROM stall WITH (NOLOCK) WHERE flag NOT IN ('k','x') and keterangan like '" & keterangan & "%'")
52                If Not rsJob.EOF Then
53                    jamTemSkrg = CInt(Left(Trim(rsJob("adjust")), InStr(rsJob("adjust"), ":") - 1))
54                    menitTemSkrg = Mid(Trim(rsJob("adjust")), InStr(Trim(rsJob("adjust")), ":") + 1, Len(Trim(rsJob("adjust"))) - InStr(Trim(rsJob("adjust")), ":"))
55                    totalJob = (jamTemSkrg * 60) + menitTemSkrg
56                    jamMulai = Trim(rsJob("jammulai"))
57                End If
58            End If
      '        If Trim(text(114)) <> "" Then
      '            totalJob = CDbl(CDbl(Trim(text(114))) * 60)
      '        End If
59        End If
          
60        totalMenit = CInt(totalSublet) + CInt(totalJob) + CInt(Log.estInv)
61        hasilJam = hitungTotalPenyerahan(CInt(totalMenit), jamMulai, Text(23))
              
62        Text(45) = Left(hasilJam, InStr(hasilJam, "|") - 1)
63        Text(59) = TglDDMMYYYY(Right(hasilJam, Len(hasilJam) - InStr(hasilJam, "|")))
          
64        Text(83) = Left(hasilJam, InStr(hasilJam, "|") - 1)
65        Text(84) = TglDDMMYYYY(Right(hasilJam, Len(hasilJam) - InStr(hasilJam, "|")))

66       On Error GoTo 0
67       Exit Sub

hitungEstimasiSelesai_Error:

68        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hitungEstimasiSelesai of Form frmBooking")
          
End Sub


Public Function getTotalPart(noOrder As String, isOrder As Boolean) As Double
          Dim rsPart As ADODB.Recordset
          Dim total As Double

1        On Error GoTo getTotalPart_Error

2         total = 0
3         If Log.isFill = 1 Then
4             Set rsPart = objO.FetchDataACC("SELECT QTY,HARGAJUAL FROM SOORD WHERE NOSO = '" & noOrder & "'")
5             If rsPart.RecordCount > 0 Then
6                 rsPart.MoveFirst
7                 Do While Not rsPart.EOF
8                     total = total + (CDbl(rsPart("QTY")) * CDbl(rsPart("HARGAJUAL")))
9                     rsPart.MoveNext
10                Loop
11            End If
12        Else
              'Set rsPart = objO.FetchDataACC("")
              
13        End If
14        getTotalPart = total

15       On Error GoTo 0
16       Exit Function

getTotalPart_Error:

17        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getTotalPart of Form frmBooking")

End Function

Private Sub prosesDSS(nActRate As Double)
1        On Error GoTo prosesDSS_Error

2         frameDSS.Visible = True
3         progDSS.Visible = False
4         cmdSimpanAdjustDSS.Enabled = True
5         cmdBatalAdjustDMS.Enabled = True
6         txtNRate.Text = nActRate
7         If Trim(nActRate) <> 0 Then
8             If InStr(nActRate, ",") > 0 Then
9                 Text(122) = Format(Left(nActRate, InStr(nActRate, ",") - 1), "0#") & ":" & Format(Val((CDbl(nActRate) - Left(nActRate, InStr(nActRate, ",") - 1)) * 60), "0#")
10            Else
11                Text(122) = Format(nActRate, "0#") & ":00"
12            End If
13        End If

14       On Error GoTo 0
15       Exit Sub

prosesDSS_Error:

16        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure prosesDSS of Form frmBooking")
End Sub

Private Sub pilihStallDSS()
1        On Error GoTo pilihStallDSS_Error
          Dim strSql As String
2         cmdSimpanAdjustDSS.Enabled = False
3         cmdBatalAdjustDMS.Enabled = False
          
          ''2016-05-03 | rate diambil dari adjust
4         txtNRate.Text = Round(cariJumlahMenit(Trim(Text(122))) / 60, 2)
          
          '---------- PENGECEKAN MINIMAL STALL -------------
5         strSql = "SELECT B.TIPEBP, C.SKILLMIN, E.LEVEL AS LvlTech, C.STALLMIN, D.level AS LvlStall, C.LiftNeed " & _
                  "FROM  HO_Flate B " & _
                  "LEFT JOIN MASTERDSS C ON B.TIPEBP = C.KODEDSS " & _
                  "LEFT JOIN LevelStall D ON C.StallMin = D.KodeStall " & _
                  "LEFT JOIN T_MECHLEV E ON C.SkillMin = E.LEVELMECH " & _
                  "WHERE B.MDL = '" & Trim(Text(10)) & "' AND (" & jobQuery & ") AND ISNULL(C.SkillMin,'') <> '' AND dss = 1 " & _
                  "order by LvlStall desc ,LvlTech desc "
          '-------------------------------------------------

6         Set rsTemp = objO.FetchData(strSql)
          
7         If Not rsTemp.EOF Then
              'jika nowo tsb masuk dss
8             rsTemp.MoveFirst
9             If Log.newJpcb = 1 Then
                  
10                strSql = "SELECT convert(integer,REPLACE(noStall,'Stall ','')) nos,*,"
11                strSql = strSql & " case when tek1 <> '' then (select top 1 tanggal + pause from tempProgress where MECH = st.tek1 order by TANGGAL + PAUSE desc) else '01/01/3000' end as idle1,"
12                strSql = strSql & " case when tek2 <> '' then (select top 1 tanggal + pause from tempProgress where MECH = st.tek2 order by TANGGAL + PAUSE desc) else '01/01/3000' end as idle2,"
13                strSql = strSql & " case when tek3 <> '' then (select top 1 tanggal + pause from tempProgress where MECH = st.tek3 order by TANGGAL + PAUSE desc) else '01/01/3000' end as idle3"
14                strSql = strSql & " FROM ("
15                strSql = strSql & " SELECT *,"
16                strSql = strSql & " CASE WHEN S.tek1 <> '' THEN (SELECT LEVEL FROM T_MECHLEV WHERE LEVELMECH = S.LEVEL1) else 0 end AS lvlTech1,"
17                strSql = strSql & " case WHEN S.tek2 <> '' THEN (SELECT LEVEL FROM T_MECHLEV WHERE LEVELMECH = S.LEVEL2) else 0 end AS lvltech2,"
18                strSql = strSql & " case WHEN S.tek3 <> '' THEN (SELECT LEVEL FROM T_MECHLEV WHERE LEVELMECH = S.LEVEL3) ELSE 0 END AS LvlTech3"
19                strSql = strSql & " FROM ("
20                strSql = strSql & " SELECT DISTINCT B.NOStall,b.activation, "
21                strSql = strSql & " b.LevelStall as LvlStall,"
22                strSql = strSql & " case when b.KodeStall = 'EM' then 1 else 0 end as EM,"
23                strSql = strSql & " case when b.KodeStall = 'SB' then 1 else 0 end as SB,"
24                strSql = strSql & " case when b.KodeStall = 'GR' then 1 else 0 end as GR,"
25                strSql = strSql & " case when b.KodeStall = 'SC' then 1 else 0 end as SC,"
26                strSql = strSql & " case when b.KodeStall = 'NP' then 1 else 0 end as NP,"
27                strSql = strSql & " case when b.lift = '1' then 1 else 0 end as Lifted,"
28                strSql = strSql & " tek1, ISNULL(C.levMech,'') AS LEVEL1,"
29                strSql = strSql & " tek2, ISNULL(D.levMech,'') AS LEVEL2,"
30                strSql = strSql & " TEK3, ISNULL(E.levMech,'') AS LEVEL3"
31                strSql = strSql & " FROM SettingStall B"
32                strSql = strSql & " LEFT JOIN MEKANIK C ON B.tek1 = C.MECH"
33                strSql = strSql & " LEFT JOIN MEKANIK D ON B.tek2 = D.MECH"
34                strSql = strSql & " LEFT JOIN MEKANIK E ON B.tek3 = E.MECH  ) S ) ST"
35                strSql = strSql & " WHERE st.activation = 1 and ST.Lifted = '" & rsTemp("LiftNeed") & "' AND (ST.LvlTech1 >= '" & rsTemp("LvlTech") & "' or st.lvltech2 >= '" & rsTemp("LvlTech") & "' or st.LvlTech3 >= '" & rsTemp("LvlTech") & "') "
36                strSql = strSql & " ORDER BY st.lvlTech1,st.lvltech2,st.LvlTech3,st.LvlStall,st.noStall  "
37            Else
38                strSql = "SELECT convert(integer,REPLACE(noStall,'Stall ','')) nos,*,"
39                strSql = strSql & " case when tek1 <> '' then (select top 1 tanggal + pause from tempProgress where MECH = st.tek1 order by TANGGAL + PAUSE desc) else '01/01/3000' end as idle1,"
40                strSql = strSql & " case when tek2 <> '' then (select top 1 tanggal + pause from tempProgress where MECH = st.tek2 order by TANGGAL + PAUSE desc) else '01/01/3000' end as idle2, "
41                strSql = strSql & " case when tek3 <> '' then (select top 1 tanggal + pause from tempProgress where MECH = st.tek3 order by TANGGAL + PAUSE desc) else '01/01/3000' end as idle3 " & _
                           "FROM ( " & _
                           "SELECT *, " & _
                           "CASE WHEN S.EM = 1 THEN (SELECT LEVEL FROM LevelStall WHERE KodeStall = 'EM') " & _
                           "WHEN S.SB = 1 THEN (SELECT LEVEL FROM LevelStall WHERE KodeStall = 'SB') " & _
                           "WHEN S.NP = 1 THEN (SELECT LEVEL FROM LevelStall WHERE KodeStall = 'NP') " & _
                           "WHEN S.SC = 1 THEN (SELECT LEVEL FROM LevelStall WHERE KodeStall = 'SC') " & _
                           "WHEN S.GR = 1 THEN (SELECT LEVEL FROM LevelStall WHERE KodeStall = 'GR') ELSE 0 END AS LvlStall, " & _
                           "CASE WHEN S.tek1 <> '' THEN (SELECT LEVEL FROM T_MECHLEV WHERE LEVELMECH = S.LEVEL1) else 0 end AS lvlTech1, " & _
                           "case WHEN S.tek2 <> '' THEN (SELECT LEVEL FROM T_MECHLEV WHERE LEVELMECH = S.LEVEL2) else 0 end AS lvltech2, " & _
                           "case WHEN S.tek3 <> '' THEN (SELECT LEVEL FROM T_MECHLEV WHERE LEVELMECH = S.LEVEL3) ELSE 0 END AS LvlTech3 " & _
                           "FROM ( " & _
                           "SELECT DISTINCT B.NOStall, " & _
                           "ISNULL((SELECT 1 FROM Lift WHERE Stall = B.NOSTALL AND TIPE = 'L'),'')  AS Lifted, " & _
                           "ISNULL((SELECT 1 FROM Lift WHERE Stall = B.NOSTALL AND TIPE = 'N'),'')  AS NP, " & _
                           "ISNULL((SELECT 1 FROM Lift WHERE Stall = B.NOSTALL AND TIPE = 'E' ),'')  AS EM, " & _
                           "ISNULL((SELECT 1 FROM Lift WHERE Stall = B.NOSTALL AND TIPE = 'S' ),'')  AS SB, " & _
                           "ISNULL((SELECT 1 FROM Lift WHERE Stall = B.NOSTALL AND TIPE = 'X' ),'')  AS SC, " & _
                           "CASE WHEN ISNULL(A.Stall,'') = '' THEN 1 ELSE 0 END AS GR, " & _
                           "tek1, ISNULL(C.levMech,'') AS LEVEL1, tek2, ISNULL(D.levMech,'') AS LEVEL2, TEK3, ISNULL(E.levMech,'') AS LEVEL3 " & _
                           "FROM SettingStall B LEFT JOIN Lift A ON B.noStall = A.Stall AND A.Tipe <> 'L' " & _
                           "LEFT JOIN MEKANIK C ON B.tek1 = C.MECH " & _
                           "LEFT JOIN MEKANIK D ON B.tek2 = D.MECH " & _
                           "LEFT JOIN MEKANIK E ON B.tek3 = E.MECH  " & _
                           ") S ) ST "
42                strSql = strSql & " WHERE ST.Lifted = '" & rsTemp("LiftNeed") & "' AND (ST.LvlTech1 >= '" & rsTemp("LvlTech") & "' or st.lvltech2 >= '" & rsTemp("LvlTech") & "' or st.LvlTech3 >= '" & rsTemp("LvlTech") & "') "
43                strSql = strSql & " ORDER BY st.lvlTech1,st.lvltech2,st.LvlTech3,st.LvlStall,st.noStall  "
44            End If
              
45            Set rsHlp = objO.FetchData(strSql)
46            If Not rsHlp.EOF Then
47                rsHlp.MoveFirst
48                progDSS.Visible = True
49                progDSS.Max = rsHlp.RecordCount + 1
                  Dim konter As Integer
50                konter = 1
                  Dim rsStall As ADODB.Recordset
                  Dim rsAbsen As ADODB.Recordset
                  Dim rsMek As ADODB.Recordset
                              
                  Dim waktuMulai As String
                  Dim waktuSelesai As String
                  Dim waktuMulaiIstirahat As String
                  Dim waktuSelesaiIstirahat As String
                  Dim totalIstirahat As Double
                  Dim totalAdjust As Double
                  Dim tanggalWo As String
                  
                  Dim tempJamSelesai As String
                  
                  Dim jamMulaiStall As String
                  Dim suggested As Boolean
                  Dim jamSelesaiStall As String
                  Dim minLvlMek As Integer
                  Dim jamMenit As String
                  
                  Dim mekanik As String
                  
51                If DateDiff("d", Format(isnow, "MM/DD/YYYY"), Format(Trim(Text(23)), "MM/DD/YYYY")) > 0 Then
52                Set rsAbsen = objO.FetchData("SELECT DISTINCT A.mech, name, '1' as masuk, ISNULL(HARI,'') AS HARI, ISNULL(TANGGAL,'') AS TANGGAL, " _
                                   & "ISNULL(SHIFT,0) AS SHIFT, ISNULL(JAMMULAI,'') AS JAMMULAI, ISNULL(JAMSELESAI,'') AS JAMSELESAI, ISNULL(MENITKERJA,0) AS MENITKERJA " _
                                   & ",ISNULL(S1.noStall,'') NOSTALL1,ISNULL(S2.noStall,'') NOSTALL2,ISNULL(S3.noStall,'') NOSTALL3 FROM MEKANIK A WITH(NOLOCK) LEFT OUTER JOIN SettingStall S1 WITH (NOLOCK) ON S1.tek1 = A.MECH LEFT OUTER JOIN SettingStall S2 WITH (NOLOCK) ON S2.tek2 = A.MECH LEFT OUTER JOIN SettingStall S3 WITH (NOLOCK) ON S3.tek3 = A.MECH LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH AND TANGGAL = '" & Trim(gTanggalStall) & "' AND HARI = '" & IIf(Weekday(gTanggalStall) = 1, 6, Weekday(gTanggalStall) - 2) & WeekdayName(IIf(Weekday(gTanggalStall) = 1, 8, Weekday(gTanggalStall)) - 1) & "'  ORDER BY a.MECH")
                  
53                Else
54                Set rsAbsen = objO.FetchData("SELECT DISTINCT A.mech, name, isnull(masuk,0) as masuk, ISNULL(HARI,'') AS HARI, ISNULL(TANGGAL,'') AS TANGGAL, " _
                                   & "ISNULL(SHIFT,0) AS SHIFT, ISNULL(JAMMULAI,'') AS JAMMULAI, ISNULL(JAMSELESAI,'') AS JAMSELESAI, ISNULL(MENITKERJA,0) AS MENITKERJA " _
                                   & ",ISNULL(S1.noStall,'') NOSTALL1,ISNULL(S2.noStall,'') NOSTALL2,ISNULL(S3.noStall,'') NOSTALL3 FROM MEKANIK A WITH(NOLOCK) LEFT OUTER JOIN SettingStall S1 WITH (NOLOCK) ON S1.tek1 = A.MECH LEFT OUTER JOIN SettingStall S2 WITH (NOLOCK) ON S2.tek2 = A.MECH LEFT OUTER JOIN SettingStall S3 WITH (NOLOCK) ON S3.tek3 = A.MECH LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH AND TANGGAL = '" & Trim(gTanggalStall) & "' AND HARI = '" & IIf(Weekday(gTanggalStall) = 1, 6, Weekday(gTanggalStall) - 2) & WeekdayName(IIf(Weekday(gTanggalStall) = 1, 8, Weekday(gTanggalStall)) - 1) & "' where isnull(masuk,'0') = '1' ORDER BY a.MECH")
                  
55                End If
                  
                 ' Set rsAbsen = objO.FetchData("SELECT DISTINCT A.mech, name, isnull(masuk,0) as masuk, ISNULL(HARI,'') AS HARI, ISNULL(TANGGAL,'') AS TANGGAL, " _
                                   & "ISNULL(SHIFT,0) AS SHIFT, ISNULL(JAMMULAI,'') AS JAMMULAI, ISNULL(JAMSELESAI,'') AS JAMSELESAI, ISNULL(MENITKERJA,0) AS MENITKERJA " _
                                   & ",ISNULL(S1.noStall,'') NOSTALL1,ISNULL(S2.noStall,'') NOSTALL2,ISNULL(S3.noStall,'') NOSTALL3 FROM MEKANIK A WITH(NOLOCK) LEFT OUTER JOIN SettingStall S1 WITH (NOLOCK) ON S1.tek1 = A.MECH LEFT OUTER JOIN SettingStall S2 WITH (NOLOCK) ON S2.tek2 = A.MECH LEFT OUTER JOIN SettingStall S3 WITH (NOLOCK) ON S3.tek3 = A.MECH LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH AND TANGGAL = '" & Trim(gTanggalStall) & "' AND HARI = '" & IIf(Weekday(gTanggalStall) = 1, 6, Weekday(gTanggalStall) - 2) & WeekdayName(IIf(Weekday(gTanggalStall) = 1, 8, Weekday(gTanggalStall)) - 1) & "' where isnull(masuk,'0') = '1' ORDER BY a.MECH")
56                If Log.newJpcb = 1 Then
57                    Set rsStall = objO.FetchData("SELECT WAKTUMULAI,WAKTUSELESAI,WAKTUMULAIISTIRAHAT,WAKTUSELESAIISTIRAHAT FROM T_STALL WITH(NOLOCK) WHERE hari = '" & IIf(Weekday(Trim(Text(23))) = 1, 6, Weekday(Trim(Text(23))) - 2) & WeekdayName(IIf(Weekday(Trim(Text(23))) = 1, 8, Weekday(Trim(Text(23)))) - 1) & "'")
58                Else
59                    Set rsStall = objO.FetchData("SELECT WAKTUMULAI,WAKTUSELESAI FROM T_STALL WITH(NOLOCK) WHERE hari = '" & IIf(Weekday(Trim(Text(23))) = 1, 6, Weekday(Trim(Text(23))) - 2) & WeekdayName(IIf(Weekday(Trim(Text(23))) = 1, 8, Weekday(Trim(Text(23)))) - 1) & "'")
60                End If
61                If rsStall.EOF Then
62                    MsgBox "Waktu Mulai dan Waktu Selesai Stall belum diset!", vbInformation
63                    frameDSS.Visible = False
64                    Exit Sub
65                Else
66                    rsStall.MoveFirst
67                    waktuMulai = Trim(rsStall("waktumulai"))
68                    waktuSelesai = Trim(rsStall("waktuselesai"))
69                    If Log.newJpcb = 1 Then
70                        waktuMulaiIstirahat = Trim(rsStall!waktuMulaiIstirahat)
71                        waktuSelesaiIstirahat = Trim(rsStall!waktuSelesaiIstirahat) '
72                        totalIstirahat = cariJumlahMenit(Trim(rsStall!waktuMulaiIstirahat)) - cariJumlahMenit(Trim(rsStall!waktuMulaiIstirahat))
73                    End If
74                End If
                  'frameHasilDSS.Visible = True
75                Call createRepoStall
76                suggested = True
77                Do While Not rsHlp.EOF
78                    progDSS.value = konter
79                    DoEvents
                      '-------------------------------------------------
                      'mencari jam tersedia pada stall
80                    If Trim(Text(23)) <> "" Then
81                        tanggalWo = (Format(Trim(Text(23)), "MM/DD/YYYY"))
82                    Else
83                        tanggalWo = (Format(isnow, "MM/DD/YYYY"))
84                    End If
                      
85                    If DateDiff("d", Format(isnow, "DD/MM/YYYY"), Format(Trim(Text(23)), "DD/MM/YYYY")) > 0 Then
86                        jamMenit = waktuMulai 'Trim(Text(24))
87                    Else
88                        jamMenit = Format(gTglServer, "HH:MM")
89                    End If
                      
                      'mencari teknisi yang absen pada stall tersebut
90                    If Not rsAbsen.EOF Then
91                        mekanik = ""
92                        minLvlMek = 99
93                        Set rsMek = objO.FetchData("SELECT * FROM SETTINGSTALL WITH (NOLOCK) WHERE NOSTALL = '" & rsHlp("NoStall") & "'")
94                        If Not rsMek.EOF Then
95                            If rsHlp("lvltech1") >= rsTemp("lvlTech") And rsMek.Fields(1).value <> "" And rsMek.Fields(1).value = rsHlp("tek1") Then
96                                If DateDiff("d", Format(isnow, "MM/DD/YYYY"), Format(Trim(Text(23)), "MM/DD/YYYY")) > 0 Then
97                                    If minLvlMek > rsHlp("lvltech1") Then minLvlMek = rsHlp("lvltech1")
98                                    mekanik = mekanik & UCase(Left(rsMek.Fields(1).value, 1)) & LCase(Mid(rsMek.Fields(1).value, 2, 2))
99                                Else
100                                   rsAbsen.filter = "mech = '" & rsMek.Fields(1).value & "' and masuk = 1"
                                      
101                                   If Not rsAbsen.EOF Then
102                                       If hitungMenit(Trim(Text(24))) >= hitungMenit(rsAbsen("jammulai")) And hitungMenit(Trim(Text(24))) <= hitungMenit(rsAbsen("jamselesai")) Then
103                                           If minLvlMek > rsHlp("lvltech1") Then minLvlMek = rsHlp("lvltech1")
104                                           mekanik = mekanik & UCase(Left(rsMek.Fields(1).value, 1)) & LCase(Mid(rsMek.Fields(1).value, 2, 2))
105                                           waktuSelesai = Trim(rsAbsen("JAMSELESAI"))
106                                       End If
107                                   End If
108                                   rsAbsen.filter = ""
109                               End If
110                           End If
111                           If rsHlp("lvltech2") >= rsTemp("lvlTech") And rsMek.Fields(2).value <> "" And rsMek.Fields(2).value = rsHlp("tek2") Then
112                               If DateDiff("d", Format(isnow, "MM/DD/YYYY"), Format(Trim(Text(23)), "MM/DD/YYYY")) > 0 Then
113                                   If minLvlMek > rsHlp("lvltech2") Then minLvlMek = rsHlp("lvltech2")
114                                   mekanik = mekanik & IIf(mekanik = "", "", " ,") & UCase(Left(rsMek.Fields(2).value, 1)) & LCase(Mid(rsMek.Fields(2).value, 2, 2))
115                               Else
116                                   rsAbsen.filter = "mech = '" & rsMek.Fields(2).value & "' and masuk = 1"
                                      
117                                   If Not rsAbsen.EOF Then
118                                       If hitungMenit(Trim(Text(24))) >= hitungMenit(rsAbsen("jammulai")) And hitungMenit(Trim(Text(24))) <= hitungMenit(rsAbsen("jamselesai")) Then
119                                           If minLvlMek > rsHlp("lvltech2") Then minLvlMek = rsHlp("lvltech2")
120                                           mekanik = mekanik & IIf(mekanik = "", "", " ,") & UCase(Left(rsMek.Fields(2).value, 1)) & LCase(Mid(rsMek.Fields(2).value, 2, 2))
121                                           waktuSelesai = Trim(rsAbsen("JAMSELESAI"))
122                                       End If
123                                   End If
124                                   rsAbsen.filter = ""
125                               End If
126                           End If
127                           If rsHlp("lvltech3") >= rsTemp("lvlTech") And rsMek.Fields(3).value <> "" And rsMek.Fields(3).value = rsHlp("tek3") Then
128                               If DateDiff("d", Format(isnow, "MM/DD/YYYY"), Format(Trim(Text(23)), "MM/DD/YYYY")) > 0 Then
129                                   If minLvlMek > rsHlp("lvltech3") Then minLvlMek = rsHlp("lvltech3")
130                                   mekanik = mekanik & IIf(mekanik = "", "", " ,") & UCase(Left(rsMek.Fields(3).value, 1)) & LCase(Mid(rsMek.Fields(3).value, 2, 2))
131                               Else
132                                   rsAbsen.filter = "mech = '" & rsMek.Fields(3).value & "' and masuk = 1"
                                      
133                                   If Not rsAbsen.EOF Then
134                                       If hitungMenit(Trim(Text(24))) >= hitungMenit(rsAbsen("jammulai")) And hitungMenit(Trim(Text(24))) <= hitungMenit(rsAbsen("jamselesai")) Then
135                                           If minLvlMek > rsHlp("lvltech3") Then minLvlMek = rsHlp("lvltech3")
136                                           mekanik = mekanik & IIf(mekanik = "", "", " ,") & UCase(Left(rsMek.Fields(3).value, 1)) & LCase(Mid(rsMek.Fields(3).value, 2, 2))
137                                           waktuSelesai = Trim(rsAbsen("JAMSELESAI"))
138                                       End If
139                                   End If
140                                   rsAbsen.filter = ""
141                               End If
142                           End If
143                       End If
                          
144                       If mekanik <> "" Then
                              'jika mekanik absen
145                           strSql = "SELECT * FROM STALL WITH(NOLOCK) WHERE FLAG <> 'X' AND TANGGAL = '" & tanggalWo & "' and NoStall = '" & rsHlp("NoStall") & "' "
146                           strSql = strSql & " UNION "
147                           strSql = strSql & " SELECT * FROM TEMPSTALL WITH(NOLOCK) WHERE FLAG <> 'X' AND TANGGAL = '" & tanggalWo & "' and NoStall = '" & rsHlp("NoStall") & "' "
148                           strSql = strSql & " order by jammulai,jamselesai"
149                           Set rsStall = objO.FetchData(strSql)
150                           jamMulaiStall = "00:00"
                              
151                           jamSelesaiStall = "00:00"
152                           If Not rsStall.EOF Then
                                  'lakukan pengecekan jam pengerjaan yang sesuai
153                               rsStall.MoveFirst
154                               tempJamSelesai = "00:00"
155                               Do While Not rsStall.EOF
156                                   jamMulaiStall = "00:00"
157                                   jamSelesaiStall = "00:00"
                                      '-------------------------------------------------------------------
                                      '- 2015-04-07 | RUDI | JAM MENIT WO DIGANTI BERDASARKAN JAM SERVER -
                                      '-------------------------------------------------------------------
                                      
                                      'If (cariJumlahMenit(rsStall("jamselesai"))) >= (cariJumlahMenit(Trim(text(66)))) Then
                                      
158                                   If (cariJumlahMenit(rsStall("jamselesai"))) >= (cariJumlahMenit(jamMenit)) Then
                                          'yang dicari jam wo keatas
159                                       If tempJamSelesai = "00:00" Then tempJamSelesai = waktuMulai
160                                       If (cariJumlahMenit(jamMenit)) >= cariJumlahMenit(tempJamSelesai) Then
161                                           tempJamSelesai = jamMenit 'Format(gTglServer, "HH:MM")
162                                       End If
                                          
163                                       If Log.newJpcb = 1 Then
164                                           If cariJumlahMenit(waktuMulaiIstirahat) >= cariJumlahMenit(rsStall("jamMulai")) And cariJumlahMenit(waktuSelesaiIstirahat) <= cariJumlahMenit(rsStall("jamMulai")) Then
165                                               jamMulaiStall = waktuSelesaiIstirahat
166                                           End If
167                                       End If
168                                       If (cariJumlahMenit(rsStall("jammulai")) - cariJumlahMenit(Trim(tempJamSelesai))) >= cariJumlahMenit(Trim(Text(122))) Then
                                              'jika ada waktu tersedia dari jam wo sampai jam mulai pada stall tersebut
169                                           jamMulaiStall = Trim(tempJamSelesai)
170                                           jamSelesaiStall = Trim(rsStall("jamMulai"))
171                                           With rsRecoStall
172                                               .AddNew
173                                               !Recom = IIf(suggested, "Sugested", "Alternative")
174                                               !stall = rsHlp("NoStall")
175                                               !Tanggal = Format(tanggalWo, "DD.MM.YYYY") & " " & jamMulaiStall & " - " & jamSelesaiStall
176                                               !Teknisi = mekanik
177                                               !waktuAwal = jamMulaiStall
178                                               !waktuAkhir = jamSelesaiStall
179                                               !minmek = minLvlMek
180                                               !nos = rsHlp("nos")
181                                               .Update
182                                               .MoveNext
183                                           End With
184                                           suggested = False
185                                           tempJamSelesai = Trim(rsStall("jamSelesai"))
186                                       End If
                                          
187                                       If jamMulaiStall = "00:00" And jamSelesaiStall = "00:00" And tempJamSelesai <> "00:00" Then
                                              'jika belum menemukan, lakukan perbandingan tempjamselesai dengan jam mulai
188                                           If Log.newJpcb = 1 Then
189                                               If cariJumlahMenit(waktuMulaiIstirahat) >= cariJumlahMenit(tempJamSelesai) And cariJumlahMenit(waktuSelesaiIstirahat) <= cariJumlahMenit(tempJamSelesai) Then
                                                      'jamMulaiStall = waktuSelesaiIstirahat
190                                                   tempJamSelesai = waktuSelesaiIstirahat
191                                               End If
192                                           End If
193                                           If (cariJumlahMenit(rsStall("jamMulai")) - cariJumlahMenit(tempJamSelesai)) >= cariJumlahMenit(Text(122)) Then
194                                               jamMulaiStall = jamMenit 'Format(gTglServer, "HH:MM") 'Trim(text(66))
195                                               jamSelesaiStall = Trim(rsStall("jamMulai"))
196                                               With rsRecoStall
197                                                   .AddNew
198                                                   !Recom = IIf(suggested, "Sugested", "Alternative")
199                                                   !stall = rsHlp("NoStall")
200                                                   !Tanggal = Format(tanggalWo, "DD.MM.YYYY") & " " & jamMulaiStall & " - " & jamSelesaiStall
201                                                   !Teknisi = mekanik
202                                                   !waktuAwal = jamMulaiStall
203                                                   !waktuAkhir = jamSelesaiStall
204                                                   !minmek = minLvlMek
205                                                   !nos = rsHlp("nos")
206                                                   .Update
207                                                   .MoveNext
208                                               End With
209                                               If suggested Then
210                                                   suggested = False
211                                               End If
212                                           End If
213                                           tempJamSelesai = Trim(rsStall("jamselesai"))
214                                       ElseIf jamMulaiStall = "00:00" And jamSelesaiStall = "00:00" And tempJamSelesai = "00:00" Then
                                              'jika pertama kali dan belum ada data yang diisi / dipilih kandidat
215                                           tempJamSelesai = Trim(rsStall("jamselesai"))
216                                       End If
217                                   End If
218                                   tempJamSelesai = Trim(rsStall("jamselesai"))
219                                   rsStall.MoveNext
220                               Loop
                                  'lakukan pengecekan jam selesai terakhir dengan jam terakhir stall tersebut
221                               If (cariJumlahMenit(jamMenit)) >= cariJumlahMenit(tempJamSelesai) Then
222                                   tempJamSelesai = jamMenit 'Format(gTglServer, "HH:MM")
223                               End If
224                               If Log.newJpcb = 1 Then
225                                   If cariJumlahMenit(waktuMulaiIstirahat) >= cariJumlahMenit(tempJamSelesai) And cariJumlahMenit(waktuSelesaiIstirahat) <= cariJumlahMenit(tempJamSelesai) Then
                                          'jamMulaiStall = waktuSelesaiIstirahat
226                                       tempJamSelesai = waktuSelesaiIstirahat
227                                   End If
228                               End If
229                               If (cariJumlahMenit(waktuSelesai) - cariJumlahMenit(tempJamSelesai)) >= cariJumlahMenit(Text(122)) Then
                                      
230                                   jamMulaiStall = Trim(tempJamSelesai)
231                                   jamSelesaiStall = Trim(waktuSelesai)
232                                   With rsRecoStall
233                                       .AddNew
234                                       !Recom = IIf(suggested, "Sugested", "Alternative")
235                                       !stall = rsHlp("NoStall")
236                                       !Tanggal = Format(tanggalWo, "DD.MM.YYYY") & " " & jamMulaiStall & " - " & jamSelesaiStall
237                                       !Teknisi = mekanik
238                                       !waktuAwal = jamMulaiStall
239                                       !waktuAkhir = jamSelesaiStall
240                                       !minmek = minLvlMek
241                                       !nos = rsHlp("nos")
242                                       .Update
243                                       .MoveNext
244                                   End With
                                      
245                                   If suggested Then
246                                       suggested = False
247                                   End If
248                               End If
249                           Else
                                  'jika stall belum ada data yang diisikan
                                  
      '                            If cariJumlahMenit(Trim(text(66))) > cariJumlahMenit(waktuMulai) Then
      '                                jamMulaiStall = Trim(text(66))
250                               If cariJumlahMenit(jamMenit) > cariJumlahMenit(waktuMulai) Then
                                  
251                                   jamMulaiStall = jamMenit 'Format(gTglServer, "HH:MM")
252                               Else
253                                   jamMulaiStall = waktuMulai
254                               End If
                                  '----------------------------------------------------------------------
                                  ' 2015-06-12 | RUDI | MENAMBAHKAN PENGECEKAN JAM MULAI MASUK ISTIRAHAT
                                  '----------------------------------------------------------------------
255                               If Log.newJpcb = 1 Then
256                                   If cariJumlahMenit(waktuMulaiIstirahat) >= cariJumlahMenit(jamMulaiStall) And cariJumlahMenit(waktuSelesaiIstirahat) <= cariJumlahMenit(jamMulaiStall) Then
257                                       jamMulaiStall = waktuSelesaiIstirahat
258                                   End If
259                               End If
260                               jamSelesaiStall = waktuSelesai
                                  
                                  
                                  
                                  'mengecek ketersediaan waktu
261                               If (cariJumlahMenit(jamSelesaiStall) - cariJumlahMenit(jamMulaiStall)) >= (cariJumlahMenit(Trim(Text(122)))) Then
                                      
262                                   With rsRecoStall
263                                       .AddNew
264                                       !Recom = IIf(suggested, "Sugested", "Alternative")
265                                       !stall = rsHlp("NoStall")
266                                       !Tanggal = Format(tanggalWo, "DD.MM.YYYY") & " " & jamMulaiStall & " - " & jamSelesaiStall
267                                       !Teknisi = mekanik
268                                       !waktuAwal = jamMulaiStall
269                                       !waktuAkhir = jamSelesaiStall
270                                       !minmek = minLvlMek
271                                       !nos = rsHlp("nos")
272                                       .Update
273                                       .MoveNext
274                                   End With
275                               End If
276                           End If
                          
                          
277                       End If
278                   End If
                      'MsgBox rsRecoStall.RecordCount
                      '-------------------------------------------------
                      
279                   konter = konter + 1
280                   rsHlp.MoveNext
281               Loop
282                   If rsRecoStall.RecordCount = 0 Then
283                       GoTo pilihStallManualDMSNonDSS
284                   Else
285                       With rsRecoStall
286                           .AddNew
287                           !Recom = "Pilih Manual Stall"
288                           !stall = "Stall Manual"
289                           !Tanggal = Format(tanggalWo, "DD.MM.YYYY")
290                           !Teknisi = ""
291                           !waktuAwal = ""
292                           !waktuAkhir = ""
293                           !nos = 99
294                           .Update
295                           .MoveNext
296                       End With
297                       rsRecoStall.MoveFirst
298                   End If
                      
299               If rsRecoStall.RecordCount > 0 Then
300                   frameDSS.Visible = False
                      'lakukan pengurutan data
301                   rsRecoStall.filter = "RECOM <> 'Pilih Manual Stall'"
302                   rsRecoStall.Sort = "waktuAwal,minmek,nos"
303                   rsRecoStall.MoveFirst
304                   rsRecoStall!Recom = "Sugested"
305                   rsRecoStall.Update
306                   rsRecoStall.MoveNext
                      
307                   Do While Not rsRecoStall.EOF
308                       If rsRecoStall!Recom <> "Pilih Manual Stall" Then
309                           rsRecoStall!Recom = "Alternative"
310                       End If
311                       rsRecoStall.MoveNext
312                   Loop
313                   rsRecoStall.filter = ""
314                   rsRecoStall.MoveFirst
      '                Set gridDSS.DataSource = rsRecoStall
      '
      '                backgroundDSS.Visible = True
      '                backgroundDSS.Top = 0
      '
315                   frameHasilDSS.Visible = True
316                   frameHasilDSS.Top = 0 '5400
                      'frameDSS.Top = SSTab1.Top
317               End If
318           Else
319               If MsgBox("Tidak ada data yang di proses di DSS, apakah ingin memilih stall manual?", vbYesNo) = vbYes Then
                      'Call prosesPilihStallManual(False)
320                   GoTo pilihStallManualDMSNonDSS
                      
321               Else
322                   Exit Sub
323               End If
324           End If
325       Else
              'lakukan pemilihan stall manual
              'Call prosesPilihStallManual(False)
326           GoTo pilihStallManualDMSNonDSS
327       End If
328       Exit Sub
pilihStallManualDMSNonDSS:
329       sgnDSS.Caption = "NON DSS"
330       With frmStall
331           .Text(13).Enabled = True
332           .Text(7) = Text(23)
333           .Text(14) = Trim(Text(122)) ''Trim(Text(85))
334           .Text(4).Mask = ""
335           .Text(4) = Trim(Text(0)) & IIf(Trim(Text(0)) <> "", "-", "") & Trim(Text(56)) & IIf(Trim(Text(57)) <> "", "-", "") & Trim(Text(57)) ' noPol
336           .Text(5) = Text(2)
337           .Text(6) = Text(3)
338           .Text(12) = Text(1)
339           .Text(9) = IIf(chEM = Checked, 0.5, IIf(chSBNP = Checked, 1, nRate))
340           .Text(8) = nMech
341           .chkThs.value = chkThs.value
342           .chEM.value = IIf(chEM = Checked, chEM.value, IIf(chSBNP = Checked, chSBNP.value, 0))
343           .frmStal.Visible = False
344           If isBookByMrs = False Then
345               NotChild frmStall.hWnd
346           Else
347               TOChild frmBooking.hWnd, frmStall.hWnd
                  '.Show
348           End If
349           .WindowState = vbMaximized
              'frmStall.Show
350       End With
351       Exit Sub
          
352      On Error GoTo 0
353      Exit Sub

pilihStallDSS_Error:

354       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure pilihStallDSS of Form frmBooking")
End Sub

Private Sub createRepoStall()
1         On Error GoTo vb_error
      '  Set rsRecoStall = New ADODB.Recordset
2         If rsRecoStall.State Then rsRecoStall.Close
3         Set rsRecoStall = Nothing
4         With rsRecoStall
5             .Fields.Append "Recom", adVarChar, 25
6             .Fields.Append "Stall", adVarChar, 25
7             .Fields.Append "Tanggal", adVarChar, 25
8             .Fields.Append "Teknisi", adVarChar, 25
9             .Fields.Append "WaktuAwal", adVarChar, 25
10            .Fields.Append "WaktuAkhir", adVarChar, 25
11            .Fields.Append "MinMek", adInteger
12            .Fields.Append "nos", adInteger
13            .Open
14        End With
15        Set gridDSS.DataSource = rsRecoStall
16        Call headGrid
17        Exit Sub
vb_error:
18        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Sub

Private Sub headGrid()
1        On Error GoTo headGrid_Error

2         With gridDSS
3             .Columns(5).Caption = "WAKTU AKHIR"
4             .Columns(5).DataField = "WaktuAkhir"
5             .Columns(5).Width = 2000
6             .Columns(5).Locked = False
              
7             .Columns(0).Caption = "REKOMENDASI"
8             .Columns(0).DataField = "Recom"
9             .Columns(0).Width = 2500
10            .Columns(0).Locked = True
              
11            .Columns(1).Caption = "STALL"
12            .Columns(1).DataField = "Stall"
13            .Columns(1).Width = 2000
14            .Columns(1).Locked = True
              
15            .Columns(2).Caption = "TANGGAL"
16            .Columns(2).DataField = "Tanggal"
17            .Columns(2).Width = 3000
18            .Columns(2).Locked = True
              
19            .Columns(3).Caption = "TEKNISI"
20            .Columns(3).DataField = "Teknisi"
21            .Columns(3).Width = 1000
22            .Columns(3).Locked = True
              
23            .Columns(4).Caption = "WAKTU AWAL"
24            .Columns(4).DataField = "WaktuAwal"
25            .Columns(4).Width = 2000
26            .Columns(4).Locked = False
27          End With

28       On Error GoTo 0
29       Exit Sub

headGrid_Error:

30        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headGrid of Form frmBooking")
End Sub

Private Sub saveStallDSS(stl As String, Tgl As Date, jmMul As String, jmSel As String, rte As Double, grp As String, mch As String, Ket As String, adjust As String)
          Dim rsTempSaveStall As ADODB.Recordset
          Dim rsTempSaveStall2 As ADODB.Recordset

1        On Error GoTo saveStallDSS_Error

2         Set rsTempSaveStall = objO.FetchData("SELECT * FROM Tempstall WITH(NOLOCK) WHERE nostall = '" & stl & "' AND tanggal = '" & Format(Tgl, "MM/DD/YYYY") & "' AND ((jamMulai BETWEEN '" & jmMul & "' AND '" & jmSel & "') OR (jamSelesai BETWEEN '" & jmMul & "' AND '" & jmSel & "'))")
3         If rsTempSaveStall.EOF Then
4             Set rsTempSaveStall2 = objO.FetchData("SELECT * FROM tempstall WITH(NOLOCK) WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan LIKE '" & Left(Ket, InStr(Ket, "|") - 2) & " |%' ")
5             If rsTempSaveStall2.EOF Then
6                 Call objHT.InsertData_BengkelOnly("INSERT INTO tempstall VALUES ('" & stl & "','" & Format(Tgl, "mm/dd/yyyy") & "','" & jmMul & "','" & jmSel & "','" & Replace(CDbl(rte), ",", ".") & "','','" & mch & "','','" & Ket & "','" & Trim(adjust) & "')")     '" & grp & "
7             Else
8                 Call objHT.UpdateData_BengkelOnly("UPDATE tempstall SET noStall = '" & stl & "', jamMulai = '" & jmMul & "', jamSelesai = '" & jmSel & "', rate = '" & Replace(CDbl(rte), ",", ".") & "', KETERANGAN = '" & Ket & "',ADJUST = '" & Trim(adjust) & "' WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan LIKE '" & Left(Ket, InStr(Ket, "|") - 2) & " |%'  ")
9             End If
10        Else
11            If Left(Trim(rsTempSaveStall("KETERANGAN")), InStr(rsTempSaveStall("KETERANGAN"), "|")) = Left(Ket, InStr(Ket, "|")) Then
12                Call objHT.UpdateData_BengkelOnly("UPDATE tempstall SET noStall = '" & stl & "', jamMulai = '" & jmMul & "', jamSelesai = '" & jmSel & "', mech = '" & mch & "', rate = '" & Replace(CDbl(rte), ",", ".") & "', keterangan = '" & Ket & "',ADJUST = '" & Trim(adjust) & "' WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan = '" & rsTempSaveStall("KETERANGAN") & "'  ")
13            Else
14                MsgBox "Alokasi Stall dan Waktu sudah dipakai" & vbCr & "Kendaraan dengan No. Polisi : " & Trim(rsTempSaveStall!keterangan)
15                erTranAll = True
16                Exit Sub
17            End If
18        End If
      '    If gStall = "order" Then
      '        Call objHT.UpdateData_BengkelOnly("UPDATE tempStall SET flag = 'N' WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan = '" & Ket & "'")
      '    End If

19       On Error GoTo 0
20       Exit Sub

saveStallDSS_Error:

21        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure saveStallDSS of Form frmBooking")
End Sub

Private Function cariJumlahMenit(jam As String) As Double
          Dim jJam As Integer
          Dim jMenit As Integer
1        On Error GoTo cariJumlahMenit_Error

2         jam = Trim(jam)
3         jJam = CInt(Left(jam, InStr(1, jam, ":") - 1))
4         jMenit = CInt(Right(jam, InStr(1, jam, ":") - 1))
          
5         cariJumlahMenit = CDbl(((jJam * 60) + (jMenit)))

6        On Error GoTo 0
7        Exit Function

cariJumlahMenit_Error:

8         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariJumlahMenit of Form frmBooking")
End Function


Private Sub cetakDetailFIR(keluhan As Byte)
         Dim rpt As New rptDetailFIR
1        Set rsTemp = objO.FetchData("SELECT J.FRONT,J.FMDL,J.MDL,D.* FROM DETAILFIR d WITH (NOLOCK) LEFT JOIN BOOK J WITH (NOLOCK) ON J.nofak = D.nofak WHERE d.NOFAK = '" & Trim(Text(1)) & "' AND IDKELUHAN = '" & keluhan & "'")
         ''kosongkan data
         ''data kendaraan
2        rpt.txtTanggal.SetText ""
3        rpt.txtNoPolisi.SetText ""
4        rpt.txtModel.SetText ""
5        rpt.txtKodeModel.SetText ""
6        rpt.txtNoRangka.SetText ""
7        rpt.Text8.SetText ""
         ''data keluhan
8        rpt.txtKeluhanPelanggan.SetText ""
         
         ''data diagnosa
9        rpt.Text23.SetText ""
10       rpt.Text25.SetText ""
11       rpt.Text27.SetText ""
12       rpt.Text29.SetText "... Hari yang lalu"
         
13       rpt.Text30.SetText ""
14       rpt.Text31.SetText ""
15       rpt.Text32.SetText ""
         
16       rpt.Text36.SetText ""
17       rpt.Text37.SetText ""
18       rpt.Text38.SetText ""
         
19       rpt.Text43.SetText ""
20       rpt.Text44.SetText ""
21       rpt.Text45.SetText ""
         
22       rpt.Text50.SetText ""
23       rpt.Text51.SetText ""
24       rpt.Text52.SetText ""
         
25       rpt.Text57.SetText ""
26       rpt.Text58.SetText ""
27       rpt.Text59.SetText ""
28       rpt.Text60.SetText ""
29       rpt.Text61.SetText ""
30       rpt.Text62.SetText ""
31       rpt.Text63.SetText ""
32       rpt.Text64.SetText ""
33       rpt.Text65.SetText ""
34       rpt.Text66.SetText ""
35       rpt.Text67.SetText ""
         
         ''kondisi saat gejala terjadi
36       rpt.Text83.SetText ""
37       rpt.Text85.SetText ""
38       rpt.Text89.SetText ""
39       rpt.Text7.SetText ""
40       rpt.Text94.SetText ""
41       rpt.Text96.SetText ""
42       rpt.Text99.SetText ""
43       rpt.Text101.SetText ""
44       rpt.Text9.SetText ""
         
45       If Not rsTemp.EOF Then
46          rsTemp.MoveFirst
            ''data kendaraan
47          rpt.Text8.SetText "SA : " & Trim(rsTemp!FRONT)
48          rpt.txtTanggal.SetText TglDDMMYYYY(Trim(rsTemp!Tanggal))
49          rpt.txtNoPolisi.SetText Trim(rsTemp!NO_POLISI)
50          rpt.txtModel.SetText Trim(rsTemp!Fmdl)
51          rpt.txtKodeModel.SetText Trim(rsTemp!MDL)
52          rpt.txtNoRangka.SetText Trim(rsTemp!NO_CHASSIS)
            
            ''data keluhan
53          rpt.txtKeluhanPelanggan.SetText Trim(rsTemp!keluhan)
            
            ''data diagnosa
            
            ''gejala pertama kali
54          If Trim(rsTemp!gejalapertama) <> "" Then
55             Select Case UCase(Trim(rsTemp!gejalapertama))
                  Case "HARI INI"
56                   rpt.Text23.SetText " X"
57                Case "MINGGU LALU"
58                   rpt.Text25.SetText " X"
59                Case Else
60                   rpt.Text27.SetText " X"
61                   rpt.Text29.SetText Mid(Trim(rsTemp!gejalapertama), 1, Len(rsTemp!gejalapertama) - 15) & " Hari yang lalu"
62             End Select
63          End If
            
64          If Trim(rsTemp!frekuensi) <> "" Then
65             Select Case Trim(rsTemp!frekuensi)
                   Case "Sekali"
66                   rpt.Text30.SetText " X"
67                Case "Kadang"
68                   rpt.Text31.SetText " X"
69                Case Else
70                   rpt.Text32.SetText " X"
71             End Select
72          End If
            
73          If Trim(rsTemp!indikator) <> "" Then
74             Select Case Trim(rsTemp!indikator)
                   Case "ON"
75                   rpt.Text36.SetText " X"
76                Case "Selalu"
77                   rpt.Text37.SetText " X"
78                Case Else
79                   rpt.Text38.SetText " X"
80             End Select
81          End If
             
82          If Trim(rsTemp!saatmil) <> "" Then
83             Select Case Trim(rsTemp!saatmil)
                   Case "Idling"
84                   rpt.Text43.SetText " X"
85                Case "Starting"
86                   rpt.Text44.SetText " X"
87                Case Else
88                   rpt.Text45.SetText " X"
89             End Select
90          End If
            
91          If Trim(rsTemp!kondisiMesin) <> "" Then
92             Select Case Trim(rsTemp!kondisiMesin)
                  Case "Panas"
93                   rpt.Text50.SetText " X"
94                Case "Dingin"
95                   rpt.Text51.SetText " X"
96                Case Else
97                   rpt.Text52.SetText " X"
98             End Select
99          End If
            
            Dim hasil() As String
100         If Trim(rsTemp!shiftlevel) <> "" Then
101               hasil() = Split(Trim(rsTemp!shiftlevel), ",")
102               For i = 0 To UBound(hasil)
103                  Select Case hasil(i)
                         Case "1"
104                        rpt.Text57.SetText " X"
105                     Case "2"
106                        rpt.Text60.SetText " X"
107                     Case "3"
108                        rpt.Text62.SetText " X"
109                     Case "4"
110                        rpt.Text64.SetText " X"
111                     Case "5"
112                        rpt.Text66.SetText " X"
113                     Case "P"
114                        rpt.Text58.SetText " X"
115                     Case "R"
116                        rpt.Text61.SetText " X"
117                     Case "N"
118                        rpt.Text63.SetText " X"
119                     Case "D"
120                        rpt.Text65.SetText " X"
121                     Case "S"
122                        rpt.Text67.SetText " X"
123                     Case Else
124                        rpt.Text59.SetText " X"
125                  End Select
126               Next
127         End If
            
            ''kondisi saat gejala terjadi
             
128         If Trim(rsTemp!Speed) <> "" Then
129            rpt.Text83.SetText Trim(rsTemp!Speed)
130         End If
                  
131         If Trim(rsTemp!rpm) <> "" Then
132            rpt.Text85.SetText Trim(rsTemp!rpm)
133         End If
                  
134         If Trim(rsTemp!beban) <> "" Then
135            rpt.Text89.SetText Trim(rsTemp!beban)
136         End If
                  
137         If Trim(rsTemp!penumpang) <> "" Then
138            rpt.Text91.SetText Round(Trim(rsTemp!penumpang), 0)
139         End If
                  
140         If Trim(rsTemp!kondisiJalan) <> "" Then
141            rpt.Text7.SetText Replace(Trim(rsTemp!kondisiJalan), ",", " / ")
142         End If
                  
143         If Trim(rsTemp!laluLintas) <> "" Then
144            Select Case Trim(rsTemp!laluLintas)
                  Case "Macet Stop & Go"
145                  rpt.Text94.SetText Trim(rsTemp!laluLintas)
146               Case "Lancar"
147                  rpt.Text94.SetText Trim(rsTemp!laluLintas)
148               Case Else
149                  rpt.Text94.SetText "Lainnya : " & Trim(rsTemp!laluLintas)
150            End Select
151         End If
                   
152         If Trim(rsTemp!cuaca) <> "" Then
153            rpt.Text96.SetText Trim(rsTemp!cuaca)
154         End If
            
            
155         If Trim(rsTemp!blowerspeedac) <> "" Then
156            rpt.Text99.SetText Trim(rsTemp!blowerspeedac)
157         End If
              
158         If Trim(rsTemp!tempsetting) <> "" Then
159            rpt.Text101.SetText Round(Trim(rsTemp!tempsetting), 0)
160         End If
            
161         rpt.Text9.SetText keluhan
            
162      End If
163      Set FrmctkOrder.Report = rpt
164      With FrmctkOrder.CRViewer1
165          .ReportSource = rpt
166          .ViewReport
167          .Zoom 100
168          .Visible = True
169      End With
170      FrmctkOrder.cmdTutup.Visible = True
171      FrmctkOrder.cmdSettingPrinter.Visible = True
172      Screen.MousePointer = vbDefault
173      FrmctkOrder.WindowState = vbMaximized
174      FrmctkOrder.show vbModal

End Sub


Private Sub cekSubletFOP(kodeModel As String, pekerjaan As String)
          Dim rsSublet As ADODB.Recordset
1         sSQL = "select distinct c.KDSUBLET,c.NAMASUBLET,a.QTY,c.HARGAJUAL,c.kodesatuan"
2         sSQL = sSQL & " from [" & dbReplikasiHO & "]..HO_FOP_Program a"
3         sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..HO_SUBLETPACK b on a.item_part = b.KDSUBLET and b.FS_LOG_DEL = ''"
4         sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..MSTSUBLET c  on c.KDSUBLET = a.item_part and c.FS_LOG_DEL = ''"
5         sSQL = sSQL & " where a.FS_LOG_DEL = '' and a.MDL = '" & kodeModel & "' and a.FS_ITEM = '" & pekerjaan & "'"
6         Set rsSublet = objO.FetchData(sSQL)
7         If rsSublet.RecordCount > 0 And tagihanFOP <> "" Then
8             GoTo loopSubletTerkait
9         Else
10            sSQL = "select c.KDSUBLET,c.NAMASUBLET,b.QTY,c.HARGAJUAL,c.kodesatuan"
11            sSQL = sSQL & " from [" & dbReplikasiHO & "]..HO_SUBLETPACK b "
12            sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..MSTSUBLET c  on c.KDSUBLET = b.KDSUBLET and c.FS_LOG_DEL = ''"
13            sSQL = sSQL & " where b.FS_LOG_DEL = '' and b.MDL = '" & kodeModel & "' and b.FS_ITEM = '" & pekerjaan & "'"
14            Set rsSublet = objO.FetchData(sSQL)
15            If rsSublet.RecordCount > 0 Then
16                GoTo loopSubletTerkait
17            Else
18                Exit Sub
19            End If
20        End If
loopSubletTerkait:
          
21        rsSublet.MoveFirst
          Dim isBoleh     As Boolean
22        Do While Not rsSublet.EOF
23            If MaxBrs >= barisSublet Then
24                isBoleh = True
25                For i = 0 To barisSublet
26                    If NulSID(subletGrid.TextMatrix(i, 1)) = NulSID(rsSublet!namaSublet) Then
27                        isBoleh = False
28                        'GoTo lanjutBoleh
                          ''isBoleh = True
29                    End If
30                Next
lanjutBoleh:
                If isBoleh Then
31                With subletGrid
32                    .TextMatrix(barisSublet, 1) = NulSID(rsSublet!namaSublet)                   ' NAMA SUBLET
33                    .TextMatrix(barisSublet, 2) = CDbl(NulSID(rsSublet!qty))                    ' QTY
34                    .TextMatrix(barisSublet, 3) = NulSID(rsSublet!KODESATUAN)                                            ' SATUAN
35                    .TextMatrix(barisSublet, 4) = Format(CDbl(NulSID(rsSublet!hargaJual)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
36                    .TextMatrix(barisSublet, 5) = 1                ' STATUS PART
37                End With
38                barisSublet = barisSublet + 1
                End If
39                Text(52) = Format(CDbl(Text(52)) + (CDbl(NulSID(rsSublet!qty) * NulSID(rsSublet!hargaJual))), "#0,#") '* ((100 + gUpSublet) / 100)
40                Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
                  
41            End If
42            rsSublet.MoveNext
43        Loop
          
End Sub

Private Sub hapusSubletFOP(kodeModel As String, pekerjaan As String)
          Dim rsSublet As ADODB.Recordset
1         sSQL = "select distinct c.KDSUBLET,c.NAMASUBLET,a.QTY,c.HARGAJUAL,c.kodesatuan"
2         sSQL = sSQL & " from [" & dbReplikasiHO & "]..HO_FOP_Program a"
3         sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..HO_SUBLETPACK b on a.item_part = b.KDSUBLET and b.FS_LOG_DEL = ''"
4         sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..MSTSUBLET c  on c.KDSUBLET = a.item_part and c.FS_LOG_DEL = ''"
5         sSQL = sSQL & " where a.FS_LOG_DEL = '' and a.MDL = '" & kodeModel & "' and a.FS_ITEM = '" & pekerjaan & "'"
6         Set rsSublet = objO.FetchData(sSQL)
7         If rsSublet.RecordCount > 0 And tagihanFOP <> "" Then
8             GoTo loopSubletTerkait
9         Else
10            sSQL = "select c.KDSUBLET,c.NAMASUBLET,b.QTY,c.HARGAJUAL,c.kodesatuan"
11            sSQL = sSQL & " from [" & dbReplikasiHO & "]..HO_SUBLETPACK b "
12            sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..MSTSUBLET c  on c.KDSUBLET = b.KDSUBLET and c.FS_LOG_DEL = ''"
13            sSQL = sSQL & " where b.FS_LOG_DEL = '' and b.MDL = '" & kodeModel & "' and b.FS_ITEM = '" & pekerjaan & "'"
14            Set rsSublet = objO.FetchData(sSQL)
15            If rsSublet.RecordCount > 0 Then
16                GoTo loopSubletTerkait
17            Else
18                Exit Sub
19            End If
20        End If
loopSubletTerkait:
          
21        rsSublet.MoveFirst
          Dim isBoleh     As Boolean
          Dim bariske     As Integer
          Dim nilaiAsli   As Double
22        Do While Not rsSublet.EOF
23            If MaxBrs >= barisSublet Then
24                isBoleh = False
                  nilaiAsli = 0
25                For i = 0 To barisSublet
26                    If NulSID(subletGrid.TextMatrix(i, 1)) = NulSID(rsSublet!namaSublet) Then
27                        bariske = i
                          nilaiAsli = subletGrid.TextMatrix(i, 2) * subletGrid.TextMatrix(i, 4)
28                        isBoleh = True
29                        GoTo lanjutBoleh
                          ''isBoleh = True
30                    End If
31                Next
lanjutBoleh:
32              If isBoleh Then
33                For j = bariske To barisSublet
34                With subletGrid
35                    .TextMatrix(j, 1) = .TextMatrix(j + 1, 1)
36                    .TextMatrix(j, 2) = .TextMatrix(j + 1, 2)
37                    .TextMatrix(j, 3) = .TextMatrix(j + 1, 3)
38                    .TextMatrix(j, 4) = .TextMatrix(j + 1, 4)
39                    .TextMatrix(j, 5) = .TextMatrix(j + 1, 5)
40                End With
41                Next
42                barisSublet = barisSublet - 1
                  
43                Text(52) = Format(CDbl(Text(52)) - (CDbl(nilaiAsli)), "#0,#") '* ((100 + gUpSublet) / 100)
44                Text(53) = Format(CDbl(Text(39)) + CDbl(Text(40)) + CDbl(Text(52)) + CDbl(Text(55)), "#0,#")
                  
45              End If
                  
46            End If
47            rsSublet.MoveNext
48        Loop
          
End Sub


Private Sub panggilWarranty(noRangka As String, modL As String)
      Dim rsRec As ADODB.Recordset
      Dim terlibat As Boolean

      '' 20180606DEV01'NOTIFSCC
      Dim warranty As String
      Dim No As Integer
1        On Error GoTo panggilWarranty_error
 
          ''20180704 AGAR WARRANTY TIDAK DOUBLE2
          
2         If Log.RPLHistori = 0 Then
3             Set rsTemp = objO.FetchData("SELECT DISTINCT A.ID_WARRANTY,a.warranty FROM ho_warranty A WITH (NOLOCK) INNER JOIN ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
                           & "WHERE no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
                           & "ORDER BY A.id_warranty")
4         Else
5             Set rsTemp = objO.FetchData("SELECT DISTINCT A.ID_WARRANTY,a.warranty FROM [" & dbReplikasiHO & "]..ho_warranty A WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
                           & "WHERE no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
                           & "ORDER BY A.id_warranty")
6         End If
          
7         terlibat = False
8         No = 1
9         If Not rsTemp.EOF Then
10            Do While Not rsTemp.EOF
              
                  ''Set rsRec = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(noRangka) & "' AND KODETIPE = '" & NulSID(rsTemp("ID_WARRANTY")) & "' AND VALIDTWC <> 'NOK'")
                  ''20200316 || pergantian req mb thania
11                Set rsRec = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(noRangka) & "' AND KODETIPE = '" & NulSID(rsTemp("ID_WARRANTY")) & "' AND VALIDTWC <> ''")
12                If Not rsRec.EOF Then
13                    terlibat = False
14                    Exit Do
15                Else
16                    terlibat = True
17                End If
                  
18                Set rsRec = objO.FetchData("SELECT NO_RANGKA FROM TEMPWAR WITH (NOLOCK) WHERE NO_RANGKA = '" & noRangka & "' AND ID_WARRANTY = '" & rsTemp("ID_WARRANTY") & "'")
19                If Not rsRec.EOF Then
20                    terlibat = False
21                    Exit Do
22                Else
23                    terlibat = True
24                End If
                  
25                Set rsRec = objO.FetchData("SELECT no_chassis FROM History WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
26                If Not rsRec.EOF Then
27                    terlibat = False
28                    Exit Do
29                Else
30                    terlibat = True
31                End If
                  
32                If Log.RPLHistori = 1 Then
33                    Set rsRec = objO.FetchData("SELECT no_chassis FROM [" & dbReplikasiHO & "]..Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
34                    If Not rsRec.EOF Then
35                        terlibat = False
36                        Exit Do
37                    Else
38                        terlibat = True
39                    End If
40                Else
41                    If Log.LogHOConn = 1 Then
42                        If ConHO = True Then
                              
43                            Set rsRec = objO.FetchDataHO("SELECT no_chassis FROM Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
44                            If Not rsRec.EOF Then
45                                terlibat = False
46                                Exit Do
47                            Else
48                                terlibat = True
49                            End If
50                        End If
51                    End If
52                End If
                  
53                warranty = warranty & No & ". " & NulSID(rsTemp!warranty) & vbCrLf
                        
54                No = No + 1
55                rsTemp.MoveNext
56            Loop
              '' 20180606DEV01'NOTIFSCC
57            If terlibat Then
58                MsgBox "UNIT TERLIBAT WARRANTY !!!" & vbCrLf & warranty, vbCritical, "UNIT SSC"
59                ';cmdWarranty.BackColor = &HFF&
                  '
60                ';cmdWarranty.Enabled = True
61                'opTpSvc(2).Enabled = True
62            End If
63        Else
64            'cmdWarranty.BackColor = &HC0C0C0
65            'cmdWarranty.Enabled = False
66        End If

         'On Error GoTo 0
67       Exit Sub

panggilWarranty_error:
      '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
68    varErlError = Erl
69    varErrNumber = err.Number
      '' varErlError,varErrNumber
70        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure panggilWarranty of Form frmBooking")
End Sub

Private Sub TampilWaktuNHS(Tggl As Date)
Dim rsWaktu As ADODB.Recordset
Dim nStall As String
Dim i As Integer
Dim keterangan As String

   On Error GoTo TampilWaktu_Error
    
    GridWaktu.Clear
    GridWaktu.cols = 12
    GridWaktu.TextMatrix(0, 0) = "..."
    GridWaktu.TextMatrix(0, 1) = "STALL"
    GridWaktu.TextMatrix(0, 2) = "08:00"
    GridWaktu.TextMatrix(0, 3) = "09:00"
    GridWaktu.TextMatrix(0, 4) = "10:00"
    GridWaktu.TextMatrix(0, 5) = "11:00"
    GridWaktu.TextMatrix(0, 6) = "12:00"
    GridWaktu.TextMatrix(0, 7) = "13:00"
    GridWaktu.TextMatrix(0, 8) = "14:00"
    GridWaktu.TextMatrix(0, 9) = "15:00"
    GridWaktu.TextMatrix(0, 10) = "16:00"
    GridWaktu.TextMatrix(0, 11) = "17:00"
     
    sSQL = "select a.FS_KD_PETUGAS , a.KodeUnit,b.TIME,b.ESTTIME akhir,b.NOFAK,b.REP1 from PlottingBookingNHS a "
    sSQL = sSQL & " inner join BOOK b on a.NOFAK = b.NOFAK and b.FS_LOG_DEL = ''"
    sSQL = sSQL & " where b.date = '" & Format(Tggl, "yyyy-mm-dd") & "'"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select a.FS_KD_PETUGAS , a.KodeUnit,c.TIME,d.TIMESERAH akhir,c.NO_WO,c.REP1 from PlottingBookingNHS a"
    sSQL = sSQL & " inner join BOOK b on a.NOFAK = b.NOFAK and b.FS_LOG_DEL = ''"
    sSQL = sSQL & " inner join JOBDTA c on c.NOFAK = b.NOFAK and c.FS_LOG_DEL = '' and c.JOB = '3'"
    sSQL = sSQL & " inner join JOBDTL d on d.NO_WO = c.NO_WO"
    sSQL = sSQL & " where c.date = '" & Format(Tggl, "yyyy-mm-dd") & "'"
    ''Set rsWaktu = objO.FetchData("SELECT * FROM STALLWKF WITH (NOLOCK) ORDER BY LEFT(NOSTALL,1),CASE LEFT(NOSTALL,1) WHEN 'S' THEN CAST(SUBSTRING(NOSTALL, 7, LEN(NOSTALL)-6) AS INTEGER) ELSE CAST(SUBSTRING(NOSTALL, 5, LEN(NOSTALL)-4) AS INTEGER) END, [H-2], [H-1], H, [H+1]")
    Set rsWaktu = objO.FetchData(sSQL)
    
    If Not rsWaktu.EOF Then
        i = 1
        GridWaktu.Rows = rsWaktu.RecordCount + 1
        GridWaktu.ColWidth(0) = 200
        GridWaktu.ColWidth(1) = 1500
'        GridWaktu.ColWidth(2) = 1200
'        GridWaktu.ColWidth(3) = 1200
'        GridWaktu.ColWidth(4) = 1200
'        GridWaktu.ColWidth(5) = 1200

        
        Do While Not rsWaktu.EOF
            GridWaktu.TextMatrix(i, 0) = "*"
            GridWaktu.TextMatrix(i, 1) = NulSID(rsWaktu("FS_KD_PETUGAS")) & " " & NulSID(rsWaktu("KodeUnit"))
            keterangan = ""
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("08:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("09:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("08:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("09:00")) Or (cariJumlahMenit("08:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("08:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 2) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 2) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 2) = "#"
                End If
                
                
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("09:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("10:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("09:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("10:00")) Or (cariJumlahMenit("09:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("09:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 3) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 3) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 3) = "#"
                End If
                'GridWaktu.TextMatrix(i, 3) = "#"
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("10:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("11:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("10:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("11:00")) Or (cariJumlahMenit("10:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("10:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                'GridWaktu.TextMatrix(i, 4) = "#"
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 4) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 4) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 4) = "#"
                End If
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("11:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("12:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("11:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("12:00")) Or (cariJumlahMenit("11:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("11:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                'GridWaktu.TextMatrix(i, 5) = "#"
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 5) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 5) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 5) = "#"
                End If
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("12:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("13:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("12:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("13:00")) Or (cariJumlahMenit("12:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("12:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                'GridWaktu.TextMatrix(i, 6) = "#"
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 6) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 6) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 6) = "#"
                End If
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("13:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("14:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("13:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("14:00")) Or (cariJumlahMenit("13:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("13:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                ''GridWaktu.TextMatrix(i, 7) = "#"
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 7) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 7) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 7) = "#"
                End If
                
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("14:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("15:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("14:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("15:00")) Or (cariJumlahMenit("14:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("14:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                ''GridWaktu.TextMatrix(i, 8) = "#"
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 8) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 8) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 8) = "#"
                End If
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("15:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("16:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("15:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("16:00")) Or (cariJumlahMenit("15:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("15:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                ''GridWaktu.TextMatrix(i, 9) = "#"
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 9) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 9) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 9) = "#"
                End If
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("16:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("17:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("16:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("17:00")) Or (cariJumlahMenit("16:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("16:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                ''GridWaktu.TextMatrix(i, 10) = "#"
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 10) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 10) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 10) = "#"
                End If
            End If
            If (cariJumlahMenit(NulSID(rsWaktu("TIME"))) >= cariJumlahMenit("17:00") And cariJumlahMenit(NulSID(rsWaktu("TIME"))) < cariJumlahMenit("18:00")) Or (cariJumlahMenit(NulSID(rsWaktu("akhir"))) >= cariJumlahMenit("17:00") And cariJumlahMenit(NulSID(rsWaktu("akhir"))) < cariJumlahMenit("18:00")) Or (cariJumlahMenit("17:00") >= cariJumlahMenit(NulSID(rsWaktu("TIME"))) And cariJumlahMenit("17:00") <= cariJumlahMenit(NulSID(rsWaktu("akhir")))) Then
                ''GridWaktu.TextMatrix(i, 11) = "#"
                If NulSID(keterangan) = "" Then
                    GridWaktu.TextMatrix(i, 11) = NulSID(rsWaktu!NoFak)
                    keterangan = "1"
                ElseIf NulSID(keterangan) = "1" Then
                    GridWaktu.TextMatrix(i, 11) = NulSID(rsWaktu!rep1)
                    keterangan = "#"
                Else
                    GridWaktu.TextMatrix(i, 11) = "#"
                End If
            End If
            i = i + 1
            rsWaktu.MoveNext
        Loop
    End If
    'GridWaktu.Refresh
   On Error GoTo 0
   Exit Sub

TampilWaktu_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TampilWaktu of Form frmBooking")
End Sub
 

