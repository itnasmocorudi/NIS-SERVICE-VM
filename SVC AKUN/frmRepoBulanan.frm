VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmRepoBulanan 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   19080
   Tag             =   "  "
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
      Height          =   270
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8640
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   480
      TabIndex        =   45
      Top             =   8280
      Visible         =   0   'False
      Width           =   7575
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame FrRate 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   8535
      Left            =   480
      TabIndex        =   70
      Top             =   0
      Visible         =   0   'False
      Width           =   14535
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   7305
         Left            =   120
         TabIndex        =   78
         Top             =   360
         Visible         =   0   'False
         Width           =   14295
         Begin TrueOleDBGrid70.TDBGrid GridFR1 
            Height          =   7245
            Left            =   1920
            TabIndex        =   79
            Top             =   0
            Width           =   10410
            _ExtentX        =   18362
            _ExtentY        =   12779
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Tanggal"
            Columns(0).DataField=   "Tgl"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Jml WO"
            Columns(1).DataField=   "NoWO"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Full Sup"
            Columns(2).DataField=   "FullSup"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Order SA"
            Columns(3).DataField=   "QtyOrderSA"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "FR"
            Columns(4).DataField=   "FR"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Li Supply"
            Columns(5).DataField=   "Li_Supply"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Li BO"
            Columns(6).DataField=   "Li_BO"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Li Cancel"
            Columns(7).DataField=   "Li_Cancel"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "SR"
            Columns(8).DataField=   "SR"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   9
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
            Splits(0)._ColumnProps(0)=   "Columns.Count=9"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=1773"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=1667"
            Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(10)=   "Column(2).Width=1773"
            Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1667"
            Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(14)=   "Column(3).Width=1773"
            Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=1667"
            Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=1"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=1773"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=1667"
            Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=1"
            Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(24)=   "Column(5).Width=1773"
            Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=1667"
            Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=1"
            Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(29)=   "Column(6).Width=1773"
            Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=1667"
            Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=1"
            Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(34)=   "Column(7).Width=1773"
            Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=1667"
            Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=1"
            Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(39)=   "Column(8).Width=1773"
            Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=1667"
            Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=1"
            Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowUpdate     =   0   'False
            Appearance      =   2
            ColumnFooters   =   -1  'True
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=104,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=101,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=102,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=103,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.alignment=2"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=78,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=75,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=76,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=77,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=86,.parent=13,.alignment=2"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=83,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=84,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=85,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=90,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=87,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=88,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=89,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=98,.parent=13,.alignment=2"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=95,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=96,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=97,.parent=17"
            _StyleDefs(72)  =   "Named:id=33:Normal"
            _StyleDefs(73)  =   ":id=33,.parent=0"
            _StyleDefs(74)  =   "Named:id=34:Heading"
            _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(76)  =   ":id=34,.wraptext=-1"
            _StyleDefs(77)  =   "Named:id=35:Footing"
            _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(79)  =   "Named:id=36:Selected"
            _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=37:Caption"
            _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(83)  =   "Named:id=38:HighlightRow"
            _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(85)  =   "Named:id=39:EvenRow"
            _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(87)  =   "Named:id=40:OddRow"
            _StyleDefs(88)  =   ":id=40,.parent=33"
            _StyleDefs(89)  =   "Named:id=41:RecordSelector"
            _StyleDefs(90)  =   ":id=41,.parent=34"
            _StyleDefs(91)  =   "Named:id=42:FilterBar"
            _StyleDefs(92)  =   ":id=42,.parent=33"
            _StyleDefs(93)  =   "Named:id=99:Normal0"
            _StyleDefs(94)  =   ":id=99,.parent=42,.bgcolor=&HFFFFCE&,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(95)  =   ":id=99,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(96)  =   ":id=99,.fontname=MS Sans Serif"
            _StyleDefs(97)  =   "Named:id=100:Normal1"
            _StyleDefs(98)  =   ":id=100,.parent=99,.bgcolor=&H55FFFF&"
            _StyleDefs(99)  =   "Named:id=0:"
            _StyleDefs(100) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(101) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(102) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(103) =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(104) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(105) =   ":id=0,.fontname=MS Sans Serif"
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   615
         Left            =   120
         TabIndex        =   75
         Top             =   7680
         Width           =   4335
         Begin VB.CheckBox cexAll 
            Caption         =   "View Summary Harian"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   240
            Width           =   2655
         End
         Begin ucXPButton.XPButton cmdRef 
            Height          =   495
            Left            =   2880
            TabIndex        =   77
            Top             =   120
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
            MICON           =   "frmRepoBulanan.frx":0000
            PICN            =   "frmRepoBulanan.frx":001C
            PICH            =   "frmRepoBulanan.frx":0598
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000080FF&
            X1              =   0
            X2              =   4320
            Y1              =   0
            Y2              =   0
         End
      End
      Begin TrueOleDBGrid70.TDBGrid GridFR 
         Height          =   7245
         Left            =   120
         TabIndex        =   72
         Top             =   360
         Width           =   14250
         _ExtentX        =   25135
         _ExtentY        =   12779
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Tanggal"
         Columns(0).DataField=   "Tgl"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No WO"
         Columns(1).DataField=   "NoWO"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Job Desc"
         Columns(2).DataField=   "Job"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "No Polisi"
         Columns(3).DataField=   "NoPol"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "No Rangka"
         Columns(4).DataField=   "NoRangka"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Model"
         Columns(5).DataField=   "Model"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Kode Part"
         Columns(6).DataField=   "Kode"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Nama Part"
         Columns(7).DataField=   "Nama"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Order SA"
         Columns(8).DataField=   "QtyOrderSA"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Supply"
         Columns(9).DataField=   "QtySupply"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "FR_Li"
         Columns(10).DataField=   "FR_Li"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "FR"
         Columns(11).DataField=   "FR"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "Li Supply"
         Columns(12).DataField=   "Li_Supply"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "Li BO"
         Columns(13).DataField=   "Li_BO"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "Li Cancel"
         Columns(14).DataField=   "Li_Cancel"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "SR_Li"
         Columns(15).DataField=   "SR_Li"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "SR"
         Columns(16).DataField=   "SR"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   17
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=17"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2831"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2725"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2461"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2355"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=1773"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=1667"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=2302"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2196"
         Splits(0)._ColumnProps(17)=   "Column(3)._ColStyle=0"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(19)=   "Column(4).Width=4048"
         Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=3942"
         Splits(0)._ColumnProps(22)=   "Column(4)._ColStyle=0"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(24)=   "Column(5).Width=5292"
         Splits(0)._ColumnProps(25)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(26)=   "Column(5)._WidthInPix=5186"
         Splits(0)._ColumnProps(27)=   "Column(5)._ColStyle=0"
         Splits(0)._ColumnProps(28)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(29)=   "Column(6).Width=3175"
         Splits(0)._ColumnProps(30)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(6)._WidthInPix=3069"
         Splits(0)._ColumnProps(32)=   "Column(6)._ColStyle=0"
         Splits(0)._ColumnProps(33)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(34)=   "Column(7).Width=3519"
         Splits(0)._ColumnProps(35)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(7)._WidthInPix=3413"
         Splits(0)._ColumnProps(37)=   "Column(7)._ColStyle=0"
         Splits(0)._ColumnProps(38)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(39)=   "Column(8).Width=1773"
         Splits(0)._ColumnProps(40)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(41)=   "Column(8)._WidthInPix=1667"
         Splits(0)._ColumnProps(42)=   "Column(8)._ColStyle=1"
         Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(44)=   "Column(9).Width=1773"
         Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=1667"
         Splits(0)._ColumnProps(47)=   "Column(9)._ColStyle=1"
         Splits(0)._ColumnProps(48)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(49)=   "Column(10).Width=1773"
         Splits(0)._ColumnProps(50)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(10)._WidthInPix=1667"
         Splits(0)._ColumnProps(52)=   "Column(10)._ColStyle=1"
         Splits(0)._ColumnProps(53)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(54)=   "Column(11).Width=1773"
         Splits(0)._ColumnProps(55)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(11)._WidthInPix=1667"
         Splits(0)._ColumnProps(57)=   "Column(11)._ColStyle=1"
         Splits(0)._ColumnProps(58)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(59)=   "Column(12).Width=1773"
         Splits(0)._ColumnProps(60)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(61)=   "Column(12)._WidthInPix=1667"
         Splits(0)._ColumnProps(62)=   "Column(12)._ColStyle=1"
         Splits(0)._ColumnProps(63)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(64)=   "Column(13).Width=1773"
         Splits(0)._ColumnProps(65)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(13)._WidthInPix=1667"
         Splits(0)._ColumnProps(67)=   "Column(13)._ColStyle=1"
         Splits(0)._ColumnProps(68)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(69)=   "Column(14).Width=1773"
         Splits(0)._ColumnProps(70)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(14)._WidthInPix=1667"
         Splits(0)._ColumnProps(72)=   "Column(14)._ColStyle=1"
         Splits(0)._ColumnProps(73)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(74)=   "Column(15).Width=1773"
         Splits(0)._ColumnProps(75)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(76)=   "Column(15)._WidthInPix=1667"
         Splits(0)._ColumnProps(77)=   "Column(15)._ColStyle=1"
         Splits(0)._ColumnProps(78)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(79)=   "Column(16).Width=1773"
         Splits(0)._ColumnProps(80)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(81)=   "Column(16)._WidthInPix=1667"
         Splits(0)._ColumnProps(82)=   "Column(16)._ColStyle=1"
         Splits(0)._ColumnProps(83)=   "Column(16).Order=17"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   2
         ColumnFooters   =   -1  'True
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=62,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=59,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=60,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=61,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=104,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=101,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=102,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=103,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=32,.parent=13,.alignment=0"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=0"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13,.alignment=0"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.alignment=0"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=58,.parent=13,.alignment=0"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=55,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=56,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=57,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=66,.parent=13,.alignment=2"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=70,.parent=13,.alignment=2"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=67,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=68,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=69,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13,.alignment=2"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=78,.parent=13,.alignment=2"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=75,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=76,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=77,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=82,.parent=13,.alignment=2"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=79,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=80,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=81,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=86,.parent=13,.alignment=2"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=83,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=84,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=85,.parent=17"
         _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=90,.parent=13,.alignment=2"
         _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=87,.parent=14"
         _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=88,.parent=15"
         _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=89,.parent=17"
         _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=94,.parent=13,.alignment=2"
         _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=91,.parent=14"
         _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=92,.parent=15"
         _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=93,.parent=17"
         _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=98,.parent=13,.alignment=2"
         _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=95,.parent=14"
         _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=96,.parent=15"
         _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=97,.parent=17"
         _StyleDefs(104) =   "Named:id=33:Normal"
         _StyleDefs(105) =   ":id=33,.parent=0"
         _StyleDefs(106) =   "Named:id=34:Heading"
         _StyleDefs(107) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(108) =   ":id=34,.wraptext=-1"
         _StyleDefs(109) =   "Named:id=35:Footing"
         _StyleDefs(110) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(111) =   "Named:id=36:Selected"
         _StyleDefs(112) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(113) =   "Named:id=37:Caption"
         _StyleDefs(114) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(115) =   "Named:id=38:HighlightRow"
         _StyleDefs(116) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(117) =   "Named:id=39:EvenRow"
         _StyleDefs(118) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(119) =   "Named:id=40:OddRow"
         _StyleDefs(120) =   ":id=40,.parent=33"
         _StyleDefs(121) =   "Named:id=41:RecordSelector"
         _StyleDefs(122) =   ":id=41,.parent=34"
         _StyleDefs(123) =   "Named:id=42:FilterBar"
         _StyleDefs(124) =   ":id=42,.parent=33"
         _StyleDefs(125) =   "Named:id=99:Normal0"
         _StyleDefs(126) =   ":id=99,.parent=42,.bgcolor=&HFFFFCE&,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(127) =   ":id=99,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(128) =   ":id=99,.fontname=MS Sans Serif"
         _StyleDefs(129) =   "Named:id=100:Normal1"
         _StyleDefs(130) =   ":id=100,.parent=99,.bgcolor=&H55FFFF&"
         _StyleDefs(131) =   "Named:id=0:"
         _StyleDefs(132) =   ":id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(133) =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(134) =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(135) =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(136) =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(137) =   ":id=0,.fontname=MS Sans Serif"
      End
      Begin ucXPButton.XPButton cmdClose 
         Height          =   495
         Left            =   13200
         TabIndex        =   73
         Top             =   7800
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
         MICON           =   "frmRepoBulanan.frx":0BCF
         PICN            =   "frmRepoBulanan.frx":0BEB
         PICH            =   "frmRepoBulanan.frx":10BF
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
         Left            =   12000
         TabIndex        =   74
         Top             =   7800
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
         MICON           =   "frmRepoBulanan.frx":160A
         PICN            =   "frmRepoBulanan.frx":1626
         PICH            =   "frmRepoBulanan.frx":1BBB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000080FF&
         X1              =   14400
         X2              =   120
         Y1              =   7680
         Y2              =   7680
      End
      Begin VB.Label Ket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".::  SERVICE RATE  DAN FILL RATE ::."
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
         Index           =   9
         Left            =   5295
         TabIndex        =   71
         Top             =   60
         Width           =   3885
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   14295
      End
   End
   Begin VB.ComboBox cboBlnSR 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9960
      Style           =   2  'Dropdown List
      TabIndex        =   66
      Top             =   7640
      Width           =   855
   End
   Begin VB.ComboBox cboThnSR 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10800
      Style           =   2  'Dropdown List
      TabIndex        =   65
      Top             =   7640
      Width           =   1095
   End
   Begin VB.ComboBox cboThnP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10080
      Style           =   2  'Dropdown List
      TabIndex        =   57
      Top             =   5475
      Width           =   1095
   End
   Begin VB.ComboBox cboBlnP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9240
      Style           =   2  'Dropdown List
      TabIndex        =   56
      Top             =   5475
      Width           =   855
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoBulanan.frx":21F6
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   54
      Top             =   0
      Width           =   450
   End
   Begin VB.PictureBox PicPS 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7920
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   46
      Top             =   3960
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar BarPS 
         Height          =   495
         Left            =   0
         TabIndex        =   47
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nPS 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox PicNEA 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   7920
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   41
      Top             =   1800
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar BarNEA 
         Height          =   495
         Left            =   0
         TabIndex        =   42
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nNEA 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   43
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.PictureBox PicBars 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   22
      Top             =   3960
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar nBars 
         Height          =   495
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nLabs 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.ComboBox cboBulan1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1155
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   1155
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin VB.PictureBox PicBar 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   6975
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   840
         ScaleHeight     =   255
         ScaleWidth      =   5295
         TabIndex        =   5
         Top             =   480
         Width           =   5295
         Begin VB.Label nKet 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   5295
         End
      End
      Begin MSComctlLib.ProgressBar nBar 
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nLab 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   4935
      End
      Begin VB.Label nPro 
         Alignment       =   1  'Right Justify
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   0
         Width           =   1785
      End
   End
   Begin VB.ComboBox cboBulan2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1155
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   405
      Index           =   1
      Left            =   3720
      TabIndex        =   3
      Top             =   1155
      Visible         =   0   'False
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   714
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">9999"
      PromptChar      =   ">"
   End
   Begin ucXPButton.XPButton TblSimpan 
      Height          =   495
      Left            =   4680
      TabIndex        =   14
      Top             =   1080
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
      MICON           =   "frmRepoBulanan.frx":42FE
      PICN            =   "frmRepoBulanan.frx":431A
      PICH            =   "frmRepoBulanan.frx":4771
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdPrint 
      Height          =   495
      Left            =   3960
      TabIndex        =   16
      Top             =   3240
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
      MICON           =   "frmRepoBulanan.frx":4C5B
      PICN            =   "frmRepoBulanan.frx":4C77
      PICH            =   "frmRepoBulanan.frx":50CE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cboBln 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3315
      Width           =   855
   End
   Begin VB.ComboBox cboThn 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   3315
      Width           =   1095
   End
   Begin ucXPButton.XPButton cmdPrintS 
      Height          =   495
      Left            =   5520
      TabIndex        =   29
      Top             =   5400
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
      MICON           =   "frmRepoBulanan.frx":55B8
      PICN            =   "frmRepoBulanan.frx":55D4
      PICH            =   "frmRepoBulanan.frx":5A2B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox PicBarEM 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   600
      ScaleHeight     =   855
      ScaleWidth      =   6975
      TabIndex        =   26
      Top             =   6120
      Visible         =   0   'False
      Width           =   6975
      Begin MSComctlLib.ProgressBar nBarEM 
         Height          =   495
         Left            =   0
         TabIndex        =   27
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label nLabEM 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   4935
      End
   End
   Begin VB.ComboBox cboBlnS 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   5475
      Width           =   855
   End
   Begin VB.ComboBox cboThnS 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   5475
      Width           =   1095
   End
   Begin ucXPButton.XPButton TblKeluar 
      Height          =   495
      Left            =   13800
      TabIndex        =   13
      Top             =   8400
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
      MICON           =   "frmRepoBulanan.frx":5F15
      PICN            =   "frmRepoBulanan.frx":5F31
      PICH            =   "frmRepoBulanan.frx":6405
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   4080
      TabIndex        =   34
      Top             =   5520
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.OptionButton Opt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "EM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4680
      TabIndex        =   35
      Top             =   5520
      Width           =   615
   End
   Begin ucXPButton.XPButton cmdPrintNEA 
      Height          =   495
      Left            =   13560
      TabIndex        =   36
      Top             =   1080
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
      MICON           =   "frmRepoBulanan.frx":6950
      PICN            =   "frmRepoBulanan.frx":696C
      PICH            =   "frmRepoBulanan.frx":6DC3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBDate6Ctl.TDBDate dt0 
      Height          =   315
      Left            =   9480
      TabIndex        =   38
      Top             =   1155
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      Calendar        =   "frmRepoBulanan.frx":72AD
      Caption         =   "frmRepoBulanan.frx":73D9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoBulanan.frx":7435
      Keys            =   "frmRepoBulanan.frx":7453
      Spin            =   "frmRepoBulanan.frx":74B1
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      ForeColor       =   0
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   2
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
      ValueVT         =   6815745
      Value           =   39183
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate dt1 
      Height          =   315
      Left            =   11520
      TabIndex        =   39
      Top             =   1155
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      Calendar        =   "frmRepoBulanan.frx":74D9
      Caption         =   "frmRepoBulanan.frx":7605
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoBulanan.frx":7661
      Keys            =   "frmRepoBulanan.frx":767F
      Spin            =   "frmRepoBulanan.frx":76DD
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      ForeColor       =   0
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   2
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
      ValueVT         =   6815745
      Value           =   39183
      CenturyMode     =   0
   End
   Begin ucXPButton.XPButton cmdPrintPS 
      Height          =   495
      Left            =   13560
      TabIndex        =   49
      Top             =   3240
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
      MICON           =   "frmRepoBulanan.frx":7705
      PICN            =   "frmRepoBulanan.frx":7721
      PICH            =   "frmRepoBulanan.frx":7B78
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TDBDate6Ctl.TDBDate dtP0 
      Height          =   315
      Left            =   9480
      TabIndex        =   50
      Top             =   3315
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      Calendar        =   "frmRepoBulanan.frx":8062
      Caption         =   "frmRepoBulanan.frx":818E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoBulanan.frx":81EA
      Keys            =   "frmRepoBulanan.frx":8208
      Spin            =   "frmRepoBulanan.frx":8266
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      ForeColor       =   0
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   2
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
      ValueVT         =   6815745
      Value           =   39183
      CenturyMode     =   0
   End
   Begin TDBDate6Ctl.TDBDate dtP1 
      Height          =   315
      Left            =   11520
      TabIndex        =   51
      Top             =   3315
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   556
      Calendar        =   "frmRepoBulanan.frx":828E
      Caption         =   "frmRepoBulanan.frx":83BA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmRepoBulanan.frx":8416
      Keys            =   "frmRepoBulanan.frx":8434
      Spin            =   "frmRepoBulanan.frx":8492
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
      BackColor       =   16777215
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
      ForeColor       =   0
      Format          =   "dd/mm/yyyy"
      HighlightText   =   0
      IMEMode         =   3
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxDate         =   2958465
      MinDate         =   2
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
      ValueVT         =   6815745
      Value           =   39183
      CenturyMode     =   0
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   120
      Width           =   12015
      _ExtentX        =   21193
      _ExtentY        =   661
      BackColor       =   12577002
   End
   Begin ucXPButton.XPButton cmdPrintP 
      Height          =   495
      Left            =   11280
      TabIndex        =   55
      Top             =   5400
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
      MICON           =   "frmRepoBulanan.frx":84BA
      PICN            =   "frmRepoBulanan.frx":84D6
      PICH            =   "frmRepoBulanan.frx":892D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdView 
      Height          =   495
      Left            =   3600
      TabIndex        =   62
      Top             =   7560
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&v"
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
      MICON           =   "frmRepoBulanan.frx":8E17
      PICN            =   "frmRepoBulanan.frx":8E33
      PICH            =   "frmRepoBulanan.frx":93A1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdPrintSR 
      Height          =   495
      Left            =   12120
      TabIndex        =   69
      Top             =   7560
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
      MICON           =   "frmRepoBulanan.frx":9995
      PICN            =   "frmRepoBulanan.frx":99B1
      PICH            =   "frmRepoBulanan.frx":9E08
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label12 
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
      Left            =   9840
      TabIndex        =   68
      Top             =   7695
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
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
      Left            =   8880
      TabIndex        =   67
      Top             =   7695
      Width           =   840
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   7800
      Top             =   7440
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  SERVICE RATE DAN FILL RATE  ::."
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
      Index           =   7
      Left            =   9600
      TabIndex        =   64
      Top             =   7140
      Width           =   3885
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   7
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  VIEW  PROSES  WARANTY  ::."
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
      Index           =   6
      Left            =   2400
      TabIndex        =   63
      Top             =   7140
      Width           =   3435
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      Top             =   7440
      Width           =   7215
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   7800
      Top             =   5280
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  LAPORAN  BULANAN  PART  ::."
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
      Index           =   5
      Left            =   9720
      TabIndex        =   61
      Top             =   4980
      Width           =   3525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
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
      Left            =   8160
      TabIndex        =   60
      Top             =   5520
      Width           =   840
   End
   Begin VB.Label Label12 
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
      Left            =   9120
      TabIndex        =   59
      Top             =   5520
      Width           =   60
   End
   Begin VB.Label nWaitPart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ... (^_^)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   12720
      TabIndex        =   58
      Top             =   5520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000080FF&
      X1              =   360
      X2              =   15000
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8415
      Left            =   0
      Picture         =   "frmRepoBulanan.frx":A2F2
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   375
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: LAPORAN  PROFIT  SHARING  ::."
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
      Index           =   4
      Left            =   9660
      TabIndex        =   53
      Top             =   2820
      Width           =   3645
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   7800
      Top             =   3120
      Width           =   7215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Range Invoice :  "
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
      Left            =   7920
      TabIndex        =   52
      Top             =   3360
      Width           =   1635
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   11280
      X2              =   11400
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   11280
      X2              =   11400
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Range Invoice :  "
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
      Left            =   7920
      TabIndex        =   40
      Top             =   1200
      Width           =   1635
   End
   Begin VB.Shape Shape6 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   7800
      Top             =   960
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: LAPORAN  N E A  ::."
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
      Index           =   3
      Left            =   10320
      TabIndex        =   37
      Top             =   660
      Width           =   2325
   End
   Begin VB.Label nWait 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait ... (^_^)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5400
      TabIndex        =   21
      Top             =   3360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label12 
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
      Index           =   0
      Left            =   1800
      TabIndex        =   19
      Top             =   3360
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
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
      Left            =   840
      TabIndex        =   20
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  LAPORAN  BULANAN  DYNA  ::."
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
      Index           =   0
      Left            =   2370
      TabIndex        =   15
      Top             =   2820
      Width           =   3585
   End
   Begin VB.Label Label12 
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
      TabIndex        =   12
      Top             =   1200
      Width           =   60
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
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
      Left            =   840
      TabIndex        =   11
      Top             =   1200
      Width           =   840
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  K P I   LAPORAN  BULANAN  ::."
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
      Left            =   2400
      TabIndex        =   10
      Top             =   660
      Width           =   3525
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      Top             =   960
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   7215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      Top             =   3120
      Width           =   7215
   End
   Begin VB.Label Ket 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".::  LAPORAN  BULANAN  SB - NP / EM  ::."
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
      Index           =   2
      Left            =   1995
      TabIndex        =   25
      Top             =   4980
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   7215
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      FillColor       =   &H00FF0000&
      Height          =   735
      Left            =   480
      Top             =   5280
      Width           =   7215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PERIODE"
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
      Left            =   840
      TabIndex        =   33
      Top             =   5520
      Width           =   840
   End
   Begin VB.Label Label12 
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
      Index           =   1
      Left            =   1800
      TabIndex        =   32
      Top             =   5520
      Width           =   60
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H000080FF&
      Height          =   375
      Left            =   3960
      Top             =   5460
      Width           =   1455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   7800
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   480
      Shape           =   4  'Rounded Rectangle
      Top             =   7080
      Width           =   7215
   End
End
Attribute VB_Name = "frmRepoBulanan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADodb.Recordset
Dim vNamaTabel As String
Dim objBilling As New BClsBilling
Private obj As New ClsKPI
Private objR As New clsRepo
Private rsNEA As New ADodb.Recordset
Private rsPS As New ADodb.Recordset

Private noinv  As String
Private nUnit  As Double
Private nUnitR As Double
Private nINV   As Double
Private nInvR  As Double

Private rsSR  As New ADodb.Recordset
Private rsSR1 As New ADodb.Recordset
Private rsFR  As New ADodb.Recordset

Private Sub hapusForm()
    For i = 0 To 1
        Text(i) = 0
    Next
    
    cboBulan1.Clear
    cboBulan2.Clear
    For i = 1 To 12
        cboBulan1.AddItem Format(i, "00") & " | " & GetNamaBulan(i)
        cboBulan2.AddItem Format(i, "00") & " | " & GetNamaBulan(i)
    Next
    cboBulan1 = Format(Month(isnow), "00") & " | " & GetNamaBulan(Month(isnow))
    cboBulan2 = Format(Month(isnow), "00") & " | " & GetNamaBulan(Month(isnow))
    Text(0) = Year(isnow)
    Text(1) = Year(isnow)

    cboThn.Clear
    cboThnS.Clear
    cboThnP.Clear
    cboThnP.AddItem "2012"
    For i = 2013 To Year(isnow)
        cboThn.AddItem i
        cboThnS.AddItem i
        cboThnP.AddItem i
        If i <> 2013 Then
            cboThnSR.AddItem i
        End If
    Next i
    cboThn.Text = Year(isnow)
    cboThnS.Text = Year(isnow)
    cboThnP.Text = Year(isnow)
    cboThnSR.Text = Year(isnow)
    cboBln.Clear
    cboBlnS.Clear
    For i = 1 To 12
        cboBln.AddItem i
        cboBlnS.AddItem i
        cboBlnP.AddItem i
        cboBlnSR.AddItem i
    Next i
    cboBln.Text = Month(isnow)
    cboBlnS.Text = Month(isnow)
    cboBlnP.Text = Month(isnow)
    cboBlnSR.Text = Month(isnow)
    

    dt0.Text = Replace(TglDDMMYYYY("01/" + Format(Month(isnow), "00") + "/" + Format(Year(isnow), "0000")), "-", "/")
    dt1.Text = Replace(GetTglAkhir(Month(isnow), Year(isnow), 0), "-", "/")

    dtP0.Text = Replace(TglDDMMYYYY("01/" + Format(Month(isnow), "00") + "/" + Format(Year(isnow), "0000")), "-", "/")
    dtP1.Text = Replace(GetTglAkhir(Month(isnow), Year(isnow), 0), "-", "/")

End Sub

Private Function cekInsert(nID As Integer, nKet As String) As Boolean
Dim rsHlp As ADodb.Recordset

    cekInsert = False
    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE urut = '" & nID & "' AND KETERANGAN = '" & nKet & "'")
    If rsHlp.EOF Then
        cekInsert = True
    End If
    
End Function

Private Function kpi6_b(Bln As Integer, Thn As Integer) As Integer
    Dim rs2 As ADodb.Recordset
    Dim P2 As Integer
    
    On Error Resume Next
    kpi6_b = 0
    Set rs2 = obj.KPI_6(False, Bln, Thn)
    If rs2.RecordCount <> 0 Then P2 = rs2.Fields(0).value
    Set rs2 = obj.KPI_6b(True, Bln, Thn)
    If rs2.RecordCount <> 0 Then P2 = P2 + CInt(rs2.Fields(0).value)
    
    kpi6_b = CDbl(P2)
End Function

Private Function kpi6_a(Bln As Integer, Thn As Integer) As Integer
    Dim rs2 As ADodb.Recordset
    Dim P1 As Integer
    
    On Error Resume Next
    kpi6_a = 0
    Set rs1 = obj.KPI_6(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then P1 = rs1.Fields(0).value
    Set rs1 = obj.KPI_6a(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then P1 = P1 + CInt(rs1.Fields(0).value)
    
    kpi6_a = CDbl(P1)
End Function

Private Function kpi9_a(Bln As Integer, Thn As Integer) As Integer
    Dim rs2 As ADodb.Recordset
    Dim Pda As Integer
    
    On Error Resume Next
    kpi9_a = 0
    Set rs1 = obj.KPI_9(True, Bln, Thn)
    If rs1.RecordCount <> 0 Then Pda = rs1.Fields(0).value
    Set rs1 = obj.KPI_9(False, Bln, Thn)
    If rs1.RecordCount <> 0 Then Pda = Pda + rs1.Fields(0).value
    
    kpi9_a = CDbl(Pda)
End Function

Private Function insertTabel(nBulan As String, nBaris As Integer, nKet As String, nJml As Double, nkode As Integer, Optional diBold As Boolean = False) As String
    If nkode = 1 Then
        insertTabel = "INSERT INTO [TEMPDB]..[" & vNamaTabel & "] (URUT, KETERANGAN, BOLD, " & nBulan & ")  VALUES (" & nBaris & ",'" & nKet & "','" & IIf(diBold, 1, 0) & "'," & Replace(CDbl(nJml), ",", ".") & ")"
    ElseIf nkode = 2 Then
        insertTabel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & nBulan & " = " & Replace(CDbl(nJml), ",", ".") & " WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    ElseIf nkode = 3 Then
        insertTabel = "UPDATE [TEMPDB]..[" & vNamaTabel & "] SET " & nBulan & " = " & nBulan & " + " & Replace(CDbl(nJml), ",", ".") & " WHERE URUT = " & nBaris & " AND KETERANGAN = '" & nKet & "'"
    End If
    
End Function

Private Sub cboBulan1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cboBulan2_KeyDown(KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub cmdPrint_Click()
    nWait.Visible = True

        '2 April 2013
        nBars.Min = 0
        nBars.Max = 20
        nBars.value = 0
        nLabs.Caption = "0 Of 20  Get Data Report"
        PicBars.Visible = True
        DoEvents


    DoEvents
    
    Dim nThn As Double
    Dim nBln As Double
    nThn = Val(cboThn.Text)
    nBln = Val(cboBln.Text)
    
    Dim nIC As Double
    For nIC = 1 To 12
        Call objR.FetchDYNA_Bulanan(nThn, nIC)
        nLabs.Caption = nIC & " Of 20  Cex Existing Data"
        nBars.value = nIC
        DoEvents
    Next nIC
    
    Dim jSPK As Double
        jSPK = objR.FetchSPK_DYNA(nThn, nBln)
        nLabs.Caption = "13 Of 20  Penjualan (Retail SPK)"
        nBars.value = 13
        DoEvents
    Dim jSB1 As Double
        jSB1 = objR.FetchSB_DYNA(nThn, nBln, "R001")
        nLabs.Caption = "14 Of 20  Service 1.000 Km"
        nBars.value = 14
        DoEvents
    Dim jSB5 As Double
        jSB5 = objR.FetchSB_DYNA(nThn, nBln, "R005")
        nLabs.Caption = "15 Of 20  Service 5.000 Km"
        nBars.value = 15
        DoEvents
    Dim jSBE As Double
        jSBE = objR.FetchSBE_DYNA(nThn, nBln)
        nLabs.Caption = "16 Of 20  Service Berkala External"
        nBars.value = 16
        DoEvents
    Dim jGRP As Double
        jGRP = objR.FetchSB_DYNA(nThn, nBln, "")
        nLabs.Caption = "17 Of 20  General Repair"
        nBars.value = 17
        DoEvents
    Dim jTHS As Double
        jTHS = objR.FetchTHS_DYNA(nThn, nBln)
        nLabs.Caption = "18 Of 20  Service - Dealer n Visit"
        nBars.value = 18
        DoEvents
    Dim jDLR As Double
        jDLR = (jSB1 + jSB5 + jSBE + jGRP) - jTHS
        Call objR.UpdateItemDYNA(nThn, nBln, jSPK, jSB1, jSB5, jSBE, jGRP, jDLR, jTHS)
        nLabs.Caption = "19 Of 20  Service - Dealer n Visit"
        nBars.value = 19
        DoEvents
    Call Print_DYNA
End Sub

Private Sub Print_DYNA()
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptDYNA_HO

    sSQL = "Select A.*,B.*,C.*,D.*,E.*,F.*,G.*,H.*,I.*,J.*,K.*,L.* from " & _
           "(select Tahun as Tahun1,SPK as SPK1,SB1 as SB11,SB5 as SB51,SBE as SBE1,GRP as GRP1,SVCD as SVCD1,SVCV as SVCV1 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 1) A Left Join " & _
           "(select Tahun as Tahun2,SPK as SPK2,SB1 as SB12,SB5 as SB52,SBE as SBE2,GRP as GRP2,SVCD as SVCD2,SVCV as SVCV2 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 2) B On A.Tahun1  = B.Tahun2 Left Join " & _
           "(select Tahun as Tahun3,SPK as SPK3,SB1 as SB13,SB5 as SB53,SBE as SBE3,GRP as GRP3,SVCD as SVCD3,SVCV as SVCV3 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 3) C On A.Tahun1  = C.Tahun3 Left Join " & _
           "(select Tahun as Tahun4,SPK as SPK4,SB1 as SB14,SB5 as SB54,SBE as SBE4,GRP as GRP4,SVCD as SVCD4,SVCV as SVCV4 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 4) D On A.Tahun1  = D.Tahun4 Left Join " & _
           "(select Tahun as Tahun5,SPK as SPK5,SB1 as SB15,SB5 as SB55,SBE as SBE5,GRP as GRP5,SVCD as SVCD5,SVCV as SVCV5 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 5) E On A.Tahun1  = E.Tahun5 Left Join " & _
           "(select Tahun as Tahun6,SPK as SPK6,SB1 as SB16,SB5 as SB56,SBE as SBE6,GRP as GRP6,SVCD as SVCD6,SVCV as SVCV6 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 6) F On A.Tahun1  = F.Tahun6 Left Join " & _
           "(select Tahun as Tahun7,SPK as SPK7,SB1 as SB17,SB5 as SB57,SBE as SBE7,GRP as GRP7,SVCD as SVCD7,SVCV as SVCV7 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 7) G On A.Tahun1  = G.Tahun7 Left Join " & _
           "(select Tahun as Tahun8,SPK as SPK8,SB1 as SB18,SB5 as SB58,SBE as SBE8,GRP as GRP8,SVCD as SVCD8,SVCV as SVCV8 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 8) H On A.Tahun1  = H.Tahun8 Left Join " & _
           "(select Tahun as Tahun9,SPK as SPK9,SB1 as SB19,SB5 as SB59,SBE as SBE9,GRP as GRP9,SVCD as SVCD9,SVCV as SVCV9 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 9) I On A.Tahun1  = I.Tahun9 Left Join " & _
           "(select Tahun as Tahun10,SPK as SPK10,SB1 as SB110,SB5 as SB510,SBE as SBE10,GRP as GRP10,SVCD as SVCD10,SVCV as SVCV10 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 10) J On A.Tahun1  = J.Tahun10 Left Join " & _
           "(select Tahun as Tahun11,SPK as SPK11,SB1 as SB111,SB5 as SB511,SBE as SBE11,GRP as GRP11,SVCD as SVCD11,SVCV as SVCV11 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 11) K On A.Tahun1  = K.Tahun11 Left Join " & _
           "(select Tahun as Tahun12,SPK as SPK12,SB1 as SB112,SB5 as SB512,SBE as SBE12,GRP as GRP12,SVCD as SVCD12,SVCV as SVCV12 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 12) L On A.Tahun1  = L.Tahun12 Where A.Tahun1 = " & cboThn.Text

            
    Set oRs = objO.FetchData(sSQL)
    Report.Database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamatvk
    
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    PicBars.Visible = False
    nWait.Visible = False
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Print_Part()
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptDYNA_HO

    sSQL = "Select A.*,B.*,C.*,D.*,E.*,F.*,G.*,H.*,I.*,J.*,K.*,L.* from " & _
           "(select Tahun as Tahun1,SPK as SPK1,SB1 as SB11,SB5 as SB51,SBE as SBE1,GRP as GRP1,SVCD as SVCD1,SVCV as SVCV1 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 1) A Left Join " & _
           "(select Tahun as Tahun2,SPK as SPK2,SB1 as SB12,SB5 as SB52,SBE as SBE2,GRP as GRP2,SVCD as SVCD2,SVCV as SVCV2 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 2) B On A.Tahun1  = B.Tahun2 Left Join " & _
           "(select Tahun as Tahun3,SPK as SPK3,SB1 as SB13,SB5 as SB53,SBE as SBE3,GRP as GRP3,SVCD as SVCD3,SVCV as SVCV3 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 3) C On A.Tahun1  = C.Tahun3 Left Join " & _
           "(select Tahun as Tahun4,SPK as SPK4,SB1 as SB14,SB5 as SB54,SBE as SBE4,GRP as GRP4,SVCD as SVCD4,SVCV as SVCV4 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 4) D On A.Tahun1  = D.Tahun4 Left Join " & _
           "(select Tahun as Tahun5,SPK as SPK5,SB1 as SB15,SB5 as SB55,SBE as SBE5,GRP as GRP5,SVCD as SVCD5,SVCV as SVCV5 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 5) E On A.Tahun1  = E.Tahun5 Left Join " & _
           "(select Tahun as Tahun6,SPK as SPK6,SB1 as SB16,SB5 as SB56,SBE as SBE6,GRP as GRP6,SVCD as SVCD6,SVCV as SVCV6 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 6) F On A.Tahun1  = F.Tahun6 Left Join " & _
           "(select Tahun as Tahun7,SPK as SPK7,SB1 as SB17,SB5 as SB57,SBE as SBE7,GRP as GRP7,SVCD as SVCD7,SVCV as SVCV7 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 7) G On A.Tahun1  = G.Tahun7 Left Join " & _
           "(select Tahun as Tahun8,SPK as SPK8,SB1 as SB18,SB5 as SB58,SBE as SBE8,GRP as GRP8,SVCD as SVCD8,SVCV as SVCV8 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 8) H On A.Tahun1  = H.Tahun8 Left Join " & _
           "(select Tahun as Tahun9,SPK as SPK9,SB1 as SB19,SB5 as SB59,SBE as SBE9,GRP as GRP9,SVCD as SVCD9,SVCV as SVCV9 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 9) I On A.Tahun1  = I.Tahun9 Left Join " & _
           "(select Tahun as Tahun10,SPK as SPK10,SB1 as SB110,SB5 as SB510,SBE as SBE10,GRP as GRP10,SVCD as SVCD10,SVCV as SVCV10 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 10) J On A.Tahun1  = J.Tahun10 Left Join " & _
           "(select Tahun as Tahun11,SPK as SPK11,SB1 as SB111,SB5 as SB511,SBE as SBE11,GRP as GRP11,SVCD as SVCD11,SVCV as SVCV11 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 11) K On A.Tahun1  = K.Tahun11 Left Join " & _
           "(select Tahun as Tahun12,SPK as SPK12,SB1 as SB112,SB5 as SB512,SBE as SBE12,GRP as GRP12,SVCD as SVCD12,SVCV as SVCV12 from DYNA_HO WHERE Tahun = " & cboThn.Text & " and Bulan = 12) L On A.Tahun1  = L.Tahun12 Where A.Tahun1 = " & cboThn.Text

            
    Set oRs = objO.FetchData(sSQL)
    Report.Database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamatvk
    
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    PicBars.Visible = False
    nWait.Visible = False
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdTutup_Click()
    On Error GoTo errNulDB
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(sSQL)
errNulDB:
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub cmdView_Click()
    Dim f As New FrmTwc_ViewCabang
    f.Show
End Sub

Private Sub Form_Activate()
    vNamaTabel = "RepoBulanan" & DriveSerialNumber("c:") & Second(Now)
End Sub

Private Sub Form_Load()
    Me.WindowState = vbMaximized
    Call hapusForm
End Sub


Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub BarProses(labA As Integer, labB As Integer, proA As Integer, proB As Integer, bar As Integer, Ket As String)
    nBar.value = bar
    nLab.Caption = "Collecting Data Report : " & labA & " Of " & labB & ""
    nPro.Caption = "Process :  " & proA & " Of " & proB & ""
    nKet.Caption = Ket
    DoEvents
End Sub

Private Sub TblSimpan_Click()
cboBulan2.Text = cboBulan1.Text
Text(1).Text = Text(0).Text
Screen.MousePointer = vbHourglass
Dim pesan As String
Dim temp As String
Dim boleh As Boolean
Dim nI As Integer
Dim nTotal As Double
Dim nTotal2 As Double
Dim nSubTotal As Double
Dim namaBulan As String
Dim nUrut As Integer
Dim nKeterangan As String
''========================''

Dim nilJasaPPn As Double
Dim nilJasaNonPPn As Double
Dim nilJasaTotal As Double

''========================''
Dim nHrsFollow As Integer
Dim nDiFollow As Integer
Dim nNoQ1 As Integer
Dim nNoQ2 As Integer
Dim nNoQ3 As Integer
Dim nRTJ As Integer

''========================''
Dim nHrsUndang  As Integer
Dim nDiUndang   As Integer
Dim nDtgUndang  As Integer

''========================''
Dim GTG         As Double
Dim UE          As Double
Dim UEnT        As Double
Dim PN          As Double
Dim GTGs        As Double
Dim UEs         As Double
Dim UEsTHS      As Double
Dim PNs         As Double
Dim TDN         As Double
Dim SBESBI      As Double
Dim PTh         As Double
Dim TTh         As Double
Dim jTHS        As Double
Dim TotalSBE    As Double
Dim totalUnitEntryTHS As Double
Dim totalUnitEntry As Double
Dim TotalUnitSBE As Double
Dim Jumlah1K As Double
Dim i As Integer
    boleh = True
    If Text(0) = "" Then
        pesan = pesan + "Tahun Awal. "
        boleh = False
    End If
    If Text(1) = "" Then
        pesan = pesan + "Tahun Akhir. "
        boleh = False
    End If
    If CDbl(Text(0)) <> CDbl(Text(1)) Then
        pesan = pesan + "Tahun harus sama. "
        boleh = False
    End If
    If cboBulan1 = "" Then
        pesan = pesan + "Bulan Awal. "
        boleh = False
    End If
    If cboBulan2 = "" Then
        pesan = pesan + "Bulan Akhir. "
        boleh = False
    End If
    
    If boleh Then
        Shape1.Height = 1575
        sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
             & " drop table [tempdb]..[" & vNamaTabel & "]"
        Call objO.UpdateData(sSQL)
        
        temp = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] ([URUT][NUMERIC](9) NOT NULL DEFAULT 0, [KETERANGAN][varchar](100) NOT NULL DEFAULT '', [BOLD][char](1) NOT NULL DEFAULT '0' "
        For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
            temp = temp + ", [" & GetNamaBulan(i) & "][money] NOT NULL DEFAULT 0"
        Next
        temp = temp + ") ON [PRIMARY]"
        Call objO.UpdateData(temp)
        
        '6 Agustus 2012
        nBar.Min = 0
        nBar.Max = 16
        nBar.value = 0
        nLab.Caption = "Get Data Report : 0 Of 16"
        nPro.Caption = "Process :  1 Of 2"
        PicBar.Visible = True
        DoEvents
        
        'For i = Left(cboBulan1, InStr(cboBulan1, " |") - 1) To Left(cboBulan2, InStr(cboBulan2, " |") - 1)
            i = Left(cboBulan1, InStr(cboBulan1, " |") - 1)
            'Set rsTemp = objO.FetchData("EXEC REPOBULANAN3 " & Left(cboBulan1, InStr(cboBulan1, " |") - 1) & ", " & Text(0) & "")
            
            Call objO.DeleteData("Delete FROM BatchBulanan Where Bulan = '" & i & "' and Tahun = '" & Text(0).Text & "'")
            
           '6 Agustus 2012
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3A " & i & "," & Text(0) & "")
            Call BarProses(1, 16, 1, 2, 1, "Get Data Unit Entry ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3B " & i & "," & Text(0) & "")
            Call BarProses(2, 16, 1, 2, 2, "Get Data Unit Entry ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3C " & i & "," & Text(0) & "")
            Call BarProses(3, 16, 1, 2, 3, "Get Data Unit Entry ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3D " & i & "," & Text(0) & "")
            Call BarProses(4, 16, 1, 2, 4, "Get Data Revenue ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3E " & i & "," & Text(0) & "")
            Call BarProses(5, 16, 1, 2, 5, "Get Data Revenue ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3F " & i & "," & Text(0) & "")
            Call BarProses(6, 16, 1, 2, 6, "Get Data Service Time ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3G " & i & "," & Text(0) & "")
            Call BarProses(7, 16, 1, 2, 7, "Get Data Follow Up ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3H " & i & "," & Text(0) & "")
            Call BarProses(8, 16, 1, 2, 8, "Get Data Lead Time ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3I " & i & "," & Text(0) & "")
            Call BarProses(9, 16, 1, 2, 9, "Get Data Lead Time Notifcation ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3J " & i & "," & Text(0) & "")
            Call BarProses(10, 16, 1, 2, 10, "Get Data Lead Time Total ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3K " & i & "," & Text(0) & "")
            Call BarProses(11, 16, 1, 2, 11, "Get Data Lead Time Total ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3L " & i & "," & Text(0) & "")
            Call BarProses(12, 16, 1, 2, 12, "Get Data Lead Time Total ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3M " & i & "," & Text(0) & "")
            Call BarProses(13, 16, 1, 2, 13, "Get Data Lead Time Total ...")
            DoEvents
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3N " & i & "," & Text(0) & "")
            Call objO.InsertData("Insert Into BatchBulanan exec repobulanan3O " & i & "," & Text(0) & "")
            Call BarProses(14, 16, 1, 2, 14, "Get Data Lead Time Total ...")
            DoEvents

            Set rsTemp = objO.FetchData("Select * from BatchBulanan  Where Bulan = '" & i & "' and Tahun = '" & Text(0).Text & "'")
            Call BarProses(15, 16, 1, 2, 15, "Get Data Report ...")
            DoEvents
            
            If Not rsTemp.EOF Then
                nTotal = 0
                nTotal2 = 0
                nSubTotal = 0
                nI = 1
                namaBulan = GetNamaBulan(Left(cboBulan1, InStr(cboBulan1, " |") - 1))

                nBar.Min = 0
                nBar.Max = 105
                nBar.value = 0
                nLab.Caption = "Calculation Data : 0 Of 105"
                nPro.Caption = "Process :  2 Of 3"
                DoEvents
                
                Do While nI <= 102
                    Select Case nI
                        Case 1
                            nUrut = nI
                            nKeterangan = "10K saja "
                            rsTemp.Find "KETERANGAN = '10'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 10 K ...")
                        Case 2
                            nUrut = nI
                            nKeterangan = "20K saja "
                            rsTemp.Find "KETERANGAN = '20'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 20 K ...")
                        Case 3
                            nUrut = nI
                            nKeterangan = "30K saja "
                            rsTemp.Find "KETERANGAN = '30'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 30 K ...")
                        Case 4
                            nUrut = nI
                            nKeterangan = "40K saja "
                            rsTemp.Find "KETERANGAN = '40'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 40 K ...")
                        Case 5
                            nUrut = nI
                            nKeterangan = "50K saja "
                            rsTemp.Find "KETERANGAN = '50'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 50 K ...")
                        Case 6
                            nUrut = nI
                            nKeterangan = "60K saja "
                            rsTemp.Find "KETERANGAN = '60'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 60 K ...")
                        Case 7
                            nUrut = nI
                            nKeterangan = "70K saja "
                            rsTemp.Find "KETERANGAN = '70'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 70 K ...")
                        Case 8
                            nUrut = nI
                            nKeterangan = "80K saja "
                            rsTemp.Find "KETERANGAN = '80'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 80 K ...")
                        Case 9
                            nUrut = nI
                            nKeterangan = "90K saja "
                            rsTemp.Find "KETERANGAN = '90'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 90 K ...")
                        Case 10
                            nUrut = nI
                            nKeterangan = "100K saja "
                            rsTemp.Find "KETERANGAN = '100'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 100 K ...")
                        Case 11
                            nUrut = nI
                            nKeterangan = "Kelipatan 5K dibawah 100K "
                            rsTemp.Find "KETERANGAN = 'Kelipatan 5K'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Kelipatan 5 K < 100 K ...")
                        Case 12
                            nUrut = nI
                            nKeterangan = "SBE lainnya diatas 100K "
                            rsTemp.Find "KETERANGAN = 'Lainnya Diatas 100K'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data > 100 K ...")
                        Case 13
                            ' TOTAL SBE
                            nUrut = nI
                            nKeterangan = "  TOTAL SERVIS BERKALA EKSTERNAL "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
                            End If
                            nSubTotal = nSubTotal + nTotal
                            TotalSBE = nSubTotal
                            TotalUnitSBE = TotalSBE
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Total SBE ...")
                        Case 14
                            nUrut = nI
                            nKeterangan = "Unit Ganti Oli Mesin "
                            rsTemp.Find "KETERANGAN = 'Unit Ganti Oli Mesin'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Ganti Oli Mesin ...")
                        Case 15
                            nUrut = nI
                            nKeterangan = "General Repair Lainnya "
                            rsTemp.Find "KETERANGAN = 'Lainnya General Repair1'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Other GR ...")
                        Case 16
                            nUrut = nI - 1
                            nKeterangan = "General Repair Lainnya "
                            rsTemp.Find "KETERANGAN = 'Lainnya General Repair2'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 3)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 3)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Other GR ...")
                        Case 17
                            ' SUB TOTAL GRP
                            nUrut = nI - 1
                            nKeterangan = "  SUBTOTAL GENERAL REPAIR "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
                            End If
                            nSubTotal = nSubTotal + nTotal
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Sub Total GR ...")
                        Case 18
                            ' TOTAL CPUS
                            nUrut = nI - 1
                            nKeterangan = "  TOTAL CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 2, True)
                            End If
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Total CPUS ...")
                        Case 19
                            nUrut = nI - 1
                            nKeterangan = "1K Penjualan Outlet Sama "
                            rsTemp.Find "KETERANGAN = 'SBI 1 dealer'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 1K Local ...")
                        Case 20
                            nUrut = nI - 1
                            nKeterangan = "1K Total "
                            rsTemp.Find "KETERANGAN = 'SBI beda dealer'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Jumlah1K = nTotal
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data 1K Total ...")
                        Case 21
                            nUrut = nI - 1
                            nKeterangan = "Warranty "
                            rsTemp.Find "KETERANGAN = 'nilai WARRANTY'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Warranty ...")
                        Case 22
                            nUrut = nI - 1
                            nKeterangan = "Return Job "
                            rsTemp.Find "KETERANGAN = 'Return Job'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data RTJ ...")
                        Case 23
                            nUrut = nI - 1
                            nKeterangan = "Pre Delivery Servis "
                            rsTemp.Find "KETERANGAN = 'PDS'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data PDS ...")
                        Case 24
                            ' SUBTOTAL NONCPUS
                            nUrut = nI - 1
                            nKeterangan = "  SUBTOTAL NON CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
                            End If
                            nSubTotal = nSubTotal + nTotal
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Sub Total Non CPUS ...")
                        Case 25
                            nUrut = nI - 1
                            nKeterangan = "Jumlah Unit THS "
                            rsTemp.Find "KETERANGAN = 'THS'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                totalUnitEntryTHS = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                totalUnitEntryTHS = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data THS ...")
                        Case 26
                            ' JUMLAH TOTAL UNIT
                            nUrut = nI - 1
                            nKeterangan = "  JUMLAH TOTAL UNIT "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nSubTotal), 2, True)
                            End If
                            totalUnitEntry = CDbl(nSubTotal)
                            nSubTotal = 0
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Total Unit ...")
                        Case 27
                            nUrut = nI - 1
                            nKeterangan = "Pelanggan Booking yang datang pada hari Booking yg sama "
                            rsTemp.Find "KETERANGAN = 'Booking Hari Sama'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Booking On Day ...")
                        Case 28
                            nUrut = nI - 1
                            nKeterangan = "Labor "
                            rsTemp.Find "KETERANGAN = 'Revenue Jasa'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Revenue Jasa ...")
                        Case 29
                            nUrut = nI - 1
                            nKeterangan = "Labor Non Ppn"
                            rsTemp.Find "KETERANGAN = 'Revenue Jasa Non ppn'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Revenue Non PPn ...")
                        Case 30
                            nUrut = nI - 1
                            nKeterangan = "Labor Total (CPUS + Non CPUS) "
                            rsTemp.Find "KETERANGAN = 'Revenue Jasa Total'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal2 = nTotal2 + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Revenue Total ...")
                        Case 31
                            nUrut = nI - 1
                            nKeterangan = "Part Toyota & Oli TGMO "
                            rsTemp.Find "KETERANGAN = 'Revenue Part dan Oli TGMO'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Part & Oli Genuine ...")
                        Case 32
                            nUrut = nI - 1
                            nKeterangan = "Part Toyota & Oli TGMO (CPUS + Non CPUS)"
                            rsTemp.Find "KETERANGAN = 'Revenue Part dan Oli TGMO Total'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal2 = nTotal2 + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Part & Oli Genuine ...")
                        Case 33
                            nUrut = nI - 1
                            nKeterangan = "Oli / Bahan Non Toyota "
                            rsTemp.Find "KETERANGAN = 'Revenue Oli dan Bahan Non Toyota'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Part & Oli Non Genuine ...")
                        Case 34
                            nUrut = nI - 1
                            nKeterangan = "Oli / Bahan Non Toyota (CPUS + Non CPUS)"
                            rsTemp.Find "KETERANGAN = 'Revenue Oli dan Bahan Non Toyota Total'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal2 = nTotal2 + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Part & Oli Non Genuine ...")
                        Case 35
                            nUrut = nI - 1
                            nKeterangan = "Pekerjaan Luar "
                            rsTemp.Find "KETERANGAN = 'Revenue Pekerjaan Luar'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal = nTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal = nTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Sublet ...")
                        Case 36
                            nUrut = nI - 1
                            nKeterangan = "Pekerjaan Luar (CPUS + Non CPUS)"
                            rsTemp.Find "KETERANGAN = 'Revenue Pekerjaan Luar Total'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nTotal2 = nTotal2 + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Sublet ...")
                        Case 37
                            ' SUB TOTAL REVENUE CPUS
                            nUrut = nI - 1
                            nKeterangan = "  SUBTOTAL REVENUE CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
                            End If
                            nSubTotal = 0
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Revenue CPUS ...")
                        Case 38
                            ' SUB TOTAL REVENUE
                            nUrut = nI - 1
                            nKeterangan = "  SUBTOTAL REVENUE CPUS + Non CPUS "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal2), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal2), 2, True)
                            End If
                            nSubTotal = 0
                            nTotal2 = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Revenue CPUS & Non CPUS ...")
                        

                        Case 39
                            nUrut = nI - 1
                            nKeterangan = "Total Servis Berkala Eks Revenue "
                            rsTemp.Find "KETERANGAN = 'Revenue SBE'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Revenue SBE ...")
                            
                        Case 40
                            ' JML STALl EM
                            nUrut = nI - 1
                            nKeterangan = "Stall Express Maintenance (EM)"
                            rsTemp.Find "KETERANGAN = 'EM'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Stall ...")
                        Case 41
                            ' JML STALl Non EM ( Produktif )
                            nUrut = nI - 1
                            nKeterangan = "Stall Produktif Lainnya (Non EM)"
                            rsTemp.Find "KETERANGAN = 'NonEM'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Stall ...")
                        Case 42
                            ' Sub Total Stall Produktif
                            nUrut = nI - 1
                            nKeterangan = "  SUBTOTAL STALL PRODUKTIF "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
                            End If
                            nSubTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Stall ...")
                        Case 43
                            ' JML STALL Non Produktif
                            nUrut = nI - 1
                            nKeterangan = "Stall Non Produktif"
                            rsTemp.Find "KETERANGAN = 'NonPro'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Stall ...")
                        Case 44
                            ' JML MOBIL THS
                            nUrut = nI - 1
                            nKeterangan = "Jumlah Kendaraan THS"
                            rsTemp.Find "KETERANGAN = 'mobilTHS'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi THS ...")
                            
                        Case 45
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Teknisi PSG (Magang)"
                            rsTemp.Find "KETERANGAN = 'PSG'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
                        Case 46
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Apprentice Technician"
                            rsTemp.Find "KETERANGAN = 'APPMech'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
                        Case 47
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Toyota Technician"
                            rsTemp.Find "KETERANGAN = 'TTMech'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
                        Case 48
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Pro Technician"
                            rsTemp.Find "KETERANGAN = 'PTMech'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
                        Case 49
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Diagnosa Technician"
                            rsTemp.Find "KETERANGAN = 'DTMech'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
                        Case 50
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Diagnosa Master Technician"
                            rsTemp.Find "KETERANGAN = 'DMTMech'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
                        Case 51
                            ' Total Tech
                            nUrut = nI - 1
                            nKeterangan = "  TOTAL TEKNISI "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
                            End If
                            nTotal2 = nSubTotal
                            nSubTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Teknisi ...")
                            
                            
                        Case 52
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "App SA"
                            rsTemp.Find "KETERANGAN = 'APPSA'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi SA ...")
                        Case 53
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "TSA"
                            rsTemp.Find "KETERANGAN = 'TSA'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi SA ...")
                        Case 54
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "PSA"
                            rsTemp.Find "KETERANGAN = 'PSA'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi SA ...")
                        Case 55
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "MSA"
                            rsTemp.Find "KETERANGAN = 'MSA'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi SA ...")
                        Case 56
                            ' Total Tech
                            nUrut = nI - 1
                            nKeterangan = "  TOTAL SA"
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
                            End If
                            nTotal2 = nTotal2 + nSubTotal
                            nSubTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi SA ...")
                            
                            
                            
                        Case 57
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Toyota Foreman Training"
                            rsTemp.Find "KETERANGAN = 'TFT'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Foreman ...")
                        Case 58
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Non Toyota Foreman Training"
                            rsTemp.Find "KETERANGAN = 'NonTFT'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Foreman ...")
                        Case 59
                            ' Total FO
                            nUrut = nI - 1
                            nKeterangan = "  TOTAL FOREMAN "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
                            End If
                            nTotal2 = nTotal2 + nSubTotal
                            nSubTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Foreman ...")
                            
                            
                        Case 60
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Team 21 PT dan Lainnya"
                            rsTemp.Find "KETERANGAN = 'T21 PT'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Foreman Team 21 ...")
                        Case 61
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Team 21 DT"
                            rsTemp.Find "KETERANGAN = 'T21 DT'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Foreman Team 21 ...")
                        Case 62
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Team 21 DMT"
                            rsTemp.Find "KETERANGAN = 'T21 DMT'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Foreman Team 21 ...")
                            nSubTotal = 0
                    
                        
                        Case 63
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "APP Partman"
                            rsTemp.Find "KETERANGAN = 'APPpart'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Partman ...")
                        Case 64
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Partman Level 1"
                            rsTemp.Find "KETERANGAN = 'Part1'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Partman ...")
                        Case 65
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Partman Level 2"
                            rsTemp.Find "KETERANGAN = 'Part2'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nSubTotal = nSubTotal + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                nSubTotal = nSubTotal + 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Partman ...")
                        Case 66
                            ' Total Partman
                            nUrut = nI - 1
                            nKeterangan = "  TOTAL PARTMAN "
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nSubTotal, 2, True)
                            End If
                            nTotal2 = nTotal2 + nSubTotal
                            nSubTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Partman ...")
                            
                            
                        Case 67
                            ' Man Power
                            nUrut = nI - 1
                            nKeterangan = "Personel Bengkel Lainnya"
                            rsTemp.Find "KETERANGAN = 'PBLain'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nTotal2 = nTotal2 + CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Personel ...")
                        Case 68
                            ' Total Personel ALL
                            nUrut = nI - 1
                            nKeterangan = "  TOTAL PERSONEL"
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nTotal2, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nTotal2, 2, True)
                            End If
                            nTotal2 = 0
                            nSubTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Komposisi Personel ...")
                        
                        
                        Case 69
                            nUrut = nI - 1
                            nKeterangan = "Jam Tersedia (Biasa & Lembur) "
                            rsTemp.Find "Keterangan = 'Jam Tersedia'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH")) / 60, 2), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, Round(CDbl(rsTemp("JUMLAH")) / 60, 2), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Service Time ...")
                        Case 70
                            nUrut = nI - 1
                            nKeterangan = "Jam Terpakai "
                            rsTemp.Find "Keterangan = 'Jam Terpakai'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Service Time ...")
                        Case 71
                            nUrut = nI - 1
                            nKeterangan = "Jam Terjual "
                            rsTemp.Find "Keterangan = 'Jam Terjual'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Service Time ...")
                        Case 72
                            nUrut = nI
                            nHrsFollow = 0
                            nKeterangan = "Jumlah Unit yang harus di Follow Up "
                            rsTemp.Find "Keterangan = 'Unit hrs Follow'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nHrsFollow = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Follow Up ...")
                        Case 73
                            nUrut = nI
                            nDiFollow = 0
                            nKeterangan = "Jumlah Unit yang berhasil di Follow Up "
                            rsTemp.Find "Keterangan = 'Unit di Follow'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nDiFollow = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Follow Up ...")
                        Case 74
                            ' PSFU Success Call Rate
                            nUrut = nI - 3
                            nKeterangan = "  PSFU Success Call Rate (%) "
                            nTotal = 0
                            If nHrsFollow > 0 Then
                                nTotal = Round(nDiFollow / nHrsFollow, 2) * 100
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
                            End If
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Follow Up ...")
                        Case 75
                            nUrut = nI - 1
                            nKeterangan = "Jawaban Tidak untuk Q1 "
                            nNoQ1 = 0
                            rsTemp.Find "Keterangan = 'Unit NO Q1'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nNoQ1 = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Follow Up ...")
                        Case 76
                            nUrut = nI - 1
                            nKeterangan = "Jawaban Tidak untuk Q2 "
                            nNoQ2 = 0
                            rsTemp.Find "Keterangan = 'Unit NO Q2'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nNoQ2 = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Follow Up ...")
                        Case 77
                            nUrut = nI - 1
                            nKeterangan = "Jawaban Tidak untuk Q3 "
                            nNoQ3 = 0
                            rsTemp.Find "Keterangan = 'Unit NO Q3'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nNoQ3 = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Follow Up ...")
                        Case 78
                            nUrut = nI - 1
                            nKeterangan = "Jumlah Unit RTJ "
                            nRTJ = 0
                            rsTemp.Find "Keterangan = 'Unit Follow Rtj'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                nRTJ = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data RTJ Total ...")
                        Case 79
                            ' FIR Ratio (dari hasil PSFU)
                            nUrut = nI - 1
                            nKeterangan = "  FIR Ratio (dari hasil PSFU) (%) "
                            nTotal = 0
                            If nDiFollow > 0 Then
                                nTotal = Round(1 - (Round(nNoQ1 / nDiFollow, 2) + Round(nNoQ2 / nDiFollow, 2) + Round(nNoQ3 / nDiFollow, 2)), 2) * 100
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2, True)
                            End If
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data FIR ...")
                        Case 80
                            nUrut = nI - 1
                            nKeterangan = "Non Fix It Right (Q1) (%) "
                            nTotal = 0
                            If nDiFollow > 0 Then
                                nTotal = Round(nNoQ1 / nDiFollow, 2) * 100
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data FIR ...")
                        Case 81
                            nUrut = nI - 1
                            nKeterangan = "Non On Time Delivery Rate (Q2) (%) "
                            nTotal = 0
                            If nDiFollow > 0 Then
                                nTotal = Round(nNoQ2 / nDiFollow, 2) * 100
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data OTD ...")
                        Case 82
                            nUrut = nI - 1
                            nKeterangan = "Non Reasonable Time (Q3) (%) "
                            nTotal = 0
                            If nDiFollow > 0 Then
                                nTotal = Round(nNoQ3 / nDiFollow, 2) * 100
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data OTD ...")
                        Case 83
                            nUrut = nI - 1
                            nKeterangan = "Gate To Gate"
                            If gHrSama = True Then
                                rsTemp.Find "Keterangan = 'Lead Time Gate Total 1Day'", , adSearchForward, 1
                            Else
                                rsTemp.Find "Keterangan = 'Lead Time Gate Total '", , adSearchForward, 1
                            End If
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                GTG = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                GTG = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Gate To Gate ...")
                        Case 84
                            nUrut = nI - 1
                            nKeterangan = "Unit Entry"
                            Dim cUE As Double
                            Dim cTHS As Double
                            Dim nUE_THS As Double
                            If gHrSama = True Then
                                rsTemp.Find "Keterangan = 'Unit Gate Total 1Day'", , adSearchForward, 1
                                If Not rsTemp.EOF Then cUE = CDbl(rsTemp("JUMLAH"))
                                rsTemp.Find "Keterangan = 'Unit Gate Total 1Day THS'", , adSearchForward, 1
                                If Not rsTemp.EOF Then cTHS = CDbl(rsTemp("JUMLAH"))
                            Else
                                cUE = totalUnitEntry
                                cTHS = totalUnitEntryTHS
                            End If
                            
                            jTHS = cTHS
                            nUE_THS = cUE - cTHS
                            
                            Dim cUENotif As Double
                            rsTemp.Find "Keterangan = 'Unit Notif Total 1Day'", , adSearchForward, 1
                            If Not rsTemp.EOF Then cUENotif = CDbl(rsTemp("JUMLAH"))
                            Dim cUENotifths As Double
                            rsTemp.Find "Keterangan = 'Unit Notif Total 1Day THS'", , adSearchForward, 1
                            If Not rsTemp.EOF Then cUENotifths = CDbl(rsTemp("JUMLAH"))
                            
                            
                            If nUE_THS > 0 Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nUE_THS, 1)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 86, nKeterangan, (cUENotif - cUENotifths), 1)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 95, "Total Unit Entry non THS", (cUENotif - cUENotifths), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nUE_THS, 2)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 86, nKeterangan, (cUENotif - cUENotifths), 2)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 95, "Total Unit Entry non THS", (cUENotif - cUENotifths), 2)
                                End If
                                UE = nUE_THS
                                UEnT = cUENotif - cUENotifths
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 86, nKeterangan, 0, 1)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 95, "Total Unit Entry non THS", 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 86, nKeterangan, 0, 2)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 95, "Total Unit Entry non THS", 0, 2)
                                End If
                                UE = 0
                                UEnT = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data All Unit Entry ...")
                        Case 85
                        'Lead Time Gate Total 1Day
                            nUrut = nI - 1
                            Dim cGate As Double
                            rsTemp.Find "Keterangan = 'Lead Time Gate Total 1Day'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                cGate = CDbl(rsTemp("JUMLAH"))
                            Else
                                cGate = 0
                            End If
                            nKeterangan = "Jumlah LT Gate to Gate (Visiting) (menit)"
                            Dim GTGpUE As Double
                            If UE = 0 Then
                                GTGpUE = 0
                            Else
                                GTGpUE = Round((cGate / UE), 0)
                            End If
                            
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, GTGpUE, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, GTGpUE, 2, True)
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Lead Time ...")
                        Case 86
                            nUrut = nI - 1
                            nKeterangan = "Produksi Notifikasi"
                            If gHrSama = True Then
                                rsTemp.Find "Keterangan = 'Lead Time Notif Total 1Day'", , adSearchForward, 1
                            Else
                                rsTemp.Find "Keterangan = 'Lead Time Notif Total'", , adSearchForward, 1
                            End If
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                PN = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                PN = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Lead Time ...")
                        Case 87
                            nUrut = nI
                            Dim cLead As Double
                            rsTemp.Find "Keterangan = 'Lead Time Notif Total 1Day'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                cLead = CDbl(rsTemp("JUMLAH"))
                            Else
                                cLead = 0
                            End If
                            nKeterangan = "Jumlah Lead Time (menit)"
                            Dim LDpUE As Double
                            If UE = 0 Then
                                LDpUE = 0
                            Else
                                LDpUE = Round((cLead / UEnT), 0)
                            End If
                            
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, LDpUE, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, LDpUE, 2, True)
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Lead Time ...")
                        Case 88
                            nUrut = nI
                            nKeterangan = "Gate To Gate SBE"
                            If gHrSama = True Then
                                rsTemp.Find "Keterangan = 'Lead Time Gate SBE Total 1Day'", , adSearchForward, 1
                            Else
                                rsTemp.Find "Keterangan = 'Lead Time SBE Gate Total '", , adSearchForward, 1
                            End If
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                GTGs = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                GTGs = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Gate To Gate SBE ...")
                        Case 89
                            nUrut = nI
                            nKeterangan = "Unit SBE Non THS"
                            Dim cUESBE As Double
                            Dim jTHS1 As Double
                            Dim nUESBE_THS As Double
                            Dim nUESBE_THSnotif As Double
                            
                            
                            Dim cUENotifSB As Double
                            rsTemp.Find "Keterangan = 'Unit Notif SBE Total 1Day'", , adSearchForward, 1
                            If Not rsTemp.EOF Then cUENotifSB = CDbl(rsTemp("JUMLAH"))
                            
                            Dim cUENotifSBths As Double
                            rsTemp.Find "Keterangan = 'Unit Notif SBE Total 1Day THS'", , adSearchForward, 1
                            If Not rsTemp.EOF Then cUENotifSBths = CDbl(rsTemp("JUMLAH"))
                            
                            nUESBE_THSnotif = cUENotifSB - cUENotifSBths
                                
                            If gHrSama = True Then
                                rsTemp.Find "Keterangan = 'Unit Gate SBE Total 1Day'", , adSearchForward, 1
                                If Not rsTemp.EOF Then cUESBE = CDbl(rsTemp("JUMLAH"))
                                rsTemp.Find "Keterangan = 'Unit Gate SBE Total 1Day THS'", , adSearchForward, 1
                                If Not rsTemp.EOF Then jTHS1 = CDbl(rsTemp("JUMLAH"))
                                nUESBE_THS = cUESBE - jTHS1
                            Else
                                cUESBE = TotalUnitSBE
                                nUESBE_THS = cUESBE - jTHS
                            End If
                            
                            If nUESBE_THS > 0 Or nUESBE_THSnotif > 0 Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nUESBE_THS, 1)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 92, nKeterangan, nUESBE_THSnotif, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nUESBE_THS, 2)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 92, nKeterangan, nUESBE_THSnotif, 2)
                                End If
                                UEs = nUESBE_THS
                                UEsTHS = nUESBE_THSnotif
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 92, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                    Call objO.UpdateData(sSQL)
                                    sSQL = insertTabel(namaBulan, 92, nKeterangan, 0, 2)
                                End If
                                UEs = 0
                                UEsTHS = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data SBE Non THS ...")
                        Case 90
                            nUrut = nI
                            Dim jGTGSBE As Double
                            Dim GTGspUEs As Double
                            nKeterangan = "Jumlah LT Gate to Gate SBE (menit)"
                            rsTemp.Find "Keterangan = 'Lead Time Gate SBE Total 1Day'", , adSearchForward, 1
                            If Not rsTemp.EOF Then jGTGSBE = CDbl(rsTemp("JUMLAH"))
                            If UEs = 0 Then
                                GTGspUEs = 0
                            Else
                                GTGspUEs = Round((jGTGSBE / UEs), 0)
                            End If
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, GTGspUEs, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, GTGspUEs, 2, True)
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Gate To Gate Lead Time ...")
                        Case 91
                            nUrut = nI
                            nKeterangan = "Produksi Notifikasi SBE"
                            If gHrSama = True Then
                                rsTemp.Find "Keterangan = 'Lead Time Notif SBE Total 1Day'", , adSearchForward, 1
                            Else
                                rsTemp.Find "Keterangan = 'Lead Time SBE Notif Total '", , adSearchForward, 1
                            End If
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                PNs = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                PNs = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Lead Time ...")
                       Case 92
                            nUrut = nI + 1
                            nKeterangan = "Jumlah Lead Time SBE (menit)"
                            Dim PNspUEs As Double
                            Dim cPNs As Double
                            rsTemp.Find "Keterangan = 'Lead Time Notif SBE Total 1Day'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                cPNs = CDbl(rsTemp("JUMLAH"))
                            Else
                                cPNs = 0
                            End If
                            If UEsTHS = 0 Then
                                PNspUEs = 0
                            Else
                                PNspUEs = Round((cPNs / UEsTHS), 0)
                            End If
                            
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, PNspUEs, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, PNspUEs, 2, True)
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Lead Time ...")
                        Case 93
                            nUrut = nI + 1
                            nKeterangan = "Time Delivery (Notifikasi)"
                            rsTemp.Find "Keterangan = 'Time Delivery Notifikasi '", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("JUMLAH")), 2)
                                End If
                                TDN = CDbl(rsTemp("JUMLAH"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                TDN = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data OTD Notification ...")
                        Case 94
                            nUrut = nI + 2
                            nKeterangan = "On Time Delivery (%)"
                            Dim xOTD As Double
                            'If totalUnitEntry - totalUnitEntryTHS > 0 Then
                            '    xOTD = Round((TDN / (totalUnitEntry - totalUnitEntryTHS)) * 100, 1)
                            'Else
                            '    xOTD = 0
                            'End If
                            xOTD = objO.FetchData_Double("select (((Select Count(A.Slisih) from " & _
                                    "(SELECT CASE WHEN DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), " & _
                                    "ISNULL(TGLPANGGIL,'')) <= 0 AND DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), ISNULL(TGLPANGGIL,'')) >= -30 THEN 1 ELSE 0 END AS SLISIH " & _
                                    "FROM JOBDTA A " & _
                                    "INNER JOIN JOBDTL B ON A.NO_WO = B.NO_WO " & _
                                    "LEFT JOIN " & Log.LogQueue & "..FINISHSERVICE C ON A.NO_POLISI = C.NOPOL AND A.INVDATE = C.TGLFINISH AND PANGGIL <> 0 " & _
                                    "WHERE NOINV <> '' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND PDS = 0 AND THS = 0 AND TPSVC <> 'W' AND " & _
                                    "MONTH(DATE) = " & i & " AND YEAR(DATE) = " & Text(0) & ") A Where A.Slisih = 1)*100) " & _
                                    "/(Select Count(A.Slisih) from " & _
                                    "(SELECT CASE WHEN DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), " & _
                                    "ISNULL(TGLPANGGIL,'')) <= 0 AND DATEDIFF(n, (CONVERT( varchar(10), DATESERAH,101)+' '+TIMESERAH), ISNULL(TGLPANGGIL,'')) >= -30 THEN 1 ELSE 0 END AS SLISIH " & _
                                    "FROM JOBDTA A " & _
                                    "INNER JOIN JOBDTL B ON A.NO_WO = B.NO_WO " & _
                                    "LEFT JOIN " & Log.LogQueue & "..FINISHSERVICE C ON A.NO_POLISI = C.NOPOL AND A.INVDATE = C.TGLFINISH AND PANGGIL <> 0 " & _
                                    "WHERE NOINV <> '' AND DELIV <> 'BTL' AND FS_LOG_DEL = '' AND PDS = 0 AND THS = 0 AND TPSVC <> 'W' AND " & _
                                    "MONTH(DATE) = " & i & " AND YEAR(DATE) = " & Text(0) & ") A)) as OTD")

                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, xOTD, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, xOTD, 2, True)
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data OTD (%) ...")
                        Case 95
                            nUrut = nI + 2
                            nKeterangan = "Jumlah pelanggan yang diundang (aktual diundang) "
                            nTotal = 0
                            nTotal = CDbl(kpi6_b(CDbl(i), Trim(Text(0))))
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nDiUndang = CDbl(nTotal)
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Pelanggan ...")
                        Case 96
                            nUrut = nI + 2
                            nKeterangan = "Jumlah pelanggan yang seharusnya diundang "
                            nTotal = 0
                            nTotal = CDbl(kpi6_a(CDbl(i), Trim(Text(0))))
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nHrsUndang = CDbl(nTotal)
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Pelanggan ...")
                        Case 97
                            nUrut = nI + 2
                            nKeterangan = "Jumlah pelanggan yang datang setelah diundang "
                            nTotal = 0
                            nTotal = CDbl(kpi9_a(CDbl(i), Trim(Text(0))))
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(nTotal), 2)
                            End If
                            nDtgUndang = CDbl(nTotal)
                            nTotal = 0
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Pelanggan ...")
                        Case 98
                            nUrut = nI + 2
                            nKeterangan = "Total SBE SBI"
                            Dim oSBE As Double
                            Dim oSBI As Double
                            oSBI = Jumlah1K
                            oSBE = totalUnitEntry
                            Dim oSBESBI As Double
                            oSBESBI = oSBE + oSBI
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, oSBESBI, 1)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, oSBESBI, 2)
                            End If
                            SBESBI = oSBESBI
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Total SB ...")
                        Case 99
                            nUrut = nI + 2
                            nKeterangan = "Kontribusi MRS - Come After Invited (%)"
                            Dim nDtgUndangpSBESBI As Double
                            If SBESBI = 0 Then
                                nDtgUndangpSBESBI = 0
                            Else
                                nDtgUndangpSBESBI = Round((nDtgUndang / SBESBI) * 100, 0)
                            End If
                            
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nDtgUndangpSBESBI, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, nDtgUndangpSBESBI, 2, True)
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data MRS - Come After Invited...")

                        Case 100   'acep
                            Set rsTemp = obj.KPI_5(True, i, CInt(Text(0)))
                            nUrut = nI + 2
                            nKeterangan = "Jumlah Telepon Terhubung (1) (Tt)"
                            'rsTemp.Find "KETERANGAN = 'KPI_5A'", , adSearchForward, 1
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("xx")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("xx")), 2)
                                End If
                                TTh = CDbl(rsTemp("xx"))
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Kontribusi MRS ...")
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                TTh = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data MRS - Call ...")
                        Case 101
                            nUrut = nI + 2
                            Set rsTemp = obj.KPI_5(False, i, CInt(Text(0)))
                            nKeterangan = "Pelanggan yang ditelepon (SBE+SBI)"
                            If Not rsTemp.EOF Then
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("xx")), 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, CDbl(rsTemp("xx")), 2)
                                End If
                                PTh = CDbl(rsTemp("xx"))
                            Else
                                If cekInsert(nUrut, nKeterangan) Then
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 1)
                                Else
                                    sSQL = insertTabel(namaBulan, nUrut, nKeterangan, 0, 2)
                                End If
                                PTh = 0
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data MRS - Call (SBE & SBI) ...")
                        Case 102
                            nUrut = nI + 2
                            nKeterangan = "Database Valid (%)"
                            Dim TThpPTh As Double
                            If PTh = 0 Then
                                TThpPTh = 0
                            Else
                                TThpPTh = Round((TTh / PTh) * 100, 0)
                            End If
                            
                            If cekInsert(nUrut, nKeterangan) Then
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, TThpPTh, 1, True)
                            Else
                                sSQL = insertTabel(namaBulan, nUrut, nKeterangan, TThpPTh, 2, True)
                            End If
                            Call BarProses(nUrut, 105, 2, 2, nUrut, "Proses Data Database Valid ...")
                    End Select
                    nI = nI + 1
                    Call objO.UpdateData(sSQL)
                Loop
            End If
            rsTemp.Close
        'Next
        nBar.Min = 0
        nBar.Max = 10
        nBar.value = 0
        nLab.Caption = "Get Data Report : 0 Of 10"
        nPro.Caption = "Process :  1 Of 2"
        nKet.Caption = "..."
        PicBar.Visible = False
        DoEvents
        Shape1.Height = 615
        Call Cetak
        Call Transfer_JAF(CInt(Left(cboBulan1, InStr(cboBulan1, " |") - 1)), Text(0).Text)
    Else
        MsgBox pesan, vbCritical, "KELENGKAPAN"
        Exit Sub
    End If
End Sub

Private Sub Cetak()
Dim oRs As Object
Dim Report As New RptRepoBulanan
    

    sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] ORDER BY urut"
    Set oRs = objO.FetchData(sSQL)
    Report.Database.SetDataSource oRs
    
    Report.txtPerusahaan.SetText gPerush
    Report.txtTahun.SetText CDbl(Text(0))
    Report.txtTanggal.SetText Mid(cboBulan1, InStr(cboBulan1, "| ") + 2, Len(cboBulan1) - InStr(cboBulan1, "| ") + 2) & CDbl(Text(0))
    
    j = CDbl(Left(cboBulan1, InStr(cboBulan1, " |") - 1))
    Report.txtBln1.SetText GetNamaBulan(j)
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Resize()
    Me.WindowState = vbMaximized
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width + 60 '0
    CRViewer1.Height = ScaleHeight - asxBars1.Height - 120
    CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 180
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Report = Nothing
End Sub

Private Function jasaPpn(Bln As Integer, Thn As Integer) As Double
Dim rsSearch As ADodb.Recordset
Dim namaTabel As String
    
    jasaPpn = 0
    namaTabel = "REVENUE"
    Call IsiTempRepo(Bln, Thn, namaTabel, "")
    
    Set rsSearch = objO.FetchData("SELECT ISNULL(SUM(ISNULL(sbi+sbe+grp+twc+lub+pdi+was,0)),0) AS TOTJASA FROM [" & namaTabel & "] WHERE TIPEINV <> 'AFFNONPPN' AND SBI = 0 AND RTJ = 0 AND PDI = 0 AND TWC = 0")
    If Not rsSearch.EOF Then
        jasaPpn = CDbl(rsSearch("TOTJASA"))
    End If
End Function

Private Function jasaNonPpn(Bln As Integer, Thn As Integer) As Double
Dim rsHlp As ADodb.Recordset
Dim rsSearch As ADodb.Recordset
Dim namaTabel As String
    
    jasaNonPpn = 0
    namaTabel = "REVENUE"
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If rsHlp.EOF Then
       Call IsiTempRepo(Bln, Thn, namaTabel, "")
    End If
    
    Set rsSearch = objO.FetchData("SELECT ISNULL(SUM(ISNULL(sbi+sbe+grp+twc+lub+pdi+was,0)),0) AS TOTJASA FROM [" & namaTabel & "] WHERE TIPEINV = 'AFFNONPPN' AND SBI = 0 AND RTJ = 0 AND PDI = 0 AND TWC = 0")
    If Not rsSearch.EOF Then
        jasaNonPpn = CDbl(rsSearch("TOTJASA"))
    End If
End Function

Private Function jasaTotal(Bln As Integer, Thn As Integer) As Double
Dim rsHlp As ADodb.Recordset
Dim rsSearch As ADodb.Recordset
Dim namaTabel As String
    
    jasaTotal = 0
    namaTabel = "REVENUE"
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If rsHlp.EOF Then
       Call IsiTempRepo(Bln, Thn, namaTabel, "")
    End If
    
    Set rsSearch = objO.FetchData("SELECT ISNULL(SUM(ISNULL(sbi+sbe+grp+twc+lub+pdi+was,0)),0) AS TOTJASA FROM [" & namaTabel & "]")
    If Not rsSearch.EOF Then
        jasaTotal = CDbl(rsSearch("TOTJASA"))
    End If
End Function

Private Sub IsiTempRepo(Bln As Integer, Thn As Integer, vNamaTabel As String, txt As String)
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim hlp5 As String
Dim hlp6 As String
Dim hlp7 As String
Dim rsHlp As ADodb.Recordset
Dim rsPro As ADodb.Recordset
Dim rsAdd As ADodb.Recordset
Dim nNilai As Double
Dim nDuoInv As Boolean
Dim subletOnly As Boolean
Dim noinv As String
Dim nmCust As String
Dim nStatus As String
'Dim vNamaTabel As  String
    
    Set rsPro = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                             & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                             & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WHERE NO_WO = A.NO_WO) AS HPP " _
                             & "FROM T_INVOICE A LEFT JOIN CUSTDAT B ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                             & "WHERE A.FS_LOG_DEL = '' AND MONTH(INVDATE) = '" & Bln & "' AND YEAR(INVDATE) = '" & Thn & "' ORDER BY NOINV")
                             ' --:: 23 AGUSTUS 2010 ; DELIV <> 'BTL' AND ; Menampilkan yg Retur ::--
    If Not rsPro.EOF Then
        Call BuatTabelRepoInv(vNamaTabel)
        Do While Not rsPro.EOF
            subletOnly = True
            For i = 1 To 8
                hlp = "REP" & i
                If rsPro(hlp) <> "" Then
                    subletOnly = False
                    hlp2 = "KLPKERJA" & i
                    hlp3 = "RTE" & i
                    hlp4 = "MEKANIK" & i
                    hlp5 = "OFF" & i
                    hlp6 = "DISC" & i
                    hlp7 = "FREEC" & i
                    If createTemp(vNamaTabel) Then
                        nDuoInv = False
                        If rsPro("NDUOINV") = 1 Then
                            nDuoInv = True
                        End If
                        '#########################################################'
                        '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                        '#########################################################'
                        nStatus = rsPro("STATUS")
                        nNilai = CDbl(rsPro(hlp5)) * (1 - (CDbl(rsPro(hlp6)) / 100))
                        noinv = Trim(rsPro("NOINV"))
                        nmCust = Trim(rsPro("CUSTOMER"))
                        
                        If nStatus = "AFFNONPPN" Then
                            nNilai = nNilai * (100 / 110)
                        End If
                        
                        Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                        If Not rsHlp.EOF Then
                            Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                         & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsPro(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsPro(hlp2) = "SBI" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "SBE = SBE + '" & IIf(rsPro(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsPro(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "TWC = TWC + '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsPro(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "PDI = PDI + '" & IIf(rsPro(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsPro(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "BOR = '" & cekDuo(CDbl(rsPro("BODY")), noinv) & "', MATERIAL = '" & cekDuo(CDbl(rsPro("MTRL")), noinv) & "', PART = '" & cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noinv) & "', " _
                                         & "OLI = '" & cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noinv) & "', SUBLET = '" & cekDuo(CDbl(rsPro("SBLT")), noinv) & "', SUBTOT = '" & cekDuo(CDbl(rsPro("SUBTOT")), noinv) & "', " _
                                         & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "', HPP = '" & cekDuo(CDbl(rsPro("HPP")), noinv) & "' " _
                                         & "WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                        Else
                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noinv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                         & "'0','" & Replace(CDbl(rsPro(hlp3)), ",", ".") & "', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '" & IIf(rsPro(hlp2) = "SBI" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & IIf(rsPro(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "','" & IIf(rsPro(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & IIf(rsPro(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & cekDuo(CDbl(rsPro("BODY")), noinv) & "', '" & cekDuo(CDbl(rsPro("MTRL")), noinv) & "', '" & cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noinv) & "', '" & cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noinv) & "', '" & cekDuo(CDbl(rsPro("SBLT")), noinv) & "', '" & cekDuo(CDbl(rsPro("SUBTOT")), noinv) & "', " _
                                         & " '" & cekDuo(CDbl(rsPro("MET")), noinv, 1) & "', '" & CDbl(rsPro("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "', '" & cekDuo(CDbl(rsPro("HPP")), noinv) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
                        End If
                        
                    End If
                End If
            Next
            If subletOnly Then
                nDuoInv = False
                noinv = Trim(rsPro("NOINV"))
                
                Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                If Not rsHlp.EOF Then
                    Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NOINV = '" & Trim(noinv) & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                 & "JOR = '0', TIPEINV = '" & Trim(rsPro("STATUS")) & "', CUSTOMER = '" & Trim(rsPro("CUSTOMER")) & "', BOR = '" & CDbl(rsPro("BODY")) & "', MATERIAL = '" & CDbl(rsPro("MTRL")) & "', " _
                                 & "PART = '" & CDbl(rsPro("PART_G") + rsPro("PART_NG")) & "', OLI = '" & CDbl(rsPro("OIL_G") + rsPro("OIL_NG")) & "', SUBLET = '" & CDbl(rsPro("SBLT")) & "', SUBTOT = '" & CDbl(rsPro("SUBTOT")) & "', " _
                                 & "LABOR = '" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsPro("NOFAKTUR")) & "', HPP = '" & CDbl(rsPro("HPP")) & "' WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                Else
                    Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noinv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                 & "'0','0', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                 & "'" & CDbl(rsPro("BODY")) & "', '" & CDbl(rsPro("MTRL")) & "', '" & CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")) & "', '" & CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")) & "', '" & CDbl(rsPro("SBLT")) & "', '" & CDbl(rsPro("SUBTOT")) & "', " _
                                 & " '" & CDbl(rsPro("MET")) & "', '" & CDbl(rsPro("TOTALL")) & "', '" & Trim(rsPro("STATUS")) & "', '" & Trim(rsPro("CUSTOMER")) & "','" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & Trim(rsPro("NOFAKTUR")) & "','" & CDbl(rsPro("HPP")) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
                End If
            End If
            
            ' DARI ADD JOB
            Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND FS_LOG_DEL = '" & Trim(rsPro("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsPro("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
            If Not rsAdd.EOF Then
                Do While Not rsAdd.EOF
                    hlp2 = "KLPKERJA"
                    hlp3 = "RTE"
                    hlp4 = "MEKANIK"
                    hlp5 = "OFFN"
                    hlp6 = "FREEC"
                    
                    nDuoInv = False
                    If rsPro("NDUOINV") = 1 Then
                        nDuoInv = True
                    End If
                    
                    '#########################################################'
                    '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                    '#########################################################'
                    nStatus = rsPro("STATUS")
                    nNilai = CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsPro("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
                    noinv = Trim(rsPro("NOINV"))
                    nmCust = Trim(rsPro("CUSTOMER"))

                    If Trim(nStatus) = "AFFNONPPN" Then
                        nNilai = nNilai * (100 / 110)
                    End If
                    
                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                    If Not rsHlp.EOF Then
                        objO.UpdateData ("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NOINV = '" & Trim(noinv) & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                     & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "TWC = TWC + '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "BOR = '" & cekDuo(CDbl(rsPro("BODY")), noinv) & "', MATERIAL = '" & cekDuo(CDbl(rsPro("MTRL")), noinv) & "', PART = '" & cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noinv) & "', " _
                                     & "OLI = '" & cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noinv) & "', SUBLET = '" & cekDuo(CDbl(rsPro("SBLT")), noinv) & "', SUBTOT = '" & CDbl(rsPro("SUBTOT")) & "', HPP = '" & cekDuo(CDbl(rsPro("HPP")), noinv) & "', " _
                                     & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "' " _
                                     & "WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                    Else
                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noinv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                     & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "SBE" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsPro("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "LUB" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & cekDuo(CDbl(rsPro("BODY")), noinv) & "', '" & cekDuo(CDbl(rsPro("MTRL")), noinv) & "', '" & cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noinv) & "', '" & cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noinv) & "', '" & cekDuo(CDbl(rsPro("SBLT")), noinv) & "', '" & CDbl(rsPro("SUBTOT")) & "', " _
                                     & " '" & cekDuo(CDbl(rsPro("MET")), noinv, 1) & "', '" & CDbl(rsPro("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "','" & cekDuo(CDbl(rsPro("HPP")), noinv) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
                    End If
                    
                    rsAdd.MoveNext
                Loop
            End If
            rsPro.MoveNext
        Loop
        Call IsiTempRepoRetur(Bln, Thn, vNamaTabel, txt)
    Else
        Exit Sub
    End If
End Sub

Private Sub IsiTempRepoRetur(Bln As Integer, Thn As Integer, vNamaTabel As String, txt As String)
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim hlp5 As String
Dim hlp6 As String
Dim hlp7 As String
Dim rsHlp As ADodb.Recordset
Dim rsPro As ADodb.Recordset
Dim rsAdd As ADodb.Recordset
Dim nNilai As Double
Dim nDuoInv As Boolean
Dim subletOnly As Boolean
Dim noinv As String
Dim nmCust As String
Dim nStatus As String
    
    Set rsPro = objO.FetchData("SELECT PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY AS SUBTOT, PART_G+PART_NG+MTRL+OIL_G+OIL_NG+SBLT+LABOR+BODY+MET AS TOTALL, " _
                             & "KODETIPE, ISNULL(C.DOUBLEINV,0) AS NDUOINV, A.*, B.FS_KD_DEALER AS KDPENJUAL, Z.FS_NM_TAGIH AS NMPENJUAL, " _
                             & "(SELECT ISNULL(SUM(ISNULL(HARGA,0)*ISNULL(QTY,0)),0) AS HPP FROM JOBSUBLET WHERE NO_WO = A.NO_WO) AS HPP " _
                             & "FROM T_INVOICE A LEFT JOIN CUSTDAT B ON A.NO_CHASSIS = B.NO_CHASSIS LEFT JOIN T_DEALER Z ON B.FS_KD_DEALER = Z.FS_KD_DEALER LEFT JOIN HO_PROGRAM C ON A.KODETIPE = C.ID_PROGRAM AND A.TPSVC = 'P' " _
                             & "WHERE DELIV = 'BTL' AND A.FS_LOG_DEL = '' AND MONTH(INVDATE) = '" & Bln & "' AND YEAR(INVDATE) = '" & Thn & "' ORDER BY NOINV")
    If Not rsPro.EOF Then
        If createTemp(vNamaTabel) Then
        Do While Not rsPro.EOF
            subletOnly = True
            For i = 1 To 8
                hlp = "REP" & i
                If rsPro(hlp) <> "" Then
                    subletOnly = False
                    hlp2 = "KLPKERJA" & i
                    hlp3 = "RTE" & i
                    hlp4 = "MEKANIK" & i
                    hlp5 = "OFF" & i
                    hlp6 = "DISC" & i
                    hlp7 = "FREEC" & i
                    If createTemp(vNamaTabel) Then
                        nDuoInv = False
                        If rsPro("NDUOINV") = 1 Then
                            nDuoInv = True
                        End If
                        '#########################################################'
                        '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                        '#########################################################'
                        nStatus = rsPro("STATUS")
                        nNilai = -1 * CDbl(rsPro(hlp5)) * (1 - (CDbl(rsPro(hlp6)) / 100))
                        noinv = Trim(rsPro("NOINV")) & "R"
                        nmCust = Trim(rsPro("CUSTOMER"))
                        
                        If nStatus = "AFFNONPPN" Then
                            nNilai = nNilai * (100 / 110)
                        End If
                        
                        Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                        If Not rsHlp.EOF Then
                            objO.UpdateData ("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                         & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsPro(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsPro(hlp2) = "SBI" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "SBE = SBE + '" & IIf(rsPro(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsPro(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "TWC = TWC + '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsPro(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "PDI = PDI + '" & IIf(rsPro(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsPro(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "BOR = '" & -1 * cekDuo(CDbl(rsPro("BODY")), noinv) & "', MATERIAL = '" & -1 * cekDuo(CDbl(rsPro("MTRL")), noinv) & "', PART = '" & -1 * cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noinv) & "', " _
                                         & "OLI = '" & -1 * cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noinv) & "', SUBLET = '" & -1 * cekDuo(CDbl(rsPro("SBLT")), noinv) & "', SUBTOT = '" & -1 * cekDuo(CDbl(rsPro("SUBTOT")), noinv) & "', " _
                                         & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "', HPP = '" & -1 * cekDuo(CDbl(rsPro("HPP")), noinv) & "' " _
                                         & "WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                        Else
                            Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noinv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                         & "'0','" & Replace(CDbl(rsPro(hlp3)), ",", ".") & "', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '" & IIf(rsPro(hlp2) = "SBI" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & IIf(rsPro(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "','" & IIf(rsPro(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & IIf(rsPro(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', '" & IIf(rsPro(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                         & "'" & -1 * cekDuo(CDbl(rsPro("BODY")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("MTRL")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("SBLT")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("SUBTOT")), noinv) & "', " _
                                         & " '" & -1 * cekDuo(CDbl(rsPro("MET")), noinv, 1) & "', '" & -1 * CDbl(rsPro("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "', '" & -1 * cekDuo(CDbl(rsPro("HPP")), noinv) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
                        End If
                        
                    End If
                End If
            Next
            If subletOnly Then
                nDuoInv = False
                If rsPro("NDUOINV") = 1 Then nDuoInv = True
                '#########################################################'
                '..:: UPDATE CODING : PISAH A-B TGL 29 SEPTEMBER 2010 ::..'
                '#########################################################'
                noinv = Trim(rsPro("NOINV")) & "R"
                
                Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                If Not rsHlp.EOF Then
                    Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NOINV = '" & Trim(noinv) & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                 & "JOR = '0', TIPEINV = '" & Trim(rsPro("STATUS")) & "', CUSTOMER = '" & Trim(rsPro("CUSTOMER")) & "', BOR = '" & -1 * CDbl(rsPro("BODY")) & "', MATERIAL = '" & -1 * CDbl(rsPro("MTRL")) & "', " _
                                 & "PART = '" & -1 * CDbl(rsPro("PART_G") + rsPro("PART_NG")) & "', OLI = '" & -1 * CDbl(rsPro("OIL_G") + rsPro("OIL_NG")) & "', SUBLET = '" & -1 * CDbl(rsPro("SBLT")) & "', SUBTOT = '" & -1 * CDbl(rsPro("SUBTOT")) & "', " _
                                 & "LABOR = '" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & Trim(rsPro("NOFAKTUR")) & "', HPP = '" & -1 * CDbl(rsPro("HPP")) & "' WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                Else
                    Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noinv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                 & "'0','0', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '0', '0','0', '0', '0', '0', '0', " _
                                 & "'" & -1 * CDbl(rsPro("BODY")) & "', '" & -1 * CDbl(rsPro("MTRL")) & "', '" & -1 * CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")) & "', '" & -1 * CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")) & "', '" & -1 * CDbl(rsPro("SBLT")) & "', '" & -1 * CDbl(rsPro("SUBTOT")) & "', " _
                                 & " '" & -1 * CDbl(rsPro("MET")) & "', '" & -1 * CDbl(rsPro("TOTALL")) & "', '" & Trim(rsPro("STATUS")) & "', '" & Trim(rsPro("CUSTOMER")) & "','" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & Trim(rsPro("NOFAKTUR")) & "','" & -1 * CDbl(rsPro("HPP")) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
                End If
            End If
            
            ' DARI ADD JOB
            Set rsAdd = objO.FetchData("SELECT * FROM HIST_JOBADD WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND FS_LOG_DEL = '" & Trim(rsPro("FS_LOG_DEL")) & "' AND FD_TGL_DEL = '" & Format(rsPro("FD_RETUR"), "MM/DD/YYYY HH:NN") & "'")
            If Not rsAdd.EOF Then
                Do While Not rsAdd.EOF
                    hlp2 = "KLPKERJA"
                    hlp3 = "RTE"
                    hlp4 = "MEKANIK"
                    hlp5 = "OFFN"
                    hlp6 = "FREEC"
                    
                    nDuoInv = False
                    If rsPro("NDUOINV") = 1 Then
                        nDuoInv = True
                    End If

                    nStatus = rsPro("STATUS")
                    nNilai = -1 * CDbl(rsAdd(hlp5)) * (1 - (IIf(Right(Trim(rsPro("NOINV")), 2) = "-B", IIf(rsAdd("FREEC"), 100, 0), rsAdd("DISC")) / 100))
                    noinv = Trim(rsPro("NOINV")) & "R"
                    nmCust = Trim(rsPro("CUSTOMER"))
                    
                    If Trim(nStatus) = "AFFNONPPN" Then
                        nNilai = nNilai * (100 / 110)
                    End If
                    
                    Set rsHlp = objO.FetchData("SELECT * FROM [TEMPDB]..[" & vNamaTabel & "] WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "'")
                    If Not rsHlp.EOF Then
                        Call objO.UpdateData("UPDATE [TEMPDB]..[" & vNamaTabel & "] SET SA = '" & rsPro("FRONT") & "', NOINV = '" & Trim(noinv) & "', NO_POLISI = '" & rsPro("NO_POLISI") & "', " _
                                     & "JOR = '0', JRTE = JRTE + '" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', SBI = SBI + '" & IIf(rsAdd(hlp2) = "SBI" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "SBE = SBE + '" & IIf(rsAdd(hlp2) = "SBE" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', GRP = GRP + '" & IIf(rsAdd(hlp2) = "GRP" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "TWC = TWC + '" & IIf(rsPro("TPSVC") = "W", Round(CDbl(nNilai)), 0) & "', LUB = LUB + '" & IIf(rsAdd(hlp2) = "LUB" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "PDI = PDI + '" & IIf(rsAdd(hlp2) = "PDS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', WAS = WAS + '" & IIf(rsAdd(hlp2) = "WAS" And rsPro("TPSVC") <> "W", Round(CDbl(nNilai)), 0) & "', " _
                                     & "BOR = '" & -1 * cekDuo(CDbl(rsPro("BODY")), noinv) & "', MATERIAL = '" & -1 * cekDuo(CDbl(rsPro("MTRL")), noinv) & "', PART = '" & -1 * cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noinv) & "', " _
                                     & "OLI = '" & -1 * cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noinv) & "', SUBLET = '" & -1 * cekDuo(CDbl(rsPro("SBLT")), noinv) & "', SUBTOT = '" & -1 * CDbl(rsPro("SUBTOT")) & "', HPP = '" & -1 * cekDuo(CDbl(rsPro("HPP")), noinv) & "', " _
                                     & "TIPEINV = '" & Trim(nStatus) & "', CUSTOMER = '" & Trim(nmCust) & "', LABOR = '" & Replace(CDbl(rsPro("LABOR")), ",", ".") & "', NOFAKTUR = '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "' " _
                                     & "WHERE NO_CHASSIS = '" & rsPro("NO_CHASSIS") & "' AND NO_WO = '" & rsPro("NO_WO") & "' AND NOINV = '" & Trim(noinv) & "'")
                    Else
                        Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsPro("FRONT") & "', '" & Trim(noinv) & "','" & rsPro("NO_WO") & "', '" & rsPro("NO_POLISI") & "','" & rsPro("NO_CHASSIS") & "', " _
                                     & "'0','" & Replace(CDbl(rsAdd(hlp3)), ",", ".") & "', '" & IIf(rsPro("FN_RETURN") = 0, 0, 1) & "', '" & IIf(rsPro("THS") = 0, 0, 1) & "', '" & IIf(rsAdd(hlp2) = "SBI" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "SBE" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "','" & IIf(rsAdd(hlp2) = "GRP" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsPro("TPSVC") = "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & IIf(rsAdd(hlp2) = "LUB" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "PDS" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', '" & IIf(rsAdd(hlp2) = "WAS" And rsPro("TPSVC") <> "W", CDbl(nNilai), 0) & "', " _
                                     & "'" & -1 * cekDuo(CDbl(rsPro("BODY")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("MTRL")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("PART_G")) + CDbl(rsPro("PART_NG")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("OIL_G")) + CDbl(rsPro("OIL_NG")), noinv) & "', '" & -1 * cekDuo(CDbl(rsPro("SBLT")), noinv) & "', '" & -1 * CDbl(rsPro("SUBTOT")) & "', " _
                                     & " '" & -1 * cekDuo(CDbl(rsPro("MET")), noinv, 1) & "', '" & -1 * CDbl(rsPro("TOTALL")) & "', '" & Trim(nStatus) & "', '" & Trim(nmCust) & "','" & -1 * Replace(CDbl(rsPro("LABOR")), ",", ".") & "', '" & IIf(nDuoInv, "", Trim(rsPro("NOFAKTUR"))) & "','" & -1 * cekDuo(CDbl(rsPro("HPP")), noinv) & "','" & Format(rsPro("INVDATE"), "MM/DD/YYYY") & "')")
                    End If
                    
                    rsAdd.MoveNext
                Loop
            End If
            rsPro.MoveNext
        Loop
        End If
    End If
End Sub

Private Function statusDealer(codeStatus As Integer)
    Select Case codeStatus
        Case 1
            statusDealer = "AFFNONPPN"
        Case 2
            statusDealer = "AFFPPN"
        Case 3
            statusDealer = "EXT"
    End Select
End Function

Private Function cekDuo(hrg As Double, nINV As String, Optional stat As Integer)
    If Right(Trim(nINV), 2) = "-B" Then
        cekDuo = 0
        If stat = 1 Then
            cekDuo = CDbl(hrg)
        End If
    Else
        cekDuo = CDbl(hrg)
        If stat = 2 Then
            cekDuo = 0
        End If
    End If
End Function

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADodb.Recordset
    createTemp = False
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTabelRepoInv(vNamaTabel)
        createTemp = True
    End If
End Function

Private Sub BuatTabelRepoInv(vNamaTabel As String)
Dim strSQL As String

    strSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.UpdateData(strSQL)

    strSQL = "CREATE TABLE [tempdb]..[" & vNamaTabel & "] (" _
         & "    [SA][char](5) NOT NULL DEFAULT '', " _
         & "    [NOINV][varchar](15) NOT NULL DEFAULT '', " _
         & "    [NO_WO][varchar](10) NOT NULL DEFAULT '', " _
         & "    [NO_POLISI][varchar](12) NOT NULL DEFAULT '', " _
         & "    [NO_CHASSIS][varchar](25) NOT NULL DEFAULT '', " _
         & "    [JOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [JRTE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [RTJ][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [THS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SBE][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [GRP][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TWC][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [LUB][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PDI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [WAS][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [BOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [MATERIAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [PART][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [OLI][decimal](18,2) NOT NULL DEFAULT 0, "
         
    strSQL = strSQL + "[SUBLET][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [SUBTOT][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [METERAI][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TOTAL][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [TIPEINV][varchar](20) NOT NULL DEFAULT '', " _
         & "    [CUSTOMER][varchar](100) NOT NULL DEFAULT '', " _
         & "    [LABOR][decimal](18,2) NOT NULL DEFAULT 0, " _
         & "    [NOFAKTUR][varchar](30) NOT NULL DEFAULT ''," _
         & "    [HPP][decimal](18,2) NOT NULL DEFAULT 0," _
         & "    [INVDATE][datetime] NOT NULL DEFAULT '01/01/3000') ON [PRIMARY]"
    Call objO.UpdateData(strSQL)
End Sub

Private Function totalUnitEntry(Bln As Integer, Thn As Integer) As Integer
Dim rsSearch As ADodb.Recordset

    totalUnitEntry = 0
    Set rsSearch = objO.FetchData("SELECT COUNT(*) AS JUMLAH FROM T_INVOICE WHERE MONTH(DATE) = '" & Bln & "' AND YEAR(DATE) = '" & Thn & "' AND DELIV <> 'BTL' AND NOINV NOT LIKE '%-B' AND FS_LOG_DEL = ''")
    If Not rsSearch.EOF Then
        totalUnitEntry = CDbl(rsSearch("JUMLAH"))
    End If
    Set rsSearch = Nothing
    
End Function

Private Sub Transfer_JAF(Bln As Integer, Thn As Integer)
Dim rsTr As New ADodb.Recordset
Set rsTr = objO.FetchData("EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'")
    If Log.LogHOConn = 1 Then
        If ConHO = False Then
            Call Saving_Log_HO("EXEC REPOTAF", "EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'", "Bulan = " & Bln & " And Tahun = " & Thn)
        Else
            If Get_ConnectionProfil_Bengkel99 = False Then
                Call Saving_Log_HO("EXEC REPOTAF", "EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'", "Bulan = " & Bln & " And Tahun = " & Thn)
            Else
                Call objHO.OpenTransactionLog_NRM
                Call objHO.AddNew_WO("Delete Bengkel99..Jam_TAF WHere Bulan = " & Bln & " And Tahun = " & Thn & " And Cabang = '" & Log.LogNamaNas & "'")
                rsTr.MoveFirst
                Do Until rsTr.EOF
                    Call objHO.AddNew_WO("Insert Into Bengkel99..Jam_TAF Values (" & rsTr.Fields(0).value & ", " & rsTr.Fields(1).value & ",'" & rsTr.Fields(2).value & "','" & rsTr.Fields(3).value & "'," & RepRp(rsTr.Fields(4).value) & ")")
                    rsTr.MoveNext
                Loop
                Call objHO.CloseTransactionLog_NRM
                If errHO = True Then Call Saving_Log_HO("EXEC REPOTAF", "EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'", "Bulan = " & Bln & " And Tahun = " & Thn)
            End If
        End If
    Else
        Call Saving_Log_HO("EXEC REPOTAF", "EXEC REPOTAF " & Bln & "," & Thn & ",'" & Log.LogNamaNas & "'", "Bulan = " & Bln & " And Tahun = " & Thn)
    End If
End Sub

Private Sub Saving_Log_HO(nStatus As String, isSQl As String, nData As String)
    Call objHO.AddNew_LogHO(nData, "EXEC", nStatus, isSQl)
End Sub

Private Sub cmdPrintS_Click()
    Dim objSt As New clsStall
    Dim nOpt As Integer
    nOpt = IIf(Opt(0).value = True, 0, 1)
    Dim nTotalUnit(31) As Double
    
    Dim rsSt As ADodb.Recordset
    Dim nStNP As Double
    Dim nStEM As Double
    Dim nStGR As Double
    
    Dim nTot(31) As Double
    Dim nJA(31) As Double
    Dim nJF(31) As Double
    Dim nRPS(31) As Double
    Dim nLTr(31) As Double
    Dim nOPS(31) As Double
    Dim nDay(31) As Double
    Dim sSQLs As String
    Dim C As Integer
    
    PicBarEM.Visible = True
    nBarEM.Min = 0
    nBarEM.Max = Day(DateSerial(Val(cboThnS.Text), Val(cboBlnS.Text) + 1, 1 - 1))
    nBarEM.value = 0
    DoEvents
    
    For C = 1 To Day(DateSerial(Val(cboThnS.Text), Val(cboBlnS.Text) + 1, 1 - 1))
    
        nLabEM.Caption = "Colecting Data " & IIf(nOpt = 0, "NP,", "EM,") & "   Periode " & C & "/" & cboBlnS.Text & "/" & cboThnS.Text
        nBarEM.value = C - 1
        DoEvents
        
        Set rsSt = objSt.FetchQtyStall_aDay(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text))
        If rsSt.RecordCount = 0 Then
            nStNP = 0
            nStEM = 0
            nStGR = 0
        Else
            nStNP = rsSt.Fields(2).value
            nStEM = rsSt.Fields(1).value
            nStGR = rsSt.Fields(3).value
        End If
        nTot(C) = objSt.FetchTotalUnit_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt)
        nJA(C) = Round((objSt.FetchTotalJA_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt) / 60), 2) ' Dalam Jam
        nJF(C) = Round((objSt.FetchTotalJF_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt) / 60), 2) ' Dalam Jam
        If rsSt.RecordCount = 0 Or nTot(C) = 0 Then
            nRPS(C) = 0
        Else
            If nOpt = 0 Then
                If Val(rsSt.Fields(2).value) <> 0 Then
                    nRPS(C) = Round(nTot(C) / Val(rsSt.Fields(2).value), 2)
                Else
                    nRPS(C) = 0
                End If
            Else   'EM
                If Val(rsSt.Fields(1).value) <> 0 Then
                    nRPS(C) = Round(nTot(C) / Val(rsSt.Fields(1).value), 2)
                Else
                    nRPS(C) = 0
                End If
            End If
        End If
        If nTot(C) = 0 Or nJA(C) = 0 Then
            nLTr(C) = 0
        Else
            nLTr(C) = Round((nJA(C) * 60) / nTot(C), 2)
        End If
        If nJA(C) = 0 Or nJF(C) = 0 Then
            nOPS(C) = 0
        Else
            nOPS(C) = (Round(nJA(C) / nJF(C), 4) * 100)
        End If
        Select Case Weekday(DateSerial(Val(cboThnS.Text), Val(cboBlnS.Text), C))
        '1 = Minggu, 2 = Senin, 3 = Selasa, 4 = Rabu, 5 = Kamis, 6 = Jum’at, 7 = Sabtu
            Case 2, 3, 4, 5, 6
                nDay(C) = 7.25
            Case 7
                nDay(C) = 4.25
            Case 1
                nDay(C) = 0
        End Select
        If nOpt = 0 Then
            sSQLs = "Select * from LapNP_HO Where Tahun = " & Val(cboThnS.Text) & " And Bulan = " & Val(cboBlnS.Text) & " And Tanggal = " & C & ""
        Else
            sSQLs = "Select * from LapEM_HO Where Tahun = " & Val(cboThnS.Text) & " And Bulan = " & Val(cboBlnS.Text) & " And Tanggal = " & C & ""
        End If
        
        If objO.FetchData_Boolean(sSQLs) = True Then
            Call objSt.Update_Sumary_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt, nTot(C), nJA(C), nJF(C), nRPS(C), nLTr(C), nOPS(C), nDay(C), nStEM, nStNP, nStGR)
        Else
            Call objSt.AddNew_Sumary_NP_EM(Val(C), Val(cboBlnS.Text), Val(cboThnS.Text), nOpt, nTot(C), nJA(C), nJF(C), nRPS(C), nLTr(C), nOPS(C), nDay(C), nStEM, nStNP, nStGR)
        End If
    Next C
    PicBarEM.Visible = False
    DoEvents
    Call Print_NP_EM(nOpt, Val(cboThnS.Text), Val(cboBlnS.Text))
End Sub

Private Sub Print_NP_EM(isEM As Integer, isThn As Double, isBln As Double)
Dim oRs As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New Rpt_NP_EM_HO
        
        sSQL = "select Tahun,Bulan,Tanggal," & IIf(isEM = 0, "StallNP", "StallEM") & " as Stall,StallGR,Tot,jA,JF,RPS,LTr,OPS,Jam from " & IIf(isEM = 0, "LapNP_HO", "LapEM_HO") & " Where " & _
               "Tahun = " & isThn & " And Bulan  = " & isBln & " Order By Tanggal ASC"

    Set oRs = objO.FetchData(sSQL)
    Report.Database.SetDataSource oRs
    Report.txtPerusahaan.SetText gPerush
    Report.Text2.SetText "----oo000oo----     SUMMARY  HARIAN  KPI  " & IIf(isEM = 0, "NP", "EM") & "     ----oo000oo----"
    Report.txtAlamat.SetText gAlamat
    Report.Text13.SetText "Periode : " & isBln & " / " & isThn
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdPrintNEA_Click()
    If dt0.Text = "__/__/____" Or dt1.Text = "__/__/____" Then Exit Sub
    nUnit = 0
    nUnitR = 0
    nINV = 0
    nInvR = 0
    noinv = ""
    Dim rsList As ADodb.Recordset
    Set rsList = objO.FetchData_LapNEA(dt0, dt1)
    createTbl_NEA True
    If rsList.RecordCount > 0 Then
        PicNEA.Visible = True
        nNEA.Caption = "Please Wait ... Processing Data NEA ..."
        BarNEA.Max = rsList.RecordCount + 1
        BarNEA.Min = 0
        BarNEA.value = 0
        C = 0
        DoEvents
        Do Until rsList.EOF
                C = C + 1
                BarNEA.value = C
                DoEvents
                rsNEA.AddNew
                rsNEA!noinv = Trim(rsList.Fields(0).value)
                
                Select Case Right(Trim(rsList.Fields(0).value), 1)
                    Case "R"
                        nInvR = nInvR + 1
                    Case Else
                        nINV = nINV + 1
                End Select
                
                If noinv <> Replace(Replace(Trim(rsList.Fields(0).value), "A", ""), "B", "") Then
                    noinv = Replace(Replace(Trim(rsList.Fields(0).value), "A", ""), "B", "")
                    If Right(Replace(Replace(Trim(rsList.Fields(0).value), "A", ""), "B", ""), 1) = "R" Then
                        nUnitR = nUnitR + 1
                    Else
                        nUnit = nUnit + 1
                    End If
                End If
                
                rsNEA!Tgl_inv = Left(Trim(rsList.Fields(1).value), Len(Trim(rsList.Fields(1).value)) - 3)
                rsNEA!No_WO = Trim(rsList.Fields(2).value)
                rsNEA!Tgl_WO = Left(Trim(rsList.Fields(3).value), Len(Trim(rsList.Fields(3).value)) - 3)
                rsNEA!NO_POLISI = Trim(rsList.Fields(4).value)
                rsNEA!Type = Trim(rsList.Fields(5).value)
                rsNEA!Problem_Keluhan = Replace(Replace(Replace(Trim(rsList.Fields(6).value), ", , ", ""), ", ,", ""), ",, ", ", ")
                
                Dim rsAdd As ADodb.Recordset
                Set rsAdd = objO.FetchData("select Rep from JobADD2 Where NO_WO = '" & Trim(rsList.Fields(2).value) & "'")
                If rsAdd.RecordCount > 0 Then
                    Dim nADD As String
                    nADD = Trim(rsList.Fields(7).value)
                    rsAdd.MoveFirst
                    Do Until rsAdd.EOF
                        nADD = nADD & ", " & Trim(rsAdd.Fields(0).value)
                        rsAdd.MoveNext
                    Loop
                    rsNEA!Solusi_Masalah = nADD
                Else
                    rsNEA!Solusi_Masalah = Trim(rsList.Fields(7).value)
                End If
                
                Dim rsSbl As ADodb.Recordset
                Set rsSbl = objO.FetchData("select namabarang from JobSublet Where NO_WO = '" & Trim(rsList.Fields(2).value) & "'")
                If rsSbl.RecordCount > 0 Then
                    Dim nSbl As String
                    nSbl = rsNEA!Solusi_Masalah
                    rsSbl.MoveFirst
                    Do Until rsSbl.EOF
                        nSbl = nSbl & ", Sblt:" & Trim(rsSbl.Fields(0).value)
                        rsSbl.MoveNext
                    Loop
                    rsNEA!Solusi_Masalah = nSbl
                End If
                rsNEA!Solusi_Masalah = Replace(Replace(Replace(rsNEA!Solusi_Masalah, ", , ", ""), ", ,", ""), ",, ", ", ")
                rsNEA!PIC = Trim(rsList.Fields(8).value)
                rsNEA!MECH = Trim(rsList.Fields(14).value)
                rsNEA!Jasa = Trim(rsList.Fields(9).value)
                rsNEA!Part = Trim(rsList.Fields(10).value)
                rsNEA!SUBLET = Trim(rsList.Fields(11).value)
                rsNEA!Revenue = Trim(rsList.Fields(12).value)
                rsNEA!Remark = Trim(rsList.Fields(13).value)
                rsNEA!NamaCust = Trim(rsList.Fields(15).value)
                rsNEA.Update
            rsList.MoveNext
        Loop
        
        nNEA.Caption = "Please Wait ... Load Report NEA ..."
        C = C + 1
        BarNEA.value = C
        DoEvents
        PicNEA.Visible = False
        Call LoadReportNea
    Else
        MsgBox "Data Tidak Tersedia !", vbInformation + vbOKOnly, "Laporan NEA"
    End If
End Sub

Private Sub createTbl_NEA(Optional edit As Boolean = False)
    If rsNEA.State Then rsNEA.Close
    With rsNEA
        .Fields.Append "Noinv", adVarChar, 25
        .Fields.Append "Tgl_inv", adVarChar, 20
        .Fields.Append "No_WO", adVarChar, 25
        .Fields.Append "Tgl_WO", adVarChar, 20
        .Fields.Append "No_Polisi", adVarChar, 15
        .Fields.Append "TYPE", adVarChar, 100
        .Fields.Append "Problem_Keluhan", adVarChar, 750
        .Fields.Append "Solusi_Masalah", adVarChar, 1000
        .Fields.Append "PIC", adVarChar, 10
        .Fields.Append "Mech", adVarChar, 10
        .Fields.Append "JASA", adDouble
        .Fields.Append "PART", adDouble
        .Fields.Append "Sublet", adDouble
        .Fields.Append "Revenue", adDouble
        .Fields.Append "Remark", adVarChar, 100
        .Fields.Append "NamaCust", adVarChar, 200
        .Open
    End With
End Sub

Private Sub LoadReportNea()
    Screen.MousePointer = vbHourglass
    Set Report = New RptNEA
    Report.Database.SetDataSource rsNEA
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.Text13.SetText "Periode Invoice : " & dt0.Text & " - " & dt1.Text
    Report.txtKbg.SetText Log.LogKbgName
    
    Report.txtInv.SetText nINV
    Report.txtInvR.SetText nInvR
    Report.txtInvOK.SetText (nINV - nInvR)
    Report.txtUnit.SetText nUnit
    Report.txtunitR.SetText nUnitR
    Report.txtUnitOK.SetText (nUnit - nUnitR)
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub


Private Sub cmdPrintPS_Click()
    If dtP0.Text = "__/__/____" Or dtP1.Text = "__/__/____" Then Exit Sub
    nUnit = 0
    nUnitR = 0
    nINV = 0
    nInvR = 0
    noinv = ""
    Dim rsList As ADodb.Recordset
    Set rsList = objO.FetchData_LapPS(dtP0, dtP1)
    createTbl_PS True
    If rsList.RecordCount > 0 Then
        PicPS.Visible = True
        nPS.Caption = "Please Wait ... Processing Data Profit Sharing ..."
        BarPS.Max = rsList.RecordCount + 1
        BarPS.Min = 0
        BarPS.value = 0
        C = 0
        DoEvents
        Do Until rsList.EOF
                C = C + 1
                BarPS.value = C
                DoEvents
                rsPS.AddNew
        
                rsPS!NO_POLISI = Trim(rsList.Fields(0).value)
                rsPS!Nama_Pelanggan = Trim(rsList.Fields(1).value)
                rsPS!Type = Trim(rsList.Fields(2).value)
                rsPS!No_WO = Trim(rsList.Fields(3).value)
                rsPS!noinv = Trim(rsList.Fields(4).value)

                Select Case Right(Trim(rsList.Fields(4).value), 1)
                    Case "R"
                        nInvR = nInvR + 1
                    Case Else
                        nINV = nINV + 1
                End Select
                
                If noinv <> Replace(Replace(Trim(rsList.Fields(4).value), "A", ""), "B", "") Then
                    noinv = Replace(Replace(Trim(rsList.Fields(4).value), "A", ""), "B", "")
                    If Right(Replace(Replace(Trim(rsList.Fields(4).value), "A", ""), "B", ""), 1) = "R" Then
                        nUnitR = nUnitR + 1
                    Else
                        nUnit = nUnit + 1
                    End If
                End If
                
                rsPS!Tgl_inv = Left(Trim(rsList.Fields(5).value), (Len(Trim(rsList.Fields(5).value)) - 3))
                
                Dim rsAdd As ADodb.Recordset
                Set rsAdd = objO.FetchData("select Rep from JobADD2 Where NO_WO = '" & Trim(rsList.Fields(3).value) & "'")
                If rsAdd.RecordCount > 0 Then
                    Dim nADD As String
                    nADD = Trim(rsList.Fields(6).value)
                    rsAdd.MoveFirst
                    Do Until rsAdd.EOF
                        nADD = nADD & ", " & Trim(rsAdd.Fields(0).value)
                        rsAdd.MoveNext
                    Loop
                    rsPS!Solusi_Masalah = nADD
                Else
                    rsPS!Solusi_Masalah = Trim(rsList.Fields(6).value)
                End If
                
                Dim rsSbl As ADodb.Recordset
                Set rsSbl = objO.FetchData("select namabarang from JobSublet Where NO_WO = '" & Trim(rsList.Fields(3).value) & "'")
                If rsSbl.RecordCount > 0 Then
                    Dim nSbl As String
                    nSbl = rsPS!Solusi_Masalah
                    rsSbl.MoveFirst
                    Do Until rsSbl.EOF
                        nSbl = nSbl & ", Sblt:" & Trim(rsSbl.Fields(0).value)
                        rsSbl.MoveNext
                    Loop
                    rsPS!Solusi_Masalah = nSbl
                End If
                rsPS!Solusi_Masalah = Replace(Replace(Replace(rsPS!Solusi_Masalah, ", , ", ""), ", ,", ""), ",, ", ", ")

                rsPS!Tgl_masuk = Trim(rsList.Fields(7).value)
                rsPS!Tgl_Mulai = Trim(objO.FetchData_String("select top 1 Tanggal + Jam from progress Where noOrder = '" & Trim(rsList.Fields(3).value) & "' and Flag = 'S' Order By Tanggal, Jam ASC"))
                rsPS!Tgl_Mulai = Left(rsPS!Tgl_Mulai, Len(rsPS!Tgl_Mulai) - 3)
                rsPS!Tgl_Selesai = Trim(objO.FetchData_String("select top 1 Tanggal + Jam from progress Where noOrder = '" & Trim(rsList.Fields(3).value) & "' and Flag = 'F' Order By Tanggal, Jam DESC"))
                rsPS!Tgl_Selesai = Left(rsPS!Tgl_Selesai, Len(rsPS!Tgl_Selesai) - 3)
                rsPS!Jasa = Trim(rsList.Fields(8).value)
                rsPS!Part = Trim(rsList.Fields(9).value)
                rsPS!SUBLET = Trim(rsList.Fields(10).value)
                rsPS!Revenue = Trim(rsList.Fields(11).value)
                rsPS.Update
            rsList.MoveNext
        Loop
        
        nPS.Caption = "Please Wait ... Load Report Profit Sharing ..."
        C = C + 1
        BarPS.value = C
        DoEvents
        PicPS.Visible = False
        Call LoadReportPS
    Else
        MsgBox "Data Tidak Tersedia !", vbInformation + vbOKOnly, "Laporan PS"
    End If
End Sub

Private Sub createTbl_PS(Optional edit As Boolean = False)
    If rsPS.State Then rsPS.Close
    With rsPS
        .Fields.Append "No_Polisi", adVarChar, 15
        .Fields.Append "Nama_Pelanggan", adVarChar, 200
        .Fields.Append "TYPE", adVarChar, 100
        .Fields.Append "No_WO", adVarChar, 25
        .Fields.Append "Noinv", adVarChar, 25
        .Fields.Append "Tgl_inv", adVarChar, 50
        .Fields.Append "Solusi_Masalah", adVarChar, 1000
        .Fields.Append "Tgl_masuk", adVarChar, 50
        .Fields.Append "Tgl_Mulai", adVarChar, 50
        .Fields.Append "Tgl_Selesai", adVarChar, 50
        .Fields.Append "JASA", adDouble
        .Fields.Append "PART", adDouble
        .Fields.Append "Sublet", adDouble
        .Fields.Append "Revenue", adDouble
        .Open
    End With
End Sub

Private Sub LoadReportPS()
    Screen.MousePointer = vbHourglass
    Set Report = New RptProfitSharing
    Report.Database.SetDataSource rsPS
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.Text13.SetText "Periode Invoice : " & dtP0.Text & " - " & dtP1.Text
    Report.txtKbg.SetText Log.LogKbgName
        
    Report.txtInv.SetText nINV
    Report.txtInvR.SetText nInvR
    Report.txtInvOK.SetText (nINV - nInvR)
    Report.txtUnit.SetText nUnit
    Report.txtunitR.SetText nUnitR
    Report.txtUnitOK.SetText (nUnit - nUnitR)
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdPrintP_Click()
    Screen.MousePointer = vbHourglass
    nWaitPart.Visible = True
    DoEvents
    Dim rsPart As ADodb.Recordset
    Set rsPart = objR.FetchLapPart(cboThnP.Text, cboBlnP.Text)
    Set Report = New RptBulPart
    Report.Database.SetDataSource rsPart
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
    Report.Text2.SetText "Periode Laporan : " & GetNamaBulan(cboBlnP.Text) & " " & cboThnP.Text
        
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    
    DoEvents
    cmdTutup.Visible = True
    Screen.MousePointer = vbDefault
    nWaitPart.Visible = False
End Sub

Private Sub createTbl_ServRate(Optional edit As Boolean = False)
On Error GoTo errIsExist
    If rsSR.State Then rsSR.Close
    With rsSR
        .Fields.Append "Tgl", adDate
        .Fields.Append "NoWO", adVarChar, 30
        .Fields.Append "NoPol", adVarChar, 30
        .Fields.Append "Job", adVarChar, 30
        .Fields.Append "NoRangka", adVarChar, 30
        .Fields.Append "Model", adVarChar, 75
        .Fields.Append "Kode", adVarChar, 30
        .Fields.Append "Nama", adVarChar, 75
        .Fields.Append "QtyOrderSA", adDouble
        .Fields.Append "QtySupply", adDouble
        .Fields.Append "FR_Li", adDouble
        .Fields.Append "FR", adVarChar, 10
        .Fields.Append "li_Supply", adDouble
        .Fields.Append "Li_BO", adDouble
        .Fields.Append "Li_Cancel", adDouble
        .Fields.Append "SR_Li", adDouble
        .Fields.Append "SR", adVarChar, 10
        .Fields.Append "FullSup", adDouble
        .Open
    End With
    Set GridFR.DataSource = rsSR
errIsExist:
End Sub

Private Sub cmdPrintSR_Click()
    If Log.isFill = 0 Then Exit Sub
    cexAll.value = 0
    Frame2.Visible = False
    FrRate.Visible = True
    FrRate.Height = 8415
    FrRate.Width = 14535
    FrRate.Left = 480
    FrRate.Top = 600
    DoEvents
    createTbl_ServRate True
    Dim rsR As ADodb.Recordset
    
    Set rsR = objO.FetchData("select A.[DAte] as Tanggal , A.No_WO as NOWO, A.No_polisi, A.No_Chassis, A.MDL + ' - ' + A.FMDL as Model , " & _
                             "B.KodeItem, B.Keterangan as NamaItem, B.Qty as Qty_OrderSA, (IsNull(B.QtyKirimBI,0) + IsNull(B.QtyKirimBL,0) + IsNull(B.QtyKeluar,0)) as Qty_Supply, " & _
                             "case B.Qty - (IsNull(B.QtyKirimBI,0) + IsNull(B.QtyKirimBL,0) + IsNull(B.QtyKeluar,0)) When 0 Then 100 Else Round((((IsNull(B.QtyKirimBI,0) + IsNull(B.QtyKirimBL,0) + IsNull(B.QtyKeluar,0))/B.Qty) * 100),2) END as FillRate_LineItem, " & _
                             "case (IsNull(B.QtyKirimBI,0) + IsNull(B.QtyKirimBL,0) + IsNull(B.QtyKeluar,0)) When 0 Then 0 Else 1 End as Li_supply, 0 as Li_BO, " & _
                             "case B.BO When 'CANCEL' Then 1 else 0 end as Li_Cancel,0,0, B.Nomor,A.FS_RC,A.PDS,A.TPSVC  " & _
                             "from T_invoice A Inner Join [ACC" & Trim(Right(gDb, 2)) & "-" & cboThnSR.Text & IIf(CInt(cboBlnSR.Text) < 10, "0" & cboBlnSR.Text, cboBlnSR.Text) & "]..SOORD B On A.No_WO = B.NoSO " & _
                             "Where Right(A.noInv,1) <> 'B' and A.Fd_retur = '01/01/1900' " & _
                             "and Month(A.[Date]) = " & CInt(cboBlnSR.Text) & " and Year(A.[Date]) = " & CInt(cboThnSR.Text) & " Order By A.[date] asc")
    
    Dim nWO1 As String
    Dim xWO As Double
    Dim nTgl1 As Date
    Dim nFrLi As Integer
    Dim isNewWO1 As Boolean
    Dim isNewTgl1 As Boolean
    nWO1 = ""
    xWO = 0
    nTgl1 = "01/01/1900"
    
    Dim sumOrdSA As Double
    Dim sumSuply As Double
    Dim sumLiSup As Double
    Dim sumLiBO As Double
    Dim sumLiCan As Double
    Dim SumFullSup As Double
    sumOrdSA = 0
    sumSuply = 0
    sumLiSup = 0
    sumLiBO = 0
    sumLiCan = 0
    SumFullSup = 0
    
    Dim nSupply As Double
    Dim nBO As Double
    Dim nCancel As Double
    Dim nFR1 As Double
    nSupply = 0
    nBO = 0
    nCancel = 0
    nFR1 = 100
    
    Dim fR100 As Double
    Dim fRALL As Double
    fR100 = 0
    fRALL = 0
    
    Dim sRSP As Double
    Dim sRBC As Double
    sRSP = 0
    sRBC = 0

    Dim fR100i As Double
    Dim fRALLi As Double
    fR100i = 0
    fRALLi = 0
    
    Dim sRSPi As Double
    Dim sRBCi As Double
    sRSPi = 0
    sRBCi = 0
    
    isNewWO1 = False
    isNewTgl1 = False
    If rsR.RecordCount > 0 Then
        rsR.MoveFirst
        Do Until rsR.EOF
            With rsSR
                .AddNew
                rsSR!Tgl = rsR.Fields(0).value
                If rsSR!Tgl <> nTgl1 Then
                    isNewTgl1 = True
                    If nTgl1 <> "01/01/1900" Then

                    rsSR!Tgl = nTgl1
                        rsSR!NoWo = xWO
                        rsSR!noPol = ""
                        rsSR!Job = ""
                        rsSR!noRangka = ""
                        rsSR!model = ""
                        rsSR!kode = ""
                        rsSR!Nama = ""
                        rsSR!QtyOrderSA = sumOrdSA
                        rsSR!QtySupply = sumSuply
                        rsSR!FR_Li = 0
                        rsSR!li_Supply = sumLiSup
                        rsSR!Li_BO = sumLiBO
                        rsSR!Li_Cancel = sumLiCan
                        rsSR!SR_Li = 0
        
                        'If fR100i <> 0 Then
                            rsSR!fr = Trim(Round((fR100i / fRALLi), 3) * 100)
                        'Else
                        '    rsSR!fr = ""
                        'End If
                        If sRSPi + sRBCi <> 0 Then
                            rsSR!SR = Trim(Round((sRSPi / (sRSPi + sRBCi)), 3) * 100)
                        Else
                            rsSR!SR = ""
                        End If
                        rsSR!FullSup = SumFullSup
                        .Update
                        xWO = 0
                        sumOrdSA = 0
                        sumSuply = 0
                        sumLiSup = 0
                        sumLiBO = 0
                        sumLiCan = 0
                        fR100i = 0
                        fRALLi = 0
                        sRSPi = 0
                        sRBCi = 0
                        SumFullSup = 0
                        
                        .AddNew
                        rsSR!Tgl = rsR.Fields(0).value
                    End If
                    nTgl1 = rsSR!Tgl
                Else
                    isNewTgl1 = False
                End If
                
                rsSR!NoWo = rsR.Fields(1).value
                If rsSR!NoWo <> nWO1 Then
                    xWO = xWO + 1
                    isNewWO1 = True
                    nWO1 = rsSR!NoWo
                    rsSR!noPol = rsR.Fields(2).value
                    If Trim(rsR!FS_RC) = "R001" Then
                        rsSR!Job = "SBI"
                    ElseIf Trim(rsR!FS_RC) = "" Then
                        If CInt(rsR!pds) = 1 Then
                            rsSR!Job = "PDS"
                        ElseIf Trim(rsR!TPSVC) = "W" Then
                            rsSR!Job = "WAR"
                        Else
                            rsSR!Job = "GRP"
                        End If
                    Else
                        If Left(Trim(rsR!FS_RC), 1) = "R" Then
                            rsSR!Job = "SBE"
                        Else
                            rsSR!Job = ""
                        End If
                    End If
                            
                    rsSR!noRangka = rsR.Fields(3).value
                    rsSR!model = rsR.Fields(4).value
                Else
                    isNewWO1 = False
                    rsSR!noPol = ""
                    rsSR!Job = ""
                    rsSR!noRangka = ""
                    rsSR!model = ""
                    rsSR!NoWo = ""
                End If
                rsSR!kode = rsR.Fields(5).value
                rsSR!Nama = rsR.Fields(6).value
                rsSR!QtyOrderSA = rsR.Fields(7).value
                sumOrdSA = sumOrdSA + 1 'rsR.Fields(7).value
                Dim rsBO1 As ADodb.Recordset
                Set rsBO1 = objO.FetchDataACC("Select Qty from BOFILE Where  KodeItem = '" & rsSR!kode & "' and (NoSO = '" & nWO1 & "' OR NoSO =(Select NoSPK From SOORD Where NOSO ='" & nWO1 & "' and KodeItem = '" & rsSR!kode & "')) and Nomor = '" & rsR!Nomor & "'")
                If rsBO1.RecordCount > 0 Then
                    rsSR!Li_BO = 1
                    rsSR!QtySupply = rsR.Fields(8).value - Val(rsBO1.Fields(0).value)
                Else
                    rsSR!Li_BO = 0
                    rsSR!QtySupply = rsR.Fields(8).value
                End If
                sumLiBO = sumLiBO + rsSR!Li_BO
                rsSR!FR_Li = IIf(rsSR!QtyOrderSA = rsSR!QtySupply, 100, 0)
                sumSuply = sumSuply + IIf(rsSR!QtySupply = 0, 0, 1) 'rsSR!QtySupply
                
                
                rsR.MoveNext
                If rsR.EOF = True Then
                    If nFR1 = 0 Then
                        rsSR!fr = "0"
                    Else
                        If rsSR!FR_Li = 100 And nFR1 = 100 Then
                            nFR1 = 100
                        Else
                            nFR1 = 0
                        End If
                    End If
                    rsSR!fr = Trim(nFR1)
                    nFR1 = 100
                Else
                    If rsR.Fields(1).value = nWO1 Then
                        If rsSR!FR_Li <> 100 Then nFR1 = 0
                    Else
                        If nFR1 = 0 Then
                            rsSR!fr = "0"
                        Else
                            If rsSR!FR_Li = 100 Then
                                nFR1 = 100
                            Else
                                nFR1 = 0
                            End If
                        End If
                        rsSR!fr = Trim(nFR1)
                        nFR1 = 100
                    End If
                End If
                rsR.MovePrevious
                
                
                SumFullSup = SumFullSup + IIf(rsSR!fr = 0, 0, 1)
                
                rsSR!li_Supply = rsR.Fields(10).value
                sRSP = sRSP + rsSR!li_Supply
                sumLiSup = sumLiSup + rsSR!li_Supply
                sRSPi = sRSPi + rsSR!li_Supply
                rsSR!Li_Cancel = rsR.Fields(12).value
                sRBC = sRBC + rsSR!Li_Cancel
                sumLiCan = sumLiCan + rsSR!Li_Cancel
                sRBCi = sRBCi + rsSR!Li_Cancel
                If isNewWO1 = False Then
                    nSupply = nSupply + rsSR!li_Supply
                    nBO = nBO + rsSR!Li_BO
                    nCancel = nCancel + rsSR!Li_Cancel
                Else
                    nSupply = rsSR!li_Supply
                    nBO = rsSR!Li_BO
                    nCancel = rsSR!Li_Cancel
                End If
                sRBC = sRBC + rsSR!Li_BO
                sRBCi = sRBCi + rsSR!Li_BO
                
                rsR.MoveNext
                If rsR.EOF = True Then
                    If nSupply = 0 Then
                        rsSR!SR = 0
                    Else
                        rsSR!SR = Trim(Round(((nSupply / (nSupply + nBO + nCancel)) * 100), 2))
                    End If
                Else
                    If rsR.Fields(1).value = nWO1 Then
                        rsSR!SR = ""
                    Else
                        If nSupply = 0 Then
                            rsSR!SR = ""
                        Else
                            rsSR!SR = Trim(Round(((nSupply / (nSupply + nBO + nCancel)) * 100), 2))
                        End If
                        nSupply = 0
                        nBO = 0
                        nCancel = 0
                    End If
                End If
                rsR.MovePrevious
                               
                If rsSR!li_Supply = 1 And rsSR!Li_BO = 0 And rsSR!Li_Cancel = 0 Then
                    rsSR!SR_Li = 100
                ElseIf rsSR!li_Supply = 1 Then
                    rsSR!SR_Li = Round(100 / (rsSR!li_Supply + rsSR!Li_BO + rsSR!Li_Cancel), 2)
                Else
                    rsSR!SR_Li = 0
                End If
                
                Select Case rsSR!fr
                    Case "100"
                        fR100 = fR100 + 1
                        fRALL = fRALL + 1
                        fR100i = fR100i + 1
                        fRALLi = fRALLi + 1
                    Case "0"
                        fRALL = fRALL + 1
                        fRALLi = fRALLi + 1
                End Select
                rsSR!FullSup = 0
                .Update
            End With
            rsR.MoveNext
        Loop
        If fR100 <> 0 Then
            GridFR.Columns(11).FooterText = Round(((fR100 / fRALL) * 100), 1)
            GridFR1.Columns(4).FooterText = Round(((fR100 / fRALL) * 100), 1)
        Else
            GridFR.Columns(11).FooterText = "0"
            GridFR1.Columns(4).FooterText = "0"
        End If
        If sRSP + sRBC <> 0 Then
            GridFR.Columns(16).FooterText = Round(((sRSP / (sRSP + sRBC)) * 100), 1)
            GridFR1.Columns(8).FooterText = Round(((sRSP / (sRSP + sRBC)) * 100), 1)
        Else
            GridFR.Columns(16).FooterText = "0"
            GridFR1.Columns(8).FooterText = "0"
        End If
        
        rsR.MoveLast
        rsSR.AddNew
        rsSR!Tgl = rsR.Fields(0).value
        rsSR!NoWo = xWO
        rsSR!noPol = ""
        rsSR!Job = ""
        rsSR!noRangka = ""
        rsSR!model = ""
        rsSR!kode = ""
        rsSR!Nama = ""
        rsSR!QtyOrderSA = sumOrdSA
        rsSR!QtySupply = sumSuply
        rsSR!FR_Li = 0
        rsSR!li_Supply = sumLiSup
        rsSR!Li_BO = sumLiBO
        rsSR!Li_Cancel = sumLiCan
        rsSR!SR_Li = 0

        If fR100i <> 0 Then
            rsSR!fr = Trim(Round(((fR100i / fRALLi) * 100), 1))
        Else
            rsSR!fr = ""
        End If
        If sRSPi + sRBCi <> 0 Then
            rsSR!SR = Trim(Round(((sRSPi / (sRSPi + sRBCi)) * 100), 1))
        Else
            rsSR!SR = ""
        End If
        rsSR!FullSup = SumFullSup
        rsSR.Update
        
        Set rsSR1 = rsSR.Clone
        Set GridFR1.DataSource = rsSR1
        rsSR1.filter = "Nama = ''"
    End If
End Sub

Private Sub GridFR_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    On Error Resume Next
    Dim RsNav   As New ADodb.Recordset
    Set RsNav = rsSR.Clone
    If Bookmark > 1 Then RsNav.Bookmark = Bookmark

    If RsNav.Fields(6).value <> "" Then
        RowStyle = "Normal0"
    Else
        RowStyle = "Normal1"
    End If

End Sub

Private Sub cmdExport_Click()
    On Error GoTo Err_Export
    If rsSR.RecordCount = 0 Then Exit Sub
    
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
    
    If cexAll.value = 0 Then
        GridFR.ExportToFile onVal, True
    Else
        GridFR1.ExportToFile onVal, True
    End If
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

Private Sub cmdRef_Click()
    If cexAll.value = 0 Then
        Frame2.Visible = False
    Else
        Frame2.Visible = True
    End If
End Sub


Private Sub cmdClose_Click()
    FrRate.Visible = False
End Sub



