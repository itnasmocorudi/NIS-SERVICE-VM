VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmOAddJob 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Penambahan Pekerjaan, Part dan Sublet"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   13080
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
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   435
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   13680
      TabIndex        =   61
      Top             =   0
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "&H000000FF&"
      Height          =   9375
      Left            =   360
      TabIndex        =   15
      Top             =   0
      Width           =   12615
      Begin VB.CommandButton cmdSublet 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Su&blet"
         Enabled         =   0   'False
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
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   3000
         Width           =   2415
      End
      Begin VB.CommandButton TblSimpan 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Simpan"
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
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   7920
         Width           =   2655
      End
      Begin VB.CommandButton TblKeluar 
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
         Height          =   495
         Left            =   9600
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   8520
         Width           =   2655
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   4515
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   8760
         Begin VB.CommandButton cmdStall 
            Caption         =   "&Stall"
            Height          =   615
            Left            =   7560
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin TrueOleDBGrid70.TDBGrid dGrid 
            Height          =   3135
            Left            =   120
            TabIndex        =   4
            Top             =   1320
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   5530
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "JOB ORDER"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "RATE"
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "GROUP"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "HARGA"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "MECH"
            Columns(4).DataField=   ""
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "ID PROGRAM"
            Columns(5).DataField=   ""
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "PROGRAM"
            Columns(6).DataField=   ""
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
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
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
            Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
            Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=8196"
            Splits(0)._ColumnProps(21)=   "Column(3).Visible=0"
            Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(23)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(24)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(25)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(26)=   "Column(4)._ColStyle=8196"
            Splits(0)._ColumnProps(27)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(28)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(29)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(5).Visible=0"
            Splits(0)._ColumnProps(33)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(34)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(35)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(36)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(37)=   "Column(6).Order=7"
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
            DeadAreaBackColor=   16777215
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=78,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.locked=-1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.locked=-1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.locked=-1"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   960
            TabIndex        =   2
            Top             =   1035
            Width           =   3450
            _ExtentX        =   6085
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
            Left            =   4455
            TabIndex        =   17
            Top             =   1035
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
            Index           =   9
            Left            =   960
            TabIndex        =   1
            Top             =   720
            Width           =   495
            _ExtentX        =   873
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
            Left            =   1485
            TabIndex        =   18
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
            Index           =   10
            Left            =   1635
            TabIndex        =   19
            Top             =   720
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   5040
            TabIndex        =   3
            Top             =   1035
            Width           =   795
            _ExtentX        =   1402
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
            Mask            =   ">CCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   270
            Index           =   0
            Left            =   5850
            TabIndex        =   20
            Top             =   795
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
            Index           =   13
            Left            =   6000
            TabIndex        =   21
            Top             =   1035
            Visible         =   0   'False
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
            Index           =   31
            Left            =   5040
            TabIndex        =   70
            Top             =   720
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
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
            Index           =   137
            Left            =   960
            TabIndex        =   0
            Top             =   390
            Width           =   1770
            _ExtentX        =   3122
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
         Begin MSMask.MaskEdBox MaskEdBox5 
            Height          =   270
            Index           =   1
            Left            =   2760
            TabIndex        =   80
            Top             =   397
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
            Index           =   138
            Left            =   3000
            TabIndex        =   81
            Top             =   390
            Width           =   4335
            _ExtentX        =   7646
            _ExtentY        =   503
            _Version        =   393216
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
            Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Program"
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
            Left            =   240
            TabIndex        =   82
            Top             =   435
            Width           =   600
         End
         Begin VB.Line Line9 
            BorderColor     =   &H000000FF&
            X1              =   6000
            X2              =   8640
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line8 
            BorderColor     =   &H000000FF&
            X1              =   6000
            X2              =   8640
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   2760
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   2760
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Order Tambahan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   3165
            TabIndex        =   63
            Top             =   75
            Width           =   2460
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   8775
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipe"
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
            Left            =   495
            TabIndex        =   24
            Top             =   765
            Width           =   300
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Job Order "
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
            Left            =   75
            TabIndex        =   23
            Top             =   1080
            Width           =   765
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "MCH"
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
            Left            =   4680
            TabIndex        =   22
            Top             =   1080
            Width           =   330
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
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
         Height          =   3855
         Left            =   9000
         TabIndex        =   28
         Top             =   3720
         Width           =   3525
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   21
            Left            =   1140
            TabIndex        =   29
            Top             =   480
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   22
            Left            =   1140
            TabIndex        =   30
            Top             =   795
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   405
            Index           =   27
            Left            =   1140
            TabIndex        =   31
            Top             =   2520
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   25
            Left            =   1140
            TabIndex        =   13
            Top             =   1695
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   23
            Left            =   1140
            TabIndex        =   32
            Top             =   1080
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   24
            Left            =   1140
            TabIndex        =   33
            Top             =   1395
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   28
            Left            =   1140
            TabIndex        =   34
            Top             =   3120
            Width           =   1440
            _ExtentX        =   2540
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
            Index           =   29
            Left            =   1140
            TabIndex        =   35
            Top             =   3480
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
            Index           =   26
            Left            =   1140
            TabIndex        =   14
            Top             =   1995
            Width           =   1860
            _ExtentX        =   3281
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
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   30
            Left            =   1920
            TabIndex        =   36
            Top             =   3480
            Visible         =   0   'False
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
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            X1              =   3000
            X2              =   3360
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            X1              =   3000
            X2              =   3360
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   480
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   480
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estimasi Tambahan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   540
            TabIndex        =   62
            Top             =   75
            Width           =   2385
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   3495
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Oli"
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
            Left            =   660
            TabIndex        =   44
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   390
            TabIndex        =   43
            Top             =   1440
            Width           =   450
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total"
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
            Left            =   195
            TabIndex        =   42
            Top             =   2565
            Width           =   645
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   540
            TabIndex        =   41
            Top             =   840
            Width           =   300
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Labour"
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
            TabIndex        =   40
            Top             =   525
            Width           =   495
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Material"
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
            Left            =   270
            TabIndex        =   39
            Top             =   1740
            Width           =   570
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Waktu Selesai"
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
            Left            =   180
            TabIndex        =   38
            Top             =   3120
            Width           =   660
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Body Rpr"
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
            Left            =   180
            TabIndex        =   37
            Top             =   2085
            Width           =   660
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   840
            X2              =   3120
            Y1              =   2400
            Y2              =   2400
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00BFE8EA&
         Height          =   3495
         Left            =   120
         TabIndex        =   25
         Top             =   5880
         Width           =   8775
         Begin VB.Frame FrameFill 
            BackColor       =   &H00C0FFFF&
            Height          =   3120
            Left            =   0
            TabIndex        =   72
            Top             =   360
            Visible         =   0   'False
            Width           =   8765
            Begin TrueOleDBGrid70.TDBGrid partGridFill 
               Height          =   2505
               Left            =   120
               TabIndex        =   73
               Top             =   600
               Width           =   8490
               _ExtentX        =   14975
               _ExtentY        =   4419
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "KODE PART"
               Columns(0).DataField=   "KODEITEM"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "NAMA PART"
               Columns(1).DataField=   "NAMAITEM"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "NAMA BY SA"
               Columns(2).DataField=   "NAMABYSA"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "SATUAN"
               Columns(3).DataField=   "SATITEM"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "QTY"
               Columns(4).DataField=   "QTY"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   4
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "OLI"
               Columns(5).DataField=   "OLI"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   6
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   688
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).DividerColor=   14215660
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=6"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2910"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2831"
               Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
               Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(6)=   "Column(1).Width=4154"
               Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=4075"
               Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
               Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
               Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
               Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(15)=   "Column(3).Width=2117"
               Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2037"
               Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8196"
               Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(20)=   "Column(4).Width=1244"
               Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1164"
               Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
               Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(25)=   "Column(5).Width=873"
               Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=794"
               Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=8193"
               Splits(0)._ColumnProps(29)=   "Column(5).Order=6"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               Appearance      =   0
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
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.locked=-1"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.locked=-1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13,.locked=-1"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
               _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13,.locked=-1"
               _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
               _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
               _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
               _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=2,.locked=-1"
               _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
               _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
               _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
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
            Begin MSMask.MaskEdBox Text 
               Height          =   285
               Index           =   119
               Left            =   960
               TabIndex        =   74
               Top             =   195
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
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
            Begin MSMask.MaskEdBox MaskEdBox17 
               Height          =   270
               Left            =   3900
               TabIndex        =   75
               Top             =   195
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
               Index           =   120
               Left            =   4080
               TabIndex        =   76
               Top             =   195
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
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
            Begin VB.Label Label29 
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
               TabIndex        =   77
               Top             =   240
               Width           =   765
            End
         End
         Begin TrueOleDBGrid70.TDBGrid partGrid 
            Height          =   2535
            Left            =   120
            TabIndex        =   12
            Top             =   840
            Width           =   8535
            _ExtentX        =   15055
            _ExtentY        =   4471
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE PART"
            Columns(0).DataField=   "KODEITEM"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NAMA PART"
            Columns(1).DataField=   "NAMAITEM"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "SATUAN"
            Columns(2).DataField=   "SATITEM"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "QTY"
            Columns(3).DataField=   "QTY"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   4
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "GANTI"
            Columns(4).DataField=   "GANTI"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "OLI"
            Columns(5).DataField=   "OLI"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
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
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=8196"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=8196"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(25)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(5)._ColStyle=8196"
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
            MultiSelect     =   2
            DeadAreaBackColor=   16777215
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.locked=-1"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.locked=-1"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13,.locked=-1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.locked=-1"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   960
            TabIndex        =   10
            Top             =   480
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
         Begin MSMask.MaskEdBox MaskEdBox11 
            Height          =   270
            Left            =   3900
            TabIndex        =   26
            Top             =   435
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
            Left            =   4005
            TabIndex        =   11
            Top             =   435
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
         Begin VB.Line Line13 
            BorderColor     =   &H000000FF&
            X1              =   6000
            X2              =   8640
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line12 
            BorderColor     =   &H000000FF&
            X1              =   6000
            X2              =   8640
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Line Line11 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   2760
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Line Line10 
            BorderColor     =   &H000000FF&
            X1              =   120
            X2              =   2760
            Y1              =   120
            Y2              =   120
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Tambahan"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   3480
            TabIndex        =   64
            Top             =   75
            Width           =   1845
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   8775
         End
         Begin VB.Label Label42 
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
            TabIndex        =   27
            Top             =   480
            Width           =   765
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1095
         TabIndex        =   45
         Top             =   120
         Width           =   3510
         _ExtentX        =   6191
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1095
         TabIndex        =   46
         Top             =   480
         Width           =   2190
         _ExtentX        =   3863
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
         Index           =   3
         Left            =   1095
         TabIndex        =   47
         Top             =   840
         Width           =   3510
         _ExtentX        =   6191
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
         Index           =   2
         Left            =   3270
         TabIndex        =   48
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   5505
         TabIndex        =   49
         Top             =   120
         Width           =   3375
         _ExtentX        =   5953
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
         Index           =   5
         Left            =   5505
         TabIndex        =   50
         Top             =   480
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   5505
         TabIndex        =   51
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox Text 
         Height          =   525
         Index           =   8
         Left            =   9975
         TabIndex        =   52
         Top             =   120
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   926
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   240
         Index           =   16
         Left            =   9000
         TabIndex        =   5
         Top             =   1680
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   423
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
         Height          =   240
         Index           =   17
         Left            =   9000
         TabIndex        =   6
         Top             =   1920
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   423
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
         Height          =   240
         Index           =   18
         Left            =   9000
         TabIndex        =   7
         Top             =   2130
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   423
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
         Height          =   240
         Index           =   19
         Left            =   9000
         TabIndex        =   8
         Top             =   2355
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   423
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
         Height          =   240
         Index           =   20
         Left            =   9000
         TabIndex        =   9
         Top             =   2580
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   423
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
         Index           =   7
         Left            =   7665
         TabIndex        =   53
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   65
         Left            =   10200
         TabIndex        =   78
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal WO"
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
         Left            =   9165
         TabIndex        =   79
         Top             =   840
         Width           =   885
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         X1              =   12120
         X2              =   12360
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         X1              =   12120
         X2              =   12360
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000FF&
         X1              =   9120
         X2              =   9360
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         X1              =   9120
         X2              =   9360
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   9440
         TabIndex        =   65
         Top             =   1395
         Width           =   2610
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   9000
         Shape           =   4  'Rounded Rectangle
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Booking"
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
         Left            =   7080
         TabIndex        =   59
         Top             =   885
         Width           =   555
      End
      Begin VB.Label Label3 
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
         Left            =   9225
         TabIndex        =   58
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
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
         Left            =   4920
         TabIndex        =   57
         Top             =   885
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   4740
         TabIndex        =   56
         Top             =   120
         Width           =   690
      End
      Begin VB.Label Label1 
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
         Left            =   540
         TabIndex        =   55
         Top             =   525
         Width           =   420
      End
      Begin VB.Label Label51 
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
         Left            =   120
         TabIndex        =   54
         Top             =   165
         Width           =   840
      End
   End
   Begin MSMask.MaskEdBox Text 
      Height          =   285
      Index           =   32
      Left            =   4800
      TabIndex        =   71
      Top             =   3840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   503
      _Version        =   393216
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
      Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9375
      Left            =   0
      Picture         =   "frmOAddJob.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmOAddJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      :
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

'Dim allowAdd As Boolean
Dim rsJob As New ADODB.Recordset
Dim rsPart As New ADODB.Recordset
Dim strSql As String
Public rateLama As Double
Dim hNamaPart As String
Dim nStok As String
Dim hHrgPart As String
Dim hSatPart As String
Dim hOliPart As Boolean
Dim eachRate As Double
Dim allRate As Double
Dim manualSet As Boolean
Dim gPO As String
Dim ntRate As Double
Private NewId As Double
Private idPart As Double
Dim rsTemp As ADODB.Recordset
Dim support     As New clsFormDealer

''20200213 || pasing data detail program
''----------------------------------------------
Private namaProgram     As String
Private namaJob         As String
Private tipeJob         As String

Public Property Let isiNamaProgram(value As String)
    namaProgram = value
    Text(138) = NulSID(namaProgram)
End Property

Public Property Let isiNamaJob(value As String)
    namaJob = value
    Text(11) = NulSID(namaJob)
End Property

Public Property Let isiTipeJob(value As String)
    tipeJob = value
    Text(9) = NulSID(tipeJob)
End Property

Private Property Let isIdPart(value As Double)
    idPart = value + 1
End Property

Public Property Let isNewID(value As Double)
    NewId = value + 1
End Property

Private Sub clearForm(potition As String)
    Select Case UCase(potition)
        Case "UP"
            For i = 0 To 8
                Text(i) = ""
            Next
            Text(12) = ""
        Case "DOWN"
            For i = 9 To 31
                Select Case i
                    Case 13, 15, 120, 21, 22, 23, 24, 25, 26, 27
                        Text(i) = "0"
                    Case 28
                        Text(i) = ">>->>->>>>"
                    Case 29, 30
                        Text(i) = ">>:>>"
                    Case 12
                    Case Else
                        Text(i) = ""
                End Select
            Next
    End Select
End Sub

Private Function createTableJob() As Boolean
1         On Error GoTo vb_error
          
2         If rsJob.State Then rsJob.Close
3         Set rsJob = Nothing
4         With rsJob
5             .Fields.Append "JOB", adVarChar, 50
6             .Fields.Append "RATE", adDouble
7             .Fields.Append "GROUP", adVarChar, 5
8             .Fields.Append "HARGA", adDouble
9             .Fields.Append "MECH", adVarChar, 5
10            .Fields.Append "TGL", adDate
              ''20200212 || menambahkan fungsi hide show object sesuai flag
              ''----------------------------------------------------------------------------
11            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
12                .Fields.Append "IDPROGRAM", adVarChar, 20
13                .Fields.Append "PROGRAM", adVarChar, 50
14            End If
15            .Open
16        End With
17        If rsJob.State Then createTableJob = True
18        Exit Function
          
vb_error:
19        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createTablePart() As Boolean
1         On Error GoTo vb_error
          
2         If rsPart.State Then rsPart.Close
3         Set rsPart = Nothing
4         With rsPart
5             .Fields.Append "KODEITEM", adVarChar, 50
6             .Fields.Append "NAMAITEM", adVarChar, 200
7             .Fields.Append "SATITEM", adVarChar, 30
8             .Fields.Append "QTY", adInteger
9             .Fields.Append "GANTI", adBoolean
10            .Fields.Append "OLI", adBoolean
11            .Fields.Append "NAMABYSA", adVarChar, 100
12            .Fields.Append "nID", adDouble
13            .Open
14        End With
15        If rsPart.State Then createTablePart = True
16        Exit Function
          
vb_error:
17        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub headColumnJob()
   On Error GoTo headColumnJob_Error

    With dGrid
        .Columns(0).Caption = "Job Order"
        .Columns(0).DataField = "JOB"
        .Columns(0).Width = 4000
        .Columns(1).Caption = "Rate"
        .Columns(1).DataField = "RATE"
        .Columns(1).Width = 800
        .Columns(2).Caption = "Tipe"
        .Columns(2).DataField = "GROUP"
        .Columns(2).Width = 800
        .Columns(2).Visible = False
        .Columns(3).Caption = "Harga"
        .Columns(3).DataField = "HARGA"
        .Columns(3).Width = 800
        .Columns(3).Visible = False
        .Columns(4).Caption = "Mekanik"
        .Columns(4).DataField = "MECH"
        .Columns(4).Width = 800
        .Columns(4).Visible = False
'        .Columns(5).Caption = "Tanggal"
'        .Columns(5).DataField = "TGL"
'        .Columns(5).Width = 800
'        .Columns(5).Visible = False

        ''20200212 || menambahkan fungsi hide show object sesuai flag
        ''----------------------------------------------------------------------------
        If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
            .Columns(5).Caption = "IDPROGRAM"
            .Columns(5).DataField = "IDPROGRAM"
            .Columns(5).Width = 800
            .Columns(5).Visible = False
            
            .Columns(6).Caption = "PROGRAM"
            .Columns(6).DataField = "PROGRAM"
            .Columns(6).Width = 2000
            .Columns(6).Visible = True
        Else
            .Columns(5).Visible = False
            .Columns(6).Visible = False
        End If
    End With

   On Error GoTo 0
   Exit Sub

headColumnJob_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumnJob of Form frmOAddJob")
End Sub

Private Sub headColumnPart()
1        On Error GoTo headColumnPart_Error

2         With partGrid
3             .Columns(0).Caption = "KODE PART"
4             .Columns(0).DataField = "KODEITEM"
5             .Columns(0).Width = 1500
6             .Columns(1).Caption = "NAMA PART"
7             .Columns(1).DataField = "NAMAITEM"
8             .Columns(1).Width = 2500
9             .Columns(2).Caption = "SATUAN PART"
10            .Columns(2).DataField = "SATITEM"
11            .Columns(2).Width = 800
12            .Columns(3).Caption = "QTY"
13            .Columns(3).DataField = "QTY"
14            .Columns(3).Width = 800
15            .Columns(4).Caption = "GANTI"
16            .Columns(4).DataField = "GANTI"
17            .Columns(4).Width = 800
18            .Columns(5).Caption = "OLI"
19            .Columns(5).DataField = "OLI"
20            .Columns(5).Width = 0
21            .Columns(5).Visible = False
22        End With

23       On Error GoTo 0
24       Exit Sub

headColumnPart_Error:

25        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumnPart of Form frmOAddJob")
End Sub

Private Sub showItem(orderNumber As String, chassisNumber As String)
      Dim rsRecord As ADODB.Recordset
1        On Error GoTo showItem_Error
2         If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
3             If createTableJob Then
4                 Set rsRecord = objO.FetchData("SELECT a.rep,a.rte,a.klpkerja,a.offn,a.mekanik,a.date,a.id_program,case when isnull(b.program,'') <> '' then b.program when isnull(c.namafp,'') <> '' then c.namafp else '' end program FROM JOBDETAIL a WITH (NOLOCK) left join [" & dbReplikasiHO & "]..ho_program b with (nolock) on a.id_program = b.id_program and b.fs_log_del = '' left join [" & dbReplikasiHO & "]..ho_FIXprice C with (nolock) on c.id_fp = a.id_program and c.fs_log_del = '' WHERE A.NO_WO = '" & Trim(orderNumber) & "' AND A.NO_CHASSIS = '" & Trim(chassisNumber) & "' AND A.ADDJOB = 1")
5             End If
6         Else
7             Set rsRecord = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & Trim(orderNumber) & "' AND NO_CHASSIS = '" & Trim(chassisNumber) & "'")
8         End If
9         If Not rsRecord.EOF Then
10            If createTableJob Then
11                ntRate = 0
12                Do While Not rsRecord.EOF
13                    With rsJob
14                        .AddNew
15                        !job = rsRecord("REP")
16                        !Rate = CDbl(rsRecord("RTE"))
17                        !Group = rsRecord("KLPKERJA")
18                        !Harga = rsRecord("OFFN")
19                        !Mech = rsRecord("MEKANIK")
20                        !Tgl = rsRecord("DATE")
21                        If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
22                            !idProgram = NulSID(rsRecord!id_program)
23                            !program = NulSID(rsRecord!program)
24                        End If
25                        .Update
26                    End With
27                    ntRate = ntRate + rsJob("RATE")
28                    Text(21) = Text(21) + CDbl(rsRecord("OFFN"))
29                    rsRecord.MoveNext
30                Loop
31                rsRecord.MoveFirst
32                dGrid.DataSource = rsJob
33                Call headColumnJob
34            End If
35        End If
          
          ' ISI PART
36        If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
37            Set rsRecord = objO.FetchDataACC("select KodeItem as KODE_ITEM, Keterangan as NAMA_ITEM, Qty, 'N' as GANTI, 'N' as Bawa, '' as KDOLI, BO as SAT_ITEM, NamaBYSA, Nomor as nID, '' as PARTBAWA from SOORD WITH (NOLOCK) Where NoSO = '" & orderNumber & "' and AddJob = 1 and (isnull(BO,'null')='null' or isnull(bo,'') <> 'CANCEL')")
38        Else
              'Set rsTmp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI, SAT_ITEM FROM jobPart A LEFT JOIN T_OLI B ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE NO_chassis = '" & chassisNumber & "' AND no_wo = '" & orderNumber & "' AND ADDJOB = 1 Order By nID asc")
39            Set rsRecord = objO.FetchData("SELECT * FROM JOBPART A WITH (NOLOCK) LEFT JOIN T_OLI B WITH (NOLOCK) ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI WHERE NO_WO = '" & orderNumber & "' AND NO_CHASSIS = '" & chassisNumber & "' AND ADDJOB = 1 Order By nID asc")
40        End If
          'Set rsRecord = objO.FetchData("SELECT * FROM JOBPART A LEFT JOIN T_OLI B ON LEFT(A.KODE_ITEM,1) = B.FS_KD_OLI WHERE NO_WO = '" & orderNumber & "' AND NO_CHASSIS = '" & chassisNumber & "' AND ADDJOB = 1")
41        If Not rsRecord.EOF Then
42            If createTablePart Then
43                Do While Not rsRecord.EOF
44                    If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
                      
45                    Else
46                        If rsRecord("TIPE") = "P" Then
47                            Text(22) = Text(22) + (HargaJualPart(rsRecord("KODE_ITEM")) * rsRecord("QTY"))
48                        ElseIf rsRecord("TIPE") = "O" Then
49                            Text(23) = Text(23) + (HargaJualPart(rsRecord("KODE_ITEM")) * rsRecord("QTY"))
50                        End If
51                    End If
                      
                      ' INSERT PART
52                    With rsPart
53                        .AddNew
54                        !KodeItem = rsRecord("KODE_ITEM")
55                        !NAMAITEM = rsRecord("NAMA_ITEM")
56                        !SATITEM = rsRecord("SAT_ITEM")
57                        !qty = CDbl(rsRecord("QTY"))
58                        !ganti = IIf(rsRecord("GANTI") = "Y", True, False)
                          '!BAWA = IIf(rsRecord("BAWA") = "Y", True, False)
59                        If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
                          
60                        Else
61                            !oli = IIf(rsRecord("TIPE") = "P", False, True)
62                        End If
63                        !namabySA = rsRecord("NAMABYSA")
64                        !nID = Val(rsRecord("nID"))
65                        .Update
66                    End With
67                    rsRecord.MoveNext
68                Loop
69                Set partGrid.DataSource = rsPart
70                Set partGridFill.DataSource = rsPart
71                Call headColumnPart
72            End If
73        End If
          
          ' ISI SUBLET
74        Call objO.UpdateData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & Trim(orderNumber) & "' AND NO_CHASSIS = '" & Trim(chassisNumber) & "' AND KODEITEM = 'A'")
75        If Log.standarisasiSublet = 1 Then
76            Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, NO_WO, NO_INV, NO_CHASSIS, NOFAK, KODEITEM, NAMABARANG, KODESATUAN, QTY, HARGA, NONOTA, KETERANGAN, FLAG, NO_POLISI,PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,estwaktu,kdpromo,ADDJOB FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(orderNumber) & "' AND NO_CHASSIS = '" & Trim(chassisNumber) & "' AND KODEITEM = 'A'")
77        Else
78            Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, NO_WO, NO_INV, NO_CHASSIS, NOFAK, KODEITEM, NAMABARANG, KODESATUAN, QTY, HARGA, NONOTA, KETERANGAN, FLAG, NO_POLISI,PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(orderNumber) & "' AND NO_CHASSIS = '" & Trim(chassisNumber) & "' AND KODEITEM = 'A'")
79        End If
80        Text(24) = objO.FetchData_Double("SELECT ISNULL(SUM(QTY*HARGA),0) AS TOTAL FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & orderNumber & "' AND NO_CHASSIS = '" & chassisNumber & "' AND KODEITEM = 'A'")
          
          ' ISI KELUHAN
          Dim nNext As Integer
81        Set rsRecord = objO.FetchData("SELECT * FROM JOBADDHDR WITH (NOLOCK) WHERE NO_WO = '" & orderNumber & "' AND NO_CHASSIS = '" & chassisNumber & "'")
82        If Not rsRecord.EOF Then
83            nNext = 0
84            Do While Not rsRecord.EOF
85                Text(16 + nNext) = Trim(rsRecord("REQ"))
86                nNext = nNext + 1
87                rsRecord.MoveNext
88            Loop
89        End If
          
90        Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))

91       On Error GoTo 0
92       Exit Sub

showItem_Error:

93        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure showItem of Form frmOAddJob")
End Sub

Public Sub isiPart(iKdItem As String, iNmItem As String, iSatItem As String, iQty As Integer, iOli As Boolean, iNmBySA As String, iID As Double)
          
1        On Error GoTo isiPart_Error

2         With rsPart
3             If .State = adStateClosed Then createTablePart
4             If Not .BOF Then .MoveFirst
              '.Find "KODEITEM = '" & Trim(iKdItem) & "'", adSearchForward, 1
              'If .EOF Then
5                 .AddNew
6                 !KodeItem = iKdItem
7                 !NAMAITEM = iNmItem
8                 !SATITEM = iSatItem
9                 !qty = CDbl(iQty)
10                !ganti = True
11                !oli = iOli
12                !namabySA = iNmBySA
13                !nID = iID
14                .Update
              'End If
15            If iOli Then
16                Text(23) = Text(23) + (HargaJualPart(iKdItem) * CDbl(iQty))
17            Else
18                Text(22) = Text(22) + (HargaJualPart(iKdItem) * CDbl(iQty))
19            End If
20            Set partGrid.DataSource = rsPart
21            Set partGridFill.DataSource = rsPart
22            Call headColumnPart
23        End With
24        NewId = NewId + 1

25       On Error GoTo 0
26       Exit Sub

isiPart_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiPart of Form frmOAddJob")
End Sub

Private Sub cmdStall_Click()
      Dim pesan As String
      Dim nActRate As Double
      Dim rsRecord As ADODB.Recordset
          
          ' ISI TEMPSTALL
1        On Error GoTo cmdStall_Click_Error

2         Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                       & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND keterangan like '" & Replace(Replace(Text(8), " ", ""), ">", "") & " |%') AND " _
                       & "tanggal = '" & Format(isnow, "mm/dd/yyyy") & "' AND keterangan like '" & Replace(Replace(Text(8), " ", ""), ">", "") & " |%' ORDER BY noStall, Tanggal, JamMulai")
          
3         boleh = True
4         If allRate = 0 Then
5             pesan = pesan + "Job Order. "
6             boleh = False
7         End If
          
8         If rsJob.State = adStateClosed Then boleh = False
          
9         If boleh Then
10            nActRate = 0
11            rsJob.MoveFirst
12            Do While Not rsJob.EOF
                  
13               ' If Log.RPLHistori = 0 Then
14                    Set rsRecord = objO.FetchData("SELECT * FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsJob("JOB")) & "' AND MDL = '" & Trim(Text(1)) & "'")
15                'Else
16                '    Set rsRecord = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsJob("JOB")) & "' AND MDL = '" & Trim(text(1)) & "'")
17                'End If
18                If Not rsRecord.EOF Then
19                    nActRate = nActRate + CDbl(rsRecord("RATE"))
20                End If
21                rsJob.MoveNext
22            Loop
              
23            nActRate = nActRate + newRate1(Trim(Text(6)))
              
24            If nActRate > 0 Then
25            gStall = "addOrder"
26            With frmStall
27                .Text(7) = TglDDMMYYYY(isnow)
28                .Text(4).Mask = ""
29                .Text(4) = Replace(Text(8), ">", " ")
30                .Text(6) = Text(4)
31                .Text(8) = Text(12)
32                .Text(9) = nActRate
33                .Text(12) = Text(6)
34                .frmStal.Visible = False
35                .chkThs.value = frmOrder.chTHS.value
36                .chWar.value = IIf(frmOrder.opTpSvc(2).value, Checked, Unchecked)
37                WindowState = vbMaximized
38                TOChild frmOAddJob.hWnd, frmStall.hWnd
39                .WindowState = vbMaximized
40                .show
41            End With
42            Else
43                MsgBox "Belum mengisi Actual Rate untuk job order tersebut atau model tersebut !!!", vbCritical, "ACTUAL RATE"
44                Exit Sub
45            End If
46        Else
47            MsgBox "Lengkapi data Order !! " & vbCr & pesan, vbCritical, "Isi Field Kosong"
48            Exit Sub
49        End If

50       On Error GoTo 0
51       Exit Sub

cmdStall_Click_Error:

52        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdStall_Click of Form frmOAddJob")
End Sub

Private Function newRate1(nOrder As String) As Double
      Dim totalRate As Integer
      Dim nSelisih As Integer
          
          Dim rs As ADODB.Recordset
              
          ' UPDATE JAM TERPAKAI MEKANIK
1        On Error GoTo newRate1_Error

2         Set rs = objO.FetchData("SELECT SUM(DATEDIFF(n,START,PAUSE)) AS SELISIH  FROM TEMPPROGRESS WITH (NOLOCK) WHERE NOORDER = '" & nOrder & "' AND DATEDIFF(n,START,PAUSE) > 0")
3         If Not rs.EOF Then
4             nSelisih = IIf(rs("SELISIH") > 0, rs("SELISIH"), 0)
5         End If
6         totalRate = CDbl(takeActualRate(nOrder)) * 60
7         sisaRate = totalRate - CDbl(nSelisih)
8         If sisaRate > 0 Then
9             newRate1 = Round(CDbl(sisaRate) / 60, 1)
10        Else
11            newRate1 = Round(CDbl(totalRate) / 60, 1)
12        End If

13       On Error GoTo 0
14       Exit Function

newRate1_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure newRate1 of Form frmOAddJob")
End Function

Private Function takeActualRate(numberOrder As String) As Double

    strSql = "SELECT ACTRATE1+ACTRATE2+ACTRATE3+ACTRATE4+ACTRATE5+ACTRATE6+ACTRATE7+ACTRATE8 AS TOTaCT FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & numberOrder & "'"
    takeActualRate = objO.FetchData_Double(strSql)
End Function

Private Sub cmdSublet_Click()
Dim rsRecord As ADODB.Recordset
Dim subletValid As Boolean
Dim rsSublet As New ADODB.Recordset
    
   On Error GoTo cmdSublet_Click_Error

    subletValid = False
    Set rsSublet = objO.FetchData("SELECT DISTINCT VALID FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(6)) & "' AND no_chassis = '" & Trim(Text(0)) & "' ORDER BY VALID")
    If Not rsSublet.EOF Then
        If Trim(rsSublet("VALID")) = "1" Then
            subletValid = True
        End If
    End If

    If Not subletValid Then
        Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, no_wo, no_inv, no_chassis, nofak, kodeitem, namabarang, kodesatuan, qty, harga, nonota, keterangan, FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM jobsublet WITH (NOLOCK) " _
                     & "WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & Trim(Text(6)) & "' AND no_chassis = '" & Trim(Text(1)) & "') and no_wo = '" & Trim(Text(6)) & "' AND no_chassis = '" & Trim(Text(1)) & "' " _
                     & "ORDER BY nonota, kodeitem")
        If Log.standarisasiSublet = 1 Then
            With frmOSubletSentral
                gTag = "addorder"
                gPO = objO.FetchData_String("SELECT NOPO FROM JOBSUBLET WITH (NOLOCK) WHERE no_wo = '" & Trim(Text(64)) & "' AND no_chassis = '" & Text(15) & "'")
                .Text(0) = Text(0)
                .Text(1) = Text(1)
                .Text(2) = Text(2)
                .Text(3) = Text(3)
                .Text(4) = Text(4)
                .Text(5) = Text(5)
                .Text(6) = Trim(Text(6))
                .Text(7) = gPO
                .Text(14) = Text(8)
                .Text(15) = Text(7)
                .Text(65) = NulSID(Text(65))
                .cmdKeluar.Caption = "Kembali ke &Add Job"
                .chAdd = Checked
                .chAdd.Enabled = False
                .Text(13).Visible = False
                .Label11.Visible = False
                .show vbModal
            End With
        Else
            With frmOSublet
                gTag = "addorder"
                gPO = objO.FetchData_String("SELECT NoPo FROM JOBSUBLET WITH (NOLOCK) WHERE no_wo = '" & Trim(Text(6)) & "' AND no_chassis = '" & Trim(Text(1)) & "'")
                .Text(0) = Text(0)
                .Text(1) = Text(1)
                .Text(2) = Text(2)
                .Text(3) = Text(3)
                .Text(4) = Text(4)
                .Text(5) = Text(5)
                .Text(6) = Text(6)
                .Text(14) = Text(8)
                '.Text(119) = Text(8)
                .Text(7) = gPO
                .cmdKeluar.Caption = "Kembali ke &Add Job"
                .chAdd = Checked
                .chAdd.Enabled = False
                .Text(13).Visible = False
                .Label11.Visible = False
                .show vbModal
            End With
        End If
        
    Else
        MsgBox "Sublet sudah divalidasi, Lakukan retur sublet !!"
    End If

   On Error GoTo 0
   Exit Sub
cmdSublet_Click_Error:
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSublet_Click of Form frmOAddJob")
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    Unload Me
End Sub

Private Sub DGrid_DblClick()
1        On Error GoTo DGrid_DblClick_Error

2         If dGrid.Columns(0) <> "" Then
              ' TAMPILKAN RECORD KE TEXTBOX
3             rsJob.Find "JOB = '" & dGrid.Columns(0) & "'", , adSearchForward, 1
4             If Not rsJob.EOF Then
5                 Text(9) = rsJob("GROUP")
6                 Text(11) = rsJob("JOB")
7                 If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
8                     Text(137) = rsJob("idprogram")
9                     Text(138) = rsJob("program")
10                End If
                  'text(12) = rsJob("MECH")
                  Dim nWKT As String
11                If Text(30) <> ">>:>>" And Text(28) <> ">>->>->>>>" Then
12                    nWKT = hitungWaktuMundur(CDbl(rsJob("RATE")) * -1, Text(29), Format(Text(28), "DD/MM/YYYY"))
13                    Text(29) = Left(nWKT, InStr(nWKT, "|") - 1)
14                    Text(28) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
15                End If
                  'Text(29) = Format(Val((((CDbl(Left(Text(29), 2)) * 60) + CDbl(Right(Text(29), 2))) - (CDbl(rsJob("RATE")) * 60)) / 60), "00") & ":" & Format(Val((((CDbl(Left(Text(29), 2)) * 60) + CDbl(Right(Text(29), 2))) - (CDbl(rsJob("RATE")) * 60)) Mod 60), "00")
16            End If
              
              ' HAPUS RECORD
17            If Not rsJob.State = adStateClosed Then
18                myChassis = Trim(Text(0))
19                With rsJob
20                    If .RecordCount >= 1 Then
21                        rsJob.Find "JOB = '" & dGrid.Columns(0) & "'", , adSearchForward, 1
22                        If Not rsJob.EOF Then
23                            Text(21) = Format(CDbl(Text(21)) - Val(hitungLabour(CDbl(rsJob("RATE")), Trim(Text(2)))), "#0,0")
24                            .Delete adAffectCurrent
25                        End If
26                    Else
                          '.Delete adAffectCurrent
27                        Set rsJob = Nothing
28                        Set dGrid.DataSource = Nothing
29                    End If
                      '.Refresh
30                End With
31            End If
              
32            Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))
33        End If

34       On Error GoTo 0
35       Exit Sub

DGrid_DblClick_Error:

36        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure DGrid_DblClick of Form frmOAddJob")
End Sub

Private Sub dGrid_GotFocus()
      Dim allowInput As Boolean
          
1        On Error GoTo dGrid_GotFocus_Error

2         allowInput = True
        Dim rsSearch    As ADODB.Recordset
        Dim rsSearch2   As ADODB.Recordset
        Dim isFixPrice  As Boolean
        Dim hargaFP     As Double
        Dim actrate     As Double
        Dim isInvPart   As Boolean
        Dim isInvSublet As Boolean
        Dim isInsertPaket       As Boolean
        Dim isTerlibatPaket     As Boolean

           
3         isInvPart = False
4         isInvSublet = False
5         isTerlibatPaket = False
6         isInsertPaket = True
7         isFixPrice = False
8         If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
9              If NulSID(Text(6)) <> "" Then
10              sSQL = "SELECT * FROM jobdta WITH (NOLOCK) WHERE no_wo = '" & NulSID(Text(6)) & "' AND NOFK <> '' "
11              Set rsCekInv = objO.FetchData(sSQL)
12              If Not rsCekInv.EOF Then
13                  isInvPart = True
14              End If
                
15              sSQL = "SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & NulSID(Text(6)) & "' AND VALID = 1"
16              Set rsCekInv = objO.FetchData(sSQL)
17              If Not rsCekInv.EOF Then
18                  isInvSublet = True
19              End If
20            End If
              
              
21              If oRs2.State <> adStateClosed Then oRs2.Close
                
22              If Log.freepart = 1 Then
23                If tagihanFOP <> "" Then
24                    Set oRs2 = support.partFopPackage(Text(9), Text(11), Text(1))
25                Else
26                    Set oRs2 = support.partPackage(Text(9), Text(11), Text(1))
27                End If
28              Else
29                Set oRs2 = support.partPackage(Text(9), Text(11), Text(1))
30              End If
31              If Not oRs2.EOF Then
32                If isInvPart Then
33                    If MsgBox("Unit kendaraan tersebut telah diproses pembuatan supply slip." & vbCrLf & "Apakah yakin tidak membutuhkan penambahan paket part? ", vbYesNo, "Part Paket") = vbYes Then
34                        isInsertPaket = False
35                    Else
36                        MsgBox "Silahkan batalkan dulu proses sebelumnya untuk dapat melakukan penambahan paket Part / Sublet", vbInformation
37                        Exit Sub
38                    End If
39                End If
40                If isInvSublet Then
41                    If MsgBox("Unit kendaraan tersebut telah diproses validasi Sublet." & vbCrLf & " Apakah yakin tidak membutuhkan penambahan paket sublet? ", vbYesNo, "Sublet Paket Paket") = vbYes Then
42                        isInsertPaket = False
43                    Else
44                        MsgBox "Silahkan batalkan dulu proses sebelumnya untuk dapat melakukan penambahan paket Part / Sublet", vbInformation
45                        Exit Sub
46                    End If
47                End If
48              End If
49              If oRs2.State <> adStateClosed Then oRs2.Close
                
              
              ''20200213 cek jika fix price job dan tipe pekerjaan harus sama dengan ketentuan
              ''--------------------------------------------------------------------------------
50            If NulSID(Text(137)) <> "" Then
51                If Log.RPLHistori = 0 Then
52                    sSQL = "SELECT top 50 AA.* from (Select DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI) AND FS_LOG_DEL = '' AND FS_LOG_DEL = '' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END) AA  WHERE AA.ID_PROGRAM = '" & NulSID(Text(137)) & "'"
53                Else
54                    sSQL = "SELECT top 50 AA.* from (Select DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI,'' KLPKERJA,'' JOBORDER, 0 AS rte,0 AS price,0 as actRte,1 isgabungprogram "
55                    sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..HO_PROGRAM A WITH (NOLOCK)"
56                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM"
57                    sSQL = sSQL & " WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0"
58                    sSQL = sSQL & " OR '" & Format(isnow, "yyyy-mm-dd") & "' BETWEEN TGLMULAI AND TGLSELESAI)"
59                    sSQL = sSQL & " AND FS_LOG_DEL = '' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
60                    sSQL = sSQL & " Union"
61                    sSQL = sSQL & " select distinct A.ID_FP ID_PROGRAM, NAMAFP PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI,KLPKERJA,REP JOBORDER,rte,price,actRte,isgabungprogram"
62                    sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..HO_FIXPRICE A WITH (NOLOCK)"
63                    sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..HO_FIXPRICECBG B WITH (NOLOCK) ON A.ID_FP = B.ID_FP"
64                    sSQL = sSQL & " WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0"
65                    sSQL = sSQL & " OR '" & Format(isnow, "yyyy-mm-dd") & "' BETWEEN TGLMULAI AND TGLSELESAI)"
66                    sSQL = sSQL & " AND FS_LOG_DEL = '' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END) AA WHERE AA.ID_PROGRAM = '" & NulSID(Text(137)) & "' "
67                End If
68                Set rsSearch = objO.FetchData(sSQL)
69                If Not rsSearch.EOF Then
70                    If NulSID(rsSearch!jobOrder) <> "" And NulSID(rsSearch!jobOrder) <> NulSID(Text(11)) Then
71                        Text(11) = NulSID(rsSearch!jobOrder)
72                        MsgBox "Pekerjaan untuk tipe Program Fix Price salah.", vbInformation
73                        Exit Sub
74                    End If
75                    If NulSID(rsSearch!jobOrder) <> "" Then
                          ''20200213 menambahkan pengecekan apakah fix price dapat digabung depan program lain
                          ''------------------------------------------------------------------------------------------
76                        If NulSID(rsSearch!isgabungprogram) = 0 Then
                              ''CEK atas job order di addjob apakah ada program lain
                              ''---------------------------------------------------------------
77                            If rsJob.State = adStateClosed Then createTableJob
                              ' CARI YANG SAMA
78                            rsJob.Find "IDPROGRAM <> ''", , adSearchForward, 1
79                            If Not rsJob.EOF Then
80                                allowInput = False
81                                MsgBox "Kode program " & NulSID(Text(137)) & " tidak diperkenankan gabungan program.", vbInformation
82                                Text(9) = ""
83                                Text(10) = ""
84                                Text(11) = ""
85                                Text(137) = ""
86                                Text(138) = ""
87                                Exit Sub
88                            End If
89                            Set rsSearch2 = objO.FetchData("SELECT * FROM jobdetail WITH (NOLOCK) WHERE NO_WO = '" & NulSID(Text(6)) & "' AND NO_CHASSIS = '" & NulSID(Text(15)) & "' AND id_program <> ''")
90                            If Not rsSearch2.EOF Then
91                                MsgBox "Kode program " & NulSID(Text(137)) & " tidak diperkenankan gabungan program.", vbInformation
92                                Exit Sub
93                            End If
94                        End If
95                        isFixPrice = True
96                        Jrate = NulSID(rsSearch!rte)
97                        hargaFP = NulSID(rsSearch!Price)
98                        actrate = NulSID(rsSearch!actrte)
99                        fOther = False
100                   Else
                          ''20200213 menambahkan pengecekan apakah fix price dapat digabung depan program lain
                          ''------------------------------------------------------------------------------------------
                          Dim kodeFPGab   As String
                          Dim rsClone     As ADODB.Recordset
101                       Set rsClone = rsJob.Clone
102                       rsClone.filter = "IDPROGRAM <> ''"
103                       If Not rsClone.EOF Then
104                           rsClone.MoveFirst
105                           Do While Not rsClone.EOF
106                               kodeFPGab = ""
107                               kodeFPGab = objO.FetchData_String("SELECT ID_FP FROM [" & dbReplikasiHO & "]..HO_FIXPRICE WITH (NOLOCK) WHERE ID_FP = '" & NulSID(rsClone!idProgram) & "' and fs_log_del = '' AND isgabungprogram = 0 ")
108                               If kodeFPGab <> "" Then
109                                   MsgBox "Kode program " & kodeFPGab & " tidak diperkenankan gabungan program.", vbInformation
110                                   Exit Sub
111                               End If
112                               rsClone.MoveNext
113                           Loop
114                       End If
                          
115                       Set rsSearch2 = objO.FetchData("SELECT * FROM jobdetail WITH (NOLOCK) WHERE NO_WO = '" & NulSID(Text(6)) & "' AND NO_CHASSIS = '" & NulSID(Text(0)) & "' AND id_program <> '' and isFP = '1'")
116                       If Not rsSearch2.EOF Then
117                           rsSearch2.MoveFirst
118                           Do While Not rsSearch2.EOF
119                               kodeFPGab = ""
120                               kodeFPGab = objO.FetchData_String("SELECT ID_FP FROM [" & dbReplikasiHO & "]..HO_FIXPRICE WITH (NOLOCK) WHERE ID_FP = '" & NulSID(rsSearch2!id_program) & "' and fs_log_del = '' AND isgabungprogram = 0 ")
121                               If kodeFPGab <> "" Then
122                                   MsgBox "Kode program " & kodeFPGab & " tidak diperkenankan gabungan program.", vbInformation
123                                   Exit Sub
124                               End If
125                               rsSearch2.MoveNext
126                           Loop
127                       End If
128                   End If
129               Else
130                   MsgBox "Kode program " & NulSID(Text(137)) & " tidak ditemukan.", vbInformation
131                   Text(9) = ""
132                   Text(10) = ""
133                   Text(11) = ""
134                   Text(137) = ""
135                   Text(138) = ""
136                   Exit Sub
137               End If
138           End If
139       End If
          
140       If Trim(Text(9)) <> "" And Trim(Text(11)) <> "" And Trim(Text(12)) <> "" Then
141           If rsJob.State = adStateClosed Then createTableJob
              ' CARI YANG SAMA
142           rsJob.Find "JOB = '" & Trim(Text(11)) & "'", , adSearchForward, 1
143           If Not rsJob.EOF Then
144               allowInput = False
145               Text(9) = ""
146               Text(10) = ""
147               Text(11) = ""
148               Exit Sub
149           End If
              'Dim rsSearch As ADODB.Recordset
              Dim actRateLama As Double
              Dim totalPengerjaan As Double
              Dim waktuMulai As String
              Dim waktuSekarang As String
              'Dim waktuMulai As String
              Dim actSekarang As Double
              
150           totalPengerjaan = 0
151           Set rsSearch = objO.FetchData("select SUM(DATEDIFF(MINUTE,start,pause)) as selisih from tempprogress where flagp <> 'M' and pause <> '' and noorder = '" & Trim(Text(6)) & "'")
152           If Not rsSearch.EOF Then
153               rsSearch.MoveFirst
154               If Nul(rsSearch!selisih) <> "" Then
155                   totalPengerjaan = rsSearch!selisih
156                   totalPengerjaan = Round((totalPengerjaan / 60), 2)
157               End If
158           End If
              
159           actRateLama = 0
160           Set rsSearch = objO.FetchData("SELECT RATE FROM TEMPSTALL WHERE KETERANGAN LIKE '%" & Trim(Text(8)) & " |" & Trim(Text(6)) & "%'")
161           If Not rsSearch.EOF Then
162               rsSearch.MoveFirst
163               actRateLama = CDbl(rsSearch("RATE"))
164           End If
              'mencari rate + addjob - pengerjaan
165           totalPengerjaan = (actRateLama + CDbl(eachRate)) - totalPengerjaan
166           actSekarang = totalPengerjaan
              'mencari menit pengerjaan
167           totalPengerjaan = totalPengerjaan * 60
              
168           Set rsSearch = objO.FetchData("SELECT * FROM TEMPSTALL WHERE KETERANGAN LIKE '%" & Trim(Text(8)) & " |" & Trim(Text(6)) & "%'")
169           If Not rsSearch.EOF Then
170               rsSearch.MoveFirst
171               waktuMulai = Trim(rsSearch!jamMulai)
172               waktuSekarang = Trim(rsSearch!jamSelesai)
173               totalPengerjaan = totalPengerjaan + cariJumlahMenit(waktuMulai)
174               waktuSekarang = Format(Val(totalPengerjaan / 60), "0#") & ":" & Format((totalPengerjaan Mod 60), "0#")
175               If Log.newJpcb = 1 Then
176                   Set rsTemp = objO.FetchData("SELECT WAKTUMULAI,WAKTUSELESAI,WAKTUMULAIISTIRAHAT,WAKTUSELESAIISTIRAHAT FROM T_STALL WITH(NOLOCK) WHERE hari = '" & IIf(Weekday(Trim(Text(65))) = 1, 6, Weekday(Trim(Text(65))) - 2) & WeekdayName(IIf(Weekday(Trim(Text(65))) = 1, 8, Weekday(Trim(Text(65)))) - 1) & "'")
177                   If rsTemp.EOF = False Then
178                       rsTemp.MoveFirst
179                       If cariJumlahMenit(waktuMulai) <= cariJumlahMenit(rsTemp!waktuMulaiIstirahat) Then
180                           If (cariJumlahMenit(Trim(rsTemp!waktuMulaiIstirahat)) <= cariJumlahMenit(waktuSekarang)) Or cariJumlahMenit(waktuSekarang) >= cariJumlahMenit(rsTemp!waktuSelesaiIstirahat) Then
181                               totalPengerjaan = cariJumlahMenit(waktuSekarang) + (cariJumlahMenit(rsTemp!waktuSelesaiIstirahat) - cariJumlahMenit(rsTemp!waktuMulaiIstirahat))
182                           End If
183                       End If
                            
184                   End If
185                   waktuSekarang = Format(Val(totalPengerjaan / 60), "0#") & ":" & Format((totalPengerjaan Mod 60), "0#")
186               End If
      '
      '            Set rsTempSaveStall = objO.FetchData("SELECT * FROM Tempstall WITH(NOLOCK) WHERE keterangan = '' and ((jamMulai BETWEEN '" & waktuMulai & "' AND '" & waktuSekarang & "') OR (jamSelesai BETWEEN '" & waktuMulai & "' AND '" & waktuSekarang & "'))")
      '            If rsTempSaveStall.EOF Then
      '                Set rsTempSaveStall2 = objO.FetchData("SELECT * FROM tempstall WITH(NOLOCK) WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan LIKE '" & Left(Ket, InStr(Ket, "|") - 2) & " |%' ")
      '                If rsTempSaveStall2.EOF Then
      '                    Call objHT.InsertData_BengkelOnly("INSERT INTO tempstall VALUES ('" & stl & "','" & Format(Tgl, "mm/dd/yyyy") & "','" & jmMul & "','" & jmSel & "','" & Replace(CDbl(rte), ",", ".") & "','','" & mch & "','','" & Ket & "','" & Trim(adjust) & "')")     '" & grp & "
      '                Else
      '                    Call objHT.UpdateData_BengkelOnly("UPDATE tempstall SET noStall = '" & stl & "', jamMulai = '" & jmMul & "', jamSelesai = '" & jmSel & "', rate = '" & Replace(CDbl(rte), ",", ".") & "', KETERANGAN = '" & Ket & "',ADJUST = '" & Trim(adjust) & "' WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan LIKE '" & Left(Ket, InStr(Ket, "|") - 2) & " |%'  ")
      '                End If
      '            Else
      '                If Left(Trim(rsTempSaveStall("KETERANGAN")), InStr(rsTempSaveStall("KETERANGAN"), "|")) = Left(Ket, InStr(Ket, "|")) Then
      '                    Call objHT.UpdateData_BengkelOnly("UPDATE tempstall SET noStall = '" & stl & "', jamMulai = '" & jmMul & "', jamSelesai = '" & jmSel & "', mech = '" & Trim(Text(8)) & "', rate = '" & Replace(CDbl(rte), ",", ".") & "', keterangan = '" & Ket & "',ADJUST = '" & Trim(adjust) & "' WHERE tanggal = '" & Format(Tgl, "mm/dd/yyyy") & "' AND keterangan = '" & rsTempSaveStall("KETERANGAN") & "'  ")
      '                Else
      '                    MsgBox "Alokasi Stall dan Waktu sudah dipakai" & vbCr & "Kendaraan dengan No. Polisi : " & Trim(rsTempSaveStall!keterangan)
      '                    Exit Sub
      '                End If
      '            End If
187           Else
188               MsgBox "Unit tidak ditemukan distall, proses dibatalkan !", vbInformation
189               Exit Sub
190           End If
               
191           If Trim(Text(9)) = "SBE" Or Trim(Text(9)) = "SBI" Then
                  '' banding last sbe service dengan yg baru masuk
                  
192               If Log.RPLHistori = 0 Then
193                   Set rsSearch = objO.FetchData("SELECT FS_NILAI FROM T_REGULAR WITH (NOLOCK) WHERE FS_KETERANGAN = '" & Mid(Text(11), InStr(Text(11), "KALA ") + 5, InStr(Text(11), "KM") - (InStr(Text(11), "KALA ") + 5)) & "'")
194               Else
195                   Set rsSearch = objO.FetchData("SELECT FS_NILAI FROM [" & dbReplikasiHO & "]..T_REGULAR WITH (NOLOCK) WHERE FS_KETERANGAN = '" & Mid(Text(11), InStr(Text(11), "KALA ") + 5, InStr(Text(11), "KM") - (InStr(Text(11), "KALA ") + 5)) & "'")
196               End If
197               If Not rsSearch.EOF Then
198                   If CDbl(Text(32)) >= CDbl(rsSearch("FS_NILAI")) Then
199                       MsgBox "Service berkala yang dipilih tidak boleh lebih kecil atau sama dengan History sebelumnya !!", vbCritical, "Salah SBE"
200                       allowInput = False
201                   Else
202                       allowInput = True
203                   End If
204               End If
205           End If

206           If allowInput Then
                  ' TAMPIL BIAYA
207               myChassis = Trim(Text(0))
208               Text(21) = Format(CDbl(Text(21)) + Val(hitungLabour(CDbl(eachRate), Trim(Text(2)))), "#0,0")
                  
                  ' SIMPAN JOB ORDER
209               If rsJob.State = adStateClosed Then createTableJob
210               With rsJob
211                   .AddNew
212                   !job = Trim(Text(11))
213                   !Rate = CDbl(eachRate)
214                   !Group = Trim(Text(9))
215                   !Harga = Val(hitungLabour(eachRate, Trim(Text(2))))
216                   !Mech = Trim(Text(12))
217                   !Tgl = isnow
218                   If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
219                       sSQL = "SELECT PRICE FROM [" & dbReplikasiHO & "]..HO_FIXPRICE WITH (NOLOCK) WHERE ID_FP = '" & NulSID(Text(137)) & "' AND FS_LOG_DEL = '' "
220                       If objO.FetchData_Double(sSQL) > 0 Then
221                           !Harga = objO.FetchData_Double(sSQL)
222                       End If
223                       sSQL = "SELECT RTE FROM [" & dbReplikasiHO & "]..HO_FIXPRICE WITH (NOLOCK) WHERE ID_FP = '" & NulSID(Text(137)) & "' AND FS_LOG_DEL = '' "
224                       If objO.FetchData_Double(sSQL) > 0 Then
225                           !Rate = objO.FetchData_Double(sSQL)
226                       End If
227                       !idProgram = NulSID(Text(137))
228                       !program = NulSID(Text(138))
229                   End If
230                   .Update
231               End With
                  
                  ' CEK & INPUT PART PACKET
232               If rsPart.State = adStateClosed Then createTablePart
233               If isInsertPaket = False Then
234                 GoTo lanjutPartPaket
235               End If
236               Call cekPartPack(Text(9), Text(11), Text(1), 1, 2, , frmOAddJob, rsPart)
                  
lanjutPartPaket:
237               Set partGrid.DataSource = rsPart
238               Call headColumnPart
                  
                  ' HITUNG TOTAL RATE
239               If rsJob.State <> adStateClosed Then
240                   allRate = 0
241                   rsJob.MoveFirst
242                   Do While Not rsJob.EOF
243                       allRate = allRate + CDbl(rsJob("RATE"))
244                       rsJob.MoveNext
245                   Loop
246               End If
                  
                  ' HITUNG JAM ESTIMASI
                  Dim nWKT As String
247               If Text(30) <> ">>:>>" And Text(28) <> ">>->>->>>>" Then
248                   nWKT = hitungWaktuJamKerja(CDbl(allRate), Text(30), Format(Text(28), "DD/MM/YYYY"))
                      'nWKT = hitungWaktu(CDbl(allRate), text(30), Format(text(28), "DD/MM/YYYY"))
249                   Text(29) = Left(nWKT, InStr(nWKT, "|") - 1)
250                   Text(28) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
251               End If
                  
252               Text(9) = ""
253               Text(10) = ""
254               Text(11) = ""
255               Text(137) = ""
256               Text(138) = ""
257               Text(9).SelLength = Len(Text(9))
258               Text(9).SetFocus
259               dGrid.DataSource = rsJob
260               Call headColumnJob
261               Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))
262           End If
263       End If

264      On Error GoTo 0
265      Exit Sub

dGrid_GotFocus_Error:

266       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure dGrid_GotFocus of Form frmOAddJob")
End Sub

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then Call DGrid_DblClick
End Sub

Private Sub Form_Activate()
1        On Error GoTo Form_Activate_Error

2         FrameFill.Visible = False
3         Set rsJob = New ADODB.Recordset
         ''20200212 || menambahkan fungsi hide show object sesuai flag
         ''----------------------------------------------------------------------------
4        If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
5             Label31(1).Visible = True
6             Text(137).Visible = True
7             Text(138).Visible = True
8             MaskEdBox5(1).Visible = True
              dGrid.Columns(6).Visible = True
9        Else
10            Label31(1).Visible = False
11            Text(137).Visible = False
12            Text(138).Visible = False
13            MaskEdBox5(1).Visible = False
14       End If
15        If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then FrameFill.Visible = True
16        partGrid.DataSource = Nothing
17        dGrid.DataSource = Nothing
18        Set rsPart = New ADODB.Recordset
19        Call showItem(Trim(Text(6)), Trim(Text(0)))
20        Text(30) = Text(30)
21        manualSet = False

22       On Error GoTo 0
23       Exit Sub

Form_Activate_Error:

24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Activate of Form frmOAddJob")
End Sub

Private Sub Form_Load()
1        On Error GoTo Form_Load_Error

2         partGrid.DataSource = Nothing
3         Set rsPart = New ADODB.Recordset
4         Set rsJob = New ADODB.Recordset
          
5         partGrid.DataSource = rsPart
6         partGridFill.DataSource = rsPart
7         FrameFill.Visible = False
8         If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then FrameFill.Visible = True
9         Call clearForm("up")
10        Call clearForm("down")
         
11       If Log.isDSS = 1 Then
12          Set rsTemp = objO.FetchData("SELECT * FROM STALL WHERE KETERANGAN LIKE '%|" & Trim(Text(6)) & "|%'")
13          If Not rsTemp.EOF Then
14             rsTemp.MoveFirst
15             Text(12) = rsTemp!Mech
16          End If
17          Label28.Visible = False
18          Text(12).Visible = False
19          MaskEdBox5(0).Visible = False
            
20       End If
         
         
21       On Error GoTo 0
22       Exit Sub

Form_Load_Error:

23        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Load of Form frmOAddJob")
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub partGrid_DblClick()
          
1        On Error GoTo partGrid_DblClick_Error

2         If Trim(partGrid.Columns(0)) <> "" Then
3             If rsPart.State <> adStateClosed Then
                  ' TAMPILKAN DATA KE TEXTBOX
4                 rsPart.Find "KODEITEM = '" & partGrid.Columns(0) & "'", , adSearchForward, 1
5                 If Not rsPart.EOF Then
6                     Text(14) = rsPart("KODEITEM")
7                     Text(119) = rsPart("KODEITEM")
8                     Text(15) = CDbl(rsPart("QTY"))
9                     Text(120) = CDbl(rsPart("QTY"))
10                    hNamaPart = rsPart("NAMAITEM")
11                    hSatPart = rsPart("SATITEM")
12                    hOliPart = rsPart("OLI")
13                End If
                  
                  ' HAPUS RECORD
14                If Not rsPart.State = adStateClosed Then
15                    With rsPart
                          ' KURANGI HARGA
16                        If rsPart("OLI") Then
17                            Text(23) = CDbl(Text(23)) - (CDbl(rsPart("QTY")) * HargaJualPart(rsPart("KODEITEM")))
18                        Else
19                            Text(22) = CDbl(Text(22)) - (CDbl(rsPart("QTY")) * HargaJualPart(rsPart("KODEITEM")))
20                        End If
21                        If .RecordCount > 1 Then
22                            .Find "KODEITEM = '" & partGrid.Columns(0) & "'", , adSearchForward, 1
23                            If Not .EOF Then
24                                .Delete adAffectCurrent
25                            End If
26                        Else
                              '.Delete adAffectCurrent
27                            Set rsPart = Nothing
28                            Set partGrid.DataSource = Nothing
29                        End If
                          '.Refresh
30                    End With
31                    partGrid.DataSource = rsPart
32                End If
                  
33                Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))
34            End If
35        End If

36       On Error GoTo 0
37       Exit Sub

partGrid_DblClick_Error:

38        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure partGrid_DblClick of Form frmOAddJob")
End Sub

Private Sub partGridFill_DblClick()
1        On Error GoTo partGridFill_DblClick_Error

2         If Trim(Text(6)) <> "" Then
              Dim rsSO As ADODB.Recordset
3             Set rsSO = objO.FetchDataACC("Select NOSO from SOORD WITH (NOLOCK) Where NOSO = '" & Trim(Text(6)) & "' and KodeItem = '" & partGridFill.Columns(0) & "' and isNull(NamaBYSA,'') = '" & partGridFill.Columns(2) & "' and nomor = " & rsPart.Fields(7).value & " and ADDJOB = 1")
4             If rsSO.RecordCount > 0 Then
5                 MsgBox "Spare Part : " & partGridFill.Columns(0) & " " & partGridFill.Columns(2) & " Sejumlah : " & partGridFill.Columns(4) & Chr(13) & "Sudah disupply !, tidak dapat dihapus/dirubah !", vbInformation + vbOKOnly
6                 Exit Sub
7             End If
8         End If
9         If partGridFill.Columns(0) <> "" Or partGridFill.Columns(2) <> "" Then
10            Text(119) = IIf(partGridFill.Columns(0) <> "", partGridFill.Columns(0), partGridFill.Columns(2))
11            hNamaPart = partGridFill.Columns(1)
12            Text(120) = CDbl(partGridFill.Columns(4))
13                Set rsTemp = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Text(119) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
14                If Not rsTemp.EOF Then
15                    nStok = CDbl(rsTemp("QTYSTOCK"))
16                Else
17                    nStok = -1
18                End If
              
19            If Not rsPart.State = adStateClosed Then
20                With rsPart
21                    If .RecordCount > 1 Then
22                        With rsPart
                              ' KURANGI HARGA
23                            If rsPart("OLI") Then
24                                Text(23) = CDbl(Text(23)) - (CDbl(rsPart("QTY")) * HargaJualPart(rsPart("KODEITEM")))
25                            Else
26                                Text(22) = CDbl(Text(22)) - (CDbl(rsPart("QTY")) * HargaJualPart(rsPart("KODEITEM")))
27                            End If
28                            If .RecordCount > 1 Then
29                                .Find "KODEITEM = '" & partGrid.Columns(0) & "'", , adSearchForward, 1
30                                If Not .EOF Then
31                                    .Delete adAffectCurrent
32                                End If
33                            Else
                                  '.Delete adAffectCurrent
34                                Set rsPart = Nothing
35                                Set partGrid.DataSource = Nothing
36                            End If
                              '.Refresh
37                        End With
                          
38                    Else
39                        Set rsPart = Nothing
40                        Set partGridFill.DataSource = Nothing
41                    End If
42                End With
43                Set partGridFill.DataSource = rsPart
44                partGridFill.Refresh
45            End If
46            Text(119).SetFocus
47        End If

48       On Error GoTo 0
49       Exit Sub

partGridFill_DblClick_Error:

50        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure partGridFill_DblClick of Form frmOAddJob")
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSimpan_Click()
      Dim rsRecord As ADODB.Recordset
      Dim rsR As ADODB.Recordset
      Dim allowInput As Boolean
      Dim pesan As String
      Dim nActRate As Double
          
          'On Error GoTo vbError
1         allowInput = True
2         nActRate = hitungActRate(rsJob, Trim(Text(1))) '+ newRate1(Trim(text(6)))
          
          ' CHECK PEKERJAAN
3         If rsJob.State = adStateClosed Then
4             If MsgBox("Pekerjaan belum diisi. Tetap akan melanjutkan ? ", vbYesNo) = vbYes Then
5                 allowInput = True
6             Else
7                 allowInput = False
8             End If
9         End If
          
          ' CHECK STALL
10        Set rsRecord = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & Trim(Text(8)) & " |" & Trim(Text(6)) & "%'")
11        If Not rsRecord.EOF Then
12            If rsRecord("FLAG") <> "P" And rsRecord("FLAG") <> "F" Then
13                Set rsR = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE LEFT(KETERANGAN, CHARINDEX(' |', KETERANGAN)) NOT LIKE '%" & Trim(Text(8)) & "%' AND NOSTALL = '" & Trim(rsRecord("NOSTALL")) & "' AND TANGGAL = '" & Format(rsRecord("TANGGAL"), "MM/DD/YYYY") & "' AND JAMMULAI BETWEEN '" & Trim(rsRecord("JAMMULAI")) & "' AND '" & hitungWaktu(nActRate, Trim(rsRecord("JAMSELESAI")), Format(isnow, "DD-MM-YYYY")) & "'")
14                If Not rsR.EOF Then
15                    MsgBox "Atur Manual Posisi Stall karena bertabrakan dengan unit lain." & vbCr & "Pindahkan Unit berikutnya "
                      'frmStallProgress.Show 'vbModal
                      'frmStallProgress.WindowState = vbMaximized
16                    manualSet = True
17                    allowInput = False
18                End If
19            End If
20        End If
          
21        If allowInput Then
22            Call saveLogMenu(gSA, ComptName, gVersi, "frmOAddJob", "simpanJob", 0)
              
23            erTranAll = False
24            Call objHT.OpenTransactionAll
25            On Error GoTo vbError
              ' HAPUS ALOKASI MEKANIK LAMA
              Dim nRate As Double
26            nRate = 0
27            Set rsRecord = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(6)) & "'")
28            If Not rsRecord.EOF Then
29                Do While Not rsRecord.EOF
30                    nRate = nRate + CDbl(rsRecord("ACTRTE"))
31                    rsRecord.MoveNext
32                Loop
33                rsRecord.MoveFirst
34                Call createMekanik
                  'Call decresRestTime(rsRecord("MEKANIK"), CDbl(nRate) * 60, 1, Format(rsRecord("DATE"), "MM/DD/YYYY"))
35            End If
              
              ' SIMPAN PEKERJAAN TAMBAHAN
              Dim actRateLama As Double
              Dim iJob As Integer
              Dim isKirimData As Boolean
36            isKirimData = False
37            iJob = 0
38            sSQL = "select max([job_ID]) as konter from ReplBengkel99..P2_MonitorSVCProgress_Job where wo_number = '" & Right(gDb, 2) & "." & NulSID(Text(6)) & "'"
39            If objHT.FetchDataBengkel_String(sSQL) <> "" Then
40                iJob = CInt(Right(objHT.FetchDataBengkel_String(sSQL), 2))
41            End If
42            actRateLama = 0
43            Set rsRecord = objO.FetchData("SELECT ISNULL(SUM(ACTRTE),0) as actual FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(6)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "'")
44            If Not rsRecord.EOF Then
45                actRateLama = CDbl(rsRecord("ACTUAL"))
46            End If
                      
47            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
48                Call objHT.InsertData_Bengkel("DELETE FROM jobdetail WHERE NO_WO = '" & Trim(Text(6)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "' and addjob = 1")
49            End If
50            Call objO.DeleteData("DELETE JOBADD2 WHERE NO_WO = '" & Trim(Text(6)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "'")
              'Call objHT.UpdateData_BengkelOnly("UPDATE STALL SET RATE = RATE - " & Replace(CDbl(actRateLama), ",", ".") & " WHERE KETERANGAN LIKE '%" & Trim(text(8)) & " |" & Trim(text(6)) & "%' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' ") 'AND NOSTALL = '" & Trim(rsR("NOSTALL")) & "'
51            If rsJob.State <> adStateClosed Then
52                If Not rsJob.BOF Then
53                    rsJob.MoveFirst
54                    Do While Not rsJob.EOF
55                        sSQL = "select * from ReplBengkel99..P2_MonitorSVCProgress_Job where wo_number = '" & Right(gDb, 2) & "." & NulSID(Text(6)) & "' and job_name= '" & Trim(rsJob("JOB")) & "'"
56                        If objHT.FetchDataBengkel_String(sSQL) = "" Then
57                            isKirimData = True
58                            iJob = iJob + 1
59                            sSQL = "INSERT INTO ReplBengkel99..P2_MonitorSVCProgress_Job ([wo_number],[job_ID],[service_type],[job_name],[labor_est_price])"
60                            sSQL = sSQL & " VALUES ('" & Right(gDb, 2) & "." & NulSID(Text(6)) & "','" & NulSID(Text(6)) & "." & Format(iJob, "0#") & "','" & NulSID(rsJob("GROUP")) & "','" & NulSID(rsJob("JOB")) & "','" & CDbl(rsJob("HARGA")) & "')"
61                            Call objHT.InsertData_Bengkel(sSQL)
                               
62                        End If
63                        If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
64                            sSQL = "INSERT INTO JOBDETAIL (no_wo,no_chassis,Date,klpkerja,rep,rte,actrte,addJob,mekanik,"
65                            sSQL = sSQL & " grp,OFFn,disc,Time,VALID,FREEC,isFP,id_program,Price) VALUES ('" & Trim(Text(6)) & "', "
66                            sSQL = sSQL & " '" & Trim(Text(0)) & "','" & Format(rsJob("TGL"), "MM/DD/YYYY") & "', '" & Trim(rsJob("GROUP")) & "', '" & Trim(rsJob("JOB")) & "', '" & Replace(CDbl(rsJob("RATE")), ",", ".") & "', " & Replace(CDbl(eachActRate(Trim(rsJob("JOB")), Trim(Text(1)))), ",", ".") & ", "
67                            sSQL = sSQL & " 1,'" & Trim(rsJob("MECH")) & "','', " & CDbl(rsJob("HARGA")) & ",0, '" & Trim(Format(isnow, "HH:NN")) & "','','0','" & objO.FetchData_Double("SELECT 1  as FP FROM [" & dbReplikasiHO & "]..HO_FIXPRICE WITH (NOLOCK) WHERE fs_log_del = '' AND iD_FP = '" & NulSID(rsJob!idProgram) & "' ") & "','" & NulSID(rsJob!idProgram) & "'," & CDbl(rsJob("HARGA")) & ") "
68                            Call objHT.InsertData_Bengkel(sSQL)
69                        End If
70                        Call objO.InsertData("INSERT INTO JOBADD2 VALUES ('" & Trim(Text(6)) & "', '" & Trim(Text(0)) & "','" & Format(rsJob("TGL"), "MM/DD/YYYY") & "', '" & Trim(rsJob("GROUP")) & "', '" & Trim(rsJob("JOB")) & "', '" & Replace(CDbl(rsJob("RATE")), ",", ".") & "', " & Replace(CDbl(eachActRate(Trim(rsJob("JOB")), Trim(Text(1)))), ",", ".") & ", '" & Trim(rsJob("MECH")) & "','', " & CDbl(rsJob("HARGA")) & ",0, '" & Trim(Format(isnow, "HH:NN")) & "','','')")
71                        rsJob.MoveNext
72                    Loop
                      
                      ' UPDATE NILAI REGULAR CHECK
                      Dim rsTemp As ADODB.Recordset
                      Dim rsTemp2 As ADODB.Recordset
                      
73                    If Log.RPLHistori = 0 Then
74                        Set rsTemp = objO.FetchData("SELECT * FROM T_REGULAR WITH (NOLOCK) WHERE FS_RC <>''")
75                    Else
76                        Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_REGULAR WITH (NOLOCK) WHERE FS_RC <>''")
77                    End If
78                    If Not rsTemp.BOF Then
79                        rsTemp.MoveFirst
80                        Do While Not rsTemp.EOF
81                            Set rsTemp2 = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & Text(6) & "' AND REP LIKE '% " & Trim(rsTemp("FS_KETERANGAN")) & "[^0]%'")
82                            If Not rsTemp2.EOF Then
83                                sSQL = "UPDATE JOBDTA SET FS_RC ='" & Trim(rsTemp("FS_RC")) & "', SBE = 'SBE' where NO_WO ='" & Text(6) & "' "
84                                Call objHT.UpdateData_Bengkel(sSQL)
85                            End If
86                            rsTemp.MoveNext
87                        Loop
88                    End If

89                End If
                  
                  ' TAMBAH ALOKASI JAM MEKANIK
90                nMK = 0
91                LoopMk = False
                  'Call countRestTime(Trim(Text(12)), Round(CDbl(nActRate) * 60, 0), 1, Format(isnow, "MM/DD/YYYY"))
                  
                  ' SIMPAN STALL
                      ' STALL AUTO, DITAMBAHKAN DARI JAMSELESAI STALL, DENGAN RATE TAMBAHAN.
92                    Set rsR = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & Trim(Text(8)) & " |" & Trim(Text(6)) & "%' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "'")
93                    If Not rsR.EOF Then
94                        If NulSID(rsR!noStall) <> "STALL FINL" Then
95                            Call objHT.UpdateData_Bengkel("UPDATE STALL SET RATE = (RATE - " & Replace(CDbl(actRateLama), ",", ".") & ") + " & Replace(CDbl(nActRate), ",", ".") & ", JAMSELESAI = '" & Left(hitungWaktu(nActRate, rsR("JAMSELESAI"), Format(rsR("TANGGAL"), "DD-MM-YYYY")), 5) & "', FLAG = '" & IIf(rsR("FLAG") = "F", "P", rsR("FLAG")) & "' WHERE KETERANGAN LIKE '%" & Trim(Text(8)) & " |" & Trim(Text(6)) & "%' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "' AND NOSTALL = '" & Trim(rsR("NOSTALL")) & "'")
96                            If rsR("FLAG") = "F" Then
97                                Call objHT.InsertData_Bengkel("INSERT INTO PROGRESS VALUES ('" & Trim(Text(6)) & "','" & Trim(Text(0)) & "','S','" & Trim(rsR("NOSTALL")) & "','" & Format(isnow, "MM/DD/YYYY") & "', '" & Format(isnow, "HH:NN") & "','" & gSA & "','ADD JOB','" & Trim(rsR("MECH")) & "')")
98                                Call objHT.InsertData_Bengkel("INSERT INTO PROGRESS VALUES ('" & Trim(Text(6)) & "','" & Trim(Text(0)) & "','P','" & Trim(rsR("NOSTALL")) & "','" & Format(isnow, "MM/DD/YYYY") & "', '" & Format(DateAdd("n", 3, isnow), "HH:NN") & "','" & gSA & "','ADD JOB','" & Trim(rsR("MECH")) & "')")
99                                Call objHT.InsertData_Bengkel("INSERT INTO TEMPPROGRESS VALUES ('" & Trim(Text(6)) & "','" & Format(isnow, "MM/DD/YYYY") & "','" & Trim(rsR("NOSTALL")) & "','" & Format(isnow, "HH:NN") & "','P','" & Format(DateAdd("n", 3, isnow), "HH:NN") & "', '" & Trim(rsR("MECH")) & "','" & Trim(Text(31)) & "')")
100                           End If
101                       End If
102                   End If
          '        End If
103           End If
              
              ' SIMPAN PART TAMBAHAN
104           Call objHT.DeleteData_Bengkel("DELETE JOBPART WHERE NO_WO = '" & Trim(Text(6)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "' AND ADDJOB = 1")
105           If rsPart.State <> adStateClosed Then
106               If Not rsPart.BOF Then
107                   rsPart.MoveFirst
108                   Do While Not rsPart.EOF
109                       Call objHT.InsertData_Bengkel("INSERT INTO JOBPART VALUES ('" & Trim(Text(6)) & "', '" & Trim(Text(0)) & "','" & Trim(Text(7)) & "', '" & Trim(rsPart("KODEITEM")) & "','" & Trim(rsPart("NAMAITEM")) & "', '" & Trim(rsPart("SATITEM")) & "', '" & CDbl(rsPart("QTY")) & "', '" & Trim(IIf(rsPart("GANTI"), "Y", "N")) & "','N','',1,'',0, '" & Trim(rsPart("NAMABYSA")) & "'," & rsPart("nID") & ")")
110                       rsPart.MoveNext
111                   Loop
112                   rsPart.MoveFirst
113               End If
114           End If
115           If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then
116               Call SimpanPartSO_AddJob
117           End If
              
              ' SIMPAN SUBLET TAMBAHAN
118           gPO = IIf(gPO <> "", gPO, isiNoPoSublet)
              
119           Call objHT.DeleteData_Bengkel("DELETE JOBSUBLET WHERE NO_WO = '" & Trim(Text(6)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "' AND KODEITEM = 'A'")
120           If Log.standarisasiSublet = 1 Then
121               Call objHT.InsertData_Bengkel("INSERT INTO JOBSUBLET SELECT '" & gPO & "', NO_WO, NO_INV, NO_CHASSIS, NOFAK, KODEITEM, NAMABARANG, KODESATUAN, QTY, HARGA, NONOTA, KETERANGAN, Flag, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU,KDPROMO,ADDJOB FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(6)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "' AND KODEITEM = 'A'")
122           Else
123               Call objHT.InsertData_Bengkel("INSERT INTO JOBSUBLET SELECT '" & gPO & "', NO_WO, NO_INV, NO_CHASSIS, NOFAK, KODEITEM, NAMABARANG, KODESATUAN, QTY, HARGA, NONOTA, KETERANGAN, Flag, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM TEMPJOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(6)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "' AND KODEITEM = 'A'")
124           End If
              
                                      
              ' SIMPAN KELUHAN PELANGGAN
125           Call objHT.DeleteData_Bengkel("DELETE JOBADDHDR WHERE NO_WO = '" & Trim(Text(6)) & "' AND NO_CHASSIS = '" & Trim(Text(0)) & "'")
126           For i = 0 To 4
127               If Trim(Text(16 + i)) <> "" Then
128                   Call objHT.InsertData_Bengkel("INSERT INTO JOBADDHDR VALUES ('" & Trim(Text(6)) & "', '" & Trim(Text(0)) & "', '" & Format(isnow, "MM/DD/YYYY") & "', '" & Trim(Text(16 + i)) & "', '" & Format(isnow, "HH:NN") & "')")
129               End If
130           Next
              
131           If isKirimData Then
132               sSQL = "INSERT INTO [" & dbReplikasiHO & "]..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,status,proses,event_no,tglentry) VALUES "
133               sSQL = sSQL & " ('MONITORINGGR','" & Right(gDb, 2) & "." & NulSID(Text(6)) & "','" & Right(gDb, 2) & "." & NulSID(Text(6)) & "','','" & Right(gDb, 2) & "','Work Order','0','',GETDATE())"
134               Call objHT.InsertData_Bengkel(sSQL)
                    
135           End If
              
136           If erTranAll = False Then
137               Call objHT.CloseTransactionAll
138               Call saveLogMenu(gSA, ComptName, gVersi, "frmOAddJob", "simpanJob", 1)
139               With frmOrder
140                   .Text(62) = Trim(Text(28))
141                   .Text(63) = Trim(Text(29))
142               End With
143               MsgBox "Additional Job berhasil Disimpan "
                  ' CETAK FORM
144               Call cetakForm
145           Else
146               Call objHT.CloseTransactionAll_Err
147               Call saveLogMenu(gSA, ComptName, gVersi, "frmOAddJob", "simpanJob", 2)
148               MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
149           End If
150       Else
151           If pesan <> "" Then MsgBox pesan
152       End If
153       Exit Sub

vbError:
154       If erTranAll = True Then
155           Call objHT.CloseTransactionAll_Err
156           Call saveLogMenu(gSA, ComptName, gVersi, "frmOAddJob", "simpanJob", 2)
157           MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblSimpan_Click of Form frmOAddJob"), vbInformation, "Gagal Simpan"
158       End If
159       Exit Sub
End Sub

Private Sub SimpanPartSO_AddJob()
          Dim rsSOno As ADODB.Recordset
1        On Error GoTo SimpanPartSO_AddJob_Error

2         Set rsSOno = objO.FetchDataACC("Select top 1 Nomor from SOORD WITH (NOLOCK) WHERE NOSO = '" & Trim(Text(6)) & "' Order By Nomor DESC")
          
          Dim rsSO As ADODB.Recordset
3         Set rsSO = objO.FetchDataACC("Select KodeItem + NamaBySA + Convert(Varchar(10),Nomor) as Kode from SOORD WITH (NOLOCK) WHERE NOSO = '" & Trim(Text(6)) & "' AND IsNull(QtyKirimBL,0) + IsNull(QtyKirimBi,0) + IsNull(QtyKeluar,0) = 0 AND ADDJOB = 1 Order By Nomor DESC")
          
          Dim rsSOSetuju As ADODB.Recordset
4         Set rsSOSetuju = objO.FetchDataACC("Select KodeItem + NamaBySA + Convert(Varchar(10),Nomor) as Kode from SOORD WITH (NOLOCK) WHERE NOSO = '" & Trim(Text(6)) & "' AND IsNull(QtyKirimBL,0) + IsNull(QtyKirimBi,0) + IsNull(QtyKeluar,0) <> 0 AND ADDJOB = 1 Order By Nomor DESC")
          
          'Call objHT.DeleteData_ACC("DELETE SOORD WHERE NOSO = '" & Trim(Text(6)) & "' and isnull(QtyKirimBL,0) + isnull(QtyKirimBi,0) + IsNull(QtyKeluar,0) = 0  and ADDJOB = '1'")
          Dim idPart As Double
5         If Not rsPart.State = adStateClosed Then
6             If Not rsPart.EOF Then
7                 rsPart.MoveFirst
                  Dim nDelIndex As String
8                 nDelIndex = GetDelIndex("SV-")
                  Dim iNo As Integer
9                 rsPart.MoveFirst
10                If rsSOno.RecordCount > 0 Then
11                    iNo = CDbl(rsSOno.Fields(0).value + 1)
12                    Set rsSOno = Nothing
13                Else
14                    iNo = 0
15                End If
                  Dim isNoExist As ADODB.Recordset
                  Dim isTambah As Boolean
16                Do While Not rsPart.EOF
17                    Set isNoExist = Nothing
18                    If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then
19                        isTambah = False
20                        If rsSO.RecordCount > 0 Then
                              'pengecekan part yang belum divalidasi ada tidak
21                            rsSO.MoveFirst
22                            rsSO.Find "Kode = '" & rsPart("KODEITEM") + rsPart("NAMABYSA") + Trim(rsPart("nID")) & "'", , adSearchForward, 1
23                            If rsSO.EOF Then
                                  'jika pengecekan part yang belum divalidasi tidak ada maka dilanjutkan ke part validasi
24                                If rsSOSetuju.RecordCount > 0 Then
25                                    rsSOSetuju.MoveFirst
26                                    rsSOSetuju.Find "Kode = '" & rsPart("KODEITEM") + rsPart("NAMABYSA") + Trim(rsPart("nID")) & "'", , adSearchForward, 1
27                                    If rsSOSetuju.EOF Then
                                          'jika tidak ada maka beri penanda penambah data
28                                        isTambah = True
29                                    End If
30                                Else
31                                    isTambah = True
32                                End If
33                            End If
34                        ElseIf rsSOSetuju.RecordCount > 0 Then
35                            rsSOSetuju.MoveFirst
36                            rsSOSetuju.Find "Kode = '" & rsPart("KODEITEM") + rsPart("NAMABYSA") + Trim(rsPart("nID")) & "'", , adSearchForward, 1
37                            If rsSOSetuju.EOF Then
                                  'jika tidak ada maka beri penanda penambah data
38                                isTambah = True
39                            End If
40                        Else
41                            isTambah = True
42                        End If
                          
43                        If isTambah Then
44                            Call objHT.InsertData_ACC("Insert Into SOORD " & _
                                  "(NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB,TglPerlu,status ) " & _
                                  "Values ('" & Trim(Text(6)) & "', '" & rsPart("KODEITEM") & "', '" & rsPart("NAMAITEM") & "', '" & rsPart("QTY") & "','1', '" & rsPart("SATITEM") & "', '" & Trim(Text(7)) & "', '" & gSA & "', Getdate(),'" & iNo + 1 & "','" & rsPart("NAMABYSA") & "', '1',getdate(),'" & gSA & "')")
45                        End If
46                    End If
47                    rsPart.MoveNext
48                    iNo = iNo + 1
49                Loop
50                If Log.isFill = 1 And Log.startAppRate <= Format(Text(65), "dd/mm/yyyy") Then
                      ''2015-10-06 | APPRATE diganti bisa menit
51                    Call objHT.UpdateData_ACC("Update SOHDR Set TglPerlu = '" & Format((DateAdd(Left(Log.AppRate, 1), Mid(Log.AppRate, 2, Len(Log.AppRate)), isnow)), "MM/DD/YYYY hh:mm:ss") & "' Where NOSO = '" & Trim(Text(6)) & "'")
                      ''Call objHT.UpdateData_ACC("Update SOHDR Set TglPerlu = '" & Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), isnow)), "MM/DD/YYYY hh:mm:ss") & "' Where NOSO = '" & Trim(Text(6)) & "'")
52                End If
53            End If
54        End If

55       On Error GoTo 0
56       Exit Sub

SimpanPartSO_AddJob_Error:

57        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure SimpanPartSO_AddJob of Form frmOAddJob")
End Sub

Public Function GetDelIndex(ModilID As String) As String
          Dim acak As Double
1        On Error GoTo GetDelIndex_Error

2         Randomize Timer
3         acak = Int(Rnd * 99999999)
4         GetDelIndex = ModilID & acak

5        On Error GoTo 0
6        Exit Function

GetDelIndex_Error:

7         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GetDelIndex of Form frmOAddJob")
End Function

Private Sub cetakForm()
      Dim sHlp1 As Object
      Dim rpt As New rptAddOrder
      Dim rsRepo As ADODB.Recordset
      Dim crxSubReportReq As CRAXDRT.Report
          
1        On Error GoTo cetakForm_Error

2         If MsgBox("Mencetak Additional Job ?", vbYesNo, "Konfirmasi") = vbYes Then
          
3             Set rsRepo = objO.FetchData("SELECT DISTINCT A.NO_WO, A.NO_CHASSIS, A.DATE, A.TIME, KLPKERJA, REP, RTE, ACTRTE, MEKANIK, GRP, OFFN, A.DISC, " _
                                       & "B.CUSTOMER, B.CUST_ADRES, B.KOTA, B.NO_POLISI, B.FMDL, NOTE, B.COLOR, B.FS_KD_DEALER, B.DDATE, B.MDL, B.NO_ENGINE, " _
                                       & "D.FRONT , FS_NM_PETUGAS, ANCUST, B.TELPON, GSM, CDMA, B.FS_KD_PLG, D.DATE AS TGLMASUK, D.TIME AS JAMMASUK FROM JOBADD2 A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS " _
                                       & "INNER JOIN JOBDTA D ON A.NO_CHASSIS = D.NO_CHASSIS AND A.NO_WO = D.NO_WO INNER JOIN T_PETUGAS C ON D.FRONT = C.FS_KD_PETUGAS WHERE A.NO_CHASSIS = '" & Trim(Text(0)) & "' AND A.NO_WO = '" & Trim(Text(6)) & "'")
4             rpt.database.SetDataSource rsRepo
              
5             sSQL = "SELECT NO_WO, NO_CHASSIS, DATE, TIME, REQ  FROM JOBADDHDR WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(Text(0)) & "' AND NO_WO = '" & Trim(Text(6)) & "'"
6             Set sHlp1 = objO.FetchData(sSQL)
7             Set crxSubReportReq = rpt.Subreport1.OpenSubreport
8             crxSubReportReq.database.SetDataSource sHlp1
              '.Database.SetDataSource sHlp
              
9             rpt.txtPT.SetText gPerush
              
10            With CRViewer1
11                .ReportSource = rpt
12                .ViewReport
13                .Zoom 100
14                .Visible = True
15                .Top = 0
16                .Left = asxPowerBanner1.Left + asxPowerBanner1.Width
17                .Height = asxPowerBanner1.Height + 200
18                .Width = Me.Width - .Left
19            End With
20            cmdTutup.Visible = True
21            cmdTutup.Left = CRViewer1.Left + 20
22            cmdTutup.Top = 435
23            Screen.MousePointer = vbDefault
24        Else
25            Unload Me
26        End If

27       On Error GoTo 0
28       Exit Sub

cetakForm_Error:

29        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakForm of Form frmOAddJob")
End Sub

Private Function isiNoPoSublet() As String
      Dim rsHlp As ADODB.Recordset
          
1        On Error GoTo isiNoPoSublet_Error

2         Set rsHlp = objO.FetchData("SELECT DISTINCT NOPO FROM JOBSUBLET WITH (NOLOCK) WHERE LEFT(NOPO,4) = '" & Format(Right(year(isnow), 2), "00") & Format(Month(isnow), "00") & "' ORDER BY NOPO DESC")
3         If Not rsHlp.EOF Then
4             isiNoPoSublet = "GR." & Format(Right(year(isnow), 2), "00") & Format(Month(isnow), "00") & CDbl(Right(rsHlp("NOPO"), 5)) + 1
5         Else
6             isiNoPoSublet = "GR." & Format(Right(year(isnow), 2), "00") & Format(Month(isnow), "00") & Format(1, "00000")  'Format(Right(Year(isnow), 2), "00")
7         End If

8        On Error GoTo 0
9        Exit Function

isiNoPoSublet_Error:

10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure isiNoPoSublet of Form frmOAddJob")
End Function

Private Function eachActRate(jobOrder As String, carModel As String) As Double
      Dim rsSearch As ADODB.Recordset
          
1        On Error GoTo eachActRate_Error
          
2        ' If Log.RPLHistori = 0 Then
3             Set rsSearch = objO.FetchData("SELECT * FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(jobOrder) & "' AND MDL = '" & carModel & "'")
4        ' Else
5         '    Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(jobOrder) & "' AND MDL = '" & carModel & "'")
6         'End If
7         If Not rsSearch.EOF Then
8             eachActRate = CDbl(rsSearch("RATE"))
9         End If

10       On Error GoTo 0
11       Exit Function

eachActRate_Error:

12        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure eachActRate of Form frmOAddJob")
End Function

Private Function hitungActRate(rsJ As ADODB.Recordset, tipeModel As String) As Double
      Dim actRateAdd As Double


1        On Error GoTo hitungActRate_Error

2         If rsJ.State <> adStateClosed Then
3             actRateAdd = 0
4             If Not rsJ.BOF Then
5                 rsJ.MoveFirst
6                 Do While Not rsJ.EOF
7                     actRateAdd = actRateAdd + eachActRate(rsJ("JOB"), tipeModel) 'rsJ("RATE")
8                     rsJ.MoveNext
9                 Loop
10            End If
11            hitungActRate = actRateAdd
12        End If

13       On Error GoTo 0
14       Exit Function

hitungActRate_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hitungActRate of Form frmOAddJob")
End Function

Private Sub Text_Change(Index As Integer)
    If Index = 25 Or Index = 26 Or Index = 30 Then
        Select Case Index
            Case 25, 26
                If Text(0) <> "" Then
                    If Text(Index) = "" Then
                        Text(Index) = 0
                    ElseIf CDbl(Text(Index)) >= 0 Then
                        Text(Index) = Format(Text(Index), "#0,#")
                        Text(Index).SelStart = Len(Text(Index))
                        If Text(26) <> "" Then
                            Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))
                        End If
                    End If
                End If
            Case 30
                If Text(30) <> ">>:>>" And Text(28) <> ">>->>->>>>" Then
                    nWKT = hitungWaktu(CDbl(ntRate), Text(30), Format(Text(28), "dd-mm-yyyy"))
                    Text(29) = Left(nWKT, InStr(nWKT, "|") - 1)
                    Text(28) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
                End If
        End Select
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
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
   On Error GoTo Text_KeyDown_Error

    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 9
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCari.show (vbModal)
            Case 11
                If Trim(Text(9)) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOBLIST2"
                    gBantuan = Text(1)
                    gBantuan2 = Text(9)
                    frmCari.show (vbModal)
                Else
                    Text(9).SetFocus
                End If
            Case 12
                gMainKeyCari = Text(Index)
                gSearchOption = "MEKANIK"
                gBantuan = "ORDER"
                gBantuan2 = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
                gBantuan3 = Format(isnow, "DD/MM/YYYY")
                gBantuan4 = ""
                frmCari.show (vbModal)
            Case 14
                gMainKeyCari = Text(Index)
                gSearchOption = "PART"
                frmCariAcess.show (vbModal)
            Case 119
                gMainKeyCari = Text(Index)
                gSearchOption = "PART"
                frmCariAcess.show (vbModal)
            Case 137
                gMainKeyCari = Text(Index)
                gBantuan = "addjob"
                gSearchOption = "MULTIPLEPROGRAM"
                frmCari.show (vbModal)
        End Select
    End If

   On Error GoTo 0
   Exit Sub

Text_KeyDown_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Text_KeyDown of Form frmOAddJob")
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As ADODB.Recordset
      Dim rsRecord As ADODB.Recordset
      Dim searchSQL As String

1        On Error GoTo text_Validate_Error

2         Select Case Index
              Case 137
                  ''---------------------------------------------------------------
                  '' 20200210DEV
                  '' 20200213 | menambahkan validate pada textbox program
                  ''---------------------------------------------------------------
3                 If Text(Index) <> "" Then
4                     If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(65), "DD/MM/YYYY")) >= 0 Then
5                         If objO.FetchData_String("SELECT TPSVC FROM JOBDTA WITH (NOLOCK) WHERE NO_WO = '" & NulSID(Text(6)) & "' AND FS_LOG_DEL = '' ") = "W" Then
6                             MsgBox "Project multiple program bukan untuk tipe warranty.", vbInformation
7                             Exit Sub
8                         Else
9                             If Log.RPLHistori = 0 Then
10                                sSQL = "SELECT top 50 AA.* from (Select DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI FROM HO_PROGRAM A WITH (NOLOCK) LEFT JOIN HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI) AND FS_LOG_DEL = '' AND FS_LOG_DEL = '' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END) AA  WHERE AA.ID_PROGRAM = '" & NulSID(Text(137)) & "'"
11                            Else
12                                sSQL = "SELECT top 50 AA.* from (Select DISTINCT A.ID_PROGRAM, PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI,'' KLPKERJA,'' JOBORDER "
13                                sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..HO_PROGRAM A WITH (NOLOCK)"
14                                sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..HO_PROCBG B WITH (NOLOCK) ON A.ID_PROGRAM = B.ID_PROGRAM"
15                                sSQL = sSQL & " WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0"
16                                sSQL = sSQL & " OR '" & Format(isnow, "yyyy-mm-dd") & "' BETWEEN TGLMULAI AND TGLSELESAI)"
17                                sSQL = sSQL & " AND FS_LOG_DEL = '' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END"
18                                sSQL = sSQL & " Union"
19                                sSQL = sSQL & " select distinct A.ID_FP ID_PROGRAM, NAMAFP PROGRAM, KETERANGAN, TGLMULAI, TGLSELESAI,KLPKERJA,REP JOBORDER"
20                                sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..HO_FIXPRICE A WITH (NOLOCK)"
21                                sSQL = sSQL & " LEFT JOIN [" & dbReplikasiHO & "]..HO_FIXPRICECBG B WITH (NOLOCK) ON A.ID_FP = B.ID_FP"
22                                sSQL = sSQL & " WHERE (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0"
23                                sSQL = sSQL & " OR '" & Format(isnow, "yyyy-mm-dd") & "' BETWEEN TGLMULAI AND TGLSELESAI)"
24                                sSQL = sSQL & " AND FS_LOG_DEL = '' AND ISNULL(KODEDEALER,'') = CASE WHEN ISNULL(KODEDEALER,'') <> '' THEN '" & Trim(gKdCbg) & "' ELSE '' END) AA WHERE AA.ID_PROGRAM = '" & NulSID(Text(137)) & "' "
25                            End If
26                            Set rsTemp = objO.FetchData(sSQL)
27                            If rsTemp.EOF Then
28                               Text(Index) = ""
29                               Text(138) = ""
30                               Text(Index).SelLength = Len(Text(Index))
31                               Sendkeys ("+{TAB}")
32                               Text(9) = ""
33                               Text(10) = ""
34                               Text(11) = ""
35                               Exit Sub
36                            Else
37                               Text(138) = NulSID(rsTemp!program)
38                               Text(9) = NulSID(rsTemp!klpkerja)
39                               Text(11) = NulSID(rsTemp!jobOrder)
40                            End If
41                        End If
42                    Else
43                        MsgBox "Project multiple program belum diaktifkan.", vbInformation
44                        Exit Sub
45                    End If
46                End If
47            Case 9
48                If Text(Index) <> "" Then
                      
                      'If Log.RPLHistori = 0 Then
49                        searchSQL = "SELECT * FROM ACTRATE WITH (NOLOCK) WHERE MDL = '" & Trim(Text(1)) & "' AND JOB = '" & Trim(Text(Index)) & "'"
                      'Else
                      '    searchSQL = "SELECT * FROM [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) WHERE MDL = '" & Trim(text(1)) & "' AND JOB = '" & Trim(text(Index)) & "'"
                      'End If
50                    Set rsSearch = objO.FetchData(searchSQL)
51                    If rsSearch.EOF Then
52                        gMainKeyCari = Text(Index)
53                        gSearchOption = "JENIS_KERJA"
54                        frmCari.show (vbModal)
55                    Else
56                        Text(10) = rsSearch("REMARK")
57                    End If
58                End If
59            Case 11
60                If Text(Index) <> "" Then
61                    searchSQL = "SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE MDL = '" & Trim(Text(1)) & "' AND FS_ITEM = '" & Trim(Text(Index)) & "'"
62                    Set rsSearch = objO.FetchData(searchSQL)
63                    If rsSearch.EOF Then
64                        gMainKeyCari = Text(Index)
65                        gSearchOption = "JOBLIST2"
66                        gBantuan = Text(1)
67                        gBantuan2 = Text(9)
68                        frmCari.show (vbModal)
69                    Else
70                        eachRate = CDbl(rsSearch("RATE"))
71                    End If
72                End If
73            Case 12
74                If Text(Index) <> "" Then
75                    If Text(9) <> "" And Text(11) <> "" Then
76                        searchSQL = "SELECT isnull(MENITKERJA,0) AS MENITKERJA, ISNULL(TERPAKAI,0) AS TERPAKAI, ISNULL(ALOKASI, 0) AS ALOKASI, LEVMECH, [GROUP] AS FO FROM mekanik A WITH (NOLOCK) LEFT JOIN T_MEKANIK B WITH (NOLOCK) ON A.MECH = B.MECH AND B.HARI = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' AND B.TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' WHERE A.mech = '" & Text(Index) & "' "
77                        Set rsSearch = objO.FetchData(searchSQL)
78                        If rsSearch.EOF Then
79                            gMainKeyCari = Text(Index)
80                            gSearchOption = "MEKANIK"
81                            gBantuan = "ORDER"
82                            gBantuan2 = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
83                            gBantuan3 = Format(isnow, "DD/MM/YYYY")
84                            gBantuan4 = ""
85                            frmCari.show (vbModal)
86                        Else
87                            Text(31) = Trim(rsSearch("FO"))
88                        End If
89                    End If
90                End If
91            Case 14
92                If Text(Index) <> "" Then
93                        searchSQL = "SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem"
94                        Set rsSearch = objO.FetchDataACC(searchSQL)
95                        If rsSearch.EOF Then
96                            gMainKeyCari = Text(Index)
97                            gSearchOption = "PART"
98                            frmCariAcess.show (vbModal)
99                        Else
100                           hNamaPart = rsSearch("NAMABARANG")
101                           nStok = rsSearch("QTYSTOCK")
102                           hHrgPart = rsSearch("HARGAJUALA")
103                           hSatPart = rsSearch("KODESATUAN")
104                           Set rsRecord = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "' AND TIPE = 'O'")
105                           If Not rsRecord.EOF Then
106                               hOliPart = True
107                           Else
108                               hOliPart = False
109                           End If
110                       End If
                      'End If
111               End If
112           Case 119
113               If Text(Index) <> "" Then
114                       Set rsOpsi = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
115                       If rsOpsi.EOF Then
116                           hNamaPart = Text(Index)
117                           nStok = -1
118                           hHrgPart = 0
119                           hSatPart = ""
120                           Set rsTemp = objO.FetchData("SELECT Top 1 Tipe FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "' AND TIPE = 'O'")
121                           If Not rsTemp.EOF Then
122                               hOliPart = True
123                           Else
124                               hOliPart = False
125                           End If
126                       Else
127                           hNamaPart = rsOpsi("NAMABARANG")
128                           nStok = rsOpsi("QTYSTOCK")
129                           hHrgPart = rsOpsi("HARGAJUALA")
130                           hSatPart = rsOpsi("KODESATUAN")
131                           Set rsTemp = objO.FetchData("SELECT Top 1 Tipe FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "' AND TIPE = 'O'")
132                           If Not rsTemp.EOF Then
133                               hOliPart = True
134                           Else
135                               hOliPart = False
136                           End If
137                       End If
138               End If
                  
139           Case 15
140               If Text(Index) = "" Then
141                   Text(Index) = 0
142                   Sendkeys ("+{TAB}")
143                   Exit Sub
144               Else
145                   If hNamaPart <> "" Then
146                       If Text(Index) > 0 Then
147                           If nStok >= CDbl(Text(Index)) Then
                                  'ISI RS PART
148                               Call isiPart(Text(14), hNamaPart, hSatPart, CDbl(Text(Index)), hOliPart, "", NewId)
149                               Text(14) = ""
150                               Text(Index) = 0
151                               Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))
152                           Else
153                               MsgBox "Stok item tidak mencukupi !!" & vbCr & "Stok : " & nStok, vbCritical, "STOK ITEM"
154                               Text(Index) = nStok
155                               Text(Index).SelLength = Len(Text(Index))
156                               Sendkeys ("+{TAB}")
157                               Exit Sub
158                           End If
159                       End If
160                   Else
161                       Text(14).SetFocus
162                       Sendkeys ("+{TAB}")
163                       Exit Sub
164                   End If
165               End If
166           Case 120
167               If Trim(Text(Index)) = "" Then
168                   Text(Index) = 0
169                   Sendkeys ("+{TAB}")
170                   Exit Sub
171               Else
172                   If hNamaPart <> "" And Trim(Text(119)) <> "" Then
173                       If Text(Index) > 0 Then
174                           If nStok >= CDbl(Text(Index)) Then
                                  'ISI RS PART
175                               Call isiPart(Text(119), hNamaPart, hSatPart, CDbl(Text(Index)), hOliPart, "", NewId)
176                               Text(119) = ""
177                               Text(Index) = 0
178                               Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))
179                           Else
180                               If nStok <> -1 Then
181                                   If MsgBox("Stok item tidak mencukupi !!" & vbCr & "Stok : " & nStok & vbCr & "Tetap inputkan data ?", vbQuestion + vbYesNo, "STOK ITEM") = vbYes Then
182                                       Call isiPart(Text(119), hNamaPart, hSatPart, CDbl(Text(Index)), hOliPart, "", NewId)
183                                       Text(119) = ""
184                                       Text(Index) = 0
185                                       Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))
186                                   Else
187                                       Text(Index) = nStok
188                                       Text(Index).SelLength = Len(Text(Index))
189                                       Sendkeys ("+{TAB}")
190                                       Exit Sub
191                                   End If
192                               Else
193                                   If MsgBox("Stok item tidak terdaftar !!" & vbCr & "Inputkan data Nama Part By SA ?", vbQuestion + vbYesNo, "STOK ITEM") = vbYes Then
194                                       Call isiPart("", "", "ITEM", CDbl(Text(Index)), hOliPart, Trim(Text(119)), NewId)
195                                       Text(119) = ""
196                                       Text(Index) = 0
197                                       Text(27) = CDbl(Text(21)) + CDbl(Text(22)) + CDbl(Text(23)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(26))
198                                   Else
199                                       Text(Index) = "0"
200                                       Text(Index).SelLength = Len(Text(Index))
201                                       Sendkeys ("+{TAB}")
202                                       Exit Sub
203                                   End If
204                               End If
205                           End If
206                       End If
207                   Else
208                       Text(119).SetFocus
209                       Sendkeys ("+{TAB}")
210                       Exit Sub
211                   End If
212               End If
213       End Select

214      On Error GoTo 0
215      Exit Sub

text_Validate_Error:

216       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmOAddJob")
End Sub

Private Function cariJumlahMenit(jam As String) As Double
1     On Error GoTo vb_error
          Dim jJam As Integer
          Dim jMenit As Integer
2         jam = Trim(jam)
3         jJam = CInt(Left(jam, InStr(1, jam, ":") - 1))
4         jMenit = CInt(Right(jam, InStr(1, jam, ":") - 1))
          
5         cariJumlahMenit = CDbl(((jJam * 60) + (jMenit)))
6         Exit Function
vb_error:
7         MsgBox "Line Number " & err.Number & vbCrLf & " descripsi : " & err.Description & vbCrLf & "Fungsi cari Jumlah Menit", vbInformation
End Function
