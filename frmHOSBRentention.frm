VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form frmHOSBRentention 
   Caption         =   "Form1"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H000000FF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5115
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4920
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5115
      Width           =   1440
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H0080FF80&
      Caption         =   "&Export"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6405
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5115
      Visible         =   0   'False
      Width           =   1185
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   3000
      TabIndex        =   17
      Top             =   5040
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
   Begin VB.Frame frmeDealer 
      BackColor       =   &H0068ECE6&
      Caption         =   $"frmHOSBRentention.frx":0000
      Height          =   3615
      Left            =   6240
      TabIndex        =   31
      Top             =   1395
      Visible         =   0   'False
      Width           =   6615
      Begin TrueOleDBGrid70.TDBGrid GridDealer 
         Height          =   3015
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5318
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KODE DEALER"
         Columns(0).DataField=   "KODEDEALER"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "NAMA DEALER"
         Columns(1).DataField=   "NAMADEALER"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   4
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   "TERLIBAT"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   6876390
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
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H808080&"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF00&"
         _StyleDefs(21)  =   ":id=9,.borderColor=&H80000001&"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Named:id=33:Normal"
         _StyleDefs(50)  =   ":id=33,.parent=0"
         _StyleDefs(51)  =   "Named:id=34:Heading"
         _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   ":id=34,.wraptext=-1"
         _StyleDefs(54)  =   "Named:id=35:Footing"
         _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=36:Selected"
         _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=37:Caption"
         _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin AIFCmp1.asxPowerButton cmdTutupDealer 
         Height          =   360
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Tut&up"
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan SB Rate ke 1-9 By Dealer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   480
      TabIndex        =   18
      Top             =   2760
      Width           =   12375
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FF8080&
         Caption         =   "Download Non Retention Data "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1035
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   20
         Top             =   360
         Width           =   8415
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal delivery"
            Height          =   255
            Index           =   2
            Left            =   5520
            TabIndex        =   23
            Top             =   120
            Width           =   2415
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal invoice"
            Height          =   255
            Index           =   1
            Left            =   105
            TabIndex        =   22
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal retail"
            Height          =   255
            Index           =   0
            Left            =   2880
            TabIndex        =   21
            Top             =   120
            Width           =   2535
         End
      End
      Begin VB.CommandButton cmdVinDealer 
         BackColor       =   &H0080FF80&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1035
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker tglDealerAwal 
         Height          =   375
         Left            =   1410
         TabIndex        =   25
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   96927745
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker tglDealerAkhir 
         Height          =   375
         Left            =   3000
         TabIndex        =   26
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   96927745
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar progbarDealer 
         Height          =   615
         Left            =   480
         TabIndex        =   27
         Top             =   1560
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2880
         TabIndex        =   30
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode : "
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
         Left            =   480
         TabIndex        =   29
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Visible         =   0   'False
         Width           =   6255
      End
   End
   Begin VB.ComboBox cboPrinter 
      Height          =   315
      Left            =   5670
      TabIndex        =   15
      Top             =   3630
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton cmdPrintOut 
      BackColor       =   &H00FFFF00&
      Caption         =   "&Cetak"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8790
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Frame frameVinByVinRetention 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Laporan SB Rate ke 1-9 By Model"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   12375
      Begin VB.CheckBox chDealer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dealer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   34
         Top             =   1095
         Width           =   1815
      End
      Begin VB.CommandButton cmdProsesVin 
         BackColor       =   &H0080FF80&
         Caption         =   "Cetak"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   8415
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal retail"
            Height          =   255
            Index           =   6
            Left            =   2880
            TabIndex        =   4
            Top             =   120
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal invoice"
            Height          =   255
            Index           =   7
            Left            =   105
            TabIndex        =   3
            Top             =   120
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton opTipe 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Berdasarkan tanggal delivery"
            Height          =   255
            Index           =   8
            Left            =   5520
            TabIndex        =   2
            Top             =   120
            Width           =   2415
         End
      End
      Begin MSComCtl2.DTPicker tglAwalVin 
         Height          =   375
         Left            =   1410
         TabIndex        =   6
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   96927745
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker tglAkhirVin 
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   1050
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   96927745
         CurrentDate     =   39852
      End
      Begin MSComctlLib.ProgressBar ProgVin 
         Height          =   615
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Visible         =   0   'False
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   1085
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   1800
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode : "
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
         Left            =   480
         TabIndex        =   10
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   2880
         TabIndex        =   9
         Top             =   1080
         Width           =   60
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   9075
      Left            =   0
      Top             =   -600
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   16007
      StartColor      =   33023
      EndColor        =   12648447
      FormatString    =   "   Laporan Retention Rate"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   4194304
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   -30
      Width           =   18540
      _ExtentX        =   32703
      _ExtentY        =   503
      BackColor       =   255
      HighlightColor  =   255
      ShadowColor     =   65535
   End
End
Attribute VB_Name = "frmHOSBRentention"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim Report As Object
Private rsTemp As ADODB.Recordset
Private rsLaporan As New ADODB.Recordset

Dim rsGrid      As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset
Dim cols        As TrueOleDBGrid70.Columns
Dim dealerPilih As String
Dim cabangPilih As String
Dim i           As Integer

Dim nilaiSBSameOutletPareto(10)     As Double
Dim nilaiSBSameOutletNonPareto(10)  As Double
Dim nilaiSBSameOutletSouth(10)      As Double
Dim nilaiSBSameOutletNorth(10)      As Double

Dim nilaiSBNRMPareto(10)     As Double
Dim nilaiSBNRMNonPareto(10)  As Double
Dim nilaiSBNRMSouth(10)      As Double
Dim nilaiSBNRMNorth(10)      As Double

Dim nilaiTotalPareto(10)     As Double
Dim nilaiTotalNonPareto(10)  As Double
Dim nilaiTotalSouth(10)      As Double
Dim nilaiTotalNorth(10)      As Double

Dim nilaiDemandPareto(10)    As Double
Dim nilaiDemandNonPareto(10) As Double
Dim nilaiDemandSouth(10)     As Double
Dim nilaiDemandNorth(10)     As Double

Private nilaiTotal1(10)          As Double
Private nilaiTotal2(10)          As Double
Private nilaiTotal3(10)          As Double
Private nilaiTotal4(10)          As Double
Private nilaiTotal5(10)          As Double
Private nilaiTotal6(10)          As Double
Private nilaiTotal7(10)          As Double

Private Sub createRS()
    
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    With rsDealer
        .Fields.Append "KODEDEALER", adVarChar, 15
        .Fields.Append "NAMADEALER", adVarChar, 25
        .Fields.Append "KODECABANG", adVarChar, 15
        .Fields.Append "TERLIBAT", adBoolean
        .Open
    End With
    
    Set GridDealer.DataSource = rsDealer
End Sub

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 3500
    GridDealer.Columns(2).Width = 1000
    GridDealer.Columns(2).Alignment = dbgCenter
End Sub

Private Sub isiDealer()
      Dim rsTmp As New ADODB.Recordset

1         Call HeaderCol
2         Set rsTmp = objO.FetchDataHO("SELECT fs_kd_dealer, fs_nm_dealer,kodecabang FROM t_dealer a WITH (NOLOCK)  WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
3         If Not rsTmp.EOF Then
4             Call createRS
5             Do While Not rsTmp.EOF
6                 rsDealer.AddNew
7                 rsDealer!KodeDealer = rsTmp("FS_KD_DEALER")
8                 rsDealer!namaDealer = rsTmp("FS_NM_DEALER")
9                 rsDealer!kodeCabang = rsTmp("kodecabang")
10                rsDealer!terlibat = False
11                rsDealer.Update
12                rsTmp.MoveNext
13            Loop
              
14            GridDealer.DataSource = rsDealer
15            GridDealer.EvenRowStyle.BackColor = vbYellow
16            GridDealer.OddRowStyle.BackColor = vbWhite
17            GridDealer.AlternatingRowStyle = True
18        End If
End Sub

Private Sub chDealer_Click()
    If chAllDealer = Unchecked Then
        frmeDealer.Visible = True
        If rsDealer.State = adStateClosed Then
            Call isiDealer
        End If
        frmeDealer.Visible = True
        GridDealer.SetFocus
    ElseIf chAllDealer = Checked Then
        frmeDealer.Visible = False
    End If
End Sub

Private Sub cmdProsesVin_Click()
    If Month(tglAwalVin) <> Month(tglAkhirVin) Or year(tglAwalVin) <> year(tglAkhirVin) Then
        MsgBox "Range periode dalam 1 bulan", vbInformation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    dealerPilih = ""
    cabangPilih = ""
    If chDealer.value = Checked Then
        If rsDealer.State <> adStateClosed Then
            rsDealer.MoveFirst
            rsDealer.Find "TERLIBAT = '" & True & "'", adSearchForward, 1
            If rsDealer.EOF Then
                MsgBox "Dealer belum dipilih.", vbInformation
                Exit Sub
            Else
                rsDealer.MoveFirst
                Do While Not rsDealer.EOF
                    dealerPilih = dealerPilih & "'" & NulSID(rsDealer!KodeDealer) & "',"
                    cabangPilih = cabangPilih & "'" & NulSID(rsDealer!kodeCabang) & "',"
                    rsDealer.MoveNext
                Loop
                dealerPilih = Left(dealerPilih, Len(dealerPilih) - 1)
                cabangPilih = Left(cabangPilih, Len(cabangPilih) - 1)
                dealerPilih = " and t.fs_kd_dealer in (" & dealerPilih & ")"
                cabangPilih = " b.dealer in (" & cabangPilih & ") and "
            End If
        Else
            MsgBox "Dealer belum dipilih.", vbInformation
            Exit Sub
        End If
    End If
    For i = 0 To 10
        nilaiTotal1(i) = 0
        nilaiTotal2(i) = 0
        nilaiTotal3(i) = 0
        nilaiTotal4(i) = 0
        nilaiTotal5(i) = 0
        nilaiTotal6(i) = 0
        nilaiTotal7(i) = 0
    Next
    If opTipe(7).value = True Then
        If dealerPilih = "" Then
            Call ambilLaporanVinRetentionTglInvoice
        Else
            Call ambilLaporanVinRetentionTglInvoiceDealer
        End If
    ElseIf opTipe(6).value = True Then
        If dealerPilih = "" Then
            Call ambilLaporanVinRetentionTglRetail
        Else
            Call ambilLaporanVinRetentionTglRetailDealer
        End If
         
    ElseIf opTipe(8).value = True Then
        If dealerPilih = "" Then
            Call ambilLaporanVinRetentionTglDelivery
        Else
            Call ambilLaporanVinRetentionTglDeliveryDealer
        End If
    End If
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
End Sub

Private Sub cmdTutupDealer_Click()
    frmeDealer.Visible = False
End Sub
  

Private Sub cmdVinDealer_Click()
    If Month(tglAwalVin) <> Month(tglAkhirVin) Or year(tglAwalVin) <> year(tglAkhirVin) Then
        MsgBox "Range periode dalam 1 bulan", vbInformation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
     
    If opTipe(1).value = True Then
        Call ambilLaporanVinDealerTglInvoice
    ElseIf opTipe(0).value = True Then
        Call ambilLaporanVinDealerTglRetail
    ElseIf opTipe(2).value = True Then
        Call ambilLaporanVinDealerTglPDS
    End If
End Sub

Private Sub Form_Activate()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > (asxBars1.Height + 100) Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100
    End If
    If ScaleWidth > (asxPowerBanner1.Width + 100) Then
        CRViewer1.Width = ScaleWidth - asxPowerBanner1.Width - 100
    End If
    
End Sub

Private Sub Form_Load()
    
    tglAwalVin = isnow
    tglAkhirVin = isnow
    
    tglDealerAwal = isnow
    tglDealerAkhir = isnow
    
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdExport.Visible = False
'    XPButton1.Visible = False
    cboPrinter.Visible = False
    cmdPrintOut.Visible = False
    
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
     
End Sub

Private Sub ambilLaporanVinRetentionTglInvoice()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         Set Report = New rptSBERateByModelVinRetention
2         Call createTableLaporanVin
          
3         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
4         Set rsModel = objO.FetchDataHO(sSQL)
        
5         sSQL = "select COUNT(noinv) total,fs_rc,"
6         sSQL = sSQL & " E.fs_grp_mdl"
7         sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
8         sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
9         sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
10        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
11        sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
12        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
13        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0" & dealerPilih
14        sSQL = sSQL & " grOUP BY FS_RC, e.fs_grp_mdl"
15        sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
           
16        Set rsSBETotal = objO.FetchDataHO(sSQL)
          
          
17        sSQL = "select COUNT(noinv) total,fs_rc,case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
18        sSQL = sSQL & " E.fs_grp_mdl"
19        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
20        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
21        sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
22        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = '' and d.internal = 1"
23        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
24        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
25        sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
26        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
27        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0" & dealerPilih
28        sSQL = sSQL & " grOUP BY FS_RC , case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end , e.fs_grp_mdl"
29        sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
            
30        Set rsSBE = objO.FetchData(sSQL)
31        If Not rsModel.EOF Then
32            rsModel.MoveFirst
33            ProgVin.Max = rsModel.RecordCount + 1
34            ProgVin.value = 1
35            ProgVin.Visible = True
36            DoEvents
37            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
38                With rsLaporan
39                    .AddNew
40                    rsSBE.filter = ""
41                    If Not rsSBE.EOF Then
42                        rsSBE.filter = ""
43                        nilaiTemp = 0
44                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
45                        If Not rsSBE.EOF Then
46                            Do While Not rsSBE.EOF
47                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
48                                rsSBE.MoveNext
49                            Loop
50                        End If
51                        !K1 = nilaiTemp
52                        rsSBE.filter = ""
53                        nilaiTemp = 0
                          
54                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
55                        If Not rsSBE.EOF Then
56                            Do While Not rsSBE.EOF
57                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
58                                rsSBE.MoveNext
59                            Loop
60                        End If
61                        !K10 = nilaiTemp
62                        rsSBE.filter = ""
63                        nilaiTemp = 0
                          
64                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
65                        If Not rsSBE.EOF Then
66                            Do While Not rsSBE.EOF
67                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
68                                rsSBE.MoveNext
69                            Loop
70                        End If
71                        !K20 = nilaiTemp
72                        rsSBE.filter = ""
73                        nilaiTemp = 0
                          
74                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
75                        If Not rsSBE.EOF Then
76                            Do While Not rsSBE.EOF
77                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
78                                rsSBE.MoveNext
79                            Loop
80                        End If
81                        !K30 = nilaiTemp
82                        rsSBE.filter = ""
83                        nilaiTemp = 0
                          
84                        rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
85                        If Not rsSBE.EOF Then
86                            Do While Not rsSBE.EOF
87                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
88                                rsSBE.MoveNext
89                            Loop
90                        End If
91                        !K40 = nilaiTemp
92                        rsSBE.filter = ""
93                        nilaiTemp = 0
                          
94                        rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
95                        If Not rsSBE.EOF Then
96                            Do While Not rsSBE.EOF
97                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
98                                rsSBE.MoveNext
99                            Loop
100                       End If
101                       !K50 = nilaiTemp
102                       rsSBE.filter = ""
103                       nilaiTemp = 0
                          
104                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
105                       If Not rsSBE.EOF Then
106                           Do While Not rsSBE.EOF
107                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
108                               rsSBE.MoveNext
109                           Loop
110                       End If
111                       !K60 = nilaiTemp
112                       rsSBE.filter = ""
113                       nilaiTemp = 0
                          
114                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
115                       If Not rsSBE.EOF Then
116                           Do While Not rsSBE.EOF
117                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
118                               rsSBE.MoveNext
119                           Loop
120                       End If
121                       !K70 = nilaiTemp
122                       rsSBE.filter = ""
123                       nilaiTemp = 0
                          
124                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
125                       If Not rsSBE.EOF Then
126                           Do While Not rsSBE.EOF
127                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
128                               rsSBE.MoveNext
129                           Loop
130                       End If
131                       !K80 = nilaiTemp
132                       rsSBE.filter = ""
133                       nilaiTemp = 0
                          
134                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
135                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
136                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
137                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
138                   Else
139                       !K1 = 0
140                       !K10 = 0
141                       !K20 = 0
142                       !K30 = 0
143                       !K40 = 0
144                       !K50 = 0
145                       !K60 = 0
146                       !K70 = 0
147                       !K80 = 0
148                       !tot16 = 0
149                       !tot19 = 0
                          
150                       !tot26 = 0
151                       !tot29 = 0
                          
152                   End If
                      
153                   nilaiTotal1(0) = nilaiTotal1(0) + !K1
154                   nilaiTotal1(1) = nilaiTotal1(1) + !K10
155                   nilaiTotal1(2) = nilaiTotal1(2) + !K20
156                   nilaiTotal1(3) = nilaiTotal1(3) + !K30
157                   nilaiTotal1(4) = nilaiTotal1(4) + !K40
158                   nilaiTotal1(5) = nilaiTotal1(5) + !K50
159                   nilaiTotal1(6) = nilaiTotal1(6) + !K60
160                   nilaiTotal1(7) = nilaiTotal1(7) + !K70
161                   nilaiTotal1(8) = nilaiTotal1(8) + !K80
                      
162                   !judul = "SB Same Group (NCM)"
163                   !model = Nul(rsModel.Fields(0).value)
164                   !Urut = 1
165                   .Update
166               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME GROUP
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
167               With rsLaporan
168                   .AddNew
                  
169                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
170                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
171                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
172                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
173                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
174                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
175                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
176                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
177                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
178                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
179                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
180                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
181                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
182                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
183                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
184                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
185                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
186                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
187                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
188                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
189                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
190                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
191                   nilaiTotal2(0) = nilaiTotal2(0) + !K1
192                   nilaiTotal2(1) = nilaiTotal2(1) + !K10
193                   nilaiTotal2(2) = nilaiTotal2(2) + !K20
194                   nilaiTotal2(3) = nilaiTotal2(3) + !K30
195                   nilaiTotal2(4) = nilaiTotal2(4) + !K40
196                   nilaiTotal2(5) = nilaiTotal2(5) + !K50
197                   nilaiTotal2(6) = nilaiTotal2(6) + !K60
198                   nilaiTotal2(7) = nilaiTotal2(7) + !K70
199                   nilaiTotal2(8) = nilaiTotal2(8) + !K80
                       
200                   !judul = "Demand Same Outlet"
201                   !model = Nul(rsModel.Fields(0).value)
202                   !Urut = 2
203                   .Update
204               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
205               For i = 0 To 10
206                   nilaiTempDemand(i) = 0
207                   nilaiTempTotal(i) = 0
208               Next
209               rsLaporan.filter = ""
210               If Not rsLaporan.EOF Then
211                   rsLaporan.filter = "judul = 'Demand Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
212                   If Not rsLaporan.EOF Then
213                       If rsLaporan!K1 > 0 Then
214                           nilaiTempDemand(0) = rsLaporan!K1
215                       End If
216                       If rsLaporan!K10 > 0 Then
217                           nilaiTempDemand(1) = rsLaporan!K10
218                       End If
219                       If rsLaporan!K20 > 0 Then
220                           nilaiTempDemand(2) = rsLaporan!K20
221                       End If
222                       If rsLaporan!K30 > 0 Then
223                           nilaiTempDemand(3) = rsLaporan!K30
224                       End If
225                       If rsLaporan!K40 > 0 Then
226                           nilaiTempDemand(4) = rsLaporan!K40
227                       End If
228                       If rsLaporan!K50 > 0 Then
229                           nilaiTempDemand(5) = rsLaporan!K50
230                       End If
231                       If rsLaporan!K60 > 0 Then
232                           nilaiTempDemand(6) = rsLaporan!K60
233                       End If
234                       If rsLaporan!K70 > 0 Then
235                           nilaiTempDemand(7) = rsLaporan!K70
236                       End If
237                       If rsLaporan!K80 > 0 Then
238                           nilaiTempDemand(8) = rsLaporan!K80
239                       End If
240                   End If
241               End If
242               rsLaporan.filter = ""
                  
243               If Not rsLaporan.EOF Then
244                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
245                   If Not rsLaporan.EOF Then
246                       If rsLaporan!K1 > 0 Then
247                           nilaiTempTotal(0) = rsLaporan!K1
248                       End If
249                       If rsLaporan!K10 > 0 Then
250                           nilaiTempTotal(1) = rsLaporan!K10
251                       End If
252                       If rsLaporan!K20 > 0 Then
253                           nilaiTempTotal(2) = rsLaporan!K20
254                       End If
255                       If rsLaporan!K30 > 0 Then
256                           nilaiTempTotal(3) = rsLaporan!K30
257                       End If
258                       If rsLaporan!K40 > 0 Then
259                           nilaiTempTotal(4) = rsLaporan!K40
260                       End If
261                       If rsLaporan!K50 > 0 Then
262                           nilaiTempTotal(5) = rsLaporan!K50
263                       End If
264                       If rsLaporan!K60 > 0 Then
265                           nilaiTempTotal(6) = rsLaporan!K60
266                       End If
267                       If rsLaporan!K70 > 0 Then
268                           nilaiTempTotal(7) = rsLaporan!K70
269                       End If
270                       If rsLaporan!K80 > 0 Then
271                           nilaiTempTotal(8) = rsLaporan!K80
272                       End If
273                   End If
274               End If
                  
275               With rsLaporan
276                   .AddNew
277                   If nilaiTempDemand(0) > 0 Then
278                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
279                   Else
280                       !K1 = 0 & " %"
281                   End If
282                   If nilaiTempDemand(1) > 0 Then
283                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
284                   Else
285                       !K10 = 0 & " %"
286                   End If
287                   If nilaiTempDemand(2) > 0 Then
288                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
289                   Else
290                       !K20 = 0 & " %"
291                   End If
292                   If nilaiTempDemand(3) > 0 Then
293                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
294                   Else
295                       !K30 = 0 & " %"
296                   End If
297                   If nilaiTempDemand(4) > 0 Then
298                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
299                   Else
300                       !K40 = 0 & " %"
301                   End If
302                   If nilaiTempDemand(5) > 0 Then
303                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
304                   Else
305                       !K50 = 0 & " %"
306                   End If
307                   If nilaiTempDemand(6) > 0 Then
308                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
309                   Else
310                       !K60 = 0 & " %"
311                   End If
312                   If nilaiTempDemand(7) > 0 Then
313                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
314                   Else
315                       !K70 = 0 & " %"
316                   End If
317                   If nilaiTempDemand(8) > 0 Then
318                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
319                   Else
320                       !K80 = 0 & " %"
321                   End If
                      
322                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
323                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
324                   Else
325                       !tot16 = 0 & " %"
326                   End If
                      
327                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
328                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
329                   Else
330                       !tot26 = 0 & " %"
331                   End If
                       
332                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
333                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
334                   Else
335                       !tot19 = 0 & " %"
336                   End If
                       
337                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
338                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
339                   Else
340                       !tot29 = 0 & " %"
341                   End If
                       
342                   !Urut = 3
343                   !judul = "Retention Rate Same Group NCM"
344                   !model = Nul(rsModel.Fields(0).value)
345                   .Update
346               End With
                   
347               With rsLaporan
348                   .AddNew
349                   If Not rsSBETotal.EOF Then
350                       rsSBETotal.filter = ""
351                       nilaiTemp = 0
352                       rsSBETotal.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
353                       If Not rsSBETotal.EOF Then
354                           Do While Not rsSBETotal.EOF
355                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
356                               rsSBETotal.MoveNext
357                           Loop
358                       End If
359                       !K1 = nilaiTemp
360                       rsSBETotal.filter = ""
361                       nilaiTemp = 0
                          
362                       rsSBETotal.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
363                       If Not rsSBETotal.EOF Then
364                           Do While Not rsSBETotal.EOF
365                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
366                               rsSBETotal.MoveNext
367                           Loop
368                       End If
369                       !K10 = nilaiTemp
370                       rsSBETotal.filter = ""
371                       nilaiTemp = 0
                          
372                       rsSBETotal.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
373                       If Not rsSBETotal.EOF Then
374                           Do While Not rsSBETotal.EOF
375                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
376                               rsSBETotal.MoveNext
377                           Loop
378                       End If
379                       !K20 = nilaiTemp
380                       rsSBETotal.filter = ""
381                       nilaiTemp = 0
                          
382                       rsSBETotal.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
383                       If Not rsSBETotal.EOF Then
384                           Do While Not rsSBETotal.EOF
385                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
386                               rsSBETotal.MoveNext
387                           Loop
388                       End If
389                       !K30 = nilaiTemp
390                       rsSBETotal.filter = ""
391                       nilaiTemp = 0
                          
392                       rsSBETotal.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
393                       If Not rsSBETotal.EOF Then
394                           Do While Not rsSBETotal.EOF
395                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
396                               rsSBETotal.MoveNext
397                           Loop
398                       End If
399                       !K40 = nilaiTemp
400                       rsSBETotal.filter = ""
401                       nilaiTemp = 0
                          
402                       rsSBETotal.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
403                       If Not rsSBETotal.EOF Then
404                           Do While Not rsSBETotal.EOF
405                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
406                               rsSBETotal.MoveNext
407                           Loop
408                       End If
409                       !K50 = nilaiTemp
410                       rsSBETotal.filter = ""
411                       nilaiTemp = 0
                          
412                       rsSBETotal.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
413                       If Not rsSBETotal.EOF Then
414                           Do While Not rsSBETotal.EOF
415                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
416                               rsSBETotal.MoveNext
417                           Loop
418                       End If
419                       !K60 = nilaiTemp
420                       rsSBETotal.filter = ""
421                       nilaiTemp = 0
                          
422                       rsSBETotal.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
423                       If Not rsSBETotal.EOF Then
424                           Do While Not rsSBETotal.EOF
425                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
426                               rsSBETotal.MoveNext
427                           Loop
428                       End If
429                       !K70 = nilaiTemp
430                       rsSBETotal.filter = ""
431                       nilaiTemp = 0
                          
432                       rsSBETotal.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
433                       If Not rsSBETotal.EOF Then
434                           Do While Not rsSBETotal.EOF
435                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
436                               rsSBETotal.MoveNext
437                           Loop
438                       End If
439                       !K80 = nilaiTemp
440                       rsSBETotal.filter = ""
441                       nilaiTemp = 0
                          
442                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
443                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
444                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
445                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
446                   Else
447                       !K1 = 0
448                       !K10 = 0
449                       !K20 = 0
450                       !K30 = 0
451                       !K40 = 0
452                       !K50 = 0
453                       !K60 = 0
454                       !K70 = 0
455                       !K80 = 0
456                       !tot16 = 0
457                       !tot19 = 0
458                       !tot29 = 0
459                       !tot26 = 0
460                   End If
                      
461                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
462                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
463                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
464                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
465                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
466                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
467                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
468                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
469                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
470                   !judul = "SB Total (NCM & Non NCM)"
471                   !model = Nul(rsModel.Fields(0).value)
472                   !Urut = 4
473                   .Update
474               End With
475               rsLaporan.filter = ""
                  
                  
476               For i = 0 To 10
477                   nilaiTempTotal(i) = 0
478               Next
479               rsLaporan.filter = ""
                  
480               If Not rsLaporan.EOF Then
481                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
482                   If Not rsLaporan.EOF Then
483                       If rsLaporan!K1 > 0 Then
484                           nilaiTempTotal(0) = rsLaporan!K1
485                       End If
486                       If rsLaporan!K10 > 0 Then
487                           nilaiTempTotal(1) = rsLaporan!K10
488                       End If
489                       If rsLaporan!K20 > 0 Then
490                           nilaiTempTotal(2) = rsLaporan!K20
491                       End If
492                       If rsLaporan!K30 > 0 Then
493                           nilaiTempTotal(3) = rsLaporan!K30
494                       End If
495                       If rsLaporan!K40 > 0 Then
496                           nilaiTempTotal(4) = rsLaporan!K40
497                       End If
498                       If rsLaporan!K50 > 0 Then
499                           nilaiTempTotal(5) = rsLaporan!K50
500                       End If
501                       If rsLaporan!K60 > 0 Then
502                           nilaiTempTotal(6) = rsLaporan!K60
503                       End If
504                       If rsLaporan!K70 > 0 Then
505                           nilaiTempTotal(7) = rsLaporan!K70
506                       End If
507                       If rsLaporan!K80 > 0 Then
508                           nilaiTempTotal(8) = rsLaporan!K80
509                       End If
510                   End If
511               End If
                  
512               With rsLaporan
513                   .AddNew
514                   If nilaiTempDemand(0) > 0 Then
515                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
516                   Else
517                       !K1 = 0 & " %"
518                   End If
519                   If nilaiTempDemand(1) > 0 Then
520                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
521                   Else
522                       !K10 = 0 & " %"
523                   End If
524                   If nilaiTempDemand(2) > 0 Then
525                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
526                   Else
527                       !K20 = 0 & " %"
528                   End If
529                   If nilaiTempDemand(3) > 0 Then
530                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
531                   Else
532                       !K30 = 0 & " %"
533                   End If
534                   If nilaiTempDemand(4) > 0 Then
535                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
536                   Else
537                       !K40 = 0 & " %"
538                   End If
539                   If nilaiTempDemand(5) > 0 Then
540                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
541                   Else
542                       !K50 = 0 & " %"
543                   End If
544                   If nilaiTempDemand(6) > 0 Then
545                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
546                   Else
547                       !K60 = 0 & " %"
548                   End If
549                   If nilaiTempDemand(7) > 0 Then
550                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
551                   Else
552                       !K70 = 0 & " %"
553                   End If
554                   If nilaiTempDemand(8) > 0 Then
555                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
556                   Else
557                       !K80 = 0 & " %"
558                   End If
                      
                      
559                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
560                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
561                   Else
562                       !tot16 = 0 & " %"
563                   End If
                      
564                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
565                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
566                   Else
567                       !tot26 = 0 & " %"
568                   End If
                       
569                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
570                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
571                   Else
572                       !tot19 = 0 & " %"
573                   End If
                       
574                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
575                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
576                   Else
577                       !tot29 = 0 & " %"
578                   End If
                      
579                   !Urut = 5
580                   !judul = "Retention Rate Total (NCM & Non NCM)"
581                   !model = Nul(rsModel.Fields(0).value)
582                   .Update
583               End With
584               rsLaporan.filter = ""
585               rsModel.MoveNext
586               ProgVin.value = ProgVin.value + 1
587               DoEvents
588           Loop
589       Else
590           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
591           Exit Sub
592       End If
          
593       Call addTotalRow
          
594       Report.database.SetDataSource rsLaporan
595       ProgVin.Visible = False
596       Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
597       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
598       With CRViewer1
599           .ReportSource = Report
600           .ViewReport
601           .Zoom 100
602           .Visible = True
603           .Left = 300
604           .Height = Me.Height - 1200
605           .Width = Me.Width - 750
606       End With
         
607       cmdTutup.Visible = True
608       cmdTutup.Left = CRViewer1.Left + 20
609       cmdTutup.Top = CRViewer1.Top + 400
610       Screen.MousePointer = vbDefault
          
611       Call DisableExport
612       Call setLokasiTombol
End Sub


Private Sub ambilLaporanVinRetentionTglInvoiceDealer()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         Set Report = New rptSBERateByModelVinRetention
2         Call createTableLaporanVin
          
3         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
4         Set rsModel = objO.FetchDataHO(sSQL)
           
5         sSQL = "select COUNT(noinv) total,fs_rc,"
6         sSQL = sSQL & " E.fs_grp_mdl"
7         sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
8         sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
9         sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
10        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
11        sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
12        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
13        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0" & dealerPilih
14        sSQL = sSQL & " grOUP BY FS_RC, e.fs_grp_mdl"
15        sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
           
16        Set rsSBETotal = objO.FetchDataHO(sSQL)
          
17        dealerPilih = Replace(dealerPilih, "t.fs_kd_dealer", "D.fs_kd_dealer")
          
           
18        sSQL = "select COUNT(noinv) total,fs_rc,case when isnull(d.FS_KD_DEALER,'') <> '' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
19        sSQL = sSQL & " E.fs_grp_mdl"
20        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
21        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
22        sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
23        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''" & dealerPilih
24        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = ''"
25        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
26        sSQL = sSQL & " AND INVDATE between '" & Format(tglAwalVin.value, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin.value, "mm/dd/yyyy") & "'"
27        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
          
28        dealerPilih = Replace(dealerPilih, "D.fs_kd_dealer", "t.fs_kd_dealer")
          
29        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0" & dealerPilih
30        sSQL = sSQL & " grOUP BY FS_RC , case when isnull(d.FS_KD_DEALER,'') <> '' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end, e.fs_grp_mdl"
31        sSQL = sSQL & " order by FS_RC,e.fs_grp_mdl"
           
32        Set rsSBE = objO.FetchDataHO(sSQL)
33        If Not rsModel.EOF Then
34            rsModel.MoveFirst
35            ProgVin.Max = rsModel.RecordCount + 1
36            ProgVin.value = 1
37            ProgVin.Visible = True
38            DoEvents
39            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
40                With rsLaporan
41                    .AddNew
42                    rsSBE.filter = ""
43                    If Not rsSBE.EOF Then
44                        rsSBE.filter = ""
45                        nilaiTemp = 0
46                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
47                        If Not rsSBE.EOF Then
48                            Do While Not rsSBE.EOF
49                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
50                                rsSBE.MoveNext
51                            Loop
52                        End If
53                        !K1 = nilaiTemp
54                        rsSBE.filter = ""
55                        nilaiTemp = 0
                          
56                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
57                        If Not rsSBE.EOF Then
58                            Do While Not rsSBE.EOF
59                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
60                                rsSBE.MoveNext
61                            Loop
62                        End If
63                        !K10 = nilaiTemp
64                        rsSBE.filter = ""
65                        nilaiTemp = 0
                          
66                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
67                        If Not rsSBE.EOF Then
68                            Do While Not rsSBE.EOF
69                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
70                                rsSBE.MoveNext
71                            Loop
72                        End If
73                        !K20 = nilaiTemp
74                        rsSBE.filter = ""
75                        nilaiTemp = 0
                          
76                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
77                        If Not rsSBE.EOF Then
78                            Do While Not rsSBE.EOF
79                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
80                                rsSBE.MoveNext
81                            Loop
82                        End If
83                        !K30 = nilaiTemp
84                        rsSBE.filter = ""
85                        nilaiTemp = 0
                          
86                        rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
87                        If Not rsSBE.EOF Then
88                            Do While Not rsSBE.EOF
89                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
90                                rsSBE.MoveNext
91                            Loop
92                        End If
93                        !K40 = nilaiTemp
94                        rsSBE.filter = ""
95                        nilaiTemp = 0
                          
96                        rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
97                        If Not rsSBE.EOF Then
98                            Do While Not rsSBE.EOF
99                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
100                               rsSBE.MoveNext
101                           Loop
102                       End If
103                       !K50 = nilaiTemp
104                       rsSBE.filter = ""
105                       nilaiTemp = 0
                          
106                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
107                       If Not rsSBE.EOF Then
108                           Do While Not rsSBE.EOF
109                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
110                               rsSBE.MoveNext
111                           Loop
112                       End If
113                       !K60 = nilaiTemp
114                       rsSBE.filter = ""
115                       nilaiTemp = 0
                          
116                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
117                       If Not rsSBE.EOF Then
118                           Do While Not rsSBE.EOF
119                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
120                               rsSBE.MoveNext
121                           Loop
122                       End If
123                       !K70 = nilaiTemp
124                       rsSBE.filter = ""
125                       nilaiTemp = 0
                          
126                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
127                       If Not rsSBE.EOF Then
128                           Do While Not rsSBE.EOF
129                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
130                               rsSBE.MoveNext
131                           Loop
132                       End If
133                       !K80 = nilaiTemp
134                       rsSBE.filter = ""
135                       nilaiTemp = 0
                          
136                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
137                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
138                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
139                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
140                   Else
141                       !K1 = 0
142                       !K10 = 0
143                       !K20 = 0
144                       !K30 = 0
145                       !K40 = 0
146                       !K50 = 0
147                       !K60 = 0
148                       !K70 = 0
149                       !K80 = 0
150                       !tot16 = 0
151                       !tot19 = 0
                          
152                       !tot26 = 0
153                       !tot29 = 0
                          
154                   End If
                      
                      
155                   nilaiTotal1(0) = nilaiTotal1(0) + !K1
156                   nilaiTotal1(1) = nilaiTotal1(1) + !K10
157                   nilaiTotal1(2) = nilaiTotal1(2) + !K20
158                   nilaiTotal1(3) = nilaiTotal1(3) + !K30
159                   nilaiTotal1(4) = nilaiTotal1(4) + !K40
160                   nilaiTotal1(5) = nilaiTotal1(5) + !K50
161                   nilaiTotal1(6) = nilaiTotal1(6) + !K60
162                   nilaiTotal1(7) = nilaiTotal1(7) + !K70
163                   nilaiTotal1(8) = nilaiTotal1(8) + !K80
                      
164                   !judul = "SB Same Outlet"
165                   !model = Nul(rsModel.Fields(0).value)
166                   !Urut = 1
167                   .Update
168               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME OUTLET
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
169               With rsLaporan
170                   .AddNew
                  
171                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
172                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
173                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
174                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
175                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
176                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
177                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
178                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
179                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
180                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
181                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
182                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
183                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
184                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
185                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
186                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
187                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
188                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
189                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
190                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
191                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
192                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
193                   nilaiTotal2(0) = nilaiTotal2(0) + !K1
194                   nilaiTotal2(1) = nilaiTotal2(1) + !K10
195                   nilaiTotal2(2) = nilaiTotal2(2) + !K20
196                   nilaiTotal2(3) = nilaiTotal2(3) + !K30
197                   nilaiTotal2(4) = nilaiTotal2(4) + !K40
198                   nilaiTotal2(5) = nilaiTotal2(5) + !K50
199                   nilaiTotal2(6) = nilaiTotal2(6) + !K60
200                   nilaiTotal2(7) = nilaiTotal2(7) + !K70
201                   nilaiTotal2(8) = nilaiTotal2(8) + !K80
                       
202                   !judul = "Demand SB Same Outlet"
203                   !model = Nul(rsModel.Fields(0).value)
204                   !Urut = 2
205                   .Update
206               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
207               For i = 0 To 10
208                   nilaiTempDemand(i) = 0
209                   nilaiTempTotal(i) = 0
210               Next
211               rsLaporan.filter = ""
212               If Not rsLaporan.EOF Then
213                   rsLaporan.filter = "judul = 'Demand SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
214                   If Not rsLaporan.EOF Then
215                       If rsLaporan!K1 > 0 Then
216                           nilaiTempDemand(0) = rsLaporan!K1
217                       End If
218                       If rsLaporan!K10 > 0 Then
219                           nilaiTempDemand(1) = rsLaporan!K10
220                       End If
221                       If rsLaporan!K20 > 0 Then
222                           nilaiTempDemand(2) = rsLaporan!K20
223                       End If
224                       If rsLaporan!K30 > 0 Then
225                           nilaiTempDemand(3) = rsLaporan!K30
226                       End If
227                       If rsLaporan!K40 > 0 Then
228                           nilaiTempDemand(4) = rsLaporan!K40
229                       End If
230                       If rsLaporan!K50 > 0 Then
231                           nilaiTempDemand(5) = rsLaporan!K50
232                       End If
233                       If rsLaporan!K60 > 0 Then
234                           nilaiTempDemand(6) = rsLaporan!K60
235                       End If
236                       If rsLaporan!K70 > 0 Then
237                           nilaiTempDemand(7) = rsLaporan!K70
238                       End If
239                       If rsLaporan!K80 > 0 Then
240                           nilaiTempDemand(8) = rsLaporan!K80
241                       End If
242                   End If
243               End If
244               rsLaporan.filter = ""
                  
245               If Not rsLaporan.EOF Then
246                   rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
247                   If Not rsLaporan.EOF Then
248                       If rsLaporan!K1 > 0 Then
249                           nilaiTempTotal(0) = rsLaporan!K1
250                       End If
251                       If rsLaporan!K10 > 0 Then
252                           nilaiTempTotal(1) = rsLaporan!K10
253                       End If
254                       If rsLaporan!K20 > 0 Then
255                           nilaiTempTotal(2) = rsLaporan!K20
256                       End If
257                       If rsLaporan!K30 > 0 Then
258                           nilaiTempTotal(3) = rsLaporan!K30
259                       End If
260                       If rsLaporan!K40 > 0 Then
261                           nilaiTempTotal(4) = rsLaporan!K40
262                       End If
263                       If rsLaporan!K50 > 0 Then
264                           nilaiTempTotal(5) = rsLaporan!K50
265                       End If
266                       If rsLaporan!K60 > 0 Then
267                           nilaiTempTotal(6) = rsLaporan!K60
268                       End If
269                       If rsLaporan!K70 > 0 Then
270                           nilaiTempTotal(7) = rsLaporan!K70
271                       End If
272                       If rsLaporan!K80 > 0 Then
273                           nilaiTempTotal(8) = rsLaporan!K80
274                       End If
275                   End If
276               End If
                  
277               With rsLaporan
278                   .AddNew
279                   If nilaiTempDemand(0) > 0 Then
280                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
281                   Else
282                       !K1 = 0 & " %"
283                   End If
284                   If nilaiTempDemand(1) > 0 Then
285                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
286                   Else
287                       !K10 = 0 & " %"
288                   End If
289                   If nilaiTempDemand(2) > 0 Then
290                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
291                   Else
292                       !K20 = 0 & " %"
293                   End If
294                   If nilaiTempDemand(3) > 0 Then
295                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
296                   Else
297                       !K30 = 0 & " %"
298                   End If
299                   If nilaiTempDemand(4) > 0 Then
300                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
301                   Else
302                       !K40 = 0 & " %"
303                   End If
304                   If nilaiTempDemand(5) > 0 Then
305                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
306                   Else
307                       !K50 = 0 & " %"
308                   End If
309                   If nilaiTempDemand(6) > 0 Then
310                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
311                   Else
312                       !K60 = 0 & " %"
313                   End If
314                   If nilaiTempDemand(7) > 0 Then
315                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
316                   Else
317                       !K70 = 0 & " %"
318                   End If
319                   If nilaiTempDemand(8) > 0 Then
320                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
321                   Else
322                       !K80 = 0 & " %"
323                   End If
                      
324                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
325                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
326                   Else
327                       !tot16 = 0 & " %"
328                   End If
                      
329                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
330                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
331                   Else
332                       !tot26 = 0 & " %"
333                   End If
                       
334                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
335                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
336                   Else
337                       !tot19 = 0 & " %"
338                   End If
                       
339                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
340                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
341                   Else
342                       !tot29 = 0 & " %"
343                   End If
                       
344                   !Urut = 3
345                   !judul = "Retention Rate Same Outlate"
346                   !model = Nul(rsModel.Fields(0).value)
347                   .Update
348               End With
                  
349               With rsLaporan
350                   .AddNew
351                   rsSBE.filter = ""
352                   If Not rsSBE.EOF Then
353                       rsSBE.filter = ""
354                       nilaiTemp = 0
355                       rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
356                       If Not rsSBE.EOF Then
357                           Do While Not rsSBE.EOF
358                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
359                               rsSBE.MoveNext
360                           Loop
361                       End If
362                       !K1 = nilaiTemp
363                       rsSBE.filter = ""
364                       nilaiTemp = 0
                          
365                       rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
366                       If Not rsSBE.EOF Then
367                           Do While Not rsSBE.EOF
368                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
369                               rsSBE.MoveNext
370                           Loop
371                       End If
372                       !K10 = nilaiTemp
373                       rsSBE.filter = ""
374                       nilaiTemp = 0
                          
375                       rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
376                       If Not rsSBE.EOF Then
377                           Do While Not rsSBE.EOF
378                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
379                               rsSBE.MoveNext
380                           Loop
381                       End If
382                       !K20 = nilaiTemp
383                       rsSBE.filter = ""
384                       nilaiTemp = 0
                          
385                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
386                       If Not rsSBE.EOF Then
387                           Do While Not rsSBE.EOF
388                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
389                               rsSBE.MoveNext
390                           Loop
391                       End If
392                       !K30 = nilaiTemp
393                       rsSBE.filter = ""
394                       nilaiTemp = 0
                          
395                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
396                       If Not rsSBE.EOF Then
397                           Do While Not rsSBE.EOF
398                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
399                               rsSBE.MoveNext
400                           Loop
401                       End If
402                       !K40 = nilaiTemp
403                       rsSBE.filter = ""
404                       nilaiTemp = 0
                          
405                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
406                       If Not rsSBE.EOF Then
407                           Do While Not rsSBE.EOF
408                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
409                               rsSBE.MoveNext
410                           Loop
411                       End If
412                       !K50 = nilaiTemp
413                       rsSBE.filter = ""
414                       nilaiTemp = 0
                          
415                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
416                       If Not rsSBE.EOF Then
417                           Do While Not rsSBE.EOF
418                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
419                               rsSBE.MoveNext
420                           Loop
421                       End If
422                       !K60 = nilaiTemp
423                       rsSBE.filter = ""
424                       nilaiTemp = 0
                          
425                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
426                       If Not rsSBE.EOF Then
427                           Do While Not rsSBE.EOF
428                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
429                               rsSBE.MoveNext
430                           Loop
431                       End If
432                       !K70 = nilaiTemp
433                       rsSBE.filter = ""
434                       nilaiTemp = 0
                          
435                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
436                       If Not rsSBE.EOF Then
437                           Do While Not rsSBE.EOF
438                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
439                               rsSBE.MoveNext
440                           Loop
441                       End If
442                       !K80 = nilaiTemp
443                       rsSBE.filter = ""
444                       nilaiTemp = 0
                           
445                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
446                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
447                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
448                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
449                   Else
450                       !K1 = 0
451                       !K10 = 0
452                       !K20 = 0
453                       !K30 = 0
454                       !K40 = 0
455                       !K50 = 0
456                       !K60 = 0
457                       !K70 = 0
458                       !K80 = 0
459                       !tot16 = 0
460                       !tot19 = 0
                          
461                       !tot26 = 0
462                       !tot29 = 0
463                   End If
                      
464                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
465                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
466                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
467                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
468                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
469                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
470                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
471                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
472                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
473                   !judul = "SB Same Group (NCM)"
474                   !model = Nul(rsModel.Fields(0).value)
475                   !Urut = 4
476                   .Update
477               End With
                  
                  
478               For i = 0 To 10
479                   nilaiTempTotal(i) = 0
480               Next
481               rsLaporan.filter = ""
                  
482               If Not rsLaporan.EOF Then
483                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
484                   If Not rsLaporan.EOF Then
485                       If rsLaporan!K1 > 0 Then
486                           nilaiTempTotal(0) = rsLaporan!K1
487                       End If
488                       If rsLaporan!K10 > 0 Then
489                           nilaiTempTotal(1) = rsLaporan!K10
490                       End If
491                       If rsLaporan!K20 > 0 Then
492                           nilaiTempTotal(2) = rsLaporan!K20
493                       End If
494                       If rsLaporan!K30 > 0 Then
495                           nilaiTempTotal(3) = rsLaporan!K30
496                       End If
497                       If rsLaporan!K40 > 0 Then
498                           nilaiTempTotal(4) = rsLaporan!K40
499                       End If
500                       If rsLaporan!K50 > 0 Then
501                           nilaiTempTotal(5) = rsLaporan!K50
502                       End If
503                       If rsLaporan!K60 > 0 Then
504                           nilaiTempTotal(6) = rsLaporan!K60
505                       End If
506                       If rsLaporan!K70 > 0 Then
507                           nilaiTempTotal(7) = rsLaporan!K70
508                       End If
509                       If rsLaporan!K80 > 0 Then
510                           nilaiTempTotal(8) = rsLaporan!K80
511                       End If
512                   End If
513               End If
                  
514               With rsLaporan
515                   .AddNew
516                   If nilaiTempDemand(0) > 0 Then
517                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
518                   Else
519                       !K1 = 0 & " %"
520                   End If
521                   If nilaiTempDemand(1) > 0 Then
522                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
523                   Else
524                       !K10 = 0 & " %"
525                   End If
526                   If nilaiTempDemand(2) > 0 Then
527                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
528                   Else
529                       !K20 = 0 & " %"
530                   End If
531                   If nilaiTempDemand(3) > 0 Then
532                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
533                   Else
534                       !K30 = 0 & " %"
535                   End If
536                   If nilaiTempDemand(4) > 0 Then
537                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
538                   Else
539                       !K40 = 0 & " %"
540                   End If
541                   If nilaiTempDemand(5) > 0 Then
542                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
543                   Else
544                       !K50 = 0 & " %"
545                   End If
546                   If nilaiTempDemand(6) > 0 Then
547                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
548                   Else
549                       !K60 = 0 & " %"
550                   End If
551                   If nilaiTempDemand(7) > 0 Then
552                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
553                   Else
554                       !K70 = 0 & " %"
555                   End If
556                   If nilaiTempDemand(8) > 0 Then
557                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
558                   Else
559                       !K80 = 0 & " %"
560                   End If
                      
                      
561                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
562                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
563                   Else
564                       !tot16 = 0 & " %"
565                   End If
                      
566                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
567                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
568                   Else
569                       !tot26 = 0 & " %"
570                   End If
                       
571                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
572                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
573                   Else
574                       !tot19 = 0 & " %"
575                   End If
                       
576                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
577                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
578                   Else
579                       !tot29 = 0 & " %"
580                   End If
                      
581                   !Urut = 5
582                   !judul = "Retention Rate Same Group (NCM)"
583                   !model = Nul(rsModel.Fields(0).value)
584                   .Update
585               End With
                   
586               With rsLaporan
587                   .AddNew
588                   If Not rsSBETotal.EOF Then
589                       rsSBETotal.filter = ""
590                       nilaiTemp = 0
591                       rsSBETotal.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
592                       If Not rsSBETotal.EOF Then
593                           Do While Not rsSBETotal.EOF
594                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
595                               rsSBETotal.MoveNext
596                           Loop
597                       End If
598                       !K1 = nilaiTemp
599                       rsSBETotal.filter = ""
600                       nilaiTemp = 0
                          
601                       rsSBETotal.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
602                       If Not rsSBETotal.EOF Then
603                           Do While Not rsSBETotal.EOF
604                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
605                               rsSBETotal.MoveNext
606                           Loop
607                       End If
608                       !K10 = nilaiTemp
609                       rsSBETotal.filter = ""
610                       nilaiTemp = 0
                          
611                       rsSBETotal.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
612                       If Not rsSBETotal.EOF Then
613                           Do While Not rsSBETotal.EOF
614                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
615                               rsSBETotal.MoveNext
616                           Loop
617                       End If
618                       !K20 = nilaiTemp
619                       rsSBETotal.filter = ""
620                       nilaiTemp = 0
                          
621                       rsSBETotal.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
622                       If Not rsSBETotal.EOF Then
623                           Do While Not rsSBETotal.EOF
624                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
625                               rsSBETotal.MoveNext
626                           Loop
627                       End If
628                       !K30 = nilaiTemp
629                       rsSBETotal.filter = ""
630                       nilaiTemp = 0
                          
631                       rsSBETotal.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
632                       If Not rsSBETotal.EOF Then
633                           Do While Not rsSBETotal.EOF
634                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
635                               rsSBETotal.MoveNext
636                           Loop
637                       End If
638                       !K40 = nilaiTemp
639                       rsSBETotal.filter = ""
640                       nilaiTemp = 0
                          
641                       rsSBETotal.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
642                       If Not rsSBETotal.EOF Then
643                           Do While Not rsSBETotal.EOF
644                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
645                               rsSBETotal.MoveNext
646                           Loop
647                       End If
648                       !K50 = nilaiTemp
649                       rsSBETotal.filter = ""
650                       nilaiTemp = 0
                          
651                       rsSBETotal.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
652                       If Not rsSBETotal.EOF Then
653                           Do While Not rsSBETotal.EOF
654                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
655                               rsSBETotal.MoveNext
656                           Loop
657                       End If
658                       !K60 = nilaiTemp
659                       rsSBETotal.filter = ""
660                       nilaiTemp = 0
                          
661                       rsSBETotal.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
662                       If Not rsSBETotal.EOF Then
663                           Do While Not rsSBETotal.EOF
664                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
665                               rsSBETotal.MoveNext
666                           Loop
667                       End If
668                       !K70 = nilaiTemp
669                       rsSBETotal.filter = ""
670                       nilaiTemp = 0
                          
671                       rsSBETotal.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
672                       If Not rsSBETotal.EOF Then
673                           Do While Not rsSBETotal.EOF
674                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
675                               rsSBETotal.MoveNext
676                           Loop
677                       End If
678                       !K80 = nilaiTemp
679                       rsSBETotal.filter = ""
680                       nilaiTemp = 0
                          
681                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
682                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
683                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
684                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
685                   Else
686                       !K1 = 0
687                       !K10 = 0
688                       !K20 = 0
689                       !K30 = 0
690                       !K40 = 0
691                       !K50 = 0
692                       !K60 = 0
693                       !K70 = 0
694                       !K80 = 0
695                       !tot16 = 0
696                       !tot19 = 0
697                       !tot29 = 0
698                       !tot26 = 0
699                   End If
                      
700                   nilaiTotal6(0) = nilaiTotal6(0) + !K1
701                   nilaiTotal6(1) = nilaiTotal6(1) + !K10
702                   nilaiTotal6(2) = nilaiTotal6(2) + !K20
703                   nilaiTotal6(3) = nilaiTotal6(3) + !K30
704                   nilaiTotal6(4) = nilaiTotal6(4) + !K40
705                   nilaiTotal6(5) = nilaiTotal6(5) + !K50
706                   nilaiTotal6(6) = nilaiTotal6(6) + !K60
707                   nilaiTotal6(7) = nilaiTotal6(7) + !K70
708                   nilaiTotal6(8) = nilaiTotal6(8) + !K80
                      
709                   !judul = "SB Total (NCM & Non NCM)"
710                   !model = Nul(rsModel.Fields(0).value)
711                   !Urut = 6
712                   .Update
713               End With
714               rsLaporan.filter = ""
                  
                  
715               For i = 0 To 10
716                   nilaiTempTotal(i) = 0
717               Next
718               rsLaporan.filter = ""
                  
719               If Not rsLaporan.EOF Then
720                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
721                   If Not rsLaporan.EOF Then
722                       If rsLaporan!K1 > 0 Then
723                           nilaiTempTotal(0) = rsLaporan!K1
724                       End If
725                       If rsLaporan!K10 > 0 Then
726                           nilaiTempTotal(1) = rsLaporan!K10
727                       End If
728                       If rsLaporan!K20 > 0 Then
729                           nilaiTempTotal(2) = rsLaporan!K20
730                       End If
731                       If rsLaporan!K30 > 0 Then
732                           nilaiTempTotal(3) = rsLaporan!K30
733                       End If
734                       If rsLaporan!K40 > 0 Then
735                           nilaiTempTotal(4) = rsLaporan!K40
736                       End If
737                       If rsLaporan!K50 > 0 Then
738                           nilaiTempTotal(5) = rsLaporan!K50
739                       End If
740                       If rsLaporan!K60 > 0 Then
741                           nilaiTempTotal(6) = rsLaporan!K60
742                       End If
743                       If rsLaporan!K70 > 0 Then
744                           nilaiTempTotal(7) = rsLaporan!K70
745                       End If
746                       If rsLaporan!K80 > 0 Then
747                           nilaiTempTotal(8) = rsLaporan!K80
748                       End If
749                   End If
750               End If
                  
751               With rsLaporan
752                   .AddNew
753                   If nilaiTempDemand(0) > 0 Then
754                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
755                   Else
756                       !K1 = 0 & " %"
757                   End If
758                   If nilaiTempDemand(1) > 0 Then
759                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
760                   Else
761                       !K10 = 0 & " %"
762                   End If
763                   If nilaiTempDemand(2) > 0 Then
764                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
765                   Else
766                       !K20 = 0 & " %"
767                   End If
768                   If nilaiTempDemand(3) > 0 Then
769                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
770                   Else
771                       !K30 = 0 & " %"
772                   End If
773                   If nilaiTempDemand(4) > 0 Then
774                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
775                   Else
776                       !K40 = 0 & " %"
777                   End If
778                   If nilaiTempDemand(5) > 0 Then
779                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
780                   Else
781                       !K50 = 0 & " %"
782                   End If
783                   If nilaiTempDemand(6) > 0 Then
784                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
785                   Else
786                       !K60 = 0 & " %"
787                   End If
788                   If nilaiTempDemand(7) > 0 Then
789                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
790                   Else
791                       !K70 = 0 & " %"
792                   End If
793                   If nilaiTempDemand(8) > 0 Then
794                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
795                   Else
796                       !K80 = 0 & " %"
797                   End If
                      
                      
798                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
799                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
800                   Else
801                       !tot16 = 0 & " %"
802                   End If
                      
803                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
804                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
805                   Else
806                       !tot26 = 0 & " %"
807                   End If
                       
808                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
809                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
810                   Else
811                       !tot19 = 0 & " %"
812                   End If
                       
813                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
814                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
815                   Else
816                       !tot29 = 0 & " %"
817                   End If
                      
818                   !Urut = 7
819                   !judul = "Retention Rate Total (NCM & Non NCM)"
820                   !model = Nul(rsModel.Fields(0).value)
821                   .Update
822               End With
823               rsLaporan.filter = ""
824               rsModel.MoveNext
825               ProgVin.value = ProgVin.value + 1
826               DoEvents
827           Loop
828       Else
829           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
830           Exit Sub
831       End If
          
832       Call addTotalRowDealer
          
833       Report.database.SetDataSource rsLaporan
834       ProgVin.Visible = False
835       Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
836       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
837       With CRViewer1
838           .ReportSource = Report
839           .ViewReport
840           .Zoom 100
841           .Visible = True
842           .Left = 300
843           .Height = Me.Height - 1200
844           .Width = Me.Width - 750
845       End With
         
846       cmdTutup.Visible = True
847       cmdTutup.Left = CRViewer1.Left + 20
848       cmdTutup.Top = CRViewer1.Top + 400
849       Screen.MousePointer = vbDefault
          
850       Call DisableExport
851       Call setLokasiTombol
End Sub

Private Sub setLokasiTombol()
    ''2017-04-28
    cmdTutup.Top = CRViewer1.Top + 345
    cmdTutup.Left = CRViewer1.Left + 30
    
    cmdSettingPrinter.Top = CRViewer1.Top + 345
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 30
    
    cmdExport.Top = CRViewer1.Top + 345
    cmdExport.Left = cmdSettingPrinter.Left + cmdSettingPrinter.Width + 30
    
'    XPButton1.Top = CRViewer1.Top + 345
'    XPButton1.Left = cmdExport.Left + cmdExport.Width + 30
    
    cboPrinter.Top = CRViewer1.Top + 345
    cboPrinter.Left = cmdExport.Left + cmdExport.Width + 30
    
    cmdPrintOut.Top = CRViewer1.Top + 345
    cmdPrintOut.Left = cboPrinter.Left + cboPrinter.Width + 30
End Sub
Private Sub DisableExport()
    Dim objPrinter As Printer
    Set objPrinter = Printer
    If InStr(UCase(objPrinter.DeviceName), "PDF") <> 0 Then
        cboPrinter = ""
    Else
        cboPrinter = objPrinter.DeviceName
    End If
    cboPrinter.Clear
    
    For Each pr In Printers
        If InStr(UCase(pr.DeviceName), "PDF") = 0 Then cboPrinter.AddItem pr.DeviceName
    Next pr
    If Log.PdfOnly = 0 Then
        CRViewer1.EnableExportButton = True
        CRViewer1.EnablePrintButton = True
        cmdExport.Visible = False
        cboPrinter.Visible = False
        cmdPrintOut.Visible = False
        'XPButton1.Visible = False
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    ElseIf Log.PdfOnly = 1 Then
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = True
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        'XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    Else
        CRViewer1.EnableExportButton = False
        CRViewer1.EnablePrintButton = False
        cmdExport.Visible = False
        cboPrinter.Visible = True
        cmdPrintOut.Visible = True
        'XPButton1.Visible = True
        
        cmdSettingPrinter.Visible = True
        cmdTutup.Visible = True
    End If

End Sub


Private Sub createTableLaporanVin()
    Set rsLaporan = New ADODB.Recordset
    If rsLaporan.State Then rsLaporan.Close
    'rsLaporan.Close
    
    
    With rsLaporan
        .Fields.Append "urut", adInteger
        .Fields.Append "keterangan", adVarChar, 100
        .Fields.Append "judul", adVarChar, 100
        .Fields.Append "model", adVarChar, 100
        .Fields.Append "K1", adVarChar, 100
        .Fields.Append "K10", adVarChar, 100
        .Fields.Append "K20", adVarChar, 100
        .Fields.Append "K30", adVarChar, 100
        .Fields.Append "K40", adVarChar, 100
        .Fields.Append "K50", adVarChar, 100
        .Fields.Append "K60", adVarChar, 100
        .Fields.Append "K70", adVarChar, 100
        .Fields.Append "K80", adVarChar, 100
        .Fields.Append "tot16", adVarChar, 100
        .Fields.Append "tot26", adVarChar, 100
        .Fields.Append "tot19", adVarChar, 100
        .Fields.Append "tot29", adVarChar, 100
        .Fields.Append "north", adInteger
        .Open
    End With
End Sub


Private Sub ambilLaporanVinRetentionTglRetailDealer()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         Set Report = New rptSBERateByModelVinRetention
2         Call createTableLaporanVin
          
3         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
4         Set rsModel = objO.FetchDataHO(sSQL)
          
          
5         dealerPilih = Replace(dealerPilih, "t.fs_kd_dealer", "D.fs_kd_dealer")
          
         
6         sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when isnull(d.FS_KD_DEALER,'') <> '' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
7         sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK) "
8         sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  "
9         sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
10        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''" & dealerPilih
11        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
12        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
13        sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [SPK99]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(tglAwalVin, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin, "mm/dd/yyyy") & "' )"
14        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
15        sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when isnull(d.FS_KD_DEALER,'') <> '' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
16        sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
           
17        Set rsSBE = objO.FetchDataHO(sSQL)
18        If Not rsModel.EOF Then
19            rsModel.MoveFirst
20            ProgVin.Max = rsModel.RecordCount + 1
21            ProgVin.value = 1
22            ProgVin.Visible = True
23            DoEvents
24            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
25                With rsLaporan
26                    .AddNew
27                    rsSBE.filter = ""
28                    If Not rsSBE.EOF Then
29                        rsSBE.filter = ""
30                        nilaiTemp = 0
31                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
32                        If Not rsSBE.EOF Then
33                            Do While Not rsSBE.EOF
34                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
35                                rsSBE.MoveNext
36                            Loop
37                        End If
38                        !K1 = nilaiTemp
39                        rsSBE.filter = ""
40                        nilaiTemp = 0
                          
41                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
42                        If Not rsSBE.EOF Then
43                            Do While Not rsSBE.EOF
44                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
45                                rsSBE.MoveNext
46                            Loop
47                        End If
48                        !K10 = nilaiTemp
49                        rsSBE.filter = ""
50                        nilaiTemp = 0
                          
51                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
52                        If Not rsSBE.EOF Then
53                            Do While Not rsSBE.EOF
54                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
55                                rsSBE.MoveNext
56                            Loop
57                        End If
58                        !K20 = nilaiTemp
59                        rsSBE.filter = ""
60                        nilaiTemp = 0
                          
61                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
62                        If Not rsSBE.EOF Then
63                            Do While Not rsSBE.EOF
64                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
65                                rsSBE.MoveNext
66                            Loop
67                        End If
68                        !K30 = nilaiTemp
69                        rsSBE.filter = ""
70                        nilaiTemp = 0
                          
71                        rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
72                        If Not rsSBE.EOF Then
73                            Do While Not rsSBE.EOF
74                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
75                                rsSBE.MoveNext
76                            Loop
77                        End If
78                        !K40 = nilaiTemp
79                        rsSBE.filter = ""
80                        nilaiTemp = 0
                          
81                        rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
82                        If Not rsSBE.EOF Then
83                            Do While Not rsSBE.EOF
84                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
85                                rsSBE.MoveNext
86                            Loop
87                        End If
88                        !K50 = nilaiTemp
89                        rsSBE.filter = ""
90                        nilaiTemp = 0
                          
91                        rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
92                        If Not rsSBE.EOF Then
93                            Do While Not rsSBE.EOF
94                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
95                                rsSBE.MoveNext
96                            Loop
97                        End If
98                        !K60 = nilaiTemp
99                        rsSBE.filter = ""
100                       nilaiTemp = 0
                          
101                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
102                       If Not rsSBE.EOF Then
103                           Do While Not rsSBE.EOF
104                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
105                               rsSBE.MoveNext
106                           Loop
107                       End If
108                       !K70 = nilaiTemp
109                       rsSBE.filter = ""
110                       nilaiTemp = 0
                          
111                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
112                       If Not rsSBE.EOF Then
113                           Do While Not rsSBE.EOF
114                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
115                               rsSBE.MoveNext
116                           Loop
117                       End If
118                       !K80 = nilaiTemp
119                       rsSBE.filter = ""
120                       nilaiTemp = 0
                          
121                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
122                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
123                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
124                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
125                   Else
126                       !K1 = 0
127                       !K10 = 0
128                       !K20 = 0
129                       !K30 = 0
130                       !K40 = 0
131                       !K50 = 0
132                       !K60 = 0
133                       !K70 = 0
134                       !K80 = 0
135                       !tot16 = 0
136                       !tot19 = 0
                          
137                       !tot26 = 0
138                       !tot29 = 0
                          
139                   End If
                      
140                   nilaiTotal1(0) = nilaiTotal1(0) + !K1
141                   nilaiTotal1(1) = nilaiTotal1(1) + !K10
142                   nilaiTotal1(2) = nilaiTotal1(2) + !K20
143                   nilaiTotal1(3) = nilaiTotal1(3) + !K30
144                   nilaiTotal1(4) = nilaiTotal1(4) + !K40
145                   nilaiTotal1(5) = nilaiTotal1(5) + !K50
146                   nilaiTotal1(6) = nilaiTotal1(6) + !K60
147                   nilaiTotal1(7) = nilaiTotal1(7) + !K70
148                   nilaiTotal1(8) = nilaiTotal1(8) + !K80
                      
149                   !judul = "SB Same Outlet"
150                   !model = Nul(rsModel.Fields(0).value)
151                   !Urut = 1
152                   .Update
153               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME OUTLET
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
154               With rsLaporan
155                   .AddNew
                  
156                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
157                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
158                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
159                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
160                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
161                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
162                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
163                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
164                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
165                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
166                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
167                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
168                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
169                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
170                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
171                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
172                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
173                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
                      
174                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
175                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
176                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
177                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
178                   nilaiTotal2(0) = nilaiTotal2(0) + !K1
179                   nilaiTotal2(1) = nilaiTotal2(1) + !K10
180                   nilaiTotal2(2) = nilaiTotal2(2) + !K20
181                   nilaiTotal2(3) = nilaiTotal2(3) + !K30
182                   nilaiTotal2(4) = nilaiTotal2(4) + !K40
183                   nilaiTotal2(5) = nilaiTotal2(5) + !K50
184                   nilaiTotal2(6) = nilaiTotal2(6) + !K60
185                   nilaiTotal2(7) = nilaiTotal2(7) + !K70
186                   nilaiTotal2(8) = nilaiTotal2(8) + !K80
                       
187                   !judul = "Demand SB Same Outlet"
188                   !model = Nul(rsModel.Fields(0).value)
189                   !Urut = 2
190                   .Update
191               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
192               For i = 0 To 10
193                   nilaiTempDemand(i) = 0
194                   nilaiTempTotal(i) = 0
195               Next
196               rsLaporan.filter = ""
197               If Not rsLaporan.EOF Then
198                   rsLaporan.filter = "judul = 'Demand SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
199                   If Not rsLaporan.EOF Then
200                       If rsLaporan!K1 > 0 Then
201                           nilaiTempDemand(0) = rsLaporan!K1
202                       End If
203                       If rsLaporan!K10 > 0 Then
204                           nilaiTempDemand(1) = rsLaporan!K10
205                       End If
206                       If rsLaporan!K20 > 0 Then
207                           nilaiTempDemand(2) = rsLaporan!K20
208                       End If
209                       If rsLaporan!K30 > 0 Then
210                           nilaiTempDemand(3) = rsLaporan!K30
211                       End If
212                       If rsLaporan!K40 > 0 Then
213                           nilaiTempDemand(4) = rsLaporan!K40
214                       End If
215                       If rsLaporan!K50 > 0 Then
216                           nilaiTempDemand(5) = rsLaporan!K50
217                       End If
218                       If rsLaporan!K60 > 0 Then
219                           nilaiTempDemand(6) = rsLaporan!K60
220                       End If
221                       If rsLaporan!K70 > 0 Then
222                           nilaiTempDemand(7) = rsLaporan!K70
223                       End If
224                       If rsLaporan!K80 > 0 Then
225                           nilaiTempDemand(8) = rsLaporan!K80
226                       End If
227                   End If
228               End If
229               rsLaporan.filter = ""
                  
230               If Not rsLaporan.EOF Then
231                   rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
232                   If Not rsLaporan.EOF Then
233                       If rsLaporan!K1 > 0 Then
234                           nilaiTempTotal(0) = rsLaporan!K1
235                       End If
236                       If rsLaporan!K10 > 0 Then
237                           nilaiTempTotal(1) = rsLaporan!K10
238                       End If
239                       If rsLaporan!K20 > 0 Then
240                           nilaiTempTotal(2) = rsLaporan!K20
241                       End If
242                       If rsLaporan!K30 > 0 Then
243                           nilaiTempTotal(3) = rsLaporan!K30
244                       End If
245                       If rsLaporan!K40 > 0 Then
246                           nilaiTempTotal(4) = rsLaporan!K40
247                       End If
248                       If rsLaporan!K50 > 0 Then
249                           nilaiTempTotal(5) = rsLaporan!K50
250                       End If
251                       If rsLaporan!K60 > 0 Then
252                           nilaiTempTotal(6) = rsLaporan!K60
253                       End If
254                       If rsLaporan!K70 > 0 Then
255                           nilaiTempTotal(7) = rsLaporan!K70
256                       End If
257                       If rsLaporan!K80 > 0 Then
258                           nilaiTempTotal(8) = rsLaporan!K80
259                       End If
260                   End If
261               End If
                  
262               With rsLaporan
263                   .AddNew
264                   If nilaiTempDemand(0) > 0 Then
265                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
266                   Else
267                       !K1 = 0 & " %"
268                   End If
269                   If nilaiTempDemand(1) > 0 Then
270                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
271                   Else
272                       !K10 = 0 & " %"
273                   End If
274                   If nilaiTempDemand(2) > 0 Then
275                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
276                   Else
277                       !K20 = 0 & " %"
278                   End If
279                   If nilaiTempDemand(3) > 0 Then
280                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
281                   Else
282                       !K30 = 0 & " %"
283                   End If
284                   If nilaiTempDemand(4) > 0 Then
285                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
286                   Else
287                       !K40 = 0 & " %"
288                   End If
289                   If nilaiTempDemand(5) > 0 Then
290                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
291                   Else
292                       !K50 = 0 & " %"
293                   End If
294                   If nilaiTempDemand(6) > 0 Then
295                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
296                   Else
297                       !K60 = 0 & " %"
298                   End If
299                   If nilaiTempDemand(7) > 0 Then
300                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
301                   Else
302                       !K70 = 0 & " %"
303                   End If
304                   If nilaiTempDemand(8) > 0 Then
305                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
306                   Else
307                       !K80 = 0 & " %"
308                   End If
                      
309                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
310                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
311                   Else
312                       !tot16 = 0 & " %"
313                   End If
                      
314                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
315                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
316                   Else
317                       !tot26 = 0 & " %"
318                   End If
                       
319                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
320                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
321                   Else
322                       !tot19 = 0 & " %"
323                   End If
                       
324                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
325                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
326                   Else
327                       !tot29 = 0 & " %"
328                   End If
                       
329                   !Urut = 3
330                   !judul = "Retention Rate Same Outlate"
331                   !model = Nul(rsModel.Fields(0).value)
332                   .Update
333               End With
                  
334               With rsLaporan
335                   .AddNew
336                   rsSBE.filter = ""
337                   If Not rsSBE.EOF Then
338                       rsSBE.filter = ""
339                       nilaiTemp = 0
340                       rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
341                       If Not rsSBE.EOF Then
342                           Do While Not rsSBE.EOF
343                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
344                               rsSBE.MoveNext
345                           Loop
346                       End If
347                       !K1 = nilaiTemp
348                       rsSBE.filter = ""
349                       nilaiTemp = 0
                          
350                       rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
351                       If Not rsSBE.EOF Then
352                           Do While Not rsSBE.EOF
353                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
354                               rsSBE.MoveNext
355                           Loop
356                       End If
357                       !K10 = nilaiTemp
358                       rsSBE.filter = ""
359                       nilaiTemp = 0
                          
360                       rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
361                       If Not rsSBE.EOF Then
362                           Do While Not rsSBE.EOF
363                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
364                               rsSBE.MoveNext
365                           Loop
366                       End If
367                       !K20 = nilaiTemp
368                       rsSBE.filter = ""
369                       nilaiTemp = 0
                          
370                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
371                       If Not rsSBE.EOF Then
372                           Do While Not rsSBE.EOF
373                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
374                               rsSBE.MoveNext
375                           Loop
376                       End If
377                       !K30 = nilaiTemp
378                       rsSBE.filter = ""
379                       nilaiTemp = 0
                          
380                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
381                       If Not rsSBE.EOF Then
382                           Do While Not rsSBE.EOF
383                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
384                               rsSBE.MoveNext
385                           Loop
386                       End If
387                       !K40 = nilaiTemp
388                       rsSBE.filter = ""
389                       nilaiTemp = 0
                          
390                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
391                       If Not rsSBE.EOF Then
392                           Do While Not rsSBE.EOF
393                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
394                               rsSBE.MoveNext
395                           Loop
396                       End If
397                       !K50 = nilaiTemp
398                       rsSBE.filter = ""
399                       nilaiTemp = 0
                          
400                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
401                       If Not rsSBE.EOF Then
402                           Do While Not rsSBE.EOF
403                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
404                               rsSBE.MoveNext
405                           Loop
406                       End If
407                       !K60 = nilaiTemp
408                       rsSBE.filter = ""
409                       nilaiTemp = 0
                          
410                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
411                       If Not rsSBE.EOF Then
412                           Do While Not rsSBE.EOF
413                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
414                               rsSBE.MoveNext
415                           Loop
416                       End If
417                       !K70 = nilaiTemp
418                       rsSBE.filter = ""
419                       nilaiTemp = 0
                          
420                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
421                       If Not rsSBE.EOF Then
422                           Do While Not rsSBE.EOF
423                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
424                               rsSBE.MoveNext
425                           Loop
426                       End If
427                       !K80 = nilaiTemp
428                       rsSBE.filter = ""
429                       nilaiTemp = 0
                           
430                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
431                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
432                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
433                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
434                   Else
435                       !K1 = 0
436                       !K10 = 0
437                       !K20 = 0
438                       !K30 = 0
439                       !K40 = 0
440                       !K50 = 0
441                       !K60 = 0
442                       !K70 = 0
443                       !K80 = 0
444                       !tot16 = 0
445                       !tot19 = 0
                          
446                       !tot26 = 0
447                       !tot29 = 0
448                   End If
                      
449                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
450                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
451                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
452                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
453                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
454                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
455                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
456                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
457                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
458                   !judul = "SB Same Group (NCM)"
459                   !model = Nul(rsModel.Fields(0).value)
460                   !Urut = 4
461                   .Update
462               End With
                  
                  
463               For i = 0 To 10
464                   nilaiTempTotal(i) = 0
465               Next
466               rsLaporan.filter = ""
                  
467               If Not rsLaporan.EOF Then
468                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
469                   If Not rsLaporan.EOF Then
470                       If rsLaporan!K1 > 0 Then
471                           nilaiTempTotal(0) = rsLaporan!K1
472                       End If
473                       If rsLaporan!K10 > 0 Then
474                           nilaiTempTotal(1) = rsLaporan!K10
475                       End If
476                       If rsLaporan!K20 > 0 Then
477                           nilaiTempTotal(2) = rsLaporan!K20
478                       End If
479                       If rsLaporan!K30 > 0 Then
480                           nilaiTempTotal(3) = rsLaporan!K30
481                       End If
482                       If rsLaporan!K40 > 0 Then
483                           nilaiTempTotal(4) = rsLaporan!K40
484                       End If
485                       If rsLaporan!K50 > 0 Then
486                           nilaiTempTotal(5) = rsLaporan!K50
487                       End If
488                       If rsLaporan!K60 > 0 Then
489                           nilaiTempTotal(6) = rsLaporan!K60
490                       End If
491                       If rsLaporan!K70 > 0 Then
492                           nilaiTempTotal(7) = rsLaporan!K70
493                       End If
494                       If rsLaporan!K80 > 0 Then
495                           nilaiTempTotal(8) = rsLaporan!K80
496                       End If
497                   End If
498               End If
                  
499               With rsLaporan
500                   .AddNew
501                   If nilaiTempDemand(0) > 0 Then
502                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
503                   Else
504                       !K1 = 0 & " %"
505                   End If
506                   If nilaiTempDemand(1) > 0 Then
507                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
508                   Else
509                       !K10 = 0 & " %"
510                   End If
511                   If nilaiTempDemand(2) > 0 Then
512                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
513                   Else
514                       !K20 = 0 & " %"
515                   End If
516                   If nilaiTempDemand(3) > 0 Then
517                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
518                   Else
519                       !K30 = 0 & " %"
520                   End If
521                   If nilaiTempDemand(4) > 0 Then
522                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
523                   Else
524                       !K40 = 0 & " %"
525                   End If
526                   If nilaiTempDemand(5) > 0 Then
527                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
528                   Else
529                       !K50 = 0 & " %"
530                   End If
531                   If nilaiTempDemand(6) > 0 Then
532                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
533                   Else
534                       !K60 = 0 & " %"
535                   End If
536                   If nilaiTempDemand(7) > 0 Then
537                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
538                   Else
539                       !K70 = 0 & " %"
540                   End If
541                   If nilaiTempDemand(8) > 0 Then
542                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
543                   Else
544                       !K80 = 0 & " %"
545                   End If
                      
                      
546                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
547                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
548                   Else
549                       !tot16 = 0 & " %"
550                   End If
                      
551                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
552                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
553                   Else
554                       !tot26 = 0 & " %"
555                   End If
                       
556                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
557                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
558                   Else
559                       !tot19 = 0 & " %"
560                   End If
                       
561                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
562                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
563                   Else
564                       !tot29 = 0 & " %"
565                   End If
                      
566                   !Urut = 5
567                   !judul = "Retention Rate Same Group (NCM)"
568                   !model = Nul(rsModel.Fields(0).value)
569                   .Update
570               End With
                   
571               With rsLaporan
572                   .AddNew
573                   If Not rsSBE.EOF Then
574                       rsSBE.filter = ""
575                       nilaiTemp = 0
576                       rsSBE.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
577                       If Not rsSBE.EOF Then
578                           Do While Not rsSBE.EOF
579                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
580                               rsSBE.MoveNext
581                           Loop
582                       End If
583                       !K1 = nilaiTemp
584                       rsSBE.filter = ""
585                       nilaiTemp = 0
                          
586                       rsSBE.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
587                       If Not rsSBE.EOF Then
588                           Do While Not rsSBE.EOF
589                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
590                               rsSBE.MoveNext
591                           Loop
592                       End If
593                       !K10 = nilaiTemp
594                       rsSBE.filter = ""
595                       nilaiTemp = 0
                          
596                       rsSBE.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
597                       If Not rsSBE.EOF Then
598                           Do While Not rsSBE.EOF
599                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
600                               rsSBE.MoveNext
601                           Loop
602                       End If
603                       !K20 = nilaiTemp
604                       rsSBE.filter = ""
605                       nilaiTemp = 0
                          
606                       rsSBE.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
607                       If Not rsSBE.EOF Then
608                           Do While Not rsSBE.EOF
609                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
610                               rsSBE.MoveNext
611                           Loop
612                       End If
613                       !K30 = nilaiTemp
614                       rsSBE.filter = ""
615                       nilaiTemp = 0
                          
616                       rsSBE.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
617                       If Not rsSBE.EOF Then
618                           Do While Not rsSBE.EOF
619                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
620                               rsSBE.MoveNext
621                           Loop
622                       End If
623                       !K40 = nilaiTemp
624                       rsSBE.filter = ""
625                       nilaiTemp = 0
                          
626                       rsSBE.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
627                       If Not rsSBE.EOF Then
628                           Do While Not rsSBE.EOF
629                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
630                               rsSBE.MoveNext
631                           Loop
632                       End If
633                       !K50 = nilaiTemp
634                       rsSBE.filter = ""
635                       nilaiTemp = 0
                          
636                       rsSBE.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
637                       If Not rsSBE.EOF Then
638                           Do While Not rsSBE.EOF
639                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
640                               rsSBE.MoveNext
641                           Loop
642                       End If
643                       !K60 = nilaiTemp
644                       rsSBE.filter = ""
645                       nilaiTemp = 0
                          
646                       rsSBE.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
647                       If Not rsSBE.EOF Then
648                           Do While Not rsSBE.EOF
649                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
650                               rsSBE.MoveNext
651                           Loop
652                       End If
653                       !K70 = nilaiTemp
654                       rsSBE.filter = ""
655                       nilaiTemp = 0
                          
656                       rsSBE.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
657                       If Not rsSBE.EOF Then
658                           Do While Not rsSBE.EOF
659                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
660                               rsSBE.MoveNext
661                           Loop
662                       End If
663                       !K80 = nilaiTemp
664                       rsSBE.filter = ""
665                       nilaiTemp = 0
                          
666                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
667                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
668                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
669                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
670                   Else
671                       !K1 = 0
672                       !K10 = 0
673                       !K20 = 0
674                       !K30 = 0
675                       !K40 = 0
676                       !K50 = 0
677                       !K60 = 0
678                       !K70 = 0
679                       !K80 = 0
680                       !tot16 = 0
681                       !tot19 = 0
682                       !tot29 = 0
683                       !tot26 = 0
684                   End If
                      
685                   nilaiTotal6(0) = nilaiTotal6(0) + !K1
686                   nilaiTotal6(1) = nilaiTotal6(1) + !K10
687                   nilaiTotal6(2) = nilaiTotal6(2) + !K20
688                   nilaiTotal6(3) = nilaiTotal6(3) + !K30
689                   nilaiTotal6(4) = nilaiTotal6(4) + !K40
690                   nilaiTotal6(5) = nilaiTotal6(5) + !K50
691                   nilaiTotal6(6) = nilaiTotal6(6) + !K60
692                   nilaiTotal6(7) = nilaiTotal6(7) + !K70
693                   nilaiTotal6(8) = nilaiTotal6(8) + !K80
                      
694                   !judul = "SB Total (NCM & Non NCM)"
695                   !model = Nul(rsModel.Fields(0).value)
696                   !Urut = 6
697                   .Update
698               End With
699               rsLaporan.filter = ""
                  
                  
700               For i = 0 To 10
701                   nilaiTempTotal(i) = 0
702               Next
703               rsLaporan.filter = ""
                  
704               If Not rsLaporan.EOF Then
705                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
706                   If Not rsLaporan.EOF Then
707                       If rsLaporan!K1 > 0 Then
708                           nilaiTempTotal(0) = rsLaporan!K1
709                       End If
710                       If rsLaporan!K10 > 0 Then
711                           nilaiTempTotal(1) = rsLaporan!K10
712                       End If
713                       If rsLaporan!K20 > 0 Then
714                           nilaiTempTotal(2) = rsLaporan!K20
715                       End If
716                       If rsLaporan!K30 > 0 Then
717                           nilaiTempTotal(3) = rsLaporan!K30
718                       End If
719                       If rsLaporan!K40 > 0 Then
720                           nilaiTempTotal(4) = rsLaporan!K40
721                       End If
722                       If rsLaporan!K50 > 0 Then
723                           nilaiTempTotal(5) = rsLaporan!K50
724                       End If
725                       If rsLaporan!K60 > 0 Then
726                           nilaiTempTotal(6) = rsLaporan!K60
727                       End If
728                       If rsLaporan!K70 > 0 Then
729                           nilaiTempTotal(7) = rsLaporan!K70
730                       End If
731                       If rsLaporan!K80 > 0 Then
732                           nilaiTempTotal(8) = rsLaporan!K80
733                       End If
734                   End If
735               End If
                  
736               With rsLaporan
737                   .AddNew
738                   If nilaiTempDemand(0) > 0 Then
739                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
740                   Else
741                       !K1 = 0 & " %"
742                   End If
743                   If nilaiTempDemand(1) > 0 Then
744                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
745                   Else
746                       !K10 = 0 & " %"
747                   End If
748                   If nilaiTempDemand(2) > 0 Then
749                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
750                   Else
751                       !K20 = 0 & " %"
752                   End If
753                   If nilaiTempDemand(3) > 0 Then
754                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
755                   Else
756                       !K30 = 0 & " %"
757                   End If
758                   If nilaiTempDemand(4) > 0 Then
759                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
760                   Else
761                       !K40 = 0 & " %"
762                   End If
763                   If nilaiTempDemand(5) > 0 Then
764                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
765                   Else
766                       !K50 = 0 & " %"
767                   End If
768                   If nilaiTempDemand(6) > 0 Then
769                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
770                   Else
771                       !K60 = 0 & " %"
772                   End If
773                   If nilaiTempDemand(7) > 0 Then
774                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
775                   Else
776                       !K70 = 0 & " %"
777                   End If
778                   If nilaiTempDemand(8) > 0 Then
779                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
780                   Else
781                       !K80 = 0 & " %"
782                   End If
                      
                      
783                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
784                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
785                   Else
786                       !tot16 = 0 & " %"
787                   End If
                      
788                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
789                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
790                   Else
791                       !tot26 = 0 & " %"
792                   End If
                       
793                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
794                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
795                   Else
796                       !tot19 = 0 & " %"
797                   End If
                       
798                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
799                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
800                   Else
801                       !tot29 = 0 & " %"
802                   End If
                      
803                   !Urut = 7
804                   !judul = "Retention Rate Total (NCM & Non NCM)"
805                   !model = Nul(rsModel.Fields(0).value)
806                   .Update
807               End With
808               rsLaporan.filter = ""
809               rsModel.MoveNext
810               ProgVin.value = ProgVin.value + 1
811               DoEvents
812           Loop
813       Else
814           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
815           Exit Sub
816       End If
          
817       Call addTotalRowDealer
          
818       Report.database.SetDataSource rsLaporan
819       ProgVin.Visible = False
820       Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
821       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
822       With CRViewer1
823           .ReportSource = Report
824           .ViewReport
825           .Zoom 100
826           .Visible = True
827           .Left = 300
828           .Height = Me.Height - 1200
829           .Width = Me.Width - 750
830       End With
         
831       cmdTutup.Visible = True
832       cmdTutup.Left = CRViewer1.Left + 20
833       cmdTutup.Top = CRViewer1.Top + 400
834       Screen.MousePointer = vbDefault
          
835       Call DisableExport
836       Call setLokasiTombol
End Sub


Private Sub ambilLaporanVinRetentionTglDelivery()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         Set Report = New rptSBERateByModelVinRetention
2         Call createTableLaporanVin
          
3         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
4         Set rsModel = objO.FetchDataHO(sSQL)
          
          
5         dealerPilih = Replace(dealerPilih, "t.fs_kd_dealer", "D.fs_kd_dealer")
          
         
6         sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
7         sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK) "
8         sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  "
9         sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
10        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''" & dealerPilih
11        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
12        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
13        sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [PDS00]..IntrnCbg WHERE Proses = 10 and Kode = 2 AND TGLSELESAI between '" & Format(tglAwalVin, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin, "mm/dd/yyyy") & "' )"
14        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
15        sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
16        sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
           
17        Set rsSBE = objO.FetchDataHO(sSQL)
18        If Not rsModel.EOF Then
19            rsModel.MoveFirst
20            ProgVin.Max = rsModel.RecordCount + 1
21            ProgVin.value = 1
22            ProgVin.Visible = True
23            DoEvents
24            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
25                With rsLaporan
26                    .AddNew
27                    rsSBE.filter = ""
28                    If Not rsSBE.EOF Then
29                        rsSBE.filter = ""
30                        nilaiTemp = 0
31                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
32                        If Not rsSBE.EOF Then
33                            Do While Not rsSBE.EOF
34                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
35                                rsSBE.MoveNext
36                            Loop
37                        End If
38                        !K1 = nilaiTemp
39                        rsSBE.filter = ""
40                        nilaiTemp = 0
                          
41                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
42                        If Not rsSBE.EOF Then
43                            Do While Not rsSBE.EOF
44                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
45                                rsSBE.MoveNext
46                            Loop
47                        End If
48                        !K10 = nilaiTemp
49                        rsSBE.filter = ""
50                        nilaiTemp = 0
                          
51                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
52                        If Not rsSBE.EOF Then
53                            Do While Not rsSBE.EOF
54                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
55                                rsSBE.MoveNext
56                            Loop
57                        End If
58                        !K20 = nilaiTemp
59                        rsSBE.filter = ""
60                        nilaiTemp = 0
                          
61                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
62                        If Not rsSBE.EOF Then
63                            Do While Not rsSBE.EOF
64                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
65                                rsSBE.MoveNext
66                            Loop
67                        End If
68                        !K30 = nilaiTemp
69                        rsSBE.filter = ""
70                        nilaiTemp = 0
                          
71                        rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
72                        If Not rsSBE.EOF Then
73                            Do While Not rsSBE.EOF
74                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
75                                rsSBE.MoveNext
76                            Loop
77                        End If
78                        !K40 = nilaiTemp
79                        rsSBE.filter = ""
80                        nilaiTemp = 0
                          
81                        rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
82                        If Not rsSBE.EOF Then
83                            Do While Not rsSBE.EOF
84                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
85                                rsSBE.MoveNext
86                            Loop
87                        End If
88                        !K50 = nilaiTemp
89                        rsSBE.filter = ""
90                        nilaiTemp = 0
                          
91                        rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
92                        If Not rsSBE.EOF Then
93                            Do While Not rsSBE.EOF
94                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
95                                rsSBE.MoveNext
96                            Loop
97                        End If
98                        !K60 = nilaiTemp
99                        rsSBE.filter = ""
100                       nilaiTemp = 0
                          
101                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
102                       If Not rsSBE.EOF Then
103                           Do While Not rsSBE.EOF
104                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
105                               rsSBE.MoveNext
106                           Loop
107                       End If
108                       !K70 = nilaiTemp
109                       rsSBE.filter = ""
110                       nilaiTemp = 0
                          
111                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
112                       If Not rsSBE.EOF Then
113                           Do While Not rsSBE.EOF
114                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
115                               rsSBE.MoveNext
116                           Loop
117                       End If
118                       !K80 = nilaiTemp
119                       rsSBE.filter = ""
120                       nilaiTemp = 0
                           
121                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
122                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
123                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
124                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
125                   Else
126                       !K1 = 0
127                       !K10 = 0
128                       !K20 = 0
129                       !K30 = 0
130                       !K40 = 0
131                       !K50 = 0
132                       !K60 = 0
133                       !K70 = 0
134                       !K80 = 0
135                       !tot16 = 0
136                       !tot19 = 0
                          
137                       !tot26 = 0
138                       !tot29 = 0
139                   End If
                      
                      
140                   nilaiTotal1(0) = nilaiTotal1(0) + !K1
141                   nilaiTotal1(1) = nilaiTotal1(1) + !K10
142                   nilaiTotal1(2) = nilaiTotal1(2) + !K20
143                   nilaiTotal1(3) = nilaiTotal1(3) + !K30
144                   nilaiTotal1(4) = nilaiTotal1(4) + !K40
145                   nilaiTotal1(5) = nilaiTotal1(5) + !K50
146                   nilaiTotal1(6) = nilaiTotal1(6) + !K60
147                   nilaiTotal1(7) = nilaiTotal1(7) + !K70
148                   nilaiTotal1(8) = nilaiTotal1(8) + !K80
                      
149                   !judul = "SB Same Group (NCM)"
150                   !model = Nul(rsModel.Fields(0).value)
151                   !Urut = 1
152                   .Update
153               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME OUTLET
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
154               With rsLaporan
155                   .AddNew
                  
156                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
157                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
158                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
159                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
160                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
161                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
162                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
163                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
164                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
165                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
166                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
167                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
168                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
169                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
170                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
171                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
172                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
173                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
                      
174                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
175                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
176                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
177                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
                      
178                   nilaiTotal2(0) = nilaiTotal2(0) + !K1
179                   nilaiTotal2(1) = nilaiTotal2(1) + !K10
180                   nilaiTotal2(2) = nilaiTotal2(2) + !K20
181                   nilaiTotal2(3) = nilaiTotal2(3) + !K30
182                   nilaiTotal2(4) = nilaiTotal2(4) + !K40
183                   nilaiTotal2(5) = nilaiTotal2(5) + !K50
184                   nilaiTotal2(6) = nilaiTotal2(6) + !K60
185                   nilaiTotal2(7) = nilaiTotal2(7) + !K70
186                   nilaiTotal2(8) = nilaiTotal2(8) + !K80
                       
187                   !judul = "Demand SB Same Group (NCM)"
188                   !model = Nul(rsModel.Fields(0).value)
189                   !Urut = 2
190                   .Update
191               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
192               For i = 0 To 10
193                   nilaiTempDemand(i) = 0
194                   nilaiTempTotal(i) = 0
195               Next
196               rsLaporan.filter = ""
197               If Not rsLaporan.EOF Then
198                   rsLaporan.filter = "judul = 'Demand SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
199                   If Not rsLaporan.EOF Then
200                       If rsLaporan!K1 > 0 Then
201                           nilaiTempDemand(0) = rsLaporan!K1
202                       End If
203                       If rsLaporan!K10 > 0 Then
204                           nilaiTempDemand(1) = rsLaporan!K10
205                       End If
206                       If rsLaporan!K20 > 0 Then
207                           nilaiTempDemand(2) = rsLaporan!K20
208                       End If
209                       If rsLaporan!K30 > 0 Then
210                           nilaiTempDemand(3) = rsLaporan!K30
211                       End If
212                       If rsLaporan!K40 > 0 Then
213                           nilaiTempDemand(4) = rsLaporan!K40
214                       End If
215                       If rsLaporan!K50 > 0 Then
216                           nilaiTempDemand(5) = rsLaporan!K50
217                       End If
218                       If rsLaporan!K60 > 0 Then
219                           nilaiTempDemand(6) = rsLaporan!K60
220                       End If
221                       If rsLaporan!K70 > 0 Then
222                           nilaiTempDemand(7) = rsLaporan!K70
223                       End If
224                       If rsLaporan!K80 > 0 Then
225                           nilaiTempDemand(8) = rsLaporan!K80
226                       End If
227                   End If
228               End If
229               rsLaporan.filter = ""
                  
230               If Not rsLaporan.EOF Then
231                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
232                   If Not rsLaporan.EOF Then
233                       If rsLaporan!K1 > 0 Then
234                           nilaiTempTotal(0) = rsLaporan!K1
235                       End If
236                       If rsLaporan!K10 > 0 Then
237                           nilaiTempTotal(1) = rsLaporan!K10
238                       End If
239                       If rsLaporan!K20 > 0 Then
240                           nilaiTempTotal(2) = rsLaporan!K20
241                       End If
242                       If rsLaporan!K30 > 0 Then
243                           nilaiTempTotal(3) = rsLaporan!K30
244                       End If
245                       If rsLaporan!K40 > 0 Then
246                           nilaiTempTotal(4) = rsLaporan!K40
247                       End If
248                       If rsLaporan!K50 > 0 Then
249                           nilaiTempTotal(5) = rsLaporan!K50
250                       End If
251                       If rsLaporan!K60 > 0 Then
252                           nilaiTempTotal(6) = rsLaporan!K60
253                       End If
254                       If rsLaporan!K70 > 0 Then
255                           nilaiTempTotal(7) = rsLaporan!K70
256                       End If
257                       If rsLaporan!K80 > 0 Then
258                           nilaiTempTotal(8) = rsLaporan!K80
259                       End If
260                   End If
261               End If
                  
262               With rsLaporan
263                   .AddNew
264                   If nilaiTempDemand(0) > 0 Then
265                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
266                   Else
267                       !K1 = 0 & " %"
268                   End If
269                   If nilaiTempDemand(1) > 0 Then
270                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
271                   Else
272                       !K10 = 0 & " %"
273                   End If
274                   If nilaiTempDemand(2) > 0 Then
275                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
276                   Else
277                       !K20 = 0 & " %"
278                   End If
279                   If nilaiTempDemand(3) > 0 Then
280                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
281                   Else
282                       !K30 = 0 & " %"
283                   End If
284                   If nilaiTempDemand(4) > 0 Then
285                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
286                   Else
287                       !K40 = 0 & " %"
288                   End If
289                   If nilaiTempDemand(5) > 0 Then
290                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
291                   Else
292                       !K50 = 0 & " %"
293                   End If
294                   If nilaiTempDemand(6) > 0 Then
295                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
296                   Else
297                       !K60 = 0 & " %"
298                   End If
299                   If nilaiTempDemand(7) > 0 Then
300                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
301                   Else
302                       !K70 = 0 & " %"
303                   End If
304                   If nilaiTempDemand(8) > 0 Then
305                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
306                   Else
307                       !K80 = 0 & " %"
308                   End If
                      
309                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
310                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
311                   Else
312                       !tot16 = 0 & " %"
313                   End If
                      
314                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
315                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
316                   Else
317                       !tot26 = 0 & " %"
318                   End If
                       
319                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
320                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
321                   Else
322                       !tot19 = 0 & " %"
323                   End If
                       
324                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
325                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
326                   Else
327                       !tot29 = 0 & " %"
328                   End If
                       
329                   !Urut = 3
330                   !judul = "Retention Rate Same Group (NCM)"
331                   !model = Nul(rsModel.Fields(0).value)
332                   .Update
333               End With
                          
                   
334               With rsLaporan
335                   .AddNew
336                   If Not rsSBE.EOF Then
337                       rsSBE.filter = ""
338                       nilaiTemp = 0
339                       rsSBE.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
340                       If Not rsSBE.EOF Then
341                           Do While Not rsSBE.EOF
342                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
343                               rsSBE.MoveNext
344                           Loop
345                       End If
346                       !K1 = nilaiTemp
347                       rsSBE.filter = ""
348                       nilaiTemp = 0
                          
349                       rsSBE.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
350                       If Not rsSBE.EOF Then
351                           Do While Not rsSBE.EOF
352                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
353                               rsSBE.MoveNext
354                           Loop
355                       End If
356                       !K10 = nilaiTemp
357                       rsSBE.filter = ""
358                       nilaiTemp = 0
                          
359                       rsSBE.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
360                       If Not rsSBE.EOF Then
361                           Do While Not rsSBE.EOF
362                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
363                               rsSBE.MoveNext
364                           Loop
365                       End If
366                       !K20 = nilaiTemp
367                       rsSBE.filter = ""
368                       nilaiTemp = 0
                          
369                       rsSBE.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
370                       If Not rsSBE.EOF Then
371                           Do While Not rsSBE.EOF
372                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
373                               rsSBE.MoveNext
374                           Loop
375                       End If
376                       !K30 = nilaiTemp
377                       rsSBE.filter = ""
378                       nilaiTemp = 0
                          
379                       rsSBE.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
380                       If Not rsSBE.EOF Then
381                           Do While Not rsSBE.EOF
382                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
383                               rsSBE.MoveNext
384                           Loop
385                       End If
386                       !K40 = nilaiTemp
387                       rsSBE.filter = ""
388                       nilaiTemp = 0
                          
389                       rsSBE.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
390                       If Not rsSBE.EOF Then
391                           Do While Not rsSBE.EOF
392                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
393                               rsSBE.MoveNext
394                           Loop
395                       End If
396                       !K50 = nilaiTemp
397                       rsSBE.filter = ""
398                       nilaiTemp = 0
                          
399                       rsSBE.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
400                       If Not rsSBE.EOF Then
401                           Do While Not rsSBE.EOF
402                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
403                               rsSBE.MoveNext
404                           Loop
405                       End If
406                       !K60 = nilaiTemp
407                       rsSBE.filter = ""
408                       nilaiTemp = 0
                          
409                       rsSBE.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
410                       If Not rsSBE.EOF Then
411                           Do While Not rsSBE.EOF
412                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
413                               rsSBE.MoveNext
414                           Loop
415                       End If
416                       !K70 = nilaiTemp
417                       rsSBE.filter = ""
418                       nilaiTemp = 0
                          
419                       rsSBE.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
420                       If Not rsSBE.EOF Then
421                           Do While Not rsSBE.EOF
422                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
423                               rsSBE.MoveNext
424                           Loop
425                       End If
426                       !K80 = nilaiTemp
427                       rsSBE.filter = ""
428                       nilaiTemp = 0
                          
429                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
430                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
431                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
432                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
433                   Else
434                       !K1 = 0
435                       !K10 = 0
436                       !K20 = 0
437                       !K30 = 0
438                       !K40 = 0
439                       !K50 = 0
440                       !K60 = 0
441                       !K70 = 0
442                       !K80 = 0
443                       !tot16 = 0
444                       !tot19 = 0
445                       !tot29 = 0
446                       !tot26 = 0
447                   End If
                      
448                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
449                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
450                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
451                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
452                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
453                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
454                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
455                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
456                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
457                   !judul = "SB Total (NCM & Non NCM)"
458                   !model = Nul(rsModel.Fields(0).value)
459                   !Urut = 4
460                   .Update
461               End With
462               rsLaporan.filter = ""
                  
                  
463               For i = 0 To 10
464                   nilaiTempTotal(i) = 0
465               Next
466               rsLaporan.filter = ""
                  
467               If Not rsLaporan.EOF Then
468                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
469                   If Not rsLaporan.EOF Then
470                       If rsLaporan!K1 > 0 Then
471                           nilaiTempTotal(0) = rsLaporan!K1
472                       End If
473                       If rsLaporan!K10 > 0 Then
474                           nilaiTempTotal(1) = rsLaporan!K10
475                       End If
476                       If rsLaporan!K20 > 0 Then
477                           nilaiTempTotal(2) = rsLaporan!K20
478                       End If
479                       If rsLaporan!K30 > 0 Then
480                           nilaiTempTotal(3) = rsLaporan!K30
481                       End If
482                       If rsLaporan!K40 > 0 Then
483                           nilaiTempTotal(4) = rsLaporan!K40
484                       End If
485                       If rsLaporan!K50 > 0 Then
486                           nilaiTempTotal(5) = rsLaporan!K50
487                       End If
488                       If rsLaporan!K60 > 0 Then
489                           nilaiTempTotal(6) = rsLaporan!K60
490                       End If
491                       If rsLaporan!K70 > 0 Then
492                           nilaiTempTotal(7) = rsLaporan!K70
493                       End If
494                       If rsLaporan!K80 > 0 Then
495                           nilaiTempTotal(8) = rsLaporan!K80
496                       End If
497                   End If
498               End If
                  
499               With rsLaporan
500                   .AddNew
501                   If nilaiTempDemand(0) > 0 Then
502                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
503                   Else
504                       !K1 = 0 & " %"
505                   End If
506                   If nilaiTempDemand(1) > 0 Then
507                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
508                   Else
509                       !K10 = 0 & " %"
510                   End If
511                   If nilaiTempDemand(2) > 0 Then
512                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
513                   Else
514                       !K20 = 0 & " %"
515                   End If
516                   If nilaiTempDemand(3) > 0 Then
517                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
518                   Else
519                       !K30 = 0 & " %"
520                   End If
521                   If nilaiTempDemand(4) > 0 Then
522                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
523                   Else
524                       !K40 = 0 & " %"
525                   End If
526                   If nilaiTempDemand(5) > 0 Then
527                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
528                   Else
529                       !K50 = 0 & " %"
530                   End If
531                   If nilaiTempDemand(6) > 0 Then
532                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
533                   Else
534                       !K60 = 0 & " %"
535                   End If
536                   If nilaiTempDemand(7) > 0 Then
537                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
538                   Else
539                       !K70 = 0 & " %"
540                   End If
541                   If nilaiTempDemand(8) > 0 Then
542                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
543                   Else
544                       !K80 = 0 & " %"
545                   End If
                      
                      
546                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
547                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
548                   Else
549                       !tot16 = 0 & " %"
550                   End If
                      
551                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
552                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
553                   Else
554                       !tot26 = 0 & " %"
555                   End If
                       
556                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
557                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
558                   Else
559                       !tot19 = 0 & " %"
560                   End If
                       
561                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
562                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
563                   Else
564                       !tot29 = 0 & " %"
565                   End If
                      
566                   !Urut = 5
567                   !judul = "Retention Rate Total (NCM & Non NCM)"
568                   !model = Nul(rsModel.Fields(0).value)
569                   .Update
570               End With
571               rsLaporan.filter = ""
572               rsModel.MoveNext
573               ProgVin.value = ProgVin.value + 1
574               DoEvents
575           Loop
576       Else
577           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
578           Exit Sub
579       End If
          
580       Call addTotalRow
          
581       Report.database.SetDataSource rsLaporan
582       ProgVin.Visible = False
583       Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
584       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
585       With CRViewer1
586           .ReportSource = Report
587           .ViewReport
588           .Zoom 100
589           .Visible = True
590           .Left = 300
591           .Height = Me.Height - 1200
592           .Width = Me.Width - 750
593       End With
         
594       cmdTutup.Visible = True
595       cmdTutup.Left = CRViewer1.Left + 20
596       cmdTutup.Top = CRViewer1.Top + 400
597       Screen.MousePointer = vbDefault
          
598       Call DisableExport
599       Call setLokasiTombol
End Sub


Private Sub ambilLaporanVinRetentionTglRetail()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         Set Report = New rptSBERateByModelVinRetention
2         Call createTableLaporanVin
          
3         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
4         Set rsModel = objO.FetchDataHO(sSQL)
          
          
5         dealerPilih = Replace(dealerPilih, "t.fs_kd_dealer", "D.fs_kd_dealer")
          
         
6         sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
7         sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK) "
8         sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  "
9         sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
10        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''" & dealerPilih
11        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
12        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
13        sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [SPK99]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(tglAwalVin, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin, "mm/dd/yyyy") & "' )"
14        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
15        sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
16        sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
           
17        Set rsSBE = objO.FetchDataHO(sSQL)
18        If Not rsModel.EOF Then
19            rsModel.MoveFirst
20            ProgVin.Max = rsModel.RecordCount + 1
21            ProgVin.value = 1
22            ProgVin.Visible = True
23            DoEvents
24            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
25                With rsLaporan
26                    .AddNew
27                    rsSBE.filter = ""
28                    If Not rsSBE.EOF Then
29                        rsSBE.filter = ""
30                        nilaiTemp = 0
31                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
32                        If Not rsSBE.EOF Then
33                            Do While Not rsSBE.EOF
34                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
35                                rsSBE.MoveNext
36                            Loop
37                        End If
38                        !K1 = nilaiTemp
39                        rsSBE.filter = ""
40                        nilaiTemp = 0
                          
41                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
42                        If Not rsSBE.EOF Then
43                            Do While Not rsSBE.EOF
44                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
45                                rsSBE.MoveNext
46                            Loop
47                        End If
48                        !K10 = nilaiTemp
49                        rsSBE.filter = ""
50                        nilaiTemp = 0
                          
51                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
52                        If Not rsSBE.EOF Then
53                            Do While Not rsSBE.EOF
54                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
55                                rsSBE.MoveNext
56                            Loop
57                        End If
58                        !K20 = nilaiTemp
59                        rsSBE.filter = ""
60                        nilaiTemp = 0
                          
61                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
62                        If Not rsSBE.EOF Then
63                            Do While Not rsSBE.EOF
64                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
65                                rsSBE.MoveNext
66                            Loop
67                        End If
68                        !K30 = nilaiTemp
69                        rsSBE.filter = ""
70                        nilaiTemp = 0
                          
71                        rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
72                        If Not rsSBE.EOF Then
73                            Do While Not rsSBE.EOF
74                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
75                                rsSBE.MoveNext
76                            Loop
77                        End If
78                        !K40 = nilaiTemp
79                        rsSBE.filter = ""
80                        nilaiTemp = 0
                          
81                        rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
82                        If Not rsSBE.EOF Then
83                            Do While Not rsSBE.EOF
84                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
85                                rsSBE.MoveNext
86                            Loop
87                        End If
88                        !K50 = nilaiTemp
89                        rsSBE.filter = ""
90                        nilaiTemp = 0
                          
91                        rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
92                        If Not rsSBE.EOF Then
93                            Do While Not rsSBE.EOF
94                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
95                                rsSBE.MoveNext
96                            Loop
97                        End If
98                        !K60 = nilaiTemp
99                        rsSBE.filter = ""
100                       nilaiTemp = 0
                          
101                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
102                       If Not rsSBE.EOF Then
103                           Do While Not rsSBE.EOF
104                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
105                               rsSBE.MoveNext
106                           Loop
107                       End If
108                       !K70 = nilaiTemp
109                       rsSBE.filter = ""
110                       nilaiTemp = 0
                          
111                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
112                       If Not rsSBE.EOF Then
113                           Do While Not rsSBE.EOF
114                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
115                               rsSBE.MoveNext
116                           Loop
117                       End If
118                       !K80 = nilaiTemp
119                       rsSBE.filter = ""
120                       nilaiTemp = 0
                           
121                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
122                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
123                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
124                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
125                   Else
126                       !K1 = 0
127                       !K10 = 0
128                       !K20 = 0
129                       !K30 = 0
130                       !K40 = 0
131                       !K50 = 0
132                       !K60 = 0
133                       !K70 = 0
134                       !K80 = 0
135                       !tot16 = 0
136                       !tot19 = 0
                          
137                       !tot26 = 0
138                       !tot29 = 0
139                   End If
                      
140                   nilaiTotal1(0) = nilaiTotal1(0) + !K1
141                   nilaiTotal1(1) = nilaiTotal1(1) + !K10
142                   nilaiTotal1(2) = nilaiTotal1(2) + !K20
143                   nilaiTotal1(3) = nilaiTotal1(3) + !K30
144                   nilaiTotal1(4) = nilaiTotal1(4) + !K40
145                   nilaiTotal1(5) = nilaiTotal1(5) + !K50
146                   nilaiTotal1(6) = nilaiTotal1(6) + !K60
147                   nilaiTotal1(7) = nilaiTotal1(7) + !K70
148                   nilaiTotal1(8) = nilaiTotal1(8) + !K80
                      
149                   !judul = "SB Same Group (NCM)"
150                   !model = Nul(rsModel.Fields(0).value)
151                   !Urut = 1
152                   .Update
153               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME OUTLET
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
154               With rsLaporan
155                   .AddNew
                  
156                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
157                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
158                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
159                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
160                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
161                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
162                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
163                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
164                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
165                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
166                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
167                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
168                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
169                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
170                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
171                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
172                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
173                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
                      
174                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
175                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
176                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
177                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
178                   nilaiTotal2(0) = nilaiTotal2(0) + !K1
179                   nilaiTotal2(1) = nilaiTotal2(1) + !K10
180                   nilaiTotal2(2) = nilaiTotal2(2) + !K20
181                   nilaiTotal2(3) = nilaiTotal2(3) + !K30
182                   nilaiTotal2(4) = nilaiTotal2(4) + !K40
183                   nilaiTotal2(5) = nilaiTotal2(5) + !K50
184                   nilaiTotal2(6) = nilaiTotal2(6) + !K60
185                   nilaiTotal2(7) = nilaiTotal2(7) + !K70
186                   nilaiTotal2(8) = nilaiTotal2(8) + !K80
                       
187                   !judul = "Demand SB Same Group (NCM)"
188                   !model = Nul(rsModel.Fields(0).value)
189                   !Urut = 2
190                   .Update
191               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
192               For i = 0 To 10
193                   nilaiTempDemand(i) = 0
194                   nilaiTempTotal(i) = 0
195               Next
196               rsLaporan.filter = ""
197               If Not rsLaporan.EOF Then
198                   rsLaporan.filter = "judul = 'Demand SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
199                   If Not rsLaporan.EOF Then
200                       If rsLaporan!K1 > 0 Then
201                           nilaiTempDemand(0) = rsLaporan!K1
202                       End If
203                       If rsLaporan!K10 > 0 Then
204                           nilaiTempDemand(1) = rsLaporan!K10
205                       End If
206                       If rsLaporan!K20 > 0 Then
207                           nilaiTempDemand(2) = rsLaporan!K20
208                       End If
209                       If rsLaporan!K30 > 0 Then
210                           nilaiTempDemand(3) = rsLaporan!K30
211                       End If
212                       If rsLaporan!K40 > 0 Then
213                           nilaiTempDemand(4) = rsLaporan!K40
214                       End If
215                       If rsLaporan!K50 > 0 Then
216                           nilaiTempDemand(5) = rsLaporan!K50
217                       End If
218                       If rsLaporan!K60 > 0 Then
219                           nilaiTempDemand(6) = rsLaporan!K60
220                       End If
221                       If rsLaporan!K70 > 0 Then
222                           nilaiTempDemand(7) = rsLaporan!K70
223                       End If
224                       If rsLaporan!K80 > 0 Then
225                           nilaiTempDemand(8) = rsLaporan!K80
226                       End If
227                   End If
228               End If
229               rsLaporan.filter = ""
                  
230               If Not rsLaporan.EOF Then
231                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
232                   If Not rsLaporan.EOF Then
233                       If rsLaporan!K1 > 0 Then
234                           nilaiTempTotal(0) = rsLaporan!K1
235                       End If
236                       If rsLaporan!K10 > 0 Then
237                           nilaiTempTotal(1) = rsLaporan!K10
238                       End If
239                       If rsLaporan!K20 > 0 Then
240                           nilaiTempTotal(2) = rsLaporan!K20
241                       End If
242                       If rsLaporan!K30 > 0 Then
243                           nilaiTempTotal(3) = rsLaporan!K30
244                       End If
245                       If rsLaporan!K40 > 0 Then
246                           nilaiTempTotal(4) = rsLaporan!K40
247                       End If
248                       If rsLaporan!K50 > 0 Then
249                           nilaiTempTotal(5) = rsLaporan!K50
250                       End If
251                       If rsLaporan!K60 > 0 Then
252                           nilaiTempTotal(6) = rsLaporan!K60
253                       End If
254                       If rsLaporan!K70 > 0 Then
255                           nilaiTempTotal(7) = rsLaporan!K70
256                       End If
257                       If rsLaporan!K80 > 0 Then
258                           nilaiTempTotal(8) = rsLaporan!K80
259                       End If
260                   End If
261               End If
                  
262               With rsLaporan
263                   .AddNew
264                   If nilaiTempDemand(0) > 0 Then
265                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
266                   Else
267                       !K1 = 0 & " %"
268                   End If
269                   If nilaiTempDemand(1) > 0 Then
270                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
271                   Else
272                       !K10 = 0 & " %"
273                   End If
274                   If nilaiTempDemand(2) > 0 Then
275                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
276                   Else
277                       !K20 = 0 & " %"
278                   End If
279                   If nilaiTempDemand(3) > 0 Then
280                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
281                   Else
282                       !K30 = 0 & " %"
283                   End If
284                   If nilaiTempDemand(4) > 0 Then
285                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
286                   Else
287                       !K40 = 0 & " %"
288                   End If
289                   If nilaiTempDemand(5) > 0 Then
290                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
291                   Else
292                       !K50 = 0 & " %"
293                   End If
294                   If nilaiTempDemand(6) > 0 Then
295                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
296                   Else
297                       !K60 = 0 & " %"
298                   End If
299                   If nilaiTempDemand(7) > 0 Then
300                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
301                   Else
302                       !K70 = 0 & " %"
303                   End If
304                   If nilaiTempDemand(8) > 0 Then
305                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
306                   Else
307                       !K80 = 0 & " %"
308                   End If
                      
309                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
310                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
311                   Else
312                       !tot16 = 0 & " %"
313                   End If
                      
314                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
315                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
316                   Else
317                       !tot26 = 0 & " %"
318                   End If
                       
319                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
320                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
321                   Else
322                       !tot19 = 0 & " %"
323                   End If
                       
324                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
325                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
326                   Else
327                       !tot29 = 0 & " %"
328                   End If
                       
329                   !Urut = 3
330                   !judul = "Retention Rate Same Group (NCM)"
331                   !model = Nul(rsModel.Fields(0).value)
332                   .Update
333               End With
                          
                   
334               With rsLaporan
335                   .AddNew
336                   If Not rsSBE.EOF Then
337                       rsSBE.filter = ""
338                       nilaiTemp = 0
339                       rsSBE.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
340                       If Not rsSBE.EOF Then
341                           Do While Not rsSBE.EOF
342                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
343                               rsSBE.MoveNext
344                           Loop
345                       End If
346                       !K1 = nilaiTemp
347                       rsSBE.filter = ""
348                       nilaiTemp = 0
                          
349                       rsSBE.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
350                       If Not rsSBE.EOF Then
351                           Do While Not rsSBE.EOF
352                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
353                               rsSBE.MoveNext
354                           Loop
355                       End If
356                       !K10 = nilaiTemp
357                       rsSBE.filter = ""
358                       nilaiTemp = 0
                          
359                       rsSBE.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
360                       If Not rsSBE.EOF Then
361                           Do While Not rsSBE.EOF
362                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
363                               rsSBE.MoveNext
364                           Loop
365                       End If
366                       !K20 = nilaiTemp
367                       rsSBE.filter = ""
368                       nilaiTemp = 0
                          
369                       rsSBE.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
370                       If Not rsSBE.EOF Then
371                           Do While Not rsSBE.EOF
372                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
373                               rsSBE.MoveNext
374                           Loop
375                       End If
376                       !K30 = nilaiTemp
377                       rsSBE.filter = ""
378                       nilaiTemp = 0
                          
379                       rsSBE.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
380                       If Not rsSBE.EOF Then
381                           Do While Not rsSBE.EOF
382                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
383                               rsSBE.MoveNext
384                           Loop
385                       End If
386                       !K40 = nilaiTemp
387                       rsSBE.filter = ""
388                       nilaiTemp = 0
                          
389                       rsSBE.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
390                       If Not rsSBE.EOF Then
391                           Do While Not rsSBE.EOF
392                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
393                               rsSBE.MoveNext
394                           Loop
395                       End If
396                       !K50 = nilaiTemp
397                       rsSBE.filter = ""
398                       nilaiTemp = 0
                          
399                       rsSBE.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
400                       If Not rsSBE.EOF Then
401                           Do While Not rsSBE.EOF
402                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
403                               rsSBE.MoveNext
404                           Loop
405                       End If
406                       !K60 = nilaiTemp
407                       rsSBE.filter = ""
408                       nilaiTemp = 0
                          
409                       rsSBE.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
410                       If Not rsSBE.EOF Then
411                           Do While Not rsSBE.EOF
412                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
413                               rsSBE.MoveNext
414                           Loop
415                       End If
416                       !K70 = nilaiTemp
417                       rsSBE.filter = ""
418                       nilaiTemp = 0
                          
419                       rsSBE.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
420                       If Not rsSBE.EOF Then
421                           Do While Not rsSBE.EOF
422                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
423                               rsSBE.MoveNext
424                           Loop
425                       End If
426                       !K80 = nilaiTemp
427                       rsSBE.filter = ""
428                       nilaiTemp = 0
                          
429                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
430                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
431                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
432                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
433                   Else
434                       !K1 = 0
435                       !K10 = 0
436                       !K20 = 0
437                       !K30 = 0
438                       !K40 = 0
439                       !K50 = 0
440                       !K60 = 0
441                       !K70 = 0
442                       !K80 = 0
443                       !tot16 = 0
444                       !tot19 = 0
445                       !tot29 = 0
446                       !tot26 = 0
447                   End If
                      
448                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
449                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
450                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
451                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
452                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
453                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
454                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
455                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
456                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
457                   !judul = "SB Total (NCM & Non NCM)"
458                   !model = Nul(rsModel.Fields(0).value)
459                   !Urut = 4
460                   .Update
461               End With
462               rsLaporan.filter = ""
                  
                  
463               For i = 0 To 10
464                   nilaiTempTotal(i) = 0
465               Next
466               rsLaporan.filter = ""
                  
467               If Not rsLaporan.EOF Then
468                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
469                   If Not rsLaporan.EOF Then
470                       If rsLaporan!K1 > 0 Then
471                           nilaiTempTotal(0) = rsLaporan!K1
472                       End If
473                       If rsLaporan!K10 > 0 Then
474                           nilaiTempTotal(1) = rsLaporan!K10
475                       End If
476                       If rsLaporan!K20 > 0 Then
477                           nilaiTempTotal(2) = rsLaporan!K20
478                       End If
479                       If rsLaporan!K30 > 0 Then
480                           nilaiTempTotal(3) = rsLaporan!K30
481                       End If
482                       If rsLaporan!K40 > 0 Then
483                           nilaiTempTotal(4) = rsLaporan!K40
484                       End If
485                       If rsLaporan!K50 > 0 Then
486                           nilaiTempTotal(5) = rsLaporan!K50
487                       End If
488                       If rsLaporan!K60 > 0 Then
489                           nilaiTempTotal(6) = rsLaporan!K60
490                       End If
491                       If rsLaporan!K70 > 0 Then
492                           nilaiTempTotal(7) = rsLaporan!K70
493                       End If
494                       If rsLaporan!K80 > 0 Then
495                           nilaiTempTotal(8) = rsLaporan!K80
496                       End If
497                   End If
498               End If
                  
499               With rsLaporan
500                   .AddNew
501                   If nilaiTempDemand(0) > 0 Then
502                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
503                   Else
504                       !K1 = 0 & " %"
505                   End If
506                   If nilaiTempDemand(1) > 0 Then
507                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
508                   Else
509                       !K10 = 0 & " %"
510                   End If
511                   If nilaiTempDemand(2) > 0 Then
512                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
513                   Else
514                       !K20 = 0 & " %"
515                   End If
516                   If nilaiTempDemand(3) > 0 Then
517                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
518                   Else
519                       !K30 = 0 & " %"
520                   End If
521                   If nilaiTempDemand(4) > 0 Then
522                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
523                   Else
524                       !K40 = 0 & " %"
525                   End If
526                   If nilaiTempDemand(5) > 0 Then
527                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
528                   Else
529                       !K50 = 0 & " %"
530                   End If
531                   If nilaiTempDemand(6) > 0 Then
532                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
533                   Else
534                       !K60 = 0 & " %"
535                   End If
536                   If nilaiTempDemand(7) > 0 Then
537                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
538                   Else
539                       !K70 = 0 & " %"
540                   End If
541                   If nilaiTempDemand(8) > 0 Then
542                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
543                   Else
544                       !K80 = 0 & " %"
545                   End If
                      
                      
546                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
547                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
548                   Else
549                       !tot16 = 0 & " %"
550                   End If
                      
551                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
552                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
553                   Else
554                       !tot26 = 0 & " %"
555                   End If
                       
556                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
557                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
558                   Else
559                       !tot19 = 0 & " %"
560                   End If
                       
561                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
562                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
563                   Else
564                       !tot29 = 0 & " %"
565                   End If
                      
566                   !Urut = 5
567                   !judul = "Retention Rate Total (NCM & Non NCM)"
568                   !model = Nul(rsModel.Fields(0).value)
569                   .Update
570               End With
571               rsLaporan.filter = ""
572               rsModel.MoveNext
573               ProgVin.value = ProgVin.value + 1
574               DoEvents
575           Loop
576       Else
577           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
578           Exit Sub
579       End If
          
580       Call addTotalRow
          
581       Report.database.SetDataSource rsLaporan
582       ProgVin.Visible = False
583       Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
584       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
585       With CRViewer1
586           .ReportSource = Report
587           .ViewReport
588           .Zoom 100
589           .Visible = True
590           .Left = 300
591           .Height = Me.Height - 1200
592           .Width = Me.Width - 750
593       End With
         
594       cmdTutup.Visible = True
595       cmdTutup.Left = CRViewer1.Left + 20
596       cmdTutup.Top = CRViewer1.Top + 400
597       Screen.MousePointer = vbDefault
          
598       Call DisableExport
599       Call setLokasiTombol
End Sub


Private Sub ambilLaporanVinRetentionTglDeliveryDealer()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         Set Report = New rptSBERateByModelVinRetention
2         Call createTableLaporanVin
          
3         sSQL = "select distinct FS_GRP_MDL as FS_GRP_MDL from model where FS_LOG_DEL = '' order by FS_GRP_MDL"
4         Set rsModel = objO.FetchDataHO(sSQL)
          
          
5         dealerPilih = Replace(dealerPilih, "t.fs_kd_dealer", "D.fs_kd_dealer")
          
         
6         sSQL = "select E.FS_GRP_MDL,COUNT(noinv) total,fs_rc,case when isnull(d.FS_KD_DEALER,'') <> '' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
7         sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK) "
8         sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  "
9         sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
10        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''" & dealerPilih
11        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
12        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
13        sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [PDS00]..IntrnCbg WHERE Proses = 10 and Kode = 2 AND Tanggal between '" & Format(tglAwalVin, "mm/dd/yyyy") & "' AND '" & Format(tglAkhirVin, "mm/dd/yyyy") & "' )"
14        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
15        sSQL = sSQL & " grOUP BY FS_RC ,E.FS_GRP_MDL, case when isnull(d.FS_KD_DEALER,'') <> '' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
16        sSQL = sSQL & " order by FS_RC,E.FS_GRP_MDL"
           
17        Set rsSBE = objO.FetchDataHO(sSQL)
18        If Not rsModel.EOF Then
19            rsModel.MoveFirst
20            ProgVin.Max = rsModel.RecordCount + 1
21            ProgVin.value = 1
22            ProgVin.Visible = True
23            DoEvents
24            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
25                With rsLaporan
26                    .AddNew
27                    rsSBE.filter = ""
28                    If Not rsSBE.EOF Then
29                        rsSBE.filter = ""
30                        nilaiTemp = 0
31                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
32                        If Not rsSBE.EOF Then
33                            Do While Not rsSBE.EOF
34                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
35                                rsSBE.MoveNext
36                            Loop
37                        End If
38                        !K1 = nilaiTemp
39                        rsSBE.filter = ""
40                        nilaiTemp = 0
                          
41                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
42                        If Not rsSBE.EOF Then
43                            Do While Not rsSBE.EOF
44                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
45                                rsSBE.MoveNext
46                            Loop
47                        End If
48                        !K10 = nilaiTemp
49                        rsSBE.filter = ""
50                        nilaiTemp = 0
                          
51                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
52                        If Not rsSBE.EOF Then
53                            Do While Not rsSBE.EOF
54                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
55                                rsSBE.MoveNext
56                            Loop
57                        End If
58                        !K20 = nilaiTemp
59                        rsSBE.filter = ""
60                        nilaiTemp = 0
                          
61                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
62                        If Not rsSBE.EOF Then
63                            Do While Not rsSBE.EOF
64                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
65                                rsSBE.MoveNext
66                            Loop
67                        End If
68                        !K30 = nilaiTemp
69                        rsSBE.filter = ""
70                        nilaiTemp = 0
                          
71                        rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
72                        If Not rsSBE.EOF Then
73                            Do While Not rsSBE.EOF
74                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
75                                rsSBE.MoveNext
76                            Loop
77                        End If
78                        !K40 = nilaiTemp
79                        rsSBE.filter = ""
80                        nilaiTemp = 0
                          
81                        rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
82                        If Not rsSBE.EOF Then
83                            Do While Not rsSBE.EOF
84                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
85                                rsSBE.MoveNext
86                            Loop
87                        End If
88                        !K50 = nilaiTemp
89                        rsSBE.filter = ""
90                        nilaiTemp = 0
                          
91                        rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
92                        If Not rsSBE.EOF Then
93                            Do While Not rsSBE.EOF
94                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
95                                rsSBE.MoveNext
96                            Loop
97                        End If
98                        !K60 = nilaiTemp
99                        rsSBE.filter = ""
100                       nilaiTemp = 0
                          
101                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
102                       If Not rsSBE.EOF Then
103                           Do While Not rsSBE.EOF
104                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
105                               rsSBE.MoveNext
106                           Loop
107                       End If
108                       !K70 = nilaiTemp
109                       rsSBE.filter = ""
110                       nilaiTemp = 0
                          
111                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
112                       If Not rsSBE.EOF Then
113                           Do While Not rsSBE.EOF
114                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
115                               rsSBE.MoveNext
116                           Loop
117                       End If
118                       !K80 = nilaiTemp
119                       rsSBE.filter = ""
120                       nilaiTemp = 0
                          
121                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
122                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
123                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
124                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
125                   Else
126                       !K1 = 0
127                       !K10 = 0
128                       !K20 = 0
129                       !K30 = 0
130                       !K40 = 0
131                       !K50 = 0
132                       !K60 = 0
133                       !K70 = 0
134                       !K80 = 0
135                       !tot16 = 0
136                       !tot19 = 0
                          
137                       !tot26 = 0
138                       !tot29 = 0
                          
139                   End If
                                      
140                   nilaiTotal1(0) = nilaiTotal1(0) + !K1
141                   nilaiTotal1(1) = nilaiTotal1(1) + !K10
142                   nilaiTotal1(2) = nilaiTotal1(2) + !K20
143                   nilaiTotal1(3) = nilaiTotal1(3) + !K30
144                   nilaiTotal1(4) = nilaiTotal1(4) + !K40
145                   nilaiTotal1(5) = nilaiTotal1(5) + !K50
146                   nilaiTotal1(6) = nilaiTotal1(6) + !K60
147                   nilaiTotal1(7) = nilaiTotal1(7) + !K70
148                   nilaiTotal1(8) = nilaiTotal1(8) + !K80
                      
149                   !judul = "SB Same Outlet"
150                   !model = Nul(rsModel.Fields(0).value)
151                   !Urut = 1
152                   .Update
153               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME OUTLET
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
154               With rsLaporan
155                   .AddNew
                  
156                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
157                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
158                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
159                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
160                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
161                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
162                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
163                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
164                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
165                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
166                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
167                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
168                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
169                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
170                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
171                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
172                   sSQL = "select isnull(sum(a.total),0) total,isnull(b.fs_grp_mdl,'') model from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where " & cabangPilih & " a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglAwalVin), "0#") & "/01/" & year(tglAwalVin) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglAkhirVin), "0#") & "/01/" & year(tglAkhirVin) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a left join [bengkel" & Right(gDb, 2) & "]..model b on b.mdl = a.kodetype where isnull(b.fs_grp_mdl,'') = '" & Nul(rsModel.Fields(0).value) & "' group by b.fs_grp_mdl"
173                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
                      
174                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
175                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
176                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
177                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
178                   nilaiTotal2(0) = nilaiTotal2(0) + !K1
179                   nilaiTotal2(1) = nilaiTotal2(1) + !K10
180                   nilaiTotal2(2) = nilaiTotal2(2) + !K20
181                   nilaiTotal2(3) = nilaiTotal2(3) + !K30
182                   nilaiTotal2(4) = nilaiTotal2(4) + !K40
183                   nilaiTotal2(5) = nilaiTotal2(5) + !K50
184                   nilaiTotal2(6) = nilaiTotal2(6) + !K60
185                   nilaiTotal2(7) = nilaiTotal2(7) + !K70
186                   nilaiTotal2(8) = nilaiTotal2(8) + !K80
                       
187                   !judul = "Demand SB Same Outlet"
188                   !model = Nul(rsModel.Fields(0).value)
189                   !Urut = 2
190                   .Update
191               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
192               For i = 0 To 10
193                   nilaiTempDemand(i) = 0
194                   nilaiTempTotal(i) = 0
195               Next
196               rsLaporan.filter = ""
197               If Not rsLaporan.EOF Then
198                   rsLaporan.filter = "judul = 'Demand SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
199                   If Not rsLaporan.EOF Then
200                       If rsLaporan!K1 > 0 Then
201                           nilaiTempDemand(0) = rsLaporan!K1
202                       End If
203                       If rsLaporan!K10 > 0 Then
204                           nilaiTempDemand(1) = rsLaporan!K10
205                       End If
206                       If rsLaporan!K20 > 0 Then
207                           nilaiTempDemand(2) = rsLaporan!K20
208                       End If
209                       If rsLaporan!K30 > 0 Then
210                           nilaiTempDemand(3) = rsLaporan!K30
211                       End If
212                       If rsLaporan!K40 > 0 Then
213                           nilaiTempDemand(4) = rsLaporan!K40
214                       End If
215                       If rsLaporan!K50 > 0 Then
216                           nilaiTempDemand(5) = rsLaporan!K50
217                       End If
218                       If rsLaporan!K60 > 0 Then
219                           nilaiTempDemand(6) = rsLaporan!K60
220                       End If
221                       If rsLaporan!K70 > 0 Then
222                           nilaiTempDemand(7) = rsLaporan!K70
223                       End If
224                       If rsLaporan!K80 > 0 Then
225                           nilaiTempDemand(8) = rsLaporan!K80
226                       End If
227                   End If
228               End If
229               rsLaporan.filter = ""
                  
230               If Not rsLaporan.EOF Then
231                   rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
232                   If Not rsLaporan.EOF Then
233                       If rsLaporan!K1 > 0 Then
234                           nilaiTempTotal(0) = rsLaporan!K1
235                       End If
236                       If rsLaporan!K10 > 0 Then
237                           nilaiTempTotal(1) = rsLaporan!K10
238                       End If
239                       If rsLaporan!K20 > 0 Then
240                           nilaiTempTotal(2) = rsLaporan!K20
241                       End If
242                       If rsLaporan!K30 > 0 Then
243                           nilaiTempTotal(3) = rsLaporan!K30
244                       End If
245                       If rsLaporan!K40 > 0 Then
246                           nilaiTempTotal(4) = rsLaporan!K40
247                       End If
248                       If rsLaporan!K50 > 0 Then
249                           nilaiTempTotal(5) = rsLaporan!K50
250                       End If
251                       If rsLaporan!K60 > 0 Then
252                           nilaiTempTotal(6) = rsLaporan!K60
253                       End If
254                       If rsLaporan!K70 > 0 Then
255                           nilaiTempTotal(7) = rsLaporan!K70
256                       End If
257                       If rsLaporan!K80 > 0 Then
258                           nilaiTempTotal(8) = rsLaporan!K80
259                       End If
260                   End If
261               End If
                  
262               With rsLaporan
263                   .AddNew
264                   If nilaiTempDemand(0) > 0 Then
265                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
266                   Else
267                       !K1 = 0 & " %"
268                   End If
269                   If nilaiTempDemand(1) > 0 Then
270                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
271                   Else
272                       !K10 = 0 & " %"
273                   End If
274                   If nilaiTempDemand(2) > 0 Then
275                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
276                   Else
277                       !K20 = 0 & " %"
278                   End If
279                   If nilaiTempDemand(3) > 0 Then
280                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
281                   Else
282                       !K30 = 0 & " %"
283                   End If
284                   If nilaiTempDemand(4) > 0 Then
285                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
286                   Else
287                       !K40 = 0 & " %"
288                   End If
289                   If nilaiTempDemand(5) > 0 Then
290                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
291                   Else
292                       !K50 = 0 & " %"
293                   End If
294                   If nilaiTempDemand(6) > 0 Then
295                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
296                   Else
297                       !K60 = 0 & " %"
298                   End If
299                   If nilaiTempDemand(7) > 0 Then
300                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
301                   Else
302                       !K70 = 0 & " %"
303                   End If
304                   If nilaiTempDemand(8) > 0 Then
305                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
306                   Else
307                       !K80 = 0 & " %"
308                   End If
                      
309                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
310                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
311                   Else
312                       !tot16 = 0 & " %"
313                   End If
                      
314                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
315                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
316                   Else
317                       !tot26 = 0 & " %"
318                   End If
                       
319                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
320                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
321                   Else
322                       !tot19 = 0 & " %"
323                   End If
                       
324                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
325                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
326                   Else
327                       !tot29 = 0 & " %"
328                   End If
                       
329                   !Urut = 3
330                   !judul = "Retention Rate Same Outlate"
331                   !model = Nul(rsModel.Fields(0).value)
332                   .Update
333               End With
                  
334               With rsLaporan
335                   .AddNew
336                   rsSBE.filter = ""
337                   If Not rsSBE.EOF Then
338                       rsSBE.filter = ""
339                       nilaiTemp = 0
340                       rsSBE.filter = "fs_rc = 'R001' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
341                       If Not rsSBE.EOF Then
342                           Do While Not rsSBE.EOF
343                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
344                               rsSBE.MoveNext
345                           Loop
346                       End If
347                       !K1 = nilaiTemp
348                       rsSBE.filter = ""
349                       nilaiTemp = 0
                          
350                       rsSBE.filter = "fs_rc = 'R010' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
351                       If Not rsSBE.EOF Then
352                           Do While Not rsSBE.EOF
353                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
354                               rsSBE.MoveNext
355                           Loop
356                       End If
357                       !K10 = nilaiTemp
358                       rsSBE.filter = ""
359                       nilaiTemp = 0
                          
360                       rsSBE.filter = "fs_rc = 'R020' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
361                       If Not rsSBE.EOF Then
362                           Do While Not rsSBE.EOF
363                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
364                               rsSBE.MoveNext
365                           Loop
366                       End If
367                       !K20 = nilaiTemp
368                       rsSBE.filter = ""
369                       nilaiTemp = 0
                          
370                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
371                       If Not rsSBE.EOF Then
372                           Do While Not rsSBE.EOF
373                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
374                               rsSBE.MoveNext
375                           Loop
376                       End If
377                       !K30 = nilaiTemp
378                       rsSBE.filter = ""
379                       nilaiTemp = 0
                          
380                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
381                       If Not rsSBE.EOF Then
382                           Do While Not rsSBE.EOF
383                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
384                               rsSBE.MoveNext
385                           Loop
386                       End If
387                       !K40 = nilaiTemp
388                       rsSBE.filter = ""
389                       nilaiTemp = 0
                          
390                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
391                       If Not rsSBE.EOF Then
392                           Do While Not rsSBE.EOF
393                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
394                               rsSBE.MoveNext
395                           Loop
396                       End If
397                       !K50 = nilaiTemp
398                       rsSBE.filter = ""
399                       nilaiTemp = 0
                          
400                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
401                       If Not rsSBE.EOF Then
402                           Do While Not rsSBE.EOF
403                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
404                               rsSBE.MoveNext
405                           Loop
406                       End If
407                       !K60 = nilaiTemp
408                       rsSBE.filter = ""
409                       nilaiTemp = 0
                          
410                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
411                       If Not rsSBE.EOF Then
412                           Do While Not rsSBE.EOF
413                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
414                               rsSBE.MoveNext
415                           Loop
416                       End If
417                       !K70 = nilaiTemp
418                       rsSBE.filter = ""
419                       nilaiTemp = 0
                          
420                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'NCM' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
421                       If Not rsSBE.EOF Then
422                           Do While Not rsSBE.EOF
423                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
424                               rsSBE.MoveNext
425                           Loop
426                       End If
427                       !K80 = nilaiTemp
428                       rsSBE.filter = ""
429                       nilaiTemp = 0
                           
430                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
431                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
432                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
433                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
434                   Else
435                       !K1 = 0
436                       !K10 = 0
437                       !K20 = 0
438                       !K30 = 0
439                       !K40 = 0
440                       !K50 = 0
441                       !K60 = 0
442                       !K70 = 0
443                       !K80 = 0
444                       !tot16 = 0
445                       !tot19 = 0
                          
446                       !tot26 = 0
447                       !tot29 = 0
448                   End If
                      
449                   nilaiTotal4(0) = nilaiTotal4(0) + !K1
450                   nilaiTotal4(1) = nilaiTotal4(1) + !K10
451                   nilaiTotal4(2) = nilaiTotal4(2) + !K20
452                   nilaiTotal4(3) = nilaiTotal4(3) + !K30
453                   nilaiTotal4(4) = nilaiTotal4(4) + !K40
454                   nilaiTotal4(5) = nilaiTotal4(5) + !K50
455                   nilaiTotal4(6) = nilaiTotal4(6) + !K60
456                   nilaiTotal4(7) = nilaiTotal4(7) + !K70
457                   nilaiTotal4(8) = nilaiTotal4(8) + !K80
                      
458                   !judul = "SB Same Group (NCM)"
459                   !model = Nul(rsModel.Fields(0).value)
460                   !Urut = 4
461                   .Update
462               End With
                  
                  
463               For i = 0 To 10
464                   nilaiTempTotal(i) = 0
465               Next
466               rsLaporan.filter = ""
                  
467               If Not rsLaporan.EOF Then
468                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
469                   If Not rsLaporan.EOF Then
470                       If rsLaporan!K1 > 0 Then
471                           nilaiTempTotal(0) = rsLaporan!K1
472                       End If
473                       If rsLaporan!K10 > 0 Then
474                           nilaiTempTotal(1) = rsLaporan!K10
475                       End If
476                       If rsLaporan!K20 > 0 Then
477                           nilaiTempTotal(2) = rsLaporan!K20
478                       End If
479                       If rsLaporan!K30 > 0 Then
480                           nilaiTempTotal(3) = rsLaporan!K30
481                       End If
482                       If rsLaporan!K40 > 0 Then
483                           nilaiTempTotal(4) = rsLaporan!K40
484                       End If
485                       If rsLaporan!K50 > 0 Then
486                           nilaiTempTotal(5) = rsLaporan!K50
487                       End If
488                       If rsLaporan!K60 > 0 Then
489                           nilaiTempTotal(6) = rsLaporan!K60
490                       End If
491                       If rsLaporan!K70 > 0 Then
492                           nilaiTempTotal(7) = rsLaporan!K70
493                       End If
494                       If rsLaporan!K80 > 0 Then
495                           nilaiTempTotal(8) = rsLaporan!K80
496                       End If
497                   End If
498               End If
                  
499               With rsLaporan
500                   .AddNew
501                   If nilaiTempDemand(0) > 0 Then
502                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
503                   Else
504                       !K1 = 0 & " %"
505                   End If
506                   If nilaiTempDemand(1) > 0 Then
507                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
508                   Else
509                       !K10 = 0 & " %"
510                   End If
511                   If nilaiTempDemand(2) > 0 Then
512                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
513                   Else
514                       !K20 = 0 & " %"
515                   End If
516                   If nilaiTempDemand(3) > 0 Then
517                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
518                   Else
519                       !K30 = 0 & " %"
520                   End If
521                   If nilaiTempDemand(4) > 0 Then
522                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
523                   Else
524                       !K40 = 0 & " %"
525                   End If
526                   If nilaiTempDemand(5) > 0 Then
527                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
528                   Else
529                       !K50 = 0 & " %"
530                   End If
531                   If nilaiTempDemand(6) > 0 Then
532                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
533                   Else
534                       !K60 = 0 & " %"
535                   End If
536                   If nilaiTempDemand(7) > 0 Then
537                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
538                   Else
539                       !K70 = 0 & " %"
540                   End If
541                   If nilaiTempDemand(8) > 0 Then
542                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
543                   Else
544                       !K80 = 0 & " %"
545                   End If
                      
                      
546                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
547                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
548                   Else
549                       !tot16 = 0 & " %"
550                   End If
                      
551                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
552                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
553                   Else
554                       !tot26 = 0 & " %"
555                   End If
                       
556                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
557                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
558                   Else
559                       !tot19 = 0 & " %"
560                   End If
                       
561                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
562                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
563                   Else
564                       !tot29 = 0 & " %"
565                   End If
                      
566                   !Urut = 5
567                   !judul = "Retention Rate Same Group (NCM)"
568                   !model = Nul(rsModel.Fields(0).value)
569                   .Update
570               End With
                   
571               With rsLaporan
572                   .AddNew
573                   If Not rsSBE.EOF Then
574                       rsSBE.filter = ""
575                       nilaiTemp = 0
576                       rsSBE.filter = "fs_rc = 'R001' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
577                       If Not rsSBE.EOF Then
578                           Do While Not rsSBE.EOF
579                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
580                               rsSBE.MoveNext
581                           Loop
582                       End If
583                       !K1 = nilaiTemp
584                       rsSBE.filter = ""
585                       nilaiTemp = 0
                          
586                       rsSBE.filter = "fs_rc = 'R010' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
587                       If Not rsSBE.EOF Then
588                           Do While Not rsSBE.EOF
589                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
590                               rsSBE.MoveNext
591                           Loop
592                       End If
593                       !K10 = nilaiTemp
594                       rsSBE.filter = ""
595                       nilaiTemp = 0
                          
596                       rsSBE.filter = "fs_rc = 'R020' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
597                       If Not rsSBE.EOF Then
598                           Do While Not rsSBE.EOF
599                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
600                               rsSBE.MoveNext
601                           Loop
602                       End If
603                       !K20 = nilaiTemp
604                       rsSBE.filter = ""
605                       nilaiTemp = 0
                          
606                       rsSBE.filter = "fs_rc = 'R030' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
607                       If Not rsSBE.EOF Then
608                           Do While Not rsSBE.EOF
609                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
610                               rsSBE.MoveNext
611                           Loop
612                       End If
613                       !K30 = nilaiTemp
614                       rsSBE.filter = ""
615                       nilaiTemp = 0
                          
616                       rsSBE.filter = "fs_rc = 'R040' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
617                       If Not rsSBE.EOF Then
618                           Do While Not rsSBE.EOF
619                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
620                               rsSBE.MoveNext
621                           Loop
622                       End If
623                       !K40 = nilaiTemp
624                       rsSBE.filter = ""
625                       nilaiTemp = 0
                          
626                       rsSBE.filter = "fs_rc = 'R050' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
627                       If Not rsSBE.EOF Then
628                           Do While Not rsSBE.EOF
629                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
630                               rsSBE.MoveNext
631                           Loop
632                       End If
633                       !K50 = nilaiTemp
634                       rsSBE.filter = ""
635                       nilaiTemp = 0
                          
636                       rsSBE.filter = "fs_rc = 'R060' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
637                       If Not rsSBE.EOF Then
638                           Do While Not rsSBE.EOF
639                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
640                               rsSBE.MoveNext
641                           Loop
642                       End If
643                       !K60 = nilaiTemp
644                       rsSBE.filter = ""
645                       nilaiTemp = 0
                          
646                       rsSBE.filter = "fs_rc = 'R070' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
647                       If Not rsSBE.EOF Then
648                           Do While Not rsSBE.EOF
649                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
650                               rsSBE.MoveNext
651                           Loop
652                       End If
653                       !K70 = nilaiTemp
654                       rsSBE.filter = ""
655                       nilaiTemp = 0
                          
656                       rsSBE.filter = "fs_rc = 'R080' and fs_grp_mdl = '" & Nul(rsModel!fs_grp_mdl) & "'"
657                       If Not rsSBE.EOF Then
658                           Do While Not rsSBE.EOF
659                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
660                               rsSBE.MoveNext
661                           Loop
662                       End If
663                       !K80 = nilaiTemp
664                       rsSBE.filter = ""
665                       nilaiTemp = 0
                          
666                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
667                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
668                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
669                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
670                   Else
671                       !K1 = 0
672                       !K10 = 0
673                       !K20 = 0
674                       !K30 = 0
675                       !K40 = 0
676                       !K50 = 0
677                       !K60 = 0
678                       !K70 = 0
679                       !K80 = 0
680                       !tot16 = 0
681                       !tot19 = 0
682                       !tot29 = 0
683                       !tot26 = 0
684                   End If
                      
685                   nilaiTotal6(0) = nilaiTotal6(0) + !K1
686                   nilaiTotal6(1) = nilaiTotal6(1) + !K10
687                   nilaiTotal6(2) = nilaiTotal6(2) + !K20
688                   nilaiTotal6(3) = nilaiTotal6(3) + !K30
689                   nilaiTotal6(4) = nilaiTotal6(4) + !K40
690                   nilaiTotal6(5) = nilaiTotal6(5) + !K50
691                   nilaiTotal6(6) = nilaiTotal6(6) + !K60
692                   nilaiTotal6(7) = nilaiTotal6(7) + !K70
693                   nilaiTotal6(8) = nilaiTotal6(8) + !K80
                      
694                   !judul = "SB Total (NCM & Non NCM)"
695                   !model = Nul(rsModel.Fields(0).value)
696                   !Urut = 6
697                   .Update
698               End With
699               rsLaporan.filter = ""
                  
                  
700               For i = 0 To 10
701                   nilaiTempTotal(i) = 0
702               Next
703               rsLaporan.filter = ""
                  
704               If Not rsLaporan.EOF Then
705                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
706                   If Not rsLaporan.EOF Then
707                       If rsLaporan!K1 > 0 Then
708                           nilaiTempTotal(0) = rsLaporan!K1
709                       End If
710                       If rsLaporan!K10 > 0 Then
711                           nilaiTempTotal(1) = rsLaporan!K10
712                       End If
713                       If rsLaporan!K20 > 0 Then
714                           nilaiTempTotal(2) = rsLaporan!K20
715                       End If
716                       If rsLaporan!K30 > 0 Then
717                           nilaiTempTotal(3) = rsLaporan!K30
718                       End If
719                       If rsLaporan!K40 > 0 Then
720                           nilaiTempTotal(4) = rsLaporan!K40
721                       End If
722                       If rsLaporan!K50 > 0 Then
723                           nilaiTempTotal(5) = rsLaporan!K50
724                       End If
725                       If rsLaporan!K60 > 0 Then
726                           nilaiTempTotal(6) = rsLaporan!K60
727                       End If
728                       If rsLaporan!K70 > 0 Then
729                           nilaiTempTotal(7) = rsLaporan!K70
730                       End If
731                       If rsLaporan!K80 > 0 Then
732                           nilaiTempTotal(8) = rsLaporan!K80
733                       End If
734                   End If
735               End If
                  
736               With rsLaporan
737                   .AddNew
738                   If nilaiTempDemand(0) > 0 Then
739                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
740                   Else
741                       !K1 = 0 & " %"
742                   End If
743                   If nilaiTempDemand(1) > 0 Then
744                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
745                   Else
746                       !K10 = 0 & " %"
747                   End If
748                   If nilaiTempDemand(2) > 0 Then
749                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
750                   Else
751                       !K20 = 0 & " %"
752                   End If
753                   If nilaiTempDemand(3) > 0 Then
754                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
755                   Else
756                       !K30 = 0 & " %"
757                   End If
758                   If nilaiTempDemand(4) > 0 Then
759                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
760                   Else
761                       !K40 = 0 & " %"
762                   End If
763                   If nilaiTempDemand(5) > 0 Then
764                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
765                   Else
766                       !K50 = 0 & " %"
767                   End If
768                   If nilaiTempDemand(6) > 0 Then
769                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
770                   Else
771                       !K60 = 0 & " %"
772                   End If
773                   If nilaiTempDemand(7) > 0 Then
774                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
775                   Else
776                       !K70 = 0 & " %"
777                   End If
778                   If nilaiTempDemand(8) > 0 Then
779                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
780                   Else
781                       !K80 = 0 & " %"
782                   End If
                      
                      
783                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
784                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
785                   Else
786                       !tot16 = 0 & " %"
787                   End If
                      
788                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
789                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
790                   Else
791                       !tot26 = 0 & " %"
792                   End If
                       
793                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
794                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
795                   Else
796                       !tot19 = 0 & " %"
797                   End If
                       
798                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
799                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
800                   Else
801                       !tot29 = 0 & " %"
802                   End If
                      
803                   !Urut = 7
804                   !judul = "Retention Rate Total (NCM & Non NCM)"
805                   !model = Nul(rsModel.Fields(0).value)
806                   .Update
807               End With
808               rsLaporan.filter = ""
809               rsModel.MoveNext
810               ProgVin.value = ProgVin.value + 1
811               DoEvents
812           Loop
813       Else
814           MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
815           Exit Sub
816       End If
          
817       Call addTotalRowDealer
          
818       Report.database.SetDataSource rsLaporan
819       ProgVin.Visible = False
820       Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
821       Report.txtPeriode.SetText "Periode : " & Format(tglAwalVin, "dd-mm-yyyy") & " s/d " & Format(tglAkhirVin, "dd-mm-yyyy")
          
822       With CRViewer1
823           .ReportSource = Report
824           .ViewReport
825           .Zoom 100
826           .Visible = True
827           .Left = 300
828           .Height = Me.Height - 1200
829           .Width = Me.Width - 750
830       End With
         
831       cmdTutup.Visible = True
832       cmdTutup.Left = CRViewer1.Left + 20
833       cmdTutup.Top = CRViewer1.Top + 400
834       Screen.MousePointer = vbDefault
          
835       Call DisableExport
836       Call setLokasiTombol
End Sub


Private Sub ambilLaporanVinDealerTglInvoice()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         For i = 0 To 10
2             nilaiSBNRMPareto(i) = 0
3             nilaiSBNRMNonPareto(i) = 0
4             nilaiSBNRMSouth(i) = 0
5             nilaiSBNRMNorth(i) = 0
              
6             nilaiDemandPareto(i) = 0
7             nilaiDemandNonPareto(i) = 0
8             nilaiDemandSouth(i) = 0
9             nilaiDemandNorth(i) = 0
              
10            nilaiTotalNonPareto(i) = 0
11            nilaiTotalPareto(i) = 0
12            nilaiTotalSouth(i) = 0
13            nilaiTotalNorth(i) = 0
              
14            nilaiSBSameOutletPareto(i) = 0
15            nilaiSBSameOutletNonPareto(i) = 0
16            nilaiSBSameOutletSouth(i) = 0
17            nilaiSBSameOutletNorth(i) = 0
18        Next
          
19        Set Report = New rptSBERateByModelVinRetentionDealer
20        Call createTableLaporanVin
          
21        sSQL = "select FS_NM_DEALER,FS_KD_DEALER,pareto,north,kodecabang from T_DEALER where INTERNAL = 1 and FS_LOG_DEL = '' and FS_KD_DEALER <> 'NR 00' order by north desc, FS_KD_DEALER"
22        Set rsModel = objO.FetchDataHO(sSQL)
        
23        sSQL = "select COUNT(noinv) total,fs_rc,"
24        sSQL = sSQL & " E.FS_NM_DEALER"
25        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
26        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
27        sSQL = sSQL & " inner JOIN T_DEALER E WITH (NOLOCK) ON e.FS_KD_DEALER = t.FS_KD_DEALER and e.fs_log_del = '' and INTERNAL = 1"
28        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
29        sSQL = sSQL & " AND INVDATE between '" & Format(tglDealerAwal.value, "mm/dd/yyyy") & "' AND '" & Format(tglDealerAkhir.value, "mm/dd/yyyy") & "'"
30        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
31        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
32        sSQL = sSQL & " grOUP BY FS_RC, e.FS_NM_DEALER"
33        sSQL = sSQL & " order by FS_RC,e.FS_NM_DEALER"
           
34        Set rsSBETotal = objO.FetchDataHO(sSQL)
          
          
35        sSQL = "select COUNT(noinv) total,fs_rc,case when isnull(d.FS_KD_DEALER,'') <> '' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
36        sSQL = sSQL & " D.FS_NM_DEALER"
37        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
38        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
39        sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
40        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''"
41        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
42        sSQL = sSQL & " AND INVDATE between '" & Format(tglDealerAwal.value, "mm/dd/yyyy") & "' AND '" & Format(tglDealerAkhir.value, "mm/dd/yyyy") & "'"
43        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
44        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
45        sSQL = sSQL & " grOUP BY FS_RC , case when isnull(d.FS_KD_DEALER,'') <> '' then 'Same Outlet' when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end, d.FS_NM_DEALER"
46        sSQL = sSQL & " order by FS_RC,d.FS_NM_DEALER"
            
47        Set rsSBE = objO.FetchDataHO(sSQL)
48        If Not rsModel.EOF Then
49            rsModel.MoveFirst
50            progbarDealer.Max = rsModel.RecordCount + 1
51            progbarDealer.value = 1
52            progbarDealer.Visible = True
53            DoEvents
54            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
55                With rsLaporan
56                    .AddNew
57                    rsSBE.filter = ""
58                    If Not rsSBE.EOF Then
59                        rsSBE.filter = ""
60                        nilaiTemp = 0
61                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
62                        If Not rsSBE.EOF Then
63                            Do While Not rsSBE.EOF
64                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
65                                rsSBE.MoveNext
66                            Loop
67                        End If
68                        !K1 = nilaiTemp
69                        rsSBE.filter = ""
70                        nilaiTemp = 0
                          
71                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
72                        If Not rsSBE.EOF Then
73                            Do While Not rsSBE.EOF
74                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
75                                rsSBE.MoveNext
76                            Loop
77                        End If
78                        !K10 = nilaiTemp
79                        rsSBE.filter = ""
80                        nilaiTemp = 0
                          
81                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
82                        If Not rsSBE.EOF Then
83                            Do While Not rsSBE.EOF
84                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
85                                rsSBE.MoveNext
86                            Loop
87                        End If
88                        !K20 = nilaiTemp
89                        rsSBE.filter = ""
90                        nilaiTemp = 0
                          
91                        rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
92                        If Not rsSBE.EOF Then
93                            Do While Not rsSBE.EOF
94                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
95                                rsSBE.MoveNext
96                            Loop
97                        End If
98                        !K30 = nilaiTemp
99                        rsSBE.filter = ""
100                       nilaiTemp = 0
                          
101                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
102                       If Not rsSBE.EOF Then
103                           Do While Not rsSBE.EOF
104                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
105                               rsSBE.MoveNext
106                           Loop
107                       End If
108                       !K40 = nilaiTemp
109                       rsSBE.filter = ""
110                       nilaiTemp = 0
                          
111                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
112                       If Not rsSBE.EOF Then
113                           Do While Not rsSBE.EOF
114                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
115                               rsSBE.MoveNext
116                           Loop
117                       End If
118                       !K50 = nilaiTemp
119                       rsSBE.filter = ""
120                       nilaiTemp = 0
                          
121                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
122                       If Not rsSBE.EOF Then
123                           Do While Not rsSBE.EOF
124                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
125                               rsSBE.MoveNext
126                           Loop
127                       End If
128                       !K60 = nilaiTemp
129                       rsSBE.filter = ""
130                       nilaiTemp = 0
                          
131                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
132                       If Not rsSBE.EOF Then
133                           Do While Not rsSBE.EOF
134                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
135                               rsSBE.MoveNext
136                           Loop
137                       End If
138                       !K70 = nilaiTemp
139                       rsSBE.filter = ""
140                       nilaiTemp = 0
                          
141                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
142                       If Not rsSBE.EOF Then
143                           Do While Not rsSBE.EOF
144                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
145                               rsSBE.MoveNext
146                           Loop
147                       End If
148                       !K80 = nilaiTemp
149                       rsSBE.filter = ""
150                       nilaiTemp = 0
                          
151                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
152                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
153                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
154                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
155                   Else
156                       !K1 = 0
157                       !K10 = 0
158                       !K20 = 0
159                       !K30 = 0
160                       !K40 = 0
161                       !K50 = 0
162                       !K60 = 0
163                       !K70 = 0
164                       !K80 = 0
165                       !tot16 = 0
166                       !tot19 = 0
                          
167                       !tot26 = 0
168                       !tot29 = 0
                          
169                   End If
170                   If NulSID(rsModel!pareto) = 1 Then
171                       nilaiSBSameOutletPareto(0) = nilaiSBSameOutletPareto(0) + !K1
172                       nilaiSBSameOutletPareto(1) = nilaiSBSameOutletPareto(1) + !K10
173                       nilaiSBSameOutletPareto(2) = nilaiSBSameOutletPareto(2) + !K20
174                       nilaiSBSameOutletPareto(3) = nilaiSBSameOutletPareto(3) + !K30
175                       nilaiSBSameOutletPareto(4) = nilaiSBSameOutletPareto(4) + !K40
176                       nilaiSBSameOutletPareto(5) = nilaiSBSameOutletPareto(5) + !K50
177                       nilaiSBSameOutletPareto(6) = nilaiSBSameOutletPareto(6) + !K60
178                       nilaiSBSameOutletPareto(7) = nilaiSBSameOutletPareto(7) + !K70
179                       nilaiSBSameOutletPareto(8) = nilaiSBSameOutletPareto(8) + !K80
180                   Else
181                       nilaiSBSameOutletNonPareto(0) = nilaiSBSameOutletNonPareto(0) + !K1
182                       nilaiSBSameOutletNonPareto(1) = nilaiSBSameOutletNonPareto(1) + !K10
183                       nilaiSBSameOutletNonPareto(2) = nilaiSBSameOutletNonPareto(2) + !K20
184                       nilaiSBSameOutletNonPareto(3) = nilaiSBSameOutletNonPareto(3) + !K30
185                       nilaiSBSameOutletNonPareto(4) = nilaiSBSameOutletNonPareto(4) + !K40
186                       nilaiSBSameOutletNonPareto(5) = nilaiSBSameOutletNonPareto(5) + !K50
187                       nilaiSBSameOutletNonPareto(6) = nilaiSBSameOutletNonPareto(6) + !K60
188                       nilaiSBSameOutletNonPareto(7) = nilaiSBSameOutletNonPareto(7) + !K70
189                       nilaiSBSameOutletNonPareto(8) = nilaiSBSameOutletNonPareto(8) + !K80
190                   End If
                      
191                   If NulSID(rsModel!north) = 1 Then
192                       !north = 1
193                       nilaiSBSameOutletNorth(0) = nilaiSBSameOutletNorth(0) + !K1
194                       nilaiSBSameOutletNorth(1) = nilaiSBSameOutletNorth(1) + !K10
195                       nilaiSBSameOutletNorth(2) = nilaiSBSameOutletNorth(2) + !K20
196                       nilaiSBSameOutletNorth(3) = nilaiSBSameOutletNorth(3) + !K30
197                       nilaiSBSameOutletNorth(4) = nilaiSBSameOutletNorth(4) + !K40
198                       nilaiSBSameOutletNorth(5) = nilaiSBSameOutletNorth(5) + !K50
199                       nilaiSBSameOutletNorth(6) = nilaiSBSameOutletNorth(6) + !K60
200                       nilaiSBSameOutletNorth(7) = nilaiSBSameOutletNorth(7) + !K70
201                       nilaiSBSameOutletNorth(8) = nilaiSBSameOutletNorth(8) + !K80
202                   Else
                      
203                       !north = 3
204                       nilaiSBSameOutletSouth(0) = nilaiSBSameOutletSouth(0) + !K1
205                       nilaiSBSameOutletSouth(1) = nilaiSBSameOutletSouth(1) + !K10
206                       nilaiSBSameOutletSouth(2) = nilaiSBSameOutletSouth(2) + !K20
207                       nilaiSBSameOutletSouth(3) = nilaiSBSameOutletSouth(3) + !K30
208                       nilaiSBSameOutletSouth(4) = nilaiSBSameOutletSouth(4) + !K40
209                       nilaiSBSameOutletSouth(5) = nilaiSBSameOutletSouth(5) + !K50
210                       nilaiSBSameOutletSouth(6) = nilaiSBSameOutletSouth(6) + !K60
211                       nilaiSBSameOutletSouth(7) = nilaiSBSameOutletSouth(7) + !K70
212                       nilaiSBSameOutletSouth(8) = nilaiSBSameOutletSouth(8) + !K80
213                   End If
214                   !judul = "SB Same Outlet"
215                   !model = Nul(rsModel.Fields(0).value)
216                   !Urut = 1
217                   .Update
218               End With
                  
                  ''LAPORAN SAME GROUP
219               With rsLaporan
220                   .AddNew
221                   rsSBE.filter = ""
222                   If Not rsSBE.EOF Then
223                       rsSBE.filter = ""
224                       nilaiTemp = 0
225                       rsSBE.filter = "fs_rc = 'R001' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
226                       If Not rsSBE.EOF Then
227                           Do While Not rsSBE.EOF
228                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
229                               rsSBE.MoveNext
230                           Loop
231                       End If
232                       !K1 = nilaiTemp
233                       rsSBE.filter = ""
234                       nilaiTemp = 0
                          
235                       rsSBE.filter = "fs_rc = 'R010' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
236                       If Not rsSBE.EOF Then
237                           Do While Not rsSBE.EOF
238                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
239                               rsSBE.MoveNext
240                           Loop
241                       End If
242                       !K10 = nilaiTemp
243                       rsSBE.filter = ""
244                       nilaiTemp = 0
                          
245                       rsSBE.filter = "fs_rc = 'R020' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
246                       If Not rsSBE.EOF Then
247                           Do While Not rsSBE.EOF
248                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
249                               rsSBE.MoveNext
250                           Loop
251                       End If
252                       !K20 = nilaiTemp
253                       rsSBE.filter = ""
254                       nilaiTemp = 0
                          
255                       rsSBE.filter = "fs_rc = 'R030' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
256                       If Not rsSBE.EOF Then
257                           Do While Not rsSBE.EOF
258                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
259                               rsSBE.MoveNext
260                           Loop
261                       End If
262                       !K30 = nilaiTemp
263                       rsSBE.filter = ""
264                       nilaiTemp = 0
                          
265                       rsSBE.filter = "fs_rc = 'R040' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
266                       If Not rsSBE.EOF Then
267                           Do While Not rsSBE.EOF
268                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
269                               rsSBE.MoveNext
270                           Loop
271                       End If
272                       !K40 = nilaiTemp
273                       rsSBE.filter = ""
274                       nilaiTemp = 0
                          
275                       rsSBE.filter = "fs_rc = 'R050' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
276                       If Not rsSBE.EOF Then
277                           Do While Not rsSBE.EOF
278                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
279                               rsSBE.MoveNext
280                           Loop
281                       End If
282                       !K50 = nilaiTemp
283                       rsSBE.filter = ""
284                       nilaiTemp = 0
                          
285                       rsSBE.filter = "fs_rc = 'R060' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
286                       If Not rsSBE.EOF Then
287                           Do While Not rsSBE.EOF
288                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
289                               rsSBE.MoveNext
290                           Loop
291                       End If
292                       !K60 = nilaiTemp
293                       rsSBE.filter = ""
294                       nilaiTemp = 0
                          
295                       rsSBE.filter = "fs_rc = 'R070' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
296                       If Not rsSBE.EOF Then
297                           Do While Not rsSBE.EOF
298                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
299                               rsSBE.MoveNext
300                           Loop
301                       End If
302                       !K70 = nilaiTemp
303                       rsSBE.filter = ""
304                       nilaiTemp = 0
                          
305                       rsSBE.filter = "fs_rc = 'R080' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
306                       If Not rsSBE.EOF Then
307                           Do While Not rsSBE.EOF
308                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
309                               rsSBE.MoveNext
310                           Loop
311                       End If
312                       !K80 = nilaiTemp
313                       rsSBE.filter = ""
314                       nilaiTemp = 0
                          
315                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
316                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
317                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
318                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
319                   Else
320                       !K1 = 0
321                       !K10 = 0
322                       !K20 = 0
323                       !K30 = 0
324                       !K40 = 0
325                       !K50 = 0
326                       !K60 = 0
327                       !K70 = 0
328                       !K80 = 0
329                       !tot16 = 0
330                       !tot19 = 0
                          
331                       !tot26 = 0
332                       !tot29 = 0
                          
333                   End If
334                   If NulSID(rsModel!pareto) = 1 Then
335                       nilaiSBNRMPareto(0) = nilaiSBNRMPareto(0) + !K1
336                       nilaiSBNRMPareto(1) = nilaiSBNRMPareto(1) + !K10
337                       nilaiSBNRMPareto(2) = nilaiSBNRMPareto(2) + !K20
338                       nilaiSBNRMPareto(3) = nilaiSBNRMPareto(3) + !K30
339                       nilaiSBNRMPareto(4) = nilaiSBNRMPareto(4) + !K40
340                       nilaiSBNRMPareto(5) = nilaiSBNRMPareto(5) + !K50
341                       nilaiSBNRMPareto(6) = nilaiSBNRMPareto(6) + !K60
342                       nilaiSBNRMPareto(7) = nilaiSBNRMPareto(7) + !K70
343                       nilaiSBNRMPareto(8) = nilaiSBNRMPareto(8) + !K80
344                   Else
345                       nilaiSBNRMNonPareto(0) = nilaiSBNRMNonPareto(0) + !K1
346                       nilaiSBNRMNonPareto(1) = nilaiSBNRMNonPareto(1) + !K10
347                       nilaiSBNRMNonPareto(2) = nilaiSBNRMNonPareto(2) + !K20
348                       nilaiSBNRMNonPareto(3) = nilaiSBNRMNonPareto(3) + !K30
349                       nilaiSBNRMNonPareto(4) = nilaiSBNRMNonPareto(4) + !K40
350                       nilaiSBNRMNonPareto(5) = nilaiSBNRMNonPareto(5) + !K50
351                       nilaiSBNRMNonPareto(6) = nilaiSBNRMNonPareto(6) + !K60
352                       nilaiSBNRMNonPareto(7) = nilaiSBNRMNonPareto(7) + !K70
353                       nilaiSBNRMNonPareto(8) = nilaiSBNRMNonPareto(8) + !K80
354                   End If
                      
355                   If NulSID(rsModel!north) = 1 Then
356                       !north = 1
357                       nilaiSBNRMNorth(0) = nilaiSBNRMNorth(0) + !K1
358                       nilaiSBNRMNorth(1) = nilaiSBNRMNorth(1) + !K10
359                       nilaiSBNRMNorth(2) = nilaiSBNRMNorth(2) + !K20
360                       nilaiSBNRMNorth(3) = nilaiSBNRMNorth(3) + !K30
361                       nilaiSBNRMNorth(4) = nilaiSBNRMNorth(4) + !K40
362                       nilaiSBNRMNorth(5) = nilaiSBNRMNorth(5) + !K50
363                       nilaiSBNRMNorth(6) = nilaiSBNRMNorth(6) + !K60
364                       nilaiSBNRMNorth(7) = nilaiSBNRMNorth(7) + !K70
365                       nilaiSBNRMNorth(8) = nilaiSBNRMNorth(8) + !K80
366                   Else
                      
367                       !north = 3
368                       nilaiSBNRMSouth(0) = nilaiSBNRMSouth(0) + !K1
369                       nilaiSBNRMSouth(1) = nilaiSBNRMSouth(1) + !K10
370                       nilaiSBNRMSouth(2) = nilaiSBNRMSouth(2) + !K20
371                       nilaiSBNRMSouth(3) = nilaiSBNRMSouth(3) + !K30
372                       nilaiSBNRMSouth(4) = nilaiSBNRMSouth(4) + !K40
373                       nilaiSBNRMSouth(5) = nilaiSBNRMSouth(5) + !K50
374                       nilaiSBNRMSouth(6) = nilaiSBNRMSouth(6) + !K60
375                       nilaiSBNRMSouth(7) = nilaiSBNRMSouth(7) + !K70
376                       nilaiSBNRMSouth(8) = nilaiSBNRMSouth(8) + !K80
377                   End If
378                   !judul = "SB Same Group (NCM)"
379                   !model = Nul(rsModel.Fields(0).value)
380                   !Urut = 2
381                   .Update
382               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME GROUP
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
383               With rsLaporan
384                   .AddNew
                  
385                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.dealer from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.dealer ) a "
386                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
387                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
388                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
389                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
390                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
391                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
392                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
393                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
394                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
395                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
396                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
397                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
398                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
399                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
400                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
401                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
402                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
403                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
404                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
405                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
406                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
407                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
408                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
409                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
410                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
411                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
412                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
413                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
414                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
415                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
416                   If NulSID(rsModel!pareto) = 1 Then
417                       nilaiDemandPareto(0) = nilaiDemandPareto(0) + !K1
418                       nilaiDemandPareto(1) = nilaiDemandPareto(1) + !K10
419                       nilaiDemandPareto(2) = nilaiDemandPareto(2) + !K20
420                       nilaiDemandPareto(3) = nilaiDemandPareto(3) + !K30
421                       nilaiDemandPareto(4) = nilaiDemandPareto(4) + !K40
422                       nilaiDemandPareto(5) = nilaiDemandPareto(5) + !K50
423                       nilaiDemandPareto(6) = nilaiDemandPareto(6) + !K60
424                       nilaiDemandPareto(7) = nilaiDemandPareto(7) + !K70
425                       nilaiDemandPareto(8) = nilaiDemandPareto(8) + !K80
426                   Else
427                       nilaiDemandNonPareto(0) = nilaiDemandNonPareto(0) + !K1
428                       nilaiDemandNonPareto(1) = nilaiDemandNonPareto(1) + !K10
429                       nilaiDemandNonPareto(2) = nilaiDemandNonPareto(2) + !K20
430                       nilaiDemandNonPareto(3) = nilaiDemandNonPareto(3) + !K30
431                       nilaiDemandNonPareto(4) = nilaiDemandNonPareto(4) + !K40
432                       nilaiDemandNonPareto(5) = nilaiDemandNonPareto(5) + !K50
433                       nilaiDemandNonPareto(6) = nilaiDemandNonPareto(6) + !K60
434                       nilaiDemandNonPareto(7) = nilaiDemandNonPareto(7) + !K70
435                       nilaiDemandNonPareto(8) = nilaiDemandNonPareto(8) + !K80
436                   End If
                      
437                   If NulSID(rsModel!north) = 1 Then
438                       !north = 1
439                       nilaiDemandNorth(0) = nilaiDemandNorth(0) + !K1
440                       nilaiDemandNorth(1) = nilaiDemandNorth(1) + !K10
441                       nilaiDemandNorth(2) = nilaiDemandNorth(2) + !K20
442                       nilaiDemandNorth(3) = nilaiDemandNorth(3) + !K30
443                       nilaiDemandNorth(4) = nilaiDemandNorth(4) + !K40
444                       nilaiDemandNorth(5) = nilaiDemandNorth(5) + !K50
445                       nilaiDemandNorth(6) = nilaiDemandNorth(6) + !K60
446                       nilaiDemandNorth(7) = nilaiDemandNorth(7) + !K70
447                       nilaiDemandNorth(8) = nilaiDemandNorth(8) + !K80
448                   Else
                      
449                       !north = 3
450                       nilaiDemandSouth(0) = nilaiDemandSouth(0) + !K1
451                       nilaiDemandSouth(1) = nilaiDemandSouth(1) + !K10
452                       nilaiDemandSouth(2) = nilaiDemandSouth(2) + !K20
453                       nilaiDemandSouth(3) = nilaiDemandSouth(3) + !K30
454                       nilaiDemandSouth(4) = nilaiDemandSouth(4) + !K40
455                       nilaiDemandSouth(5) = nilaiDemandSouth(5) + !K50
456                       nilaiDemandSouth(6) = nilaiDemandSouth(6) + !K60
457                       nilaiDemandSouth(7) = nilaiDemandSouth(7) + !K70
458                       nilaiDemandSouth(8) = nilaiDemandSouth(8) + !K80
459                   End If
                       
460                   !judul = "Demand Same Outlet"
461                   !model = Nul(rsModel.Fields(0).value)
462                   !Urut = 4
463                   .Update
464               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
465               For i = 0 To 10
466                   nilaiTempDemand(i) = 0
467                   nilaiTempTotal(i) = 0
468               Next
469               rsLaporan.filter = ""
470               If Not rsLaporan.EOF Then
471                   rsLaporan.filter = "judul = 'Demand Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
472                   If Not rsLaporan.EOF Then
473                       If rsLaporan!K1 > 0 Then
474                           nilaiTempDemand(0) = rsLaporan!K1
475                       End If
476                       If rsLaporan!K10 > 0 Then
477                           nilaiTempDemand(1) = rsLaporan!K10
478                       End If
479                       If rsLaporan!K20 > 0 Then
480                           nilaiTempDemand(2) = rsLaporan!K20
481                       End If
482                       If rsLaporan!K30 > 0 Then
483                           nilaiTempDemand(3) = rsLaporan!K30
484                       End If
485                       If rsLaporan!K40 > 0 Then
486                           nilaiTempDemand(4) = rsLaporan!K40
487                       End If
488                       If rsLaporan!K50 > 0 Then
489                           nilaiTempDemand(5) = rsLaporan!K50
490                       End If
491                       If rsLaporan!K60 > 0 Then
492                           nilaiTempDemand(6) = rsLaporan!K60
493                       End If
494                       If rsLaporan!K70 > 0 Then
495                           nilaiTempDemand(7) = rsLaporan!K70
496                       End If
497                       If rsLaporan!K80 > 0 Then
498                           nilaiTempDemand(8) = rsLaporan!K80
499                       End If
500                   End If
501               End If
502               rsLaporan.filter = ""
                  
503               If Not rsLaporan.EOF Then
504                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
505                   If Not rsLaporan.EOF Then
506                       If rsLaporan!K1 > 0 Then
507                           nilaiTempTotal(0) = rsLaporan!K1
508                       End If
509                       If rsLaporan!K10 > 0 Then
510                           nilaiTempTotal(1) = rsLaporan!K10
511                       End If
512                       If rsLaporan!K20 > 0 Then
513                           nilaiTempTotal(2) = rsLaporan!K20
514                       End If
515                       If rsLaporan!K30 > 0 Then
516                           nilaiTempTotal(3) = rsLaporan!K30
517                       End If
518                       If rsLaporan!K40 > 0 Then
519                           nilaiTempTotal(4) = rsLaporan!K40
520                       End If
521                       If rsLaporan!K50 > 0 Then
522                           nilaiTempTotal(5) = rsLaporan!K50
523                       End If
524                       If rsLaporan!K60 > 0 Then
525                           nilaiTempTotal(6) = rsLaporan!K60
526                       End If
527                       If rsLaporan!K70 > 0 Then
528                           nilaiTempTotal(7) = rsLaporan!K70
529                       End If
530                       If rsLaporan!K80 > 0 Then
531                           nilaiTempTotal(8) = rsLaporan!K80
532                       End If
533                   End If
534               End If
                  
535               With rsLaporan
536                   .AddNew
537                   If nilaiTempDemand(0) > 0 Then
538                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
539                   Else
540                       !K1 = 0 & " %"
541                   End If
542                   If nilaiTempDemand(1) > 0 Then
543                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
544                   Else
545                       !K10 = 0 & " %"
546                   End If
547                   If nilaiTempDemand(2) > 0 Then
548                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
549                   Else
550                       !K20 = 0 & " %"
551                   End If
552                   If nilaiTempDemand(3) > 0 Then
553                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
554                   Else
555                       !K30 = 0 & " %"
556                   End If
557                   If nilaiTempDemand(4) > 0 Then
558                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
559                   Else
560                       !K40 = 0 & " %"
561                   End If
562                   If nilaiTempDemand(5) > 0 Then
563                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
564                   Else
565                       !K50 = 0 & " %"
566                   End If
567                   If nilaiTempDemand(6) > 0 Then
568                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
569                   Else
570                       !K60 = 0 & " %"
571                   End If
572                   If nilaiTempDemand(7) > 0 Then
573                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
574                   Else
575                       !K70 = 0 & " %"
576                   End If
577                   If nilaiTempDemand(8) > 0 Then
578                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
579                   Else
580                       !K80 = 0 & " %"
581                   End If
                      
582                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
583                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
584                   Else
585                       !tot16 = 0 & " %"
586                   End If
                      
587                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
588                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
589                   Else
590                       !tot26 = 0 & " %"
591                   End If
                       
592                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
593                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
594                   Else
595                       !tot19 = 0 & " %"
596                   End If
                       
597                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
598                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
599                   Else
600                       !tot29 = 0 & " %"
601                   End If
                      
602                   If NulSID(rsModel!north) = 1 Then
603                       !north = 1
604                   Else
605                       !north = 3
606                   End If
607                   !Urut = 6
608                   !judul = "Retention Rate Same Group NCM"
609                   !model = Nul(rsModel.Fields(0).value)
610                   .Update
611               End With
                   
612               With rsLaporan
613                   .AddNew
614                   If Not rsSBETotal.EOF Then
615                       rsSBETotal.filter = ""
616                       nilaiTemp = 0
617                       rsSBETotal.filter = "fs_rc = 'R001' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
618                       If Not rsSBETotal.EOF Then
619                           Do While Not rsSBETotal.EOF
620                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
621                               rsSBETotal.MoveNext
622                           Loop
623                       End If
624                       !K1 = nilaiTemp
625                       rsSBETotal.filter = ""
626                       nilaiTemp = 0
                          
627                       rsSBETotal.filter = "fs_rc = 'R010' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
628                       If Not rsSBETotal.EOF Then
629                           Do While Not rsSBETotal.EOF
630                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
631                               rsSBETotal.MoveNext
632                           Loop
633                       End If
634                       !K10 = nilaiTemp
635                       rsSBETotal.filter = ""
636                       nilaiTemp = 0
                          
637                       rsSBETotal.filter = "fs_rc = 'R020' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
638                       If Not rsSBETotal.EOF Then
639                           Do While Not rsSBETotal.EOF
640                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
641                               rsSBETotal.MoveNext
642                           Loop
643                       End If
644                       !K20 = nilaiTemp
645                       rsSBETotal.filter = ""
646                       nilaiTemp = 0
                          
647                       rsSBETotal.filter = "fs_rc = 'R030' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
648                       If Not rsSBETotal.EOF Then
649                           Do While Not rsSBETotal.EOF
650                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
651                               rsSBETotal.MoveNext
652                           Loop
653                       End If
654                       !K30 = nilaiTemp
655                       rsSBETotal.filter = ""
656                       nilaiTemp = 0
                          
657                       rsSBETotal.filter = "fs_rc = 'R040' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
658                       If Not rsSBETotal.EOF Then
659                           Do While Not rsSBETotal.EOF
660                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
661                               rsSBETotal.MoveNext
662                           Loop
663                       End If
664                       !K40 = nilaiTemp
665                       rsSBETotal.filter = ""
666                       nilaiTemp = 0
                          
667                       rsSBETotal.filter = "fs_rc = 'R050' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
668                       If Not rsSBETotal.EOF Then
669                           Do While Not rsSBETotal.EOF
670                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
671                               rsSBETotal.MoveNext
672                           Loop
673                       End If
674                       !K50 = nilaiTemp
675                       rsSBETotal.filter = ""
676                       nilaiTemp = 0
                          
677                       rsSBETotal.filter = "fs_rc = 'R060' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
678                       If Not rsSBETotal.EOF Then
679                           Do While Not rsSBETotal.EOF
680                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
681                               rsSBETotal.MoveNext
682                           Loop
683                       End If
684                       !K60 = nilaiTemp
685                       rsSBETotal.filter = ""
686                       nilaiTemp = 0
                          
687                       rsSBETotal.filter = "fs_rc = 'R070' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
688                       If Not rsSBETotal.EOF Then
689                           Do While Not rsSBETotal.EOF
690                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
691                               rsSBETotal.MoveNext
692                           Loop
693                       End If
694                       !K70 = nilaiTemp
695                       rsSBETotal.filter = ""
696                       nilaiTemp = 0
                          
697                       rsSBETotal.filter = "fs_rc = 'R080' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
698                       If Not rsSBETotal.EOF Then
699                           Do While Not rsSBETotal.EOF
700                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
701                               rsSBETotal.MoveNext
702                           Loop
703                       End If
704                       !K80 = nilaiTemp
705                       rsSBETotal.filter = ""
706                       nilaiTemp = 0
                          
707                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
708                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
709                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
710                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
711                   Else
712                       !K1 = 0
713                       !K10 = 0
714                       !K20 = 0
715                       !K30 = 0
716                       !K40 = 0
717                       !K50 = 0
718                       !K60 = 0
719                       !K70 = 0
720                       !K80 = 0
721                       !tot16 = 0
722                       !tot19 = 0
723                       !tot29 = 0
724                       !tot26 = 0
725                   End If
                      
726                   If NulSID(rsModel!pareto) = 1 Then
727                       nilaiTotalPareto(0) = nilaiTotalPareto(0) + !K1
728                       nilaiTotalPareto(1) = nilaiTotalPareto(1) + !K10
729                       nilaiTotalPareto(2) = nilaiTotalPareto(2) + !K20
730                       nilaiTotalPareto(3) = nilaiTotalPareto(3) + !K30
731                       nilaiTotalPareto(4) = nilaiTotalPareto(4) + !K40
732                       nilaiTotalPareto(5) = nilaiTotalPareto(5) + !K50
733                       nilaiTotalPareto(6) = nilaiTotalPareto(6) + !K60
734                       nilaiTotalPareto(7) = nilaiTotalPareto(7) + !K70
735                       nilaiTotalPareto(8) = nilaiTotalPareto(8) + !K80
736                   Else
737                       nilaiTotalNonPareto(0) = nilaiTotalNonPareto(0) + !K1
738                       nilaiTotalNonPareto(1) = nilaiTotalNonPareto(1) + !K10
739                       nilaiTotalNonPareto(2) = nilaiTotalNonPareto(2) + !K20
740                       nilaiTotalNonPareto(3) = nilaiTotalNonPareto(3) + !K30
741                       nilaiTotalNonPareto(4) = nilaiTotalNonPareto(4) + !K40
742                       nilaiTotalNonPareto(5) = nilaiTotalNonPareto(5) + !K50
743                       nilaiTotalNonPareto(6) = nilaiTotalNonPareto(6) + !K60
744                       nilaiTotalNonPareto(7) = nilaiTotalNonPareto(7) + !K70
745                       nilaiTotalNonPareto(8) = nilaiTotalNonPareto(8) + !K80
746                   End If
                      
747                   If NulSID(rsModel!north) = 1 Then
748                       !north = 1
749                       nilaiTotalNorth(0) = nilaiTotalNorth(0) + !K1
750                       nilaiTotalNorth(1) = nilaiTotalNorth(1) + !K10
751                       nilaiTotalNorth(2) = nilaiTotalNorth(2) + !K20
752                       nilaiTotalNorth(3) = nilaiTotalNorth(3) + !K30
753                       nilaiTotalNorth(4) = nilaiTotalNorth(4) + !K40
754                       nilaiTotalNorth(5) = nilaiTotalNorth(5) + !K50
755                       nilaiTotalNorth(6) = nilaiTotalNorth(6) + !K60
756                       nilaiTotalNorth(7) = nilaiTotalNorth(7) + !K70
757                       nilaiTotalNorth(8) = nilaiTotalNorth(8) + !K80
758                   Else
                      
759                       !north = 3
760                       nilaiTotalSouth(0) = nilaiTotalSouth(0) + !K1
761                       nilaiTotalSouth(1) = nilaiTotalSouth(1) + !K10
762                       nilaiTotalSouth(2) = nilaiTotalSouth(2) + !K20
763                       nilaiTotalSouth(3) = nilaiTotalSouth(3) + !K30
764                       nilaiTotalSouth(4) = nilaiTotalSouth(4) + !K40
765                       nilaiTotalSouth(5) = nilaiTotalSouth(5) + !K50
766                       nilaiTotalSouth(6) = nilaiTotalSouth(6) + !K60
767                       nilaiTotalSouth(7) = nilaiTotalSouth(7) + !K70
768                       nilaiTotalSouth(8) = nilaiTotalSouth(8) + !K80
769                   End If
770                   !judul = "SB Total (NCM & Non NCM)"
771                   !model = Nul(rsModel.Fields(0).value)
772                   !Urut = 3
773                   .Update
774               End With
775               rsLaporan.filter = ""
                  
                  
776               For i = 0 To 10
777                   nilaiTempTotal(i) = 0
778               Next
779               rsLaporan.filter = ""
                  
780               If Not rsLaporan.EOF Then
781                   rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
782                   If Not rsLaporan.EOF Then
783                       If rsLaporan!K1 > 0 Then
784                           nilaiTempTotal(0) = rsLaporan!K1
785                       End If
786                       If rsLaporan!K10 > 0 Then
787                           nilaiTempTotal(1) = rsLaporan!K10
788                       End If
789                       If rsLaporan!K20 > 0 Then
790                           nilaiTempTotal(2) = rsLaporan!K20
791                       End If
792                       If rsLaporan!K30 > 0 Then
793                           nilaiTempTotal(3) = rsLaporan!K30
794                       End If
795                       If rsLaporan!K40 > 0 Then
796                           nilaiTempTotal(4) = rsLaporan!K40
797                       End If
798                       If rsLaporan!K50 > 0 Then
799                           nilaiTempTotal(5) = rsLaporan!K50
800                       End If
801                       If rsLaporan!K60 > 0 Then
802                           nilaiTempTotal(6) = rsLaporan!K60
803                       End If
804                       If rsLaporan!K70 > 0 Then
805                           nilaiTempTotal(7) = rsLaporan!K70
806                       End If
807                       If rsLaporan!K80 > 0 Then
808                           nilaiTempTotal(8) = rsLaporan!K80
809                       End If
810                   End If
811               End If
                  
812               With rsLaporan
813                   .AddNew
814                   If nilaiTempDemand(0) > 0 Then
815                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
816                   Else
817                       !K1 = 0 & " %"
818                   End If
819                   If nilaiTempDemand(1) > 0 Then
820                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
821                   Else
822                       !K10 = 0 & " %"
823                   End If
824                   If nilaiTempDemand(2) > 0 Then
825                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
826                   Else
827                       !K20 = 0 & " %"
828                   End If
829                   If nilaiTempDemand(3) > 0 Then
830                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
831                   Else
832                       !K30 = 0 & " %"
833                   End If
834                   If nilaiTempDemand(4) > 0 Then
835                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
836                   Else
837                       !K40 = 0 & " %"
838                   End If
839                   If nilaiTempDemand(5) > 0 Then
840                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
841                   Else
842                       !K50 = 0 & " %"
843                   End If
844                   If nilaiTempDemand(6) > 0 Then
845                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
846                   Else
847                       !K60 = 0 & " %"
848                   End If
849                   If nilaiTempDemand(7) > 0 Then
850                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
851                   Else
852                       !K70 = 0 & " %"
853                   End If
854                   If nilaiTempDemand(8) > 0 Then
855                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
856                   Else
857                       !K80 = 0 & " %"
858                   End If
                      
                      
859                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
860                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
861                   Else
862                       !tot16 = 0 & " %"
863                   End If
                      
864                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
865                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
866                   Else
867                       !tot26 = 0 & " %"
868                   End If
                       
869                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
870                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
871                   Else
872                       !tot19 = 0 & " %"
873                   End If
                       
874                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
875                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
876                   Else
877                       !tot29 = 0 & " %"
878                   End If
                      
879                   If NulSID(rsModel!north) = 1 Then
880                       !north = 1
881                   Else
882                       !north = 3
883                   End If
884                   !Urut = 5
885                   !judul = "Retention Rate Same Outlet"
886                   !model = Nul(rsModel.Fields(0).value)
887                   .Update
888               End With
                  
889               For i = 0 To 10
890                   nilaiTempTotal(i) = 0
891               Next
892               rsLaporan.filter = ""
                  
893               If Not rsLaporan.EOF Then
894                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
895                   If Not rsLaporan.EOF Then
896                       If rsLaporan!K1 > 0 Then
897                           nilaiTempTotal(0) = rsLaporan!K1
898                       End If
899                       If rsLaporan!K10 > 0 Then
900                           nilaiTempTotal(1) = rsLaporan!K10
901                       End If
902                       If rsLaporan!K20 > 0 Then
903                           nilaiTempTotal(2) = rsLaporan!K20
904                       End If
905                       If rsLaporan!K30 > 0 Then
906                           nilaiTempTotal(3) = rsLaporan!K30
907                       End If
908                       If rsLaporan!K40 > 0 Then
909                           nilaiTempTotal(4) = rsLaporan!K40
910                       End If
911                       If rsLaporan!K50 > 0 Then
912                           nilaiTempTotal(5) = rsLaporan!K50
913                       End If
914                       If rsLaporan!K60 > 0 Then
915                           nilaiTempTotal(6) = rsLaporan!K60
916                       End If
917                       If rsLaporan!K70 > 0 Then
918                           nilaiTempTotal(7) = rsLaporan!K70
919                       End If
920                       If rsLaporan!K80 > 0 Then
921                           nilaiTempTotal(8) = rsLaporan!K80
922                       End If
923                   End If
924               End If
                  
925               With rsLaporan
926                   .AddNew
927                   If nilaiTempDemand(0) > 0 Then
928                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
929                   Else
930                       !K1 = 0 & " %"
931                   End If
932                   If nilaiTempDemand(1) > 0 Then
933                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
934                   Else
935                       !K10 = 0 & " %"
936                   End If
937                   If nilaiTempDemand(2) > 0 Then
938                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
939                   Else
940                       !K20 = 0 & " %"
941                   End If
942                   If nilaiTempDemand(3) > 0 Then
943                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
944                   Else
945                       !K30 = 0 & " %"
946                   End If
947                   If nilaiTempDemand(4) > 0 Then
948                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
949                   Else
950                       !K40 = 0 & " %"
951                   End If
952                   If nilaiTempDemand(5) > 0 Then
953                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
954                   Else
955                       !K50 = 0 & " %"
956                   End If
957                   If nilaiTempDemand(6) > 0 Then
958                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
959                   Else
960                       !K60 = 0 & " %"
961                   End If
962                   If nilaiTempDemand(7) > 0 Then
963                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
964                   Else
965                       !K70 = 0 & " %"
966                   End If
967                   If nilaiTempDemand(8) > 0 Then
968                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
969                   Else
970                       !K80 = 0 & " %"
971                   End If
                      
                      
972                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
973                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
974                   Else
975                       !tot16 = 0 & " %"
976                   End If
                      
977                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
978                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
979                   Else
980                       !tot26 = 0 & " %"
981                   End If
                       
982                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
983                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
984                   Else
985                       !tot19 = 0 & " %"
986                   End If
                       
987                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
988                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
989                   Else
990                       !tot29 = 0 & " %"
991                   End If
                      
992                   If NulSID(rsModel!north) = 1 Then
993                       !north = 1
994                   Else
995                       !north = 3
996                   End If
997                   !Urut = 7
998                   !judul = "Retention Rate Total (NCM & Non NCM)"
999                   !model = Nul(rsModel.Fields(0).value)
1000                  .Update
1001              End With
1002              rsLaporan.filter = ""
1003              rsModel.MoveNext
1004              progbarDealer.value = progbarDealer.value + 1
1005              DoEvents
1006          Loop
1007          Call setNilaiKategoriDealer
              
1008      Else
1009          MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
1010          Exit Sub
              
1011      End If
1012      rsLaporan.Sort = "north asc"
1013      Report.database.SetDataSource rsLaporan
1014      progbarDealer.Visible = False
1015      Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
1016      Report.txtPeriode.SetText "Periode : " & Format(tglDealerAwal, "dd-mm-yyyy") & " s/d " & Format(tglDealerAkhir, "dd-mm-yyyy")
          
1017      With CRViewer1
1018          .ReportSource = Report
1019          .ViewReport
1020          .Zoom 100
1021          .Visible = True
1022          .Left = 300
1023          .Height = Me.Height - 1200
1024          .Width = Me.Width - 750
1025      End With
         
1026      cmdTutup.Visible = True
1027      cmdTutup.Left = CRViewer1.Left + 20
1028      cmdTutup.Top = CRViewer1.Top + 400
1029      Screen.MousePointer = vbDefault
          
1030      Call DisableExport
1031      Call setLokasiTombol
End Sub


Private Sub ambilLaporanVinDealerTglRetail()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         For i = 0 To 10
2             nilaiSBNRMPareto(i) = 0
3             nilaiSBNRMNonPareto(i) = 0
4             nilaiSBNRMSouth(i) = 0
5             nilaiSBNRMNorth(i) = 0
              
6             nilaiDemandPareto(i) = 0
7             nilaiDemandNonPareto(i) = 0
8             nilaiDemandSouth(i) = 0
9             nilaiDemandNorth(i) = 0
              
10            nilaiTotalNonPareto(i) = 0
11            nilaiTotalPareto(i) = 0
12            nilaiTotalSouth(i) = 0
13            nilaiTotalNorth(i) = 0
              
14            nilaiSBSameOutletPareto(i) = 0
15            nilaiSBSameOutletNonPareto(i) = 0
16            nilaiSBSameOutletSouth(i) = 0
17            nilaiSBSameOutletNorth(i) = 0
18        Next
          
19        Set Report = New rptSBERateByModelVinRetentionDealer
20        Call createTableLaporanVin
          
21        sSQL = "select FS_NM_DEALER,FS_KD_DEALER,pareto,north,kodecabang from T_DEALER where INTERNAL = 1 and FS_LOG_DEL = '' and FS_KD_DEALER <> 'NR 00' order by north desc, FS_KD_DEALER"
22        Set rsModel = objO.FetchDataHO(sSQL)
        
23        sSQL = "select COUNT(noinv) total,fs_rc,"
24        sSQL = sSQL & " E.FS_NM_DEALER"
25        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
26        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
27        sSQL = sSQL & " inner JOIN T_DEALER E WITH (NOLOCK) ON e.FS_KD_DEALER = t.FS_KD_DEALER and e.fs_log_del = '' and INTERNAL = 1"
28        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
29        sSQL = sSQL & " AND INVDATE between '" & Format(tglDealerAwal.value, "mm/dd/yyyy") & "' AND '" & Format(tglDealerAkhir.value, "mm/dd/yyyy") & "'"
30        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
31        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
32        sSQL = sSQL & " grOUP BY FS_RC, e.FS_NM_DEALER"
33        sSQL = sSQL & " order by FS_RC,e.FS_NM_DEALER"
           
34        Set rsSBETotal = objO.FetchDataHO(sSQL)
          
          
35        sSQL = "select COUNT(noinv) total,fs_rc,case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
36        sSQL = sSQL & " D.FS_NM_DEALER"
37        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
38        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
39        sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
40        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''"
41        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
42        sSQL = sSQL & " AND INVDATE between '" & Format(tglDealerAwal.value, "mm/dd/yyyy") & "' AND '" & Format(tglDealerAkhir.value, "mm/dd/yyyy") & "'"
43        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
44        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
45        sSQL = sSQL & " grOUP BY FS_RC , case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end , d.FS_NM_DEALER"
46        sSQL = sSQL & " order by FS_RC,d.FS_NM_DEALER"
          
47        sSQL = "select D.FS_NM_DEALER,COUNT(noinv) total,fs_rc,case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
48        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK) "
49        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  "
50        sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
51        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''" & dealerPilih
52        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
53        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
54        sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [SPK99]..Summary WHERE TipeTr = 'spk' and [Status] = '1' AND Tanggal between '" & Format(tglDealerAwal, "mm/dd/yyyy") & "' AND '" & Format(tglDealerAkhir, "mm/dd/yyyy") & "' )"
55        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
56        sSQL = sSQL & " grOUP BY FS_RC ,d.FS_NM_DEALER, case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
57        sSQL = sSQL & " order by FS_RC,d.FS_NM_DEALER"
58        Set rsSBE = objO.FetchDataHO(sSQL)
59        If Not rsModel.EOF Then
60            rsModel.MoveFirst
61            progbarDealer.Max = rsModel.RecordCount + 1
62            progbarDealer.value = 1
63            progbarDealer.Visible = True
64            DoEvents
65            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
66                With rsLaporan
67                    .AddNew
68                    rsSBE.filter = ""
69                    If Not rsSBE.EOF Then
70                        rsSBE.filter = ""
71                        nilaiTemp = 0
72                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
73                        If Not rsSBE.EOF Then
74                            Do While Not rsSBE.EOF
75                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
76                                rsSBE.MoveNext
77                            Loop
78                        End If
79                        !K1 = nilaiTemp
80                        rsSBE.filter = ""
81                        nilaiTemp = 0
                          
82                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
83                        If Not rsSBE.EOF Then
84                            Do While Not rsSBE.EOF
85                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
86                                rsSBE.MoveNext
87                            Loop
88                        End If
89                        !K10 = nilaiTemp
90                        rsSBE.filter = ""
91                        nilaiTemp = 0
                          
92                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
93                        If Not rsSBE.EOF Then
94                            Do While Not rsSBE.EOF
95                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
96                                rsSBE.MoveNext
97                            Loop
98                        End If
99                        !K20 = nilaiTemp
100                       rsSBE.filter = ""
101                       nilaiTemp = 0
                          
102                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
103                       If Not rsSBE.EOF Then
104                           Do While Not rsSBE.EOF
105                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
106                               rsSBE.MoveNext
107                           Loop
108                       End If
109                       !K30 = nilaiTemp
110                       rsSBE.filter = ""
111                       nilaiTemp = 0
                          
112                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
113                       If Not rsSBE.EOF Then
114                           Do While Not rsSBE.EOF
115                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
116                               rsSBE.MoveNext
117                           Loop
118                       End If
119                       !K40 = nilaiTemp
120                       rsSBE.filter = ""
121                       nilaiTemp = 0
                          
122                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
123                       If Not rsSBE.EOF Then
124                           Do While Not rsSBE.EOF
125                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
126                               rsSBE.MoveNext
127                           Loop
128                       End If
129                       !K50 = nilaiTemp
130                       rsSBE.filter = ""
131                       nilaiTemp = 0
                          
132                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
133                       If Not rsSBE.EOF Then
134                           Do While Not rsSBE.EOF
135                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
136                               rsSBE.MoveNext
137                           Loop
138                       End If
139                       !K60 = nilaiTemp
140                       rsSBE.filter = ""
141                       nilaiTemp = 0
                          
142                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
143                       If Not rsSBE.EOF Then
144                           Do While Not rsSBE.EOF
145                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
146                               rsSBE.MoveNext
147                           Loop
148                       End If
149                       !K70 = nilaiTemp
150                       rsSBE.filter = ""
151                       nilaiTemp = 0
                          
152                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
153                       If Not rsSBE.EOF Then
154                           Do While Not rsSBE.EOF
155                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
156                               rsSBE.MoveNext
157                           Loop
158                       End If
159                       !K80 = nilaiTemp
160                       rsSBE.filter = ""
161                       nilaiTemp = 0
                          
162                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
163                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
164                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
165                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
166                   Else
167                       !K1 = 0
168                       !K10 = 0
169                       !K20 = 0
170                       !K30 = 0
171                       !K40 = 0
172                       !K50 = 0
173                       !K60 = 0
174                       !K70 = 0
175                       !K80 = 0
176                       !tot16 = 0
177                       !tot19 = 0
                          
178                       !tot26 = 0
179                       !tot29 = 0
                          
180                   End If
181                   If NulSID(rsModel!pareto) = 1 Then
182                       nilaiSBSameOutletPareto(0) = nilaiSBSameOutletPareto(0) + !K1
183                       nilaiSBSameOutletPareto(1) = nilaiSBSameOutletPareto(1) + !K10
184                       nilaiSBSameOutletPareto(2) = nilaiSBSameOutletPareto(2) + !K20
185                       nilaiSBSameOutletPareto(3) = nilaiSBSameOutletPareto(3) + !K30
186                       nilaiSBSameOutletPareto(4) = nilaiSBSameOutletPareto(4) + !K40
187                       nilaiSBSameOutletPareto(5) = nilaiSBSameOutletPareto(5) + !K50
188                       nilaiSBSameOutletPareto(6) = nilaiSBSameOutletPareto(6) + !K60
189                       nilaiSBSameOutletPareto(7) = nilaiSBSameOutletPareto(7) + !K70
190                       nilaiSBSameOutletPareto(8) = nilaiSBSameOutletPareto(8) + !K80
191                   Else
192                       nilaiSBSameOutletNonPareto(0) = nilaiSBSameOutletNonPareto(0) + !K1
193                       nilaiSBSameOutletNonPareto(1) = nilaiSBSameOutletNonPareto(1) + !K10
194                       nilaiSBSameOutletNonPareto(2) = nilaiSBSameOutletNonPareto(2) + !K20
195                       nilaiSBSameOutletNonPareto(3) = nilaiSBSameOutletNonPareto(3) + !K30
196                       nilaiSBSameOutletNonPareto(4) = nilaiSBSameOutletNonPareto(4) + !K40
197                       nilaiSBSameOutletNonPareto(5) = nilaiSBSameOutletNonPareto(5) + !K50
198                       nilaiSBSameOutletNonPareto(6) = nilaiSBSameOutletNonPareto(6) + !K60
199                       nilaiSBSameOutletNonPareto(7) = nilaiSBSameOutletNonPareto(7) + !K70
200                       nilaiSBSameOutletNonPareto(8) = nilaiSBSameOutletNonPareto(8) + !K80
201                   End If
                      
202                   If NulSID(rsModel!north) = 1 Then
203                       !north = 1
204                       nilaiSBSameOutletNorth(0) = nilaiSBSameOutletNorth(0) + !K1
205                       nilaiSBSameOutletNorth(1) = nilaiSBSameOutletNorth(1) + !K10
206                       nilaiSBSameOutletNorth(2) = nilaiSBSameOutletNorth(2) + !K20
207                       nilaiSBSameOutletNorth(3) = nilaiSBSameOutletNorth(3) + !K30
208                       nilaiSBSameOutletNorth(4) = nilaiSBSameOutletNorth(4) + !K40
209                       nilaiSBSameOutletNorth(5) = nilaiSBSameOutletNorth(5) + !K50
210                       nilaiSBSameOutletNorth(6) = nilaiSBSameOutletNorth(6) + !K60
211                       nilaiSBSameOutletNorth(7) = nilaiSBSameOutletNorth(7) + !K70
212                       nilaiSBSameOutletNorth(8) = nilaiSBSameOutletNorth(8) + !K80
213                   Else
                      
214                       !north = 3
215                       nilaiSBSameOutletSouth(0) = nilaiSBSameOutletSouth(0) + !K1
216                       nilaiSBSameOutletSouth(1) = nilaiSBSameOutletSouth(1) + !K10
217                       nilaiSBSameOutletSouth(2) = nilaiSBSameOutletSouth(2) + !K20
218                       nilaiSBSameOutletSouth(3) = nilaiSBSameOutletSouth(3) + !K30
219                       nilaiSBSameOutletSouth(4) = nilaiSBSameOutletSouth(4) + !K40
220                       nilaiSBSameOutletSouth(5) = nilaiSBSameOutletSouth(5) + !K50
221                       nilaiSBSameOutletSouth(6) = nilaiSBSameOutletSouth(6) + !K60
222                       nilaiSBSameOutletSouth(7) = nilaiSBSameOutletSouth(7) + !K70
223                       nilaiSBSameOutletSouth(8) = nilaiSBSameOutletSouth(8) + !K80
224                   End If
225                   !judul = "SB Same Outlet"
226                   !model = Nul(rsModel.Fields(0).value)
227                   !Urut = 1
228                   .Update
229               End With
                  
                  ''LAPORAN SAME GROUP
230               With rsLaporan
231                   .AddNew
232                   rsSBE.filter = ""
233                   If Not rsSBE.EOF Then
234                       rsSBE.filter = ""
235                       nilaiTemp = 0
236                       rsSBE.filter = "fs_rc = 'R001' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
237                       If Not rsSBE.EOF Then
238                           Do While Not rsSBE.EOF
239                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
240                               rsSBE.MoveNext
241                           Loop
242                       End If
243                       !K1 = nilaiTemp
244                       rsSBE.filter = ""
245                       nilaiTemp = 0
                          
246                       rsSBE.filter = "fs_rc = 'R010' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
247                       If Not rsSBE.EOF Then
248                           Do While Not rsSBE.EOF
249                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
250                               rsSBE.MoveNext
251                           Loop
252                       End If
253                       !K10 = nilaiTemp
254                       rsSBE.filter = ""
255                       nilaiTemp = 0
                          
256                       rsSBE.filter = "fs_rc = 'R020' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
257                       If Not rsSBE.EOF Then
258                           Do While Not rsSBE.EOF
259                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
260                               rsSBE.MoveNext
261                           Loop
262                       End If
263                       !K20 = nilaiTemp
264                       rsSBE.filter = ""
265                       nilaiTemp = 0
                          
266                       rsSBE.filter = "fs_rc = 'R030' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
267                       If Not rsSBE.EOF Then
268                           Do While Not rsSBE.EOF
269                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
270                               rsSBE.MoveNext
271                           Loop
272                       End If
273                       !K30 = nilaiTemp
274                       rsSBE.filter = ""
275                       nilaiTemp = 0
                          
276                       rsSBE.filter = "fs_rc = 'R040' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
277                       If Not rsSBE.EOF Then
278                           Do While Not rsSBE.EOF
279                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
280                               rsSBE.MoveNext
281                           Loop
282                       End If
283                       !K40 = nilaiTemp
284                       rsSBE.filter = ""
285                       nilaiTemp = 0
                          
286                       rsSBE.filter = "fs_rc = 'R050' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
287                       If Not rsSBE.EOF Then
288                           Do While Not rsSBE.EOF
289                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
290                               rsSBE.MoveNext
291                           Loop
292                       End If
293                       !K50 = nilaiTemp
294                       rsSBE.filter = ""
295                       nilaiTemp = 0
                          
296                       rsSBE.filter = "fs_rc = 'R060' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
297                       If Not rsSBE.EOF Then
298                           Do While Not rsSBE.EOF
299                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
300                               rsSBE.MoveNext
301                           Loop
302                       End If
303                       !K60 = nilaiTemp
304                       rsSBE.filter = ""
305                       nilaiTemp = 0
                          
306                       rsSBE.filter = "fs_rc = 'R070' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
307                       If Not rsSBE.EOF Then
308                           Do While Not rsSBE.EOF
309                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
310                               rsSBE.MoveNext
311                           Loop
312                       End If
313                       !K70 = nilaiTemp
314                       rsSBE.filter = ""
315                       nilaiTemp = 0
                          
316                       rsSBE.filter = "fs_rc = 'R080' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
317                       If Not rsSBE.EOF Then
318                           Do While Not rsSBE.EOF
319                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
320                               rsSBE.MoveNext
321                           Loop
322                       End If
323                       !K80 = nilaiTemp
324                       rsSBE.filter = ""
325                       nilaiTemp = 0
                          
326                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
327                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
328                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
329                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
330                   Else
331                       !K1 = 0
332                       !K10 = 0
333                       !K20 = 0
334                       !K30 = 0
335                       !K40 = 0
336                       !K50 = 0
337                       !K60 = 0
338                       !K70 = 0
339                       !K80 = 0
340                       !tot16 = 0
341                       !tot19 = 0
                          
342                       !tot26 = 0
343                       !tot29 = 0
                          
344                   End If
345                   If NulSID(rsModel!pareto) = 1 Then
346                       nilaiSBNRMPareto(0) = nilaiSBNRMPareto(0) + !K1
347                       nilaiSBNRMPareto(1) = nilaiSBNRMPareto(1) + !K10
348                       nilaiSBNRMPareto(2) = nilaiSBNRMPareto(2) + !K20
349                       nilaiSBNRMPareto(3) = nilaiSBNRMPareto(3) + !K30
350                       nilaiSBNRMPareto(4) = nilaiSBNRMPareto(4) + !K40
351                       nilaiSBNRMPareto(5) = nilaiSBNRMPareto(5) + !K50
352                       nilaiSBNRMPareto(6) = nilaiSBNRMPareto(6) + !K60
353                       nilaiSBNRMPareto(7) = nilaiSBNRMPareto(7) + !K70
354                       nilaiSBNRMPareto(8) = nilaiSBNRMPareto(8) + !K80
355                   Else
356                       nilaiSBNRMNonPareto(0) = nilaiSBNRMNonPareto(0) + !K1
357                       nilaiSBNRMNonPareto(1) = nilaiSBNRMNonPareto(1) + !K10
358                       nilaiSBNRMNonPareto(2) = nilaiSBNRMNonPareto(2) + !K20
359                       nilaiSBNRMNonPareto(3) = nilaiSBNRMNonPareto(3) + !K30
360                       nilaiSBNRMNonPareto(4) = nilaiSBNRMNonPareto(4) + !K40
361                       nilaiSBNRMNonPareto(5) = nilaiSBNRMNonPareto(5) + !K50
362                       nilaiSBNRMNonPareto(6) = nilaiSBNRMNonPareto(6) + !K60
363                       nilaiSBNRMNonPareto(7) = nilaiSBNRMNonPareto(7) + !K70
364                       nilaiSBNRMNonPareto(8) = nilaiSBNRMNonPareto(8) + !K80
365                   End If
                      
366                   If NulSID(rsModel!north) = 1 Then
367                       !north = 1
368                       nilaiSBNRMNorth(0) = nilaiSBNRMNorth(0) + !K1
369                       nilaiSBNRMNorth(1) = nilaiSBNRMNorth(1) + !K10
370                       nilaiSBNRMNorth(2) = nilaiSBNRMNorth(2) + !K20
371                       nilaiSBNRMNorth(3) = nilaiSBNRMNorth(3) + !K30
372                       nilaiSBNRMNorth(4) = nilaiSBNRMNorth(4) + !K40
373                       nilaiSBNRMNorth(5) = nilaiSBNRMNorth(5) + !K50
374                       nilaiSBNRMNorth(6) = nilaiSBNRMNorth(6) + !K60
375                       nilaiSBNRMNorth(7) = nilaiSBNRMNorth(7) + !K70
376                       nilaiSBNRMNorth(8) = nilaiSBNRMNorth(8) + !K80
377                   Else
                      
378                       !north = 3
379                       nilaiSBNRMSouth(0) = nilaiSBNRMSouth(0) + !K1
380                       nilaiSBNRMSouth(1) = nilaiSBNRMSouth(1) + !K10
381                       nilaiSBNRMSouth(2) = nilaiSBNRMSouth(2) + !K20
382                       nilaiSBNRMSouth(3) = nilaiSBNRMSouth(3) + !K30
383                       nilaiSBNRMSouth(4) = nilaiSBNRMSouth(4) + !K40
384                       nilaiSBNRMSouth(5) = nilaiSBNRMSouth(5) + !K50
385                       nilaiSBNRMSouth(6) = nilaiSBNRMSouth(6) + !K60
386                       nilaiSBNRMSouth(7) = nilaiSBNRMSouth(7) + !K70
387                       nilaiSBNRMSouth(8) = nilaiSBNRMSouth(8) + !K80
388                   End If
389                   !judul = "SB Same Group (NCM)"
390                   !model = Nul(rsModel.Fields(0).value)
391                   !Urut = 2
392                   .Update
393               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME GROUP
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
394               With rsLaporan
395                   .AddNew
396                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.dealer from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.dealer ) a "
397                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
398                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
399                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
400                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
401                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
402                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
403                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
404                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
405                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
406                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
407                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
408                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
409                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
410                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
411                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
412                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
413                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
414                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
415                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
416                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
417                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
418                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
419                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
420                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
421                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
422                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
423                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
424                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
425                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
426                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
427                   If NulSID(rsModel!pareto) = 1 Then
428                       nilaiDemandPareto(0) = nilaiDemandPareto(0) + !K1
429                       nilaiDemandPareto(1) = nilaiDemandPareto(1) + !K10
430                       nilaiDemandPareto(2) = nilaiDemandPareto(2) + !K20
431                       nilaiDemandPareto(3) = nilaiDemandPareto(3) + !K30
432                       nilaiDemandPareto(4) = nilaiDemandPareto(4) + !K40
433                       nilaiDemandPareto(5) = nilaiDemandPareto(5) + !K50
434                       nilaiDemandPareto(6) = nilaiDemandPareto(6) + !K60
435                       nilaiDemandPareto(7) = nilaiDemandPareto(7) + !K70
436                       nilaiDemandPareto(8) = nilaiDemandPareto(8) + !K80
437                   Else
438                       nilaiDemandNonPareto(0) = nilaiDemandNonPareto(0) + !K1
439                       nilaiDemandNonPareto(1) = nilaiDemandNonPareto(1) + !K10
440                       nilaiDemandNonPareto(2) = nilaiDemandNonPareto(2) + !K20
441                       nilaiDemandNonPareto(3) = nilaiDemandNonPareto(3) + !K30
442                       nilaiDemandNonPareto(4) = nilaiDemandNonPareto(4) + !K40
443                       nilaiDemandNonPareto(5) = nilaiDemandNonPareto(5) + !K50
444                       nilaiDemandNonPareto(6) = nilaiDemandNonPareto(6) + !K60
445                       nilaiDemandNonPareto(7) = nilaiDemandNonPareto(7) + !K70
446                       nilaiDemandNonPareto(8) = nilaiDemandNonPareto(8) + !K80
447                   End If
                      
448                   If NulSID(rsModel!north) = 1 Then
449                       !north = 1
450                       nilaiDemandNorth(0) = nilaiDemandNorth(0) + !K1
451                       nilaiDemandNorth(1) = nilaiDemandNorth(1) + !K10
452                       nilaiDemandNorth(2) = nilaiDemandNorth(2) + !K20
453                       nilaiDemandNorth(3) = nilaiDemandNorth(3) + !K30
454                       nilaiDemandNorth(4) = nilaiDemandNorth(4) + !K40
455                       nilaiDemandNorth(5) = nilaiDemandNorth(5) + !K50
456                       nilaiDemandNorth(6) = nilaiDemandNorth(6) + !K60
457                       nilaiDemandNorth(7) = nilaiDemandNorth(7) + !K70
458                       nilaiDemandNorth(8) = nilaiDemandNorth(8) + !K80
459                   Else
                      
460                       !north = 3
461                       nilaiSBNRMSouth(0) = nilaiSBNRMSouth(0) + !K1
462                       nilaiSBNRMSouth(1) = nilaiSBNRMSouth(1) + !K10
463                       nilaiSBNRMSouth(2) = nilaiSBNRMSouth(2) + !K20
464                       nilaiSBNRMSouth(3) = nilaiSBNRMSouth(3) + !K30
465                       nilaiSBNRMSouth(4) = nilaiSBNRMSouth(4) + !K40
466                       nilaiSBNRMSouth(5) = nilaiSBNRMSouth(5) + !K50
467                       nilaiSBNRMSouth(6) = nilaiSBNRMSouth(6) + !K60
468                       nilaiSBNRMSouth(7) = nilaiSBNRMSouth(7) + !K70
469                       nilaiSBNRMSouth(8) = nilaiSBNRMSouth(8) + !K80
470                   End If
                       
471                   !judul = "Demand Same Outlet"
472                   !model = Nul(rsModel.Fields(0).value)
473                   !Urut = 4
474                   .Update
475               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
476               For i = 0 To 10
477                   nilaiTempDemand(i) = 0
478                   nilaiTempTotal(i) = 0
479               Next
480               rsLaporan.filter = ""
481               If Not rsLaporan.EOF Then
482                   rsLaporan.filter = "judul = 'Demand Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
483                   If Not rsLaporan.EOF Then
484                       If rsLaporan!K1 > 0 Then
485                           nilaiTempDemand(0) = rsLaporan!K1
486                       End If
487                       If rsLaporan!K10 > 0 Then
488                           nilaiTempDemand(1) = rsLaporan!K10
489                       End If
490                       If rsLaporan!K20 > 0 Then
491                           nilaiTempDemand(2) = rsLaporan!K20
492                       End If
493                       If rsLaporan!K30 > 0 Then
494                           nilaiTempDemand(3) = rsLaporan!K30
495                       End If
496                       If rsLaporan!K40 > 0 Then
497                           nilaiTempDemand(4) = rsLaporan!K40
498                       End If
499                       If rsLaporan!K50 > 0 Then
500                           nilaiTempDemand(5) = rsLaporan!K50
501                       End If
502                       If rsLaporan!K60 > 0 Then
503                           nilaiTempDemand(6) = rsLaporan!K60
504                       End If
505                       If rsLaporan!K70 > 0 Then
506                           nilaiTempDemand(7) = rsLaporan!K70
507                       End If
508                       If rsLaporan!K80 > 0 Then
509                           nilaiTempDemand(8) = rsLaporan!K80
510                       End If
511                   End If
512               End If
513               rsLaporan.filter = ""
                  
514               If Not rsLaporan.EOF Then
515                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
516                   If Not rsLaporan.EOF Then
517                       If rsLaporan!K1 > 0 Then
518                           nilaiTempTotal(0) = rsLaporan!K1
519                       End If
520                       If rsLaporan!K10 > 0 Then
521                           nilaiTempTotal(1) = rsLaporan!K10
522                       End If
523                       If rsLaporan!K20 > 0 Then
524                           nilaiTempTotal(2) = rsLaporan!K20
525                       End If
526                       If rsLaporan!K30 > 0 Then
527                           nilaiTempTotal(3) = rsLaporan!K30
528                       End If
529                       If rsLaporan!K40 > 0 Then
530                           nilaiTempTotal(4) = rsLaporan!K40
531                       End If
532                       If rsLaporan!K50 > 0 Then
533                           nilaiTempTotal(5) = rsLaporan!K50
534                       End If
535                       If rsLaporan!K60 > 0 Then
536                           nilaiTempTotal(6) = rsLaporan!K60
537                       End If
538                       If rsLaporan!K70 > 0 Then
539                           nilaiTempTotal(7) = rsLaporan!K70
540                       End If
541                       If rsLaporan!K80 > 0 Then
542                           nilaiTempTotal(8) = rsLaporan!K80
543                       End If
544                   End If
545               End If
                  
546               With rsLaporan
547                   .AddNew
548                   If nilaiTempDemand(0) > 0 Then
549                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
550                   Else
551                       !K1 = 0 & " %"
552                   End If
553                   If nilaiTempDemand(1) > 0 Then
554                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
555                   Else
556                       !K10 = 0 & " %"
557                   End If
558                   If nilaiTempDemand(2) > 0 Then
559                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
560                   Else
561                       !K20 = 0 & " %"
562                   End If
563                   If nilaiTempDemand(3) > 0 Then
564                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
565                   Else
566                       !K30 = 0 & " %"
567                   End If
568                   If nilaiTempDemand(4) > 0 Then
569                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
570                   Else
571                       !K40 = 0 & " %"
572                   End If
573                   If nilaiTempDemand(5) > 0 Then
574                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
575                   Else
576                       !K50 = 0 & " %"
577                   End If
578                   If nilaiTempDemand(6) > 0 Then
579                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
580                   Else
581                       !K60 = 0 & " %"
582                   End If
583                   If nilaiTempDemand(7) > 0 Then
584                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
585                   Else
586                       !K70 = 0 & " %"
587                   End If
588                   If nilaiTempDemand(8) > 0 Then
589                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
590                   Else
591                       !K80 = 0 & " %"
592                   End If
                      
593                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
594                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
595                   Else
596                       !tot16 = 0 & " %"
597                   End If
                      
598                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
599                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
600                   Else
601                       !tot26 = 0 & " %"
602                   End If
                       
603                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
604                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
605                   Else
606                       !tot19 = 0 & " %"
607                   End If
                       
608                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
609                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
610                   Else
611                       !tot29 = 0 & " %"
612                   End If
                      
613                   If NulSID(rsModel!north) = 1 Then
614                       !north = 1
615                   Else
616                       !north = 3
617                   End If
618                   !Urut = 6
619                   !judul = "Retention Rate Same Group NCM"
620                   !model = Nul(rsModel.Fields(0).value)
621                   .Update
622               End With
                   
623               With rsLaporan
624                   .AddNew
625                   If Not rsSBETotal.EOF Then
626                       rsSBETotal.filter = ""
627                       nilaiTemp = 0
628                       rsSBETotal.filter = "fs_rc = 'R001' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
629                       If Not rsSBETotal.EOF Then
630                           Do While Not rsSBETotal.EOF
631                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
632                               rsSBETotal.MoveNext
633                           Loop
634                       End If
635                       !K1 = nilaiTemp
636                       rsSBETotal.filter = ""
637                       nilaiTemp = 0
                          
638                       rsSBETotal.filter = "fs_rc = 'R010' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
639                       If Not rsSBETotal.EOF Then
640                           Do While Not rsSBETotal.EOF
641                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
642                               rsSBETotal.MoveNext
643                           Loop
644                       End If
645                       !K10 = nilaiTemp
646                       rsSBETotal.filter = ""
647                       nilaiTemp = 0
                          
648                       rsSBETotal.filter = "fs_rc = 'R020' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
649                       If Not rsSBETotal.EOF Then
650                           Do While Not rsSBETotal.EOF
651                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
652                               rsSBETotal.MoveNext
653                           Loop
654                       End If
655                       !K20 = nilaiTemp
656                       rsSBETotal.filter = ""
657                       nilaiTemp = 0
                          
658                       rsSBETotal.filter = "fs_rc = 'R030' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
659                       If Not rsSBETotal.EOF Then
660                           Do While Not rsSBETotal.EOF
661                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
662                               rsSBETotal.MoveNext
663                           Loop
664                       End If
665                       !K30 = nilaiTemp
666                       rsSBETotal.filter = ""
667                       nilaiTemp = 0
                          
668                       rsSBETotal.filter = "fs_rc = 'R040' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
669                       If Not rsSBETotal.EOF Then
670                           Do While Not rsSBETotal.EOF
671                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
672                               rsSBETotal.MoveNext
673                           Loop
674                       End If
675                       !K40 = nilaiTemp
676                       rsSBETotal.filter = ""
677                       nilaiTemp = 0
                          
678                       rsSBETotal.filter = "fs_rc = 'R050' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
679                       If Not rsSBETotal.EOF Then
680                           Do While Not rsSBETotal.EOF
681                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
682                               rsSBETotal.MoveNext
683                           Loop
684                       End If
685                       !K50 = nilaiTemp
686                       rsSBETotal.filter = ""
687                       nilaiTemp = 0
                          
688                       rsSBETotal.filter = "fs_rc = 'R060' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
689                       If Not rsSBETotal.EOF Then
690                           Do While Not rsSBETotal.EOF
691                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
692                               rsSBETotal.MoveNext
693                           Loop
694                       End If
695                       !K60 = nilaiTemp
696                       rsSBETotal.filter = ""
697                       nilaiTemp = 0
                          
698                       rsSBETotal.filter = "fs_rc = 'R070' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
699                       If Not rsSBETotal.EOF Then
700                           Do While Not rsSBETotal.EOF
701                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
702                               rsSBETotal.MoveNext
703                           Loop
704                       End If
705                       !K70 = nilaiTemp
706                       rsSBETotal.filter = ""
707                       nilaiTemp = 0
                          
708                       rsSBETotal.filter = "fs_rc = 'R080' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
709                       If Not rsSBETotal.EOF Then
710                           Do While Not rsSBETotal.EOF
711                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
712                               rsSBETotal.MoveNext
713                           Loop
714                       End If
715                       !K80 = nilaiTemp
716                       rsSBETotal.filter = ""
717                       nilaiTemp = 0
                          
718                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
719                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
720                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
721                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
722                   Else
723                       !K1 = 0
724                       !K10 = 0
725                       !K20 = 0
726                       !K30 = 0
727                       !K40 = 0
728                       !K50 = 0
729                       !K60 = 0
730                       !K70 = 0
731                       !K80 = 0
732                       !tot16 = 0
733                       !tot19 = 0
734                       !tot29 = 0
735                       !tot26 = 0
736                   End If
                      
737                   If NulSID(rsModel!pareto) = 1 Then
738                       nilaiTotalPareto(0) = nilaiTotalPareto(0) + !K1
739                       nilaiTotalPareto(1) = nilaiTotalPareto(1) + !K10
740                       nilaiTotalPareto(2) = nilaiTotalPareto(2) + !K20
741                       nilaiTotalPareto(3) = nilaiTotalPareto(3) + !K30
742                       nilaiTotalPareto(4) = nilaiTotalPareto(4) + !K40
743                       nilaiTotalPareto(5) = nilaiTotalPareto(5) + !K50
744                       nilaiTotalPareto(6) = nilaiTotalPareto(6) + !K60
745                       nilaiTotalPareto(7) = nilaiTotalPareto(7) + !K70
746                       nilaiTotalPareto(8) = nilaiTotalPareto(8) + !K80
747                   Else
748                       nilaiTotalNonPareto(0) = nilaiTotalNonPareto(0) + !K1
749                       nilaiTotalNonPareto(1) = nilaiTotalNonPareto(1) + !K10
750                       nilaiTotalNonPareto(2) = nilaiTotalNonPareto(2) + !K20
751                       nilaiTotalNonPareto(3) = nilaiTotalNonPareto(3) + !K30
752                       nilaiTotalNonPareto(4) = nilaiTotalNonPareto(4) + !K40
753                       nilaiTotalNonPareto(5) = nilaiTotalNonPareto(5) + !K50
754                       nilaiTotalNonPareto(6) = nilaiTotalNonPareto(6) + !K60
755                       nilaiTotalNonPareto(7) = nilaiTotalNonPareto(7) + !K70
756                       nilaiTotalNonPareto(8) = nilaiTotalNonPareto(8) + !K80
757                   End If
                      
758                   If NulSID(rsModel!north) = 1 Then
759                       !north = 1
760                       nilaiTotalNorth(0) = nilaiTotalNorth(0) + !K1
761                       nilaiTotalNorth(1) = nilaiTotalNorth(1) + !K10
762                       nilaiTotalNorth(2) = nilaiTotalNorth(2) + !K20
763                       nilaiTotalNorth(3) = nilaiTotalNorth(3) + !K30
764                       nilaiTotalNorth(4) = nilaiTotalNorth(4) + !K40
765                       nilaiTotalNorth(5) = nilaiTotalNorth(5) + !K50
766                       nilaiTotalNorth(6) = nilaiTotalNorth(6) + !K60
767                       nilaiTotalNorth(7) = nilaiTotalNorth(7) + !K70
768                       nilaiTotalNorth(8) = nilaiTotalNorth(8) + !K80
769                   Else
                      
770                       !north = 3
771                       nilaiTotalSouth(0) = nilaiTotalSouth(0) + !K1
772                       nilaiTotalSouth(1) = nilaiTotalSouth(1) + !K10
773                       nilaiTotalSouth(2) = nilaiTotalSouth(2) + !K20
774                       nilaiTotalSouth(3) = nilaiTotalSouth(3) + !K30
775                       nilaiTotalSouth(4) = nilaiTotalSouth(4) + !K40
776                       nilaiTotalSouth(5) = nilaiTotalSouth(5) + !K50
777                       nilaiTotalSouth(6) = nilaiTotalSouth(6) + !K60
778                       nilaiTotalSouth(7) = nilaiTotalSouth(7) + !K70
779                       nilaiTotalSouth(8) = nilaiTotalSouth(8) + !K80
780                   End If
781                   !judul = "SB Total (NCM & Non NCM)"
782                   !model = Nul(rsModel.Fields(0).value)
783                   !Urut = 3
784                   .Update
785               End With
786               rsLaporan.filter = ""
                  
                  
787               For i = 0 To 10
788                   nilaiTempTotal(i) = 0
789               Next
790               rsLaporan.filter = ""
                  
791               If Not rsLaporan.EOF Then
792                   rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
793                   If Not rsLaporan.EOF Then
794                       If rsLaporan!K1 > 0 Then
795                           nilaiTempTotal(0) = rsLaporan!K1
796                       End If
797                       If rsLaporan!K10 > 0 Then
798                           nilaiTempTotal(1) = rsLaporan!K10
799                       End If
800                       If rsLaporan!K20 > 0 Then
801                           nilaiTempTotal(2) = rsLaporan!K20
802                       End If
803                       If rsLaporan!K30 > 0 Then
804                           nilaiTempTotal(3) = rsLaporan!K30
805                       End If
806                       If rsLaporan!K40 > 0 Then
807                           nilaiTempTotal(4) = rsLaporan!K40
808                       End If
809                       If rsLaporan!K50 > 0 Then
810                           nilaiTempTotal(5) = rsLaporan!K50
811                       End If
812                       If rsLaporan!K60 > 0 Then
813                           nilaiTempTotal(6) = rsLaporan!K60
814                       End If
815                       If rsLaporan!K70 > 0 Then
816                           nilaiTempTotal(7) = rsLaporan!K70
817                       End If
818                       If rsLaporan!K80 > 0 Then
819                           nilaiTempTotal(8) = rsLaporan!K80
820                       End If
821                   End If
822               End If
                  
823               With rsLaporan
824                   .AddNew
825                   If nilaiTempDemand(0) > 0 Then
826                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
827                   Else
828                       !K1 = 0 & " %"
829                   End If
830                   If nilaiTempDemand(1) > 0 Then
831                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
832                   Else
833                       !K10 = 0 & " %"
834                   End If
835                   If nilaiTempDemand(2) > 0 Then
836                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
837                   Else
838                       !K20 = 0 & " %"
839                   End If
840                   If nilaiTempDemand(3) > 0 Then
841                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
842                   Else
843                       !K30 = 0 & " %"
844                   End If
845                   If nilaiTempDemand(4) > 0 Then
846                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
847                   Else
848                       !K40 = 0 & " %"
849                   End If
850                   If nilaiTempDemand(5) > 0 Then
851                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
852                   Else
853                       !K50 = 0 & " %"
854                   End If
855                   If nilaiTempDemand(6) > 0 Then
856                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
857                   Else
858                       !K60 = 0 & " %"
859                   End If
860                   If nilaiTempDemand(7) > 0 Then
861                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
862                   Else
863                       !K70 = 0 & " %"
864                   End If
865                   If nilaiTempDemand(8) > 0 Then
866                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
867                   Else
868                       !K80 = 0 & " %"
869                   End If
                      
                      
870                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
871                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
872                   Else
873                       !tot16 = 0 & " %"
874                   End If
                      
875                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
876                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
877                   Else
878                       !tot26 = 0 & " %"
879                   End If
                       
880                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
881                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
882                   Else
883                       !tot19 = 0 & " %"
884                   End If
                       
885                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
886                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
887                   Else
888                       !tot29 = 0 & " %"
889                   End If
                      
890                   If NulSID(rsModel!north) = 1 Then
891                       !north = 1
892                   Else
893                       !north = 3
894                   End If
895                   !Urut = 5
896                   !judul = "Retention Rate Same Outlet"
897                   !model = Nul(rsModel.Fields(0).value)
898                   .Update
899               End With
                  
900               For i = 0 To 10
901                   nilaiTempTotal(i) = 0
902               Next
903               rsLaporan.filter = ""
                  
904               If Not rsLaporan.EOF Then
905                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
906                   If Not rsLaporan.EOF Then
907                       If rsLaporan!K1 > 0 Then
908                           nilaiTempTotal(0) = rsLaporan!K1
909                       End If
910                       If rsLaporan!K10 > 0 Then
911                           nilaiTempTotal(1) = rsLaporan!K10
912                       End If
913                       If rsLaporan!K20 > 0 Then
914                           nilaiTempTotal(2) = rsLaporan!K20
915                       End If
916                       If rsLaporan!K30 > 0 Then
917                           nilaiTempTotal(3) = rsLaporan!K30
918                       End If
919                       If rsLaporan!K40 > 0 Then
920                           nilaiTempTotal(4) = rsLaporan!K40
921                       End If
922                       If rsLaporan!K50 > 0 Then
923                           nilaiTempTotal(5) = rsLaporan!K50
924                       End If
925                       If rsLaporan!K60 > 0 Then
926                           nilaiTempTotal(6) = rsLaporan!K60
927                       End If
928                       If rsLaporan!K70 > 0 Then
929                           nilaiTempTotal(7) = rsLaporan!K70
930                       End If
931                       If rsLaporan!K80 > 0 Then
932                           nilaiTempTotal(8) = rsLaporan!K80
933                       End If
934                   End If
935               End If
                  
936               With rsLaporan
937                   .AddNew
938                   If nilaiTempDemand(0) > 0 Then
939                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
940                   Else
941                       !K1 = 0 & " %"
942                   End If
943                   If nilaiTempDemand(1) > 0 Then
944                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
945                   Else
946                       !K10 = 0 & " %"
947                   End If
948                   If nilaiTempDemand(2) > 0 Then
949                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
950                   Else
951                       !K20 = 0 & " %"
952                   End If
953                   If nilaiTempDemand(3) > 0 Then
954                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
955                   Else
956                       !K30 = 0 & " %"
957                   End If
958                   If nilaiTempDemand(4) > 0 Then
959                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
960                   Else
961                       !K40 = 0 & " %"
962                   End If
963                   If nilaiTempDemand(5) > 0 Then
964                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
965                   Else
966                       !K50 = 0 & " %"
967                   End If
968                   If nilaiTempDemand(6) > 0 Then
969                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
970                   Else
971                       !K60 = 0 & " %"
972                   End If
973                   If nilaiTempDemand(7) > 0 Then
974                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
975                   Else
976                       !K70 = 0 & " %"
977                   End If
978                   If nilaiTempDemand(8) > 0 Then
979                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
980                   Else
981                       !K80 = 0 & " %"
982                   End If
                      
                      
983                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
984                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
985                   Else
986                       !tot16 = 0 & " %"
987                   End If
                      
988                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
989                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
990                   Else
991                       !tot26 = 0 & " %"
992                   End If
                       
993                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
994                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
995                   Else
996                       !tot19 = 0 & " %"
997                   End If
                       
998                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
999                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
1000                  Else
1001                      !tot29 = 0 & " %"
1002                  End If
                      
1003                  If NulSID(rsModel!north) = 1 Then
1004                      !north = 1
1005                  Else
1006                      !north = 3
1007                  End If
1008                  !Urut = 7
1009                  !judul = "Retention Rate Total (NCM & Non NCM)"
1010                  !model = Nul(rsModel.Fields(0).value)
1011                  .Update
1012              End With
1013              rsLaporan.filter = ""
1014              rsModel.MoveNext
1015              progbarDealer.value = progbarDealer.value + 1
1016              DoEvents
1017          Loop
1018          Call setNilaiKategoriDealer
              
1019      Else
1020          MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
1021          Exit Sub
              
1022      End If
1023      rsLaporan.Sort = "north asc"
1024      Report.database.SetDataSource rsLaporan
1025      progbarDealer.Visible = False
1026      Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
1027      Report.txtPeriode.SetText "Periode : " & Format(tglDealerAwal, "dd-mm-yyyy") & " s/d " & Format(tglDealerAkhir, "dd-mm-yyyy")
          
1028      With CRViewer1
1029          .ReportSource = Report
1030          .ViewReport
1031          .Zoom 100
1032          .Visible = True
1033          .Left = 300
1034          .Height = Me.Height - 1200
1035          .Width = Me.Width - 750
1036      End With
         
1037      cmdTutup.Visible = True
1038      cmdTutup.Left = CRViewer1.Left + 20
1039      cmdTutup.Top = CRViewer1.Top + 400
1040      Screen.MousePointer = vbDefault
          
1041      Call DisableExport
1042      Call setLokasiTombol
End Sub


Private Sub ambilLaporanVinDealerTglPDS()
          Dim rsModel             As ADODB.Recordset
          Dim rsSBE               As ADODB.Recordset
          Dim rsDemand(10)        As ADODB.Recordset
          Dim rsSBETotal          As ADODB.Recordset
          Dim nilaiTemp           As Double
          Dim nilaiTempDemand(10) As Double
          Dim nilaiTempTotal(10)  As Double
          
1         For i = 0 To 10
2             nilaiSBNRMPareto(i) = 0
3             nilaiSBNRMNonPareto(i) = 0
4             nilaiSBNRMSouth(i) = 0
5             nilaiSBNRMNorth(i) = 0
              
6             nilaiDemandPareto(i) = 0
7             nilaiDemandNonPareto(i) = 0
8             nilaiDemandSouth(i) = 0
9             nilaiDemandNorth(i) = 0
              
10            nilaiTotalNonPareto(i) = 0
11            nilaiTotalPareto(i) = 0
12            nilaiTotalSouth(i) = 0
13            nilaiTotalNorth(i) = 0
              
14            nilaiSBSameOutletPareto(i) = 0
15            nilaiSBSameOutletNonPareto(i) = 0
16            nilaiSBSameOutletSouth(i) = 0
17            nilaiSBSameOutletNorth(i) = 0
18        Next
          
19        Set Report = New rptSBERateByModelVinRetentionDealer
20        Call createTableLaporanVin
          
21        sSQL = "select FS_NM_DEALER,FS_KD_DEALER,pareto,north,kodecabang from T_DEALER where INTERNAL = 1 and FS_LOG_DEL = '' and FS_KD_DEALER <> 'NR 00' order by north desc, FS_KD_DEALER"
22        Set rsModel = objO.FetchDataHO(sSQL)
        
23        sSQL = "select COUNT(noinv) total,fs_rc,"
24        sSQL = sSQL & " E.FS_NM_DEALER"
25        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
26        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
27        sSQL = sSQL & " inner JOIN T_DEALER E WITH (NOLOCK) ON e.FS_KD_DEALER = t.FS_KD_DEALER and e.fs_log_del = '' and INTERNAL = 1"
28        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
29        sSQL = sSQL & " AND INVDATE between '" & Format(tglDealerAwal.value, "mm/dd/yyyy") & "' AND '" & Format(tglDealerAkhir.value, "mm/dd/yyyy") & "'"
30        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
31        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
32        sSQL = sSQL & " grOUP BY FS_RC, e.FS_NM_DEALER"
33        sSQL = sSQL & " order by FS_RC,e.FS_NM_DEALER"
           
34        Set rsSBETotal = objO.FetchDataHO(sSQL)
          
          
35        sSQL = "select COUNT(noinv) total,fs_rc,case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan,"
36        sSQL = sSQL & " D.FS_NM_DEALER"
37        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK)"
38        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1' "
39        sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
40        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''"
41        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
42        sSQL = sSQL & " AND INVDATE between '" & Format(tglDealerAwal.value, "mm/dd/yyyy") & "' AND '" & Format(tglDealerAkhir.value, "mm/dd/yyyy") & "'"
43        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080')"
44        sSQL = sSQL & " and t.TPSVC <> 'w' and t.PDS = 0"
45        sSQL = sSQL & " grOUP BY FS_RC , case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end , d.FS_NM_DEALER"
46        sSQL = sSQL & " order by FS_RC,d.FS_NM_DEALER"
          
47        sSQL = "select D.FS_NM_DEALER,COUNT(noinv) total,fs_rc,case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end as keterangan"
48        sSQL = sSQL & " From T_INVOICEnew T WITH (NOLOCK) "
49        sSQL = sSQL & " LEFT JOIN [SPK99]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA and s.tipetr = 'ND' And s.[Status] = '1'  "
50        sSQL = sSQL & " LEFT JOIN SingleIDOri..UnitCustomer C WITH (NOLOCK) ON T.NO_CHASSIS = C.norangka"
51        sSQL = sSQL & " LEFT JOIN T_DEALER D WITH (NOLOCK) ON d.FS_KD_DEALER = C.KdDlrPenjual and d.FS_LOG_DEL = ''" & dealerPilih
52        sSQL = sSQL & " LEFT JOIN model E WITH (NOLOCK) ON e.mdl = t.mdl and e.fs_log_del = '' "
53        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND FS_NO_RETUR = ''  and FS_RC <> ''"
54        sSQL = sSQL & " and T.NO_CHASSIS in (SELECT NORANGKA FROM [PDS00]..IntrnCbg WHERE Proses = 10 and Kode = 2 AND TGLSELESAI between '" & Format(tglDealerAwal, "mm/dd/yyyy") & "' AND '" & Format(tglDealerAkhir, "mm/dd/yyyy") & "' )"
55        sSQL = sSQL & " and t.FS_RC in ('R001','R010','R020','R030','R040','R050','R060','R070','R080' ) and t.TPSVC <> 'w' and t.PDS = 0"
56        sSQL = sSQL & " grOUP BY FS_RC ,d.FS_NM_DEALER, case when isnull(d.INTERNAL,0) = '1' then 'NCM' else 'Non NCM' end "
57        sSQL = sSQL & " order by FS_RC,d.FS_NM_DEALER"
58        Set rsSBE = objO.FetchDataHO(sSQL)
59        If Not rsModel.EOF Then
60            rsModel.MoveFirst
61            progbarDealer.Max = rsModel.RecordCount + 1
62            progbarDealer.value = 1
63            progbarDealer.Visible = True
64            DoEvents
65            Do While Not rsModel.EOF
                  
                  ''LAPORAN SAME OUTLET
66                With rsLaporan
67                    .AddNew
68                    rsSBE.filter = ""
69                    If Not rsSBE.EOF Then
70                        rsSBE.filter = ""
71                        nilaiTemp = 0
72                        rsSBE.filter = "fs_rc = 'R001' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
73                        If Not rsSBE.EOF Then
74                            Do While Not rsSBE.EOF
75                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
76                                rsSBE.MoveNext
77                            Loop
78                        End If
79                        !K1 = nilaiTemp
80                        rsSBE.filter = ""
81                        nilaiTemp = 0
                          
82                        rsSBE.filter = "fs_rc = 'R010' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
83                        If Not rsSBE.EOF Then
84                            Do While Not rsSBE.EOF
85                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
86                                rsSBE.MoveNext
87                            Loop
88                        End If
89                        !K10 = nilaiTemp
90                        rsSBE.filter = ""
91                        nilaiTemp = 0
                          
92                        rsSBE.filter = "fs_rc = 'R020' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
93                        If Not rsSBE.EOF Then
94                            Do While Not rsSBE.EOF
95                                nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
96                                rsSBE.MoveNext
97                            Loop
98                        End If
99                        !K20 = nilaiTemp
100                       rsSBE.filter = ""
101                       nilaiTemp = 0
                          
102                       rsSBE.filter = "fs_rc = 'R030' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
103                       If Not rsSBE.EOF Then
104                           Do While Not rsSBE.EOF
105                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
106                               rsSBE.MoveNext
107                           Loop
108                       End If
109                       !K30 = nilaiTemp
110                       rsSBE.filter = ""
111                       nilaiTemp = 0
                          
112                       rsSBE.filter = "fs_rc = 'R040' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
113                       If Not rsSBE.EOF Then
114                           Do While Not rsSBE.EOF
115                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
116                               rsSBE.MoveNext
117                           Loop
118                       End If
119                       !K40 = nilaiTemp
120                       rsSBE.filter = ""
121                       nilaiTemp = 0
                          
122                       rsSBE.filter = "fs_rc = 'R050' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
123                       If Not rsSBE.EOF Then
124                           Do While Not rsSBE.EOF
125                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
126                               rsSBE.MoveNext
127                           Loop
128                       End If
129                       !K50 = nilaiTemp
130                       rsSBE.filter = ""
131                       nilaiTemp = 0
                          
132                       rsSBE.filter = "fs_rc = 'R060' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
133                       If Not rsSBE.EOF Then
134                           Do While Not rsSBE.EOF
135                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
136                               rsSBE.MoveNext
137                           Loop
138                       End If
139                       !K60 = nilaiTemp
140                       rsSBE.filter = ""
141                       nilaiTemp = 0
                          
142                       rsSBE.filter = "fs_rc = 'R070' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
143                       If Not rsSBE.EOF Then
144                           Do While Not rsSBE.EOF
145                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
146                               rsSBE.MoveNext
147                           Loop
148                       End If
149                       !K70 = nilaiTemp
150                       rsSBE.filter = ""
151                       nilaiTemp = 0
                          
152                       rsSBE.filter = "fs_rc = 'R080' and keterangan = 'Same Outlet' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
153                       If Not rsSBE.EOF Then
154                           Do While Not rsSBE.EOF
155                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
156                               rsSBE.MoveNext
157                           Loop
158                       End If
159                       !K80 = nilaiTemp
160                       rsSBE.filter = ""
161                       nilaiTemp = 0
                          
162                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
163                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
164                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
165                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
166                   Else
167                       !K1 = 0
168                       !K10 = 0
169                       !K20 = 0
170                       !K30 = 0
171                       !K40 = 0
172                       !K50 = 0
173                       !K60 = 0
174                       !K70 = 0
175                       !K80 = 0
176                       !tot16 = 0
177                       !tot19 = 0
                          
178                       !tot26 = 0
179                       !tot29 = 0
                          
180                   End If
181                   If NulSID(rsModel!pareto) = 1 Then
182                       nilaiSBSameOutletPareto(0) = nilaiSBSameOutletPareto(0) + !K1
183                       nilaiSBSameOutletPareto(1) = nilaiSBSameOutletPareto(1) + !K10
184                       nilaiSBSameOutletPareto(2) = nilaiSBSameOutletPareto(2) + !K20
185                       nilaiSBSameOutletPareto(3) = nilaiSBSameOutletPareto(3) + !K30
186                       nilaiSBSameOutletPareto(4) = nilaiSBSameOutletPareto(4) + !K40
187                       nilaiSBSameOutletPareto(5) = nilaiSBSameOutletPareto(5) + !K50
188                       nilaiSBSameOutletPareto(6) = nilaiSBSameOutletPareto(6) + !K60
189                       nilaiSBSameOutletPareto(7) = nilaiSBSameOutletPareto(7) + !K70
190                       nilaiSBSameOutletPareto(8) = nilaiSBSameOutletPareto(8) + !K80
191                   Else
192                       nilaiSBSameOutletNonPareto(0) = nilaiSBSameOutletNonPareto(0) + !K1
193                       nilaiSBSameOutletNonPareto(1) = nilaiSBSameOutletNonPareto(1) + !K10
194                       nilaiSBSameOutletNonPareto(2) = nilaiSBSameOutletNonPareto(2) + !K20
195                       nilaiSBSameOutletNonPareto(3) = nilaiSBSameOutletNonPareto(3) + !K30
196                       nilaiSBSameOutletNonPareto(4) = nilaiSBSameOutletNonPareto(4) + !K40
197                       nilaiSBSameOutletNonPareto(5) = nilaiSBSameOutletNonPareto(5) + !K50
198                       nilaiSBSameOutletNonPareto(6) = nilaiSBSameOutletNonPareto(6) + !K60
199                       nilaiSBSameOutletNonPareto(7) = nilaiSBSameOutletNonPareto(7) + !K70
200                       nilaiSBSameOutletNonPareto(8) = nilaiSBSameOutletNonPareto(8) + !K80
201                   End If
                      
202                   If NulSID(rsModel!north) = 1 Then
203                       !north = 1
204                       nilaiSBSameOutletNorth(0) = nilaiSBSameOutletNorth(0) + !K1
205                       nilaiSBSameOutletNorth(1) = nilaiSBSameOutletNorth(1) + !K10
206                       nilaiSBSameOutletNorth(2) = nilaiSBSameOutletNorth(2) + !K20
207                       nilaiSBSameOutletNorth(3) = nilaiSBSameOutletNorth(3) + !K30
208                       nilaiSBSameOutletNorth(4) = nilaiSBSameOutletNorth(4) + !K40
209                       nilaiSBSameOutletNorth(5) = nilaiSBSameOutletNorth(5) + !K50
210                       nilaiSBSameOutletNorth(6) = nilaiSBSameOutletNorth(6) + !K60
211                       nilaiSBSameOutletNorth(7) = nilaiSBSameOutletNorth(7) + !K70
212                       nilaiSBSameOutletNorth(8) = nilaiSBSameOutletNorth(8) + !K80
213                   Else
                      
214                       !north = 3
215                       nilaiSBSameOutletSouth(0) = nilaiSBSameOutletSouth(0) + !K1
216                       nilaiSBSameOutletSouth(1) = nilaiSBSameOutletSouth(1) + !K10
217                       nilaiSBSameOutletSouth(2) = nilaiSBSameOutletSouth(2) + !K20
218                       nilaiSBSameOutletSouth(3) = nilaiSBSameOutletSouth(3) + !K30
219                       nilaiSBSameOutletSouth(4) = nilaiSBSameOutletSouth(4) + !K40
220                       nilaiSBSameOutletSouth(5) = nilaiSBSameOutletSouth(5) + !K50
221                       nilaiSBSameOutletSouth(6) = nilaiSBSameOutletSouth(6) + !K60
222                       nilaiSBSameOutletSouth(7) = nilaiSBSameOutletSouth(7) + !K70
223                       nilaiSBSameOutletSouth(8) = nilaiSBSameOutletSouth(8) + !K80
224                   End If
225                   !judul = "SB Same Outlet"
226                   !model = Nul(rsModel.Fields(0).value)
227                   !Urut = 1
228                   .Update
229               End With
                  
                  ''LAPORAN SAME GROUP
230               With rsLaporan
231                   .AddNew
232                   rsSBE.filter = ""
233                   If Not rsSBE.EOF Then
234                       rsSBE.filter = ""
235                       nilaiTemp = 0
236                       rsSBE.filter = "fs_rc = 'R001' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
237                       If Not rsSBE.EOF Then
238                           Do While Not rsSBE.EOF
239                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
240                               rsSBE.MoveNext
241                           Loop
242                       End If
243                       !K1 = nilaiTemp
244                       rsSBE.filter = ""
245                       nilaiTemp = 0
                          
246                       rsSBE.filter = "fs_rc = 'R010' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
247                       If Not rsSBE.EOF Then
248                           Do While Not rsSBE.EOF
249                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
250                               rsSBE.MoveNext
251                           Loop
252                       End If
253                       !K10 = nilaiTemp
254                       rsSBE.filter = ""
255                       nilaiTemp = 0
                          
256                       rsSBE.filter = "fs_rc = 'R020' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
257                       If Not rsSBE.EOF Then
258                           Do While Not rsSBE.EOF
259                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
260                               rsSBE.MoveNext
261                           Loop
262                       End If
263                       !K20 = nilaiTemp
264                       rsSBE.filter = ""
265                       nilaiTemp = 0
                          
266                       rsSBE.filter = "fs_rc = 'R030' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
267                       If Not rsSBE.EOF Then
268                           Do While Not rsSBE.EOF
269                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
270                               rsSBE.MoveNext
271                           Loop
272                       End If
273                       !K30 = nilaiTemp
274                       rsSBE.filter = ""
275                       nilaiTemp = 0
                          
276                       rsSBE.filter = "fs_rc = 'R040' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
277                       If Not rsSBE.EOF Then
278                           Do While Not rsSBE.EOF
279                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
280                               rsSBE.MoveNext
281                           Loop
282                       End If
283                       !K40 = nilaiTemp
284                       rsSBE.filter = ""
285                       nilaiTemp = 0
                          
286                       rsSBE.filter = "fs_rc = 'R050' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
287                       If Not rsSBE.EOF Then
288                           Do While Not rsSBE.EOF
289                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
290                               rsSBE.MoveNext
291                           Loop
292                       End If
293                       !K50 = nilaiTemp
294                       rsSBE.filter = ""
295                       nilaiTemp = 0
                          
296                       rsSBE.filter = "fs_rc = 'R060' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
297                       If Not rsSBE.EOF Then
298                           Do While Not rsSBE.EOF
299                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
300                               rsSBE.MoveNext
301                           Loop
302                       End If
303                       !K60 = nilaiTemp
304                       rsSBE.filter = ""
305                       nilaiTemp = 0
                          
306                       rsSBE.filter = "fs_rc = 'R070' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
307                       If Not rsSBE.EOF Then
308                           Do While Not rsSBE.EOF
309                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
310                               rsSBE.MoveNext
311                           Loop
312                       End If
313                       !K70 = nilaiTemp
314                       rsSBE.filter = ""
315                       nilaiTemp = 0
                          
316                       rsSBE.filter = "fs_rc = 'R080' and keterangan <> 'Non NCM' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
317                       If Not rsSBE.EOF Then
318                           Do While Not rsSBE.EOF
319                               nilaiTemp = nilaiTemp + CDbl(rsSBE!total)
320                               rsSBE.MoveNext
321                           Loop
322                       End If
323                       !K80 = nilaiTemp
324                       rsSBE.filter = ""
325                       nilaiTemp = 0
                          
326                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
327                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
328                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
329                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
330                   Else
331                       !K1 = 0
332                       !K10 = 0
333                       !K20 = 0
334                       !K30 = 0
335                       !K40 = 0
336                       !K50 = 0
337                       !K60 = 0
338                       !K70 = 0
339                       !K80 = 0
340                       !tot16 = 0
341                       !tot19 = 0
                          
342                       !tot26 = 0
343                       !tot29 = 0
                          
344                   End If
345                   If NulSID(rsModel!pareto) = 1 Then
346                       nilaiSBNRMPareto(0) = nilaiSBNRMPareto(0) + !K1
347                       nilaiSBNRMPareto(1) = nilaiSBNRMPareto(1) + !K10
348                       nilaiSBNRMPareto(2) = nilaiSBNRMPareto(2) + !K20
349                       nilaiSBNRMPareto(3) = nilaiSBNRMPareto(3) + !K30
350                       nilaiSBNRMPareto(4) = nilaiSBNRMPareto(4) + !K40
351                       nilaiSBNRMPareto(5) = nilaiSBNRMPareto(5) + !K50
352                       nilaiSBNRMPareto(6) = nilaiSBNRMPareto(6) + !K60
353                       nilaiSBNRMPareto(7) = nilaiSBNRMPareto(7) + !K70
354                       nilaiSBNRMPareto(8) = nilaiSBNRMPareto(8) + !K80
355                   Else
356                       nilaiSBNRMNonPareto(0) = nilaiSBNRMNonPareto(0) + !K1
357                       nilaiSBNRMNonPareto(1) = nilaiSBNRMNonPareto(1) + !K10
358                       nilaiSBNRMNonPareto(2) = nilaiSBNRMNonPareto(2) + !K20
359                       nilaiSBNRMNonPareto(3) = nilaiSBNRMNonPareto(3) + !K30
360                       nilaiSBNRMNonPareto(4) = nilaiSBNRMNonPareto(4) + !K40
361                       nilaiSBNRMNonPareto(5) = nilaiSBNRMNonPareto(5) + !K50
362                       nilaiSBNRMNonPareto(6) = nilaiSBNRMNonPareto(6) + !K60
363                       nilaiSBNRMNonPareto(7) = nilaiSBNRMNonPareto(7) + !K70
364                       nilaiSBNRMNonPareto(8) = nilaiSBNRMNonPareto(8) + !K80
365                   End If
                      
366                   If NulSID(rsModel!north) = 1 Then
367                       !north = 1
368                       nilaiSBNRMNorth(0) = nilaiSBNRMNorth(0) + !K1
369                       nilaiSBNRMNorth(1) = nilaiSBNRMNorth(1) + !K10
370                       nilaiSBNRMNorth(2) = nilaiSBNRMNorth(2) + !K20
371                       nilaiSBNRMNorth(3) = nilaiSBNRMNorth(3) + !K30
372                       nilaiSBNRMNorth(4) = nilaiSBNRMNorth(4) + !K40
373                       nilaiSBNRMNorth(5) = nilaiSBNRMNorth(5) + !K50
374                       nilaiSBNRMNorth(6) = nilaiSBNRMNorth(6) + !K60
375                       nilaiSBNRMNorth(7) = nilaiSBNRMNorth(7) + !K70
376                       nilaiSBNRMNorth(8) = nilaiSBNRMNorth(8) + !K80
377                   Else
                      
378                       !north = 3
379                       nilaiSBNRMSouth(0) = nilaiSBNRMSouth(0) + !K1
380                       nilaiSBNRMSouth(1) = nilaiSBNRMSouth(1) + !K10
381                       nilaiSBNRMSouth(2) = nilaiSBNRMSouth(2) + !K20
382                       nilaiSBNRMSouth(3) = nilaiSBNRMSouth(3) + !K30
383                       nilaiSBNRMSouth(4) = nilaiSBNRMSouth(4) + !K40
384                       nilaiSBNRMSouth(5) = nilaiSBNRMSouth(5) + !K50
385                       nilaiSBNRMSouth(6) = nilaiSBNRMSouth(6) + !K60
386                       nilaiSBNRMSouth(7) = nilaiSBNRMSouth(7) + !K70
387                       nilaiSBNRMSouth(8) = nilaiSBNRMSouth(8) + !K80
388                   End If
389                   !judul = "SB Same Group (NCM)"
390                   !model = Nul(rsModel.Fields(0).value)
391                   !Urut = 2
392                   .Update
393               End With
                  
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
                  ''2020-03-18 || DEMAND SAME GROUP
                  ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
394               With rsLaporan
395                   .AddNew
                  
396                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.dealer from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.dealer ) a "
397                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-1,DATEADD(MONTH,0,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,0,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
398                   !K1 = objO.FetchDataSPK_Double(sSQL)
                      
399                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
400                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-4,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
401                   !K10 = objO.FetchDataSPK_Double(sSQL)
                      
402                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
403                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-10,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
404                   !K20 = objO.FetchDataSPK_Double(sSQL)
                      
405                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
406                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-16,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
407                   !K30 = objO.FetchDataSPK_Double(sSQL)
                      
408                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
409                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-22,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
410                   !K40 = objO.FetchDataSPK_Double(sSQL)
                      
411                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
412                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-28,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
413                   !K50 = objO.FetchDataSPK_Double(sSQL)
                      
414                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
415                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-34,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
416                   !K60 = objO.FetchDataSPK_Double(sSQL)
                      
417                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
418                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-40,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
419                   !K70 = objO.FetchDataSPK_Double(sSQL)
                      
420                   sSQL = "select isnull(sum(a.total),0) total from (select (sum(case a.typetrn when 'A' then 1 else 0 end) - sum(case a.typetrn when 'K' then 1 else 0 end)) as total,MONTH(a.TglFaktur) as bulan,YEAR(a.TglFaktur) as tahun , b.kodetype from [spk00]..Sihdr a INNER JOIN [spk99]..DataSPK b ON a.nospk = b.NoSPK and b.Status = 1 where b.dealer = '" & NulSID(rsModel!kodeCabang) & "' and a.TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "')))) group by MONTH(a.TglFaktur),YEAR(a.TglFaktur) ,b.kodetype ) a "
421                   sSQL = "select COUNT(*) total from SPK00..DoDealer b inner join SPK99..summary a on b.KodeDealer = a.KodeDealer where a.KodeDealer = '" & NulSID(rsModel!kodeCabang) & "' and TipeTr = 'ND' and Status = 1 and Tanggal between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & Format(Month(tglDealerAwal), "0#") & "/01/" & year(tglDealerAwal) & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-46,'" & Format(Month(tglDealerAkhir), "0#") & "/01/" & year(tglDealerAkhir) & "'))))"
422                   !K80 = objO.FetchDataSPK_Double(sSQL)
                      
423                   !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
424                   !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
425                   !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
426                   !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                      
427                   If NulSID(rsModel!pareto) = 1 Then
428                       nilaiDemandPareto(0) = nilaiDemandPareto(0) + !K1
429                       nilaiDemandPareto(1) = nilaiDemandPareto(1) + !K10
430                       nilaiDemandPareto(2) = nilaiDemandPareto(2) + !K20
431                       nilaiDemandPareto(3) = nilaiDemandPareto(3) + !K30
432                       nilaiDemandPareto(4) = nilaiDemandPareto(4) + !K40
433                       nilaiDemandPareto(5) = nilaiDemandPareto(5) + !K50
434                       nilaiDemandPareto(6) = nilaiDemandPareto(6) + !K60
435                       nilaiDemandPareto(7) = nilaiDemandPareto(7) + !K70
436                       nilaiDemandPareto(8) = nilaiDemandPareto(8) + !K80
437                   Else
438                       nilaiDemandNonPareto(0) = nilaiDemandNonPareto(0) + !K1
439                       nilaiDemandNonPareto(1) = nilaiDemandNonPareto(1) + !K10
440                       nilaiDemandNonPareto(2) = nilaiDemandNonPareto(2) + !K20
441                       nilaiDemandNonPareto(3) = nilaiDemandNonPareto(3) + !K30
442                       nilaiDemandNonPareto(4) = nilaiDemandNonPareto(4) + !K40
443                       nilaiDemandNonPareto(5) = nilaiDemandNonPareto(5) + !K50
444                       nilaiDemandNonPareto(6) = nilaiDemandNonPareto(6) + !K60
445                       nilaiDemandNonPareto(7) = nilaiDemandNonPareto(7) + !K70
446                       nilaiDemandNonPareto(8) = nilaiDemandNonPareto(8) + !K80
447                   End If
                      
448                   If NulSID(rsModel!north) = 1 Then
449                       !north = 1
450                       nilaiDemandNorth(0) = nilaiDemandNorth(0) + !K1
451                       nilaiDemandNorth(1) = nilaiDemandNorth(1) + !K10
452                       nilaiDemandNorth(2) = nilaiDemandNorth(2) + !K20
453                       nilaiDemandNorth(3) = nilaiDemandNorth(3) + !K30
454                       nilaiDemandNorth(4) = nilaiDemandNorth(4) + !K40
455                       nilaiDemandNorth(5) = nilaiDemandNorth(5) + !K50
456                       nilaiDemandNorth(6) = nilaiDemandNorth(6) + !K60
457                       nilaiDemandNorth(7) = nilaiDemandNorth(7) + !K70
458                       nilaiDemandNorth(8) = nilaiDemandNorth(8) + !K80
459                   Else
                      
460                       !north = 3
461                       nilaiSBNRMSouth(0) = nilaiSBNRMSouth(0) + !K1
462                       nilaiSBNRMSouth(1) = nilaiSBNRMSouth(1) + !K10
463                       nilaiSBNRMSouth(2) = nilaiSBNRMSouth(2) + !K20
464                       nilaiSBNRMSouth(3) = nilaiSBNRMSouth(3) + !K30
465                       nilaiSBNRMSouth(4) = nilaiSBNRMSouth(4) + !K40
466                       nilaiSBNRMSouth(5) = nilaiSBNRMSouth(5) + !K50
467                       nilaiSBNRMSouth(6) = nilaiSBNRMSouth(6) + !K60
468                       nilaiSBNRMSouth(7) = nilaiSBNRMSouth(7) + !K70
469                       nilaiSBNRMSouth(8) = nilaiSBNRMSouth(8) + !K80
470                   End If
                       
471                   !judul = "Demand Same Outlet"
472                   !model = Nul(rsModel.Fields(0).value)
473                   !Urut = 4
474                   .Update
475               End With
                  ''------------------------------------------------------------
                  '' AKHIR DEMAND
                  ''------------------------------------------------------------
                  
                  
                  ''-----------------------------------------------------------
                  ''2020-03-18 || MENGAMBIL DATA PERSENTASE RETENTION TOTAL   =
                  ''-----------------------------------------------------------
476               For i = 0 To 10
477                   nilaiTempDemand(i) = 0
478                   nilaiTempTotal(i) = 0
479               Next
480               rsLaporan.filter = ""
481               If Not rsLaporan.EOF Then
482                   rsLaporan.filter = "judul = 'Demand Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
483                   If Not rsLaporan.EOF Then
484                       If rsLaporan!K1 > 0 Then
485                           nilaiTempDemand(0) = rsLaporan!K1
486                       End If
487                       If rsLaporan!K10 > 0 Then
488                           nilaiTempDemand(1) = rsLaporan!K10
489                       End If
490                       If rsLaporan!K20 > 0 Then
491                           nilaiTempDemand(2) = rsLaporan!K20
492                       End If
493                       If rsLaporan!K30 > 0 Then
494                           nilaiTempDemand(3) = rsLaporan!K30
495                       End If
496                       If rsLaporan!K40 > 0 Then
497                           nilaiTempDemand(4) = rsLaporan!K40
498                       End If
499                       If rsLaporan!K50 > 0 Then
500                           nilaiTempDemand(5) = rsLaporan!K50
501                       End If
502                       If rsLaporan!K60 > 0 Then
503                           nilaiTempDemand(6) = rsLaporan!K60
504                       End If
505                       If rsLaporan!K70 > 0 Then
506                           nilaiTempDemand(7) = rsLaporan!K70
507                       End If
508                       If rsLaporan!K80 > 0 Then
509                           nilaiTempDemand(8) = rsLaporan!K80
510                       End If
511                   End If
512               End If
513               rsLaporan.filter = ""
                  
514               If Not rsLaporan.EOF Then
515                   rsLaporan.filter = "judul = 'SB Same Group (NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
516                   If Not rsLaporan.EOF Then
517                       If rsLaporan!K1 > 0 Then
518                           nilaiTempTotal(0) = rsLaporan!K1
519                       End If
520                       If rsLaporan!K10 > 0 Then
521                           nilaiTempTotal(1) = rsLaporan!K10
522                       End If
523                       If rsLaporan!K20 > 0 Then
524                           nilaiTempTotal(2) = rsLaporan!K20
525                       End If
526                       If rsLaporan!K30 > 0 Then
527                           nilaiTempTotal(3) = rsLaporan!K30
528                       End If
529                       If rsLaporan!K40 > 0 Then
530                           nilaiTempTotal(4) = rsLaporan!K40
531                       End If
532                       If rsLaporan!K50 > 0 Then
533                           nilaiTempTotal(5) = rsLaporan!K50
534                       End If
535                       If rsLaporan!K60 > 0 Then
536                           nilaiTempTotal(6) = rsLaporan!K60
537                       End If
538                       If rsLaporan!K70 > 0 Then
539                           nilaiTempTotal(7) = rsLaporan!K70
540                       End If
541                       If rsLaporan!K80 > 0 Then
542                           nilaiTempTotal(8) = rsLaporan!K80
543                       End If
544                   End If
545               End If
                  
546               With rsLaporan
547                   .AddNew
548                   If nilaiTempDemand(0) > 0 Then
549                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
550                   Else
551                       !K1 = 0 & " %"
552                   End If
553                   If nilaiTempDemand(1) > 0 Then
554                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
555                   Else
556                       !K10 = 0 & " %"
557                   End If
558                   If nilaiTempDemand(2) > 0 Then
559                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
560                   Else
561                       !K20 = 0 & " %"
562                   End If
563                   If nilaiTempDemand(3) > 0 Then
564                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
565                   Else
566                       !K30 = 0 & " %"
567                   End If
568                   If nilaiTempDemand(4) > 0 Then
569                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
570                   Else
571                       !K40 = 0 & " %"
572                   End If
573                   If nilaiTempDemand(5) > 0 Then
574                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
575                   Else
576                       !K50 = 0 & " %"
577                   End If
578                   If nilaiTempDemand(6) > 0 Then
579                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
580                   Else
581                       !K60 = 0 & " %"
582                   End If
583                   If nilaiTempDemand(7) > 0 Then
584                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
585                   Else
586                       !K70 = 0 & " %"
587                   End If
588                   If nilaiTempDemand(8) > 0 Then
589                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
590                   Else
591                       !K80 = 0 & " %"
592                   End If
                      
593                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
594                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
595                   Else
596                       !tot16 = 0 & " %"
597                   End If
                      
598                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
599                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
600                   Else
601                       !tot26 = 0 & " %"
602                   End If
                       
603                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
604                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
605                   Else
606                       !tot19 = 0 & " %"
607                   End If
                       
608                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
609                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
610                   Else
611                       !tot29 = 0 & " %"
612                   End If
                      
613                   If NulSID(rsModel!north) = 1 Then
614                       !north = 1
615                   Else
616                       !north = 3
617                   End If
618                   !Urut = 6
619                   !judul = "Retention Rate Same Group NCM"
620                   !model = Nul(rsModel.Fields(0).value)
621                   .Update
622               End With
                   
623               With rsLaporan
624                   .AddNew
625                   If Not rsSBETotal.EOF Then
626                       rsSBETotal.filter = ""
627                       nilaiTemp = 0
628                       rsSBETotal.filter = "fs_rc = 'R001' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
629                       If Not rsSBETotal.EOF Then
630                           Do While Not rsSBETotal.EOF
631                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
632                               rsSBETotal.MoveNext
633                           Loop
634                       End If
635                       !K1 = nilaiTemp
636                       rsSBETotal.filter = ""
637                       nilaiTemp = 0
                          
638                       rsSBETotal.filter = "fs_rc = 'R010' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
639                       If Not rsSBETotal.EOF Then
640                           Do While Not rsSBETotal.EOF
641                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
642                               rsSBETotal.MoveNext
643                           Loop
644                       End If
645                       !K10 = nilaiTemp
646                       rsSBETotal.filter = ""
647                       nilaiTemp = 0
                          
648                       rsSBETotal.filter = "fs_rc = 'R020' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
649                       If Not rsSBETotal.EOF Then
650                           Do While Not rsSBETotal.EOF
651                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
652                               rsSBETotal.MoveNext
653                           Loop
654                       End If
655                       !K20 = nilaiTemp
656                       rsSBETotal.filter = ""
657                       nilaiTemp = 0
                          
658                       rsSBETotal.filter = "fs_rc = 'R030' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
659                       If Not rsSBETotal.EOF Then
660                           Do While Not rsSBETotal.EOF
661                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
662                               rsSBETotal.MoveNext
663                           Loop
664                       End If
665                       !K30 = nilaiTemp
666                       rsSBETotal.filter = ""
667                       nilaiTemp = 0
                          
668                       rsSBETotal.filter = "fs_rc = 'R040' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
669                       If Not rsSBETotal.EOF Then
670                           Do While Not rsSBETotal.EOF
671                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
672                               rsSBETotal.MoveNext
673                           Loop
674                       End If
675                       !K40 = nilaiTemp
676                       rsSBETotal.filter = ""
677                       nilaiTemp = 0
                          
678                       rsSBETotal.filter = "fs_rc = 'R050' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
679                       If Not rsSBETotal.EOF Then
680                           Do While Not rsSBETotal.EOF
681                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
682                               rsSBETotal.MoveNext
683                           Loop
684                       End If
685                       !K50 = nilaiTemp
686                       rsSBETotal.filter = ""
687                       nilaiTemp = 0
                          
688                       rsSBETotal.filter = "fs_rc = 'R060' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
689                       If Not rsSBETotal.EOF Then
690                           Do While Not rsSBETotal.EOF
691                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
692                               rsSBETotal.MoveNext
693                           Loop
694                       End If
695                       !K60 = nilaiTemp
696                       rsSBETotal.filter = ""
697                       nilaiTemp = 0
                          
698                       rsSBETotal.filter = "fs_rc = 'R070' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
699                       If Not rsSBETotal.EOF Then
700                           Do While Not rsSBETotal.EOF
701                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
702                               rsSBETotal.MoveNext
703                           Loop
704                       End If
705                       !K70 = nilaiTemp
706                       rsSBETotal.filter = ""
707                       nilaiTemp = 0
                          
708                       rsSBETotal.filter = "fs_rc = 'R080' and FS_NM_DEALER = '" & Nul(rsModel!fs_nm_dealer) & "'"
709                       If Not rsSBETotal.EOF Then
710                           Do While Not rsSBETotal.EOF
711                               nilaiTemp = nilaiTemp + CDbl(rsSBETotal!total)
712                               rsSBETotal.MoveNext
713                           Loop
714                       End If
715                       !K80 = nilaiTemp
716                       rsSBETotal.filter = ""
717                       nilaiTemp = 0
                          
718                       !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
719                       !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                          
720                       !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
721                       !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
722                   Else
723                       !K1 = 0
724                       !K10 = 0
725                       !K20 = 0
726                       !K30 = 0
727                       !K40 = 0
728                       !K50 = 0
729                       !K60 = 0
730                       !K70 = 0
731                       !K80 = 0
732                       !tot16 = 0
733                       !tot19 = 0
734                       !tot29 = 0
735                       !tot26 = 0
736                   End If
                      
737                   If NulSID(rsModel!pareto) = 1 Then
738                       nilaiTotalPareto(0) = nilaiTotalPareto(0) + !K1
739                       nilaiTotalPareto(1) = nilaiTotalPareto(1) + !K10
740                       nilaiTotalPareto(2) = nilaiTotalPareto(2) + !K20
741                       nilaiTotalPareto(3) = nilaiTotalPareto(3) + !K30
742                       nilaiTotalPareto(4) = nilaiTotalPareto(4) + !K40
743                       nilaiTotalPareto(5) = nilaiTotalPareto(5) + !K50
744                       nilaiTotalPareto(6) = nilaiTotalPareto(6) + !K60
745                       nilaiTotalPareto(7) = nilaiTotalPareto(7) + !K70
746                       nilaiTotalPareto(8) = nilaiTotalPareto(8) + !K80
747                   Else
748                       nilaiTotalNonPareto(0) = nilaiTotalNonPareto(0) + !K1
749                       nilaiTotalNonPareto(1) = nilaiTotalNonPareto(1) + !K10
750                       nilaiTotalNonPareto(2) = nilaiTotalNonPareto(2) + !K20
751                       nilaiTotalNonPareto(3) = nilaiTotalNonPareto(3) + !K30
752                       nilaiTotalNonPareto(4) = nilaiTotalNonPareto(4) + !K40
753                       nilaiTotalNonPareto(5) = nilaiTotalNonPareto(5) + !K50
754                       nilaiTotalNonPareto(6) = nilaiTotalNonPareto(6) + !K60
755                       nilaiTotalNonPareto(7) = nilaiTotalNonPareto(7) + !K70
756                       nilaiTotalNonPareto(8) = nilaiTotalNonPareto(8) + !K80
757                   End If
                      
758                   If NulSID(rsModel!north) = 1 Then
759                       !north = 1
760                       nilaiTotalNorth(0) = nilaiTotalNorth(0) + !K1
761                       nilaiTotalNorth(1) = nilaiTotalNorth(1) + !K10
762                       nilaiTotalNorth(2) = nilaiTotalNorth(2) + !K20
763                       nilaiTotalNorth(3) = nilaiTotalNorth(3) + !K30
764                       nilaiTotalNorth(4) = nilaiTotalNorth(4) + !K40
765                       nilaiTotalNorth(5) = nilaiTotalNorth(5) + !K50
766                       nilaiTotalNorth(6) = nilaiTotalNorth(6) + !K60
767                       nilaiTotalNorth(7) = nilaiTotalNorth(7) + !K70
768                       nilaiTotalNorth(8) = nilaiTotalNorth(8) + !K80
769                   Else
                      
770                       !north = 3
771                       nilaiTotalSouth(0) = nilaiTotalSouth(0) + !K1
772                       nilaiTotalSouth(1) = nilaiTotalSouth(1) + !K10
773                       nilaiTotalSouth(2) = nilaiTotalSouth(2) + !K20
774                       nilaiTotalSouth(3) = nilaiTotalSouth(3) + !K30
775                       nilaiTotalSouth(4) = nilaiTotalSouth(4) + !K40
776                       nilaiTotalSouth(5) = nilaiTotalSouth(5) + !K50
777                       nilaiTotalSouth(6) = nilaiTotalSouth(6) + !K60
778                       nilaiTotalSouth(7) = nilaiTotalSouth(7) + !K70
779                       nilaiTotalSouth(8) = nilaiTotalSouth(8) + !K80
780                   End If
781                   !judul = "SB Total (NCM & Non NCM)"
782                   !model = Nul(rsModel.Fields(0).value)
783                   !Urut = 3
784                   .Update
785               End With
786               rsLaporan.filter = ""
                  
                  
787               For i = 0 To 10
788                   nilaiTempTotal(i) = 0
789               Next
790               rsLaporan.filter = ""
                  
791               If Not rsLaporan.EOF Then
792                   rsLaporan.filter = "judul = 'SB Same Outlet' and model = '" & Nul(rsModel.Fields(0).value) & "'"
793                   If Not rsLaporan.EOF Then
794                       If rsLaporan!K1 > 0 Then
795                           nilaiTempTotal(0) = rsLaporan!K1
796                       End If
797                       If rsLaporan!K10 > 0 Then
798                           nilaiTempTotal(1) = rsLaporan!K10
799                       End If
800                       If rsLaporan!K20 > 0 Then
801                           nilaiTempTotal(2) = rsLaporan!K20
802                       End If
803                       If rsLaporan!K30 > 0 Then
804                           nilaiTempTotal(3) = rsLaporan!K30
805                       End If
806                       If rsLaporan!K40 > 0 Then
807                           nilaiTempTotal(4) = rsLaporan!K40
808                       End If
809                       If rsLaporan!K50 > 0 Then
810                           nilaiTempTotal(5) = rsLaporan!K50
811                       End If
812                       If rsLaporan!K60 > 0 Then
813                           nilaiTempTotal(6) = rsLaporan!K60
814                       End If
815                       If rsLaporan!K70 > 0 Then
816                           nilaiTempTotal(7) = rsLaporan!K70
817                       End If
818                       If rsLaporan!K80 > 0 Then
819                           nilaiTempTotal(8) = rsLaporan!K80
820                       End If
821                   End If
822               End If
                  
823               With rsLaporan
824                   .AddNew
825                   If nilaiTempDemand(0) > 0 Then
826                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
827                   Else
828                       !K1 = 0 & " %"
829                   End If
830                   If nilaiTempDemand(1) > 0 Then
831                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
832                   Else
833                       !K10 = 0 & " %"
834                   End If
835                   If nilaiTempDemand(2) > 0 Then
836                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
837                   Else
838                       !K20 = 0 & " %"
839                   End If
840                   If nilaiTempDemand(3) > 0 Then
841                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
842                   Else
843                       !K30 = 0 & " %"
844                   End If
845                   If nilaiTempDemand(4) > 0 Then
846                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
847                   Else
848                       !K40 = 0 & " %"
849                   End If
850                   If nilaiTempDemand(5) > 0 Then
851                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
852                   Else
853                       !K50 = 0 & " %"
854                   End If
855                   If nilaiTempDemand(6) > 0 Then
856                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
857                   Else
858                       !K60 = 0 & " %"
859                   End If
860                   If nilaiTempDemand(7) > 0 Then
861                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
862                   Else
863                       !K70 = 0 & " %"
864                   End If
865                   If nilaiTempDemand(8) > 0 Then
866                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
867                   Else
868                       !K80 = 0 & " %"
869                   End If
                      
                      
870                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
871                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
872                   Else
873                       !tot16 = 0 & " %"
874                   End If
                      
875                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
876                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
877                   Else
878                       !tot26 = 0 & " %"
879                   End If
                       
880                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
881                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
882                   Else
883                       !tot19 = 0 & " %"
884                   End If
                       
885                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
886                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
887                   Else
888                       !tot29 = 0 & " %"
889                   End If
                      
890                   If NulSID(rsModel!north) = 1 Then
891                       !north = 1
892                   Else
893                       !north = 3
894                   End If
895                   !Urut = 5
896                   !judul = "Retention Rate Same Outlet"
897                   !model = Nul(rsModel.Fields(0).value)
898                   .Update
899               End With
                  
900               For i = 0 To 10
901                   nilaiTempTotal(i) = 0
902               Next
903               rsLaporan.filter = ""
                  
904               If Not rsLaporan.EOF Then
905                   rsLaporan.filter = "judul = 'SB Total (NCM & Non NCM)' and model = '" & Nul(rsModel.Fields(0).value) & "'"
906                   If Not rsLaporan.EOF Then
907                       If rsLaporan!K1 > 0 Then
908                           nilaiTempTotal(0) = rsLaporan!K1
909                       End If
910                       If rsLaporan!K10 > 0 Then
911                           nilaiTempTotal(1) = rsLaporan!K10
912                       End If
913                       If rsLaporan!K20 > 0 Then
914                           nilaiTempTotal(2) = rsLaporan!K20
915                       End If
916                       If rsLaporan!K30 > 0 Then
917                           nilaiTempTotal(3) = rsLaporan!K30
918                       End If
919                       If rsLaporan!K40 > 0 Then
920                           nilaiTempTotal(4) = rsLaporan!K40
921                       End If
922                       If rsLaporan!K50 > 0 Then
923                           nilaiTempTotal(5) = rsLaporan!K50
924                       End If
925                       If rsLaporan!K60 > 0 Then
926                           nilaiTempTotal(6) = rsLaporan!K60
927                       End If
928                       If rsLaporan!K70 > 0 Then
929                           nilaiTempTotal(7) = rsLaporan!K70
930                       End If
931                       If rsLaporan!K80 > 0 Then
932                           nilaiTempTotal(8) = rsLaporan!K80
933                       End If
934                   End If
935               End If
                  
936               With rsLaporan
937                   .AddNew
938                   If nilaiTempDemand(0) > 0 Then
939                       !K1 = Round((nilaiTempTotal(0) / nilaiTempDemand(0)) * 100, 0) & " %"
940                   Else
941                       !K1 = 0 & " %"
942                   End If
943                   If nilaiTempDemand(1) > 0 Then
944                       !K10 = Round((nilaiTempTotal(1) / nilaiTempDemand(1)) * 100, 0) & " %"
945                   Else
946                       !K10 = 0 & " %"
947                   End If
948                   If nilaiTempDemand(2) > 0 Then
949                       !K20 = Round((nilaiTempTotal(2) / nilaiTempDemand(2)) * 100, 0) & " %"
950                   Else
951                       !K20 = 0 & " %"
952                   End If
953                   If nilaiTempDemand(3) > 0 Then
954                       !K30 = Round((nilaiTempTotal(3) / nilaiTempDemand(3)) * 100, 0) & " %"
955                   Else
956                       !K30 = 0 & " %"
957                   End If
958                   If nilaiTempDemand(4) > 0 Then
959                       !K40 = Round((nilaiTempTotal(4) / nilaiTempDemand(4)) * 100, 0) & " %"
960                   Else
961                       !K40 = 0 & " %"
962                   End If
963                   If nilaiTempDemand(5) > 0 Then
964                       !K50 = Round((nilaiTempTotal(5) / nilaiTempDemand(5)) * 100, 0) & " %"
965                   Else
966                       !K50 = 0 & " %"
967                   End If
968                   If nilaiTempDemand(6) > 0 Then
969                       !K60 = Round((nilaiTempTotal(6) / nilaiTempDemand(6)) * 100, 0) & " %"
970                   Else
971                       !K60 = 0 & " %"
972                   End If
973                   If nilaiTempDemand(7) > 0 Then
974                       !K70 = Round((nilaiTempTotal(7) / nilaiTempDemand(7)) * 100, 0) & " %"
975                   Else
976                       !K70 = 0 & " %"
977                   End If
978                   If nilaiTempDemand(8) > 0 Then
979                       !K80 = Round((nilaiTempTotal(8) / nilaiTempDemand(8)) * 100, 0) & " %"
980                   Else
981                       !K80 = 0 & " %"
982                   End If
                      
                      
983                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
984                       !tot16 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
985                   Else
986                       !tot16 = 0 & " %"
987                   End If
                      
988                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) > 0 Then
989                       !tot26 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5))) * 100, 0) & " %"
990                   Else
991                       !tot26 = 0 & " %"
992                   End If
                       
993                   If nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
994                       !tot19 = Round(((nilaiTempTotal(0) + nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(0) + nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
995                   Else
996                       !tot19 = 0 & " %"
997                   End If
                       
998                   If nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8) > 0 Then
999                       !tot29 = Round(((nilaiTempTotal(1) + nilaiTempTotal(2) + nilaiTempTotal(3) + nilaiTempTotal(4) + nilaiTempTotal(5) + nilaiTempTotal(6) + nilaiTempTotal(7) + nilaiTempTotal(8)) / (nilaiTempDemand(1) + nilaiTempDemand(2) + nilaiTempDemand(3) + nilaiTempDemand(4) + nilaiTempDemand(5) + nilaiTempDemand(6) + nilaiTempDemand(7) + nilaiTempDemand(8))) * 100, 0) & " %"
1000                  Else
1001                      !tot29 = 0 & " %"
1002                  End If
                      
1003                  If NulSID(rsModel!north) = 1 Then
1004                      !north = 1
1005                  Else
1006                      !north = 3
1007                  End If
1008                  !Urut = 7
1009                  !judul = "Retention Rate Total (NCM & Non NCM)"
1010                  !model = Nul(rsModel.Fields(0).value)
1011                  .Update
1012              End With
1013              rsLaporan.filter = ""
1014              rsModel.MoveNext
1015              progbarDealer.value = progbarDealer.value + 1
1016              DoEvents
1017          Loop
1018          Call setNilaiKategoriDealer
              
1019      Else
1020          MsgBox "Tidak dapat memuat data model kendaraan.", vbInformation
1021          Exit Sub
              
1022      End If
1023      rsLaporan.Sort = "north asc"
1024      Report.database.SetDataSource rsLaporan
1025      progbarDealer.Visible = False
1026      Report.txtHeaderLaporan.SetText "Laporan Retention Nasmoco Group"
1027      Report.txtPeriode.SetText "Periode : " & Format(tglDealerAwal, "dd-mm-yyyy") & " s/d " & Format(tglDealerAkhir, "dd-mm-yyyy")
          
1028      With CRViewer1
1029          .ReportSource = Report
1030          .ViewReport
1031          .Zoom 100
1032          .Visible = True
1033          .Left = 300
1034          .Height = Me.Height - 1200
1035          .Width = Me.Width - 750
1036      End With
         
1037      cmdTutup.Visible = True
1038      cmdTutup.Left = CRViewer1.Left + 20
1039      cmdTutup.Top = CRViewer1.Top + 400
1040      Screen.MousePointer = vbDefault
          
1041      Call DisableExport
1042      Call setLokasiTombol
End Sub

Private Sub setNilaiKategoriDealer(Optional param As String)
          '' ----------------------------------------------
          '' urut 1
          '' ----------------------------------------------
1         With rsLaporan
2             .AddNew
3             !K1 = nilaiSBSameOutletNorth(0)
4             !K10 = nilaiSBSameOutletNorth(1)
5             !K20 = nilaiSBSameOutletNorth(2)
6             !K30 = nilaiSBSameOutletNorth(3)
7             !K40 = nilaiSBSameOutletNorth(4)
8             !K50 = nilaiSBSameOutletNorth(5)
9             !K60 = nilaiSBSameOutletNorth(6)
10            !K70 = nilaiSBSameOutletNorth(7)
11            !K80 = nilaiSBSameOutletNorth(8)
12            !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
13            !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
14            !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
15            !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
16            !judul = "SB Same Outlet"
17            !north = 2
18            !model = "NORTH AREA"
19            !Urut = 1
20            .Update
21        End With
          
22        With rsLaporan
23            .AddNew
24            !K1 = nilaiSBSameOutletPareto(0)
25            !K10 = nilaiSBSameOutletPareto(1)
26            !K20 = nilaiSBSameOutletPareto(2)
27            !K30 = nilaiSBSameOutletPareto(3)
28            !K40 = nilaiSBSameOutletPareto(4)
29            !K50 = nilaiSBSameOutletPareto(5)
30            !K60 = nilaiSBSameOutletPareto(6)
31            !K70 = nilaiSBSameOutletPareto(7)
32            !K80 = nilaiSBSameOutletPareto(8)
33            !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
34            !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
35            !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
36            !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
37            !judul = "SB Same Outlet"
38            !north = 6
39            !model = "PARETO"
40            !Urut = 1
41            .Update
42        End With
          
43        With rsLaporan
44            .AddNew
45            !K1 = nilaiSBSameOutletNonPareto(0)
46            !K10 = nilaiSBSameOutletNonPareto(1)
47            !K20 = nilaiSBSameOutletNonPareto(2)
48            !K30 = nilaiSBSameOutletNonPareto(3)
49            !K40 = nilaiSBSameOutletNonPareto(4)
50            !K50 = nilaiSBSameOutletNonPareto(5)
51            !K60 = nilaiSBSameOutletNonPareto(6)
52            !K70 = nilaiSBSameOutletNonPareto(7)
53            !K80 = nilaiSBSameOutletNonPareto(8)
54            !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
55            !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
56            !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
57            !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
58            !judul = "SB Same Outlet"
59            !north = 7
60            !model = "NON PARETO"
61            !Urut = 1
62            .Update
63        End With
              
64        With rsLaporan
65            .AddNew
              
66            !K1 = nilaiSBSameOutletSouth(0)
67            !K10 = nilaiSBSameOutletSouth(1)
68            !K20 = nilaiSBSameOutletSouth(2)
69            !K30 = nilaiSBSameOutletSouth(3)
70            !K40 = nilaiSBSameOutletSouth(4)
71            !K50 = nilaiSBSameOutletSouth(5)
72            !K60 = nilaiSBSameOutletSouth(6)
73            !K70 = nilaiSBSameOutletSouth(7)
74            !K80 = nilaiSBSameOutletSouth(8)
75            !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
76            !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
77            !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
78            !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
79            !judul = "SB Same Outlet"
80            !model = "SOUTH AREA"
81            !north = 4
82            !Urut = 1
83            .Update
84        End With
          
85        With rsLaporan
86            .AddNew
              
87            !K1 = nilaiSBSameOutletSouth(0) + nilaiSBSameOutletNorth(0)
88            !K10 = nilaiSBSameOutletSouth(1) + nilaiSBSameOutletNorth(1)
89            !K20 = nilaiSBSameOutletSouth(2) + nilaiSBSameOutletNorth(2)
90            !K30 = nilaiSBSameOutletSouth(3) + nilaiSBSameOutletNorth(3)
91            !K40 = nilaiSBSameOutletSouth(4) + nilaiSBSameOutletNorth(4)
92            !K50 = nilaiSBSameOutletSouth(5) + nilaiSBSameOutletNorth(5)
93            !K60 = nilaiSBSameOutletSouth(6) + nilaiSBSameOutletNorth(6)
94            !K70 = nilaiSBSameOutletSouth(7) + nilaiSBSameOutletNorth(7)
95            !K80 = nilaiSBSameOutletSouth(8) + nilaiSBSameOutletNorth(8)
96            !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
97            !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
98            !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
99            !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
100           !judul = "SB Same Outlet"
101           !model = "NRM"
102           !north = 5
103           !Urut = 1
104           .Update
105       End With
          '' akhir urut 1 ------------------------------
          
          '' ----------------------------------------------
          '' urut 2
          '' ----------------------------------------------
          
106       With rsLaporan
107           .AddNew
              
108           !K1 = nilaiSBNRMNorth(0)
109           !K10 = nilaiSBNRMNorth(1)
110           !K20 = nilaiSBNRMNorth(2)
111           !K30 = nilaiSBNRMNorth(3)
112           !K40 = nilaiSBNRMNorth(4)
113           !K50 = nilaiSBNRMNorth(5)
114           !K60 = nilaiSBNRMNorth(6)
115           !K70 = nilaiSBNRMNorth(7)
116           !K80 = nilaiSBNRMNorth(8)
117           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
118           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
119           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
120           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
121           !judul = "SB Same Group (NCM)"
122           !north = 2
123           !model = "NORTH AREA"
124           !Urut = 2
125           .Update
126       End With
                          
127       With rsLaporan
128           .AddNew
              
129           !K1 = nilaiSBNRMSouth(0)
130           !K10 = nilaiSBNRMSouth(1)
131           !K20 = nilaiSBNRMSouth(2)
132           !K30 = nilaiSBNRMSouth(3)
133           !K40 = nilaiSBNRMSouth(4)
134           !K50 = nilaiSBNRMSouth(5)
135           !K60 = nilaiSBNRMSouth(6)
136           !K70 = nilaiSBNRMSouth(7)
137           !K80 = nilaiSBNRMSouth(8)
138           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
139           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
140           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
141           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
142           !judul = "SB Same Group (NCM)"
143           !north = 4
144           !model = "SOUTH AREA"
145           !Urut = 2
146           .Update
147       End With
              
148       With rsLaporan
149           .AddNew
              
150           !K1 = nilaiSBNRMSouth(0) + nilaiSBNRMNorth(0)
151           !K10 = nilaiSBNRMSouth(1) + nilaiSBNRMNorth(1)
152           !K20 = nilaiSBNRMSouth(2) + nilaiSBNRMNorth(2)
153           !K30 = nilaiSBNRMSouth(3) + nilaiSBNRMNorth(3)
154           !K40 = nilaiSBNRMSouth(4) + nilaiSBNRMNorth(4)
155           !K50 = nilaiSBNRMSouth(5) + nilaiSBNRMNorth(5)
156           !K60 = nilaiSBNRMSouth(6) + nilaiSBNRMNorth(6)
157           !K70 = nilaiSBNRMSouth(7) + nilaiSBNRMNorth(7)
158           !K80 = nilaiSBNRMSouth(8) + nilaiSBNRMNorth(8)
159           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
160           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
161           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
162           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
163           !judul = "SB Same Group (NCM)"
164           !model = "NRM"
165           !north = 5
166           !Urut = 2
167           .Update
168       End With
          
          
169       With rsLaporan
170           .AddNew
171           !K1 = nilaiSBNRMPareto(0)
172           !K10 = nilaiSBNRMPareto(1)
173           !K20 = nilaiSBNRMPareto(2)
174           !K30 = nilaiSBNRMPareto(3)
175           !K40 = nilaiSBNRMPareto(4)
176           !K50 = nilaiSBNRMPareto(5)
177           !K60 = nilaiSBNRMPareto(6)
178           !K70 = nilaiSBNRMPareto(7)
179           !K80 = nilaiSBNRMPareto(8)
180           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
181           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
182           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
183           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
184           !judul = "SB Same Group (NCM)"
185           !north = 6
186           !model = "PARETO"
187           !Urut = 2
188           .Update
189       End With
          
190       With rsLaporan
191           .AddNew
192           !K1 = nilaiSBNRMNonPareto(0)
193           !K10 = nilaiSBNRMNonPareto(1)
194           !K20 = nilaiSBNRMNonPareto(2)
195           !K30 = nilaiSBNRMNonPareto(3)
196           !K40 = nilaiSBNRMNonPareto(4)
197           !K50 = nilaiSBNRMNonPareto(5)
198           !K60 = nilaiSBNRMNonPareto(6)
199           !K70 = nilaiSBNRMNonPareto(7)
200           !K80 = nilaiSBNRMNonPareto(8)
201           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
202           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
203           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
204           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
205           !judul = "SB Same Group (NCM)"
206           !north = 7
207           !model = "NON PARETO"
208           !Urut = 2
209           .Update
210       End With
          
          '' --- akhir urut 2 -----------------------
          
          
          '' ----------------------------------------------
          '' urut 3
          '' ----------------------------------------------
          
211       With rsLaporan
212           .AddNew
              
213           !K1 = nilaiTotalNorth(0)
214           !K10 = nilaiTotalNorth(1)
215           !K20 = nilaiTotalNorth(2)
216           !K30 = nilaiTotalNorth(3)
217           !K40 = nilaiTotalNorth(4)
218           !K50 = nilaiTotalNorth(5)
219           !K60 = nilaiTotalNorth(6)
220           !K70 = nilaiTotalNorth(7)
221           !K80 = nilaiTotalNorth(8)
222           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
223           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
224           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
225           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
226           !judul = "SB Total (NCM & Non NCM)"
227           !north = 2
228           !model = "NORTH AREA"
229           !Urut = 3
230           .Update
231       End With
                          
232       With rsLaporan
233           .AddNew
              
234           !K1 = nilaiTotalSouth(0)
235           !K10 = nilaiTotalSouth(1)
236           !K20 = nilaiTotalSouth(2)
237           !K30 = nilaiTotalSouth(3)
238           !K40 = nilaiTotalSouth(4)
239           !K50 = nilaiTotalSouth(5)
240           !K60 = nilaiTotalSouth(6)
241           !K70 = nilaiTotalSouth(7)
242           !K80 = nilaiTotalSouth(8)
243           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
244           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
245           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
246           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
247           !judul = "SB Total (NCM & Non NCM)"
248           !north = 4
249           !model = "SOUTH AREA"
250           !Urut = 3
251           .Update
252       End With
              
253       With rsLaporan
254           .AddNew
              
255           !K1 = nilaiTotalSouth(0) + nilaiTotalNorth(0)
256           !K10 = nilaiTotalSouth(1) + nilaiTotalNorth(1)
257           !K20 = nilaiTotalSouth(2) + nilaiTotalNorth(2)
258           !K30 = nilaiTotalSouth(3) + nilaiTotalNorth(3)
259           !K40 = nilaiTotalSouth(4) + nilaiTotalNorth(4)
260           !K50 = nilaiTotalSouth(5) + nilaiTotalNorth(5)
261           !K60 = nilaiTotalSouth(6) + nilaiTotalNorth(6)
262           !K70 = nilaiTotalSouth(7) + nilaiTotalNorth(7)
263           !K80 = nilaiTotalSouth(8) + nilaiTotalNorth(8)
264           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
265           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
266           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
267           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
268           !judul = "SB Total (NCM & Non NCM)"
269           !model = "NRM"
270           !north = 5
271           !Urut = 3
272           .Update
273       End With
          
          
274       With rsLaporan
275           .AddNew
276           !K1 = nilaiTotalPareto(0)
277           !K10 = nilaiTotalPareto(1)
278           !K20 = nilaiTotalPareto(2)
279           !K30 = nilaiTotalPareto(3)
280           !K40 = nilaiTotalPareto(4)
281           !K50 = nilaiTotalPareto(5)
282           !K60 = nilaiTotalPareto(6)
283           !K70 = nilaiTotalPareto(7)
284           !K80 = nilaiTotalPareto(8)
285           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
286           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
287           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
288           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
289           !judul = "SB Total (NCM & Non NCM)"
290           !north = 6
291           !model = "PARETO"
292           !Urut = 3
293           .Update
294       End With
          
295       With rsLaporan
296           .AddNew
297           !K1 = nilaiTotalNonPareto(0)
298           !K10 = nilaiTotalNonPareto(1)
299           !K20 = nilaiTotalNonPareto(2)
300           !K30 = nilaiTotalNonPareto(3)
301           !K40 = nilaiTotalNonPareto(4)
302           !K50 = nilaiTotalNonPareto(5)
303           !K60 = nilaiTotalNonPareto(6)
304           !K70 = nilaiTotalNonPareto(7)
305           !K80 = nilaiTotalNonPareto(8)
306           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
307           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
308           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
309           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
310           !judul = "SB Total (NCM & Non NCM)"
311           !north = 7
312           !model = "NON PARETO"
313           !Urut = 3
314           .Update
315       End With
          
          '' --- akhir urut 3 -----------------------
          
          
          '' ----------------------------------------------
          '' urut 4
          '' ----------------------------------------------
          
316       With rsLaporan
317           .AddNew
              
318           !K1 = nilaiDemandNorth(0)
319           !K10 = nilaiDemandNorth(1)
320           !K20 = nilaiDemandNorth(2)
321           !K30 = nilaiDemandNorth(3)
322           !K40 = nilaiDemandNorth(4)
323           !K50 = nilaiDemandNorth(5)
324           !K60 = nilaiDemandNorth(6)
325           !K70 = nilaiDemandNorth(7)
326           !K80 = nilaiDemandNorth(8)
327           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
328           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
329           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
330           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
331           !judul = "Demand Same Outlet"
332           !north = 2
333           !model = "NORTH AREA"
334           !Urut = 4
335           .Update
336       End With
                          
337       With rsLaporan
338           .AddNew
              
339           !K1 = nilaiDemandSouth(0)
340           !K10 = nilaiDemandSouth(1)
341           !K20 = nilaiDemandSouth(2)
342           !K30 = nilaiDemandSouth(3)
343           !K40 = nilaiDemandSouth(4)
344           !K50 = nilaiDemandSouth(5)
345           !K60 = nilaiDemandSouth(6)
346           !K70 = nilaiDemandSouth(7)
347           !K80 = nilaiDemandSouth(8)
348           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
349           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
350           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
351           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
352           !judul = "Demand Same Outlet"
353           !north = 4
354           !model = "SOUTH AREA"
355           !Urut = 4
356           .Update
357       End With
              
358       With rsLaporan
359           .AddNew
              
360           !K1 = nilaiDemandSouth(0) + nilaiDemandNorth(0)
361           !K10 = nilaiDemandSouth(1) + nilaiDemandNorth(1)
362           !K20 = nilaiDemandSouth(2) + nilaiDemandNorth(2)
363           !K30 = nilaiDemandSouth(3) + nilaiDemandNorth(3)
364           !K40 = nilaiDemandSouth(4) + nilaiDemandNorth(4)
365           !K50 = nilaiDemandSouth(5) + nilaiDemandNorth(5)
366           !K60 = nilaiDemandSouth(6) + nilaiDemandNorth(6)
367           !K70 = nilaiDemandSouth(7) + nilaiDemandNorth(7)
368           !K80 = nilaiDemandSouth(8) + nilaiDemandNorth(8)
369           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
370           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
371           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
372           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
373           !judul = "Demand Same Outlet"
374           !model = "NRM"
375           !north = 5
376           !Urut = 4
377           .Update
378       End With
          
          
379       With rsLaporan
380           .AddNew
381           !K1 = nilaiDemandPareto(0)
382           !K10 = nilaiDemandPareto(1)
383           !K20 = nilaiDemandPareto(2)
384           !K30 = nilaiDemandPareto(3)
385           !K40 = nilaiDemandPareto(4)
386           !K50 = nilaiDemandPareto(5)
387           !K60 = nilaiDemandPareto(6)
388           !K70 = nilaiDemandPareto(7)
389           !K80 = nilaiDemandPareto(8)
390           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
391           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
392           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
393           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
394           !judul = "Demand Same Outlet"
395           !north = 6
396           !model = "PARETO"
397           !Urut = 4
398           .Update
399       End With
          
400       With rsLaporan
401           .AddNew
402           !K1 = nilaiDemandNonPareto(0)
403           !K10 = nilaiDemandNonPareto(1)
404           !K20 = nilaiDemandNonPareto(2)
405           !K30 = nilaiDemandNonPareto(3)
406           !K40 = nilaiDemandNonPareto(4)
407           !K50 = nilaiDemandNonPareto(5)
408           !K60 = nilaiDemandNonPareto(6)
409           !K70 = nilaiDemandNonPareto(7)
410           !K80 = nilaiDemandNonPareto(8)
411           !tot16 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
412           !tot19 = CDbl(!K1) + CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
              
413           !tot26 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50)
414           !tot29 = CDbl(!K10) + CDbl(!K20) + CDbl(!K30) + CDbl(!K40) + CDbl(!K50) + CDbl(!K60) + CDbl(!K70) + CDbl(!K80)
                       
415           !judul = "Demand Same Outlet"
416           !north = 7
417           !model = "NON PARETO"
418           !Urut = 4
419           .Update
420       End With
          
          '' --- akhir urut 4 -----------------------
          
421       Call setNilaiRateDealer
End Sub

Private Sub setNilaiRateDealer()

          
          '' -----------------------------------------
          '' urut 5
          '' -----------------------------------------
          
1         With rsLaporan
2             .AddNew
3             If nilaiDemandNorth(0) > 0 Then
4                 !K1 = Round((nilaiSBSameOutletNorth(0) / nilaiDemandNorth(0)) * 100, 0) & " %"
5             Else
6                 !K1 = 0 & " %"
7             End If
8             If nilaiDemandNorth(1) > 0 Then
9                 !K10 = Round((nilaiSBSameOutletNorth(1) / nilaiDemandNorth(1)) * 100, 0) & " %"
10            Else
11                !K10 = 0 & " %"
12            End If
13            If nilaiDemandNorth(2) > 0 Then
14                !K20 = Round((nilaiSBSameOutletNorth(2) / nilaiDemandNorth(2)) * 100, 0) & " %"
15            Else
16                !K20 = 0 & " %"
17            End If
18            If nilaiDemandNorth(3) > 0 Then
19                !K30 = Round((nilaiSBSameOutletNorth(3) / nilaiDemandNorth(3)) * 100, 0) & " %"
20            Else
21                !K30 = 0 & " %"
22            End If
23            If nilaiDemandNorth(4) > 0 Then
24                !K40 = Round((nilaiSBSameOutletNorth(4) / nilaiDemandNorth(4)) * 100, 0) & " %"
25            Else
26                !K40 = 0 & " %"
27            End If
28            If nilaiDemandNorth(5) > 0 Then
29                !K50 = Round((nilaiSBSameOutletNorth(5) / nilaiDemandNorth(5)) * 100, 0) & " %"
30            Else
31                !K50 = 0 & " %"
32            End If
33            If nilaiDemandNorth(6) > 0 Then
34                !K60 = Round((nilaiSBSameOutletNorth(6) / nilaiDemandNorth(6)) * 100, 0) & " %"
35            Else
36                !K60 = 0 & " %"
37            End If
38            If nilaiDemandNorth(7) > 0 Then
39                !K70 = Round((nilaiSBSameOutletNorth(7) / nilaiDemandNorth(7)) * 100, 0) & " %"
40            Else
41                !K70 = 0 & " %"
42            End If
43            If nilaiDemandNorth(8) > 0 Then
44                !K80 = Round((nilaiSBSameOutletNorth(8) / nilaiDemandNorth(8)) * 100, 0) & " %"
45            Else
46                !K80 = 0 & " %"
47            End If
              
48            If nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) > 0 Then
49                !tot16 = Round(((nilaiSBSameOutletNorth(0) + nilaiSBSameOutletNorth(1) + nilaiSBSameOutletNorth(2) + nilaiSBSameOutletNorth(3) + nilaiSBSameOutletNorth(4) + nilaiSBSameOutletNorth(5)) / (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5))) * 100, 0) & " %"
50            Else
51                !tot16 = 0 & " %"
52            End If
              
53            If nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) > 0 Then
54                !tot26 = Round(((nilaiSBSameOutletNorth(1) + nilaiSBSameOutletNorth(2) + nilaiSBSameOutletNorth(3) + nilaiSBSameOutletNorth(4) + nilaiSBSameOutletNorth(5)) / (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5))) * 100, 0) & " %"
55            Else
56                !tot26 = 0 & " %"
57            End If
               
58            If nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8) > 0 Then
59                !tot19 = Round(((nilaiSBSameOutletNorth(0) + nilaiSBSameOutletNorth(1) + nilaiSBSameOutletNorth(2) + nilaiSBSameOutletNorth(3) + nilaiSBSameOutletNorth(4) + nilaiSBSameOutletNorth(5) + nilaiSBSameOutletNorth(6) + nilaiSBSameOutletNorth(7) + nilaiSBSameOutletNorth(8)) / (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
60            Else
61                !tot19 = 0 & " %"
62            End If
               
63            If nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8) > 0 Then
64                !tot29 = Round(((nilaiSBSameOutletNorth(1) + nilaiSBSameOutletNorth(2) + nilaiSBSameOutletNorth(3) + nilaiSBSameOutletNorth(4) + nilaiSBSameOutletNorth(5) + nilaiSBSameOutletNorth(6) + nilaiSBSameOutletNorth(7) + nilaiSBSameOutletNorth(8)) / (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
65            Else
66                !tot29 = 0 & " %"
67            End If
               
68            !Urut = 5
69            !judul = "Retention Rate Same Outlet"
70            !model = "NORTH AREA"
71            !north = 2
72            .Update
73        End With
        
74        With rsLaporan
75            .AddNew
76            If nilaiDemandSouth(0) > 0 Then
77                !K1 = Round((nilaiSBSameOutletSouth(0) / nilaiDemandSouth(0)) * 100, 0) & " %"
78            Else
79                !K1 = 0 & " %"
80            End If
81            If nilaiDemandSouth(1) > 0 Then
82                !K10 = Round((nilaiSBSameOutletSouth(1) / nilaiDemandSouth(1)) * 100, 0) & " %"
83            Else
84                !K10 = 0 & " %"
85            End If
86            If nilaiDemandSouth(2) > 0 Then
87                !K20 = Round((nilaiSBSameOutletSouth(2) / nilaiDemandSouth(2)) * 100, 0) & " %"
88            Else
89                !K20 = 0 & " %"
90            End If
91            If nilaiDemandSouth(3) > 0 Then
92                !K30 = Round((nilaiSBSameOutletSouth(3) / nilaiDemandSouth(3)) * 100, 0) & " %"
93            Else
94                !K30 = 0 & " %"
95            End If
96            If nilaiDemandSouth(4) > 0 Then
97                !K40 = Round((nilaiSBSameOutletSouth(4) / nilaiDemandSouth(4)) * 100, 0) & " %"
98            Else
99                !K40 = 0 & " %"
100           End If
101           If nilaiDemandSouth(5) > 0 Then
102               !K50 = Round((nilaiSBSameOutletSouth(5) / nilaiDemandSouth(5)) * 100, 0) & " %"
103           Else
104               !K50 = 0 & " %"
105           End If
106           If nilaiDemandSouth(6) > 0 Then
107               !K60 = Round((nilaiSBSameOutletSouth(6) / nilaiDemandSouth(6)) * 100, 0) & " %"
108           Else
109               !K60 = 0 & " %"
110           End If
111           If nilaiDemandSouth(7) > 0 Then
112               !K70 = Round((nilaiSBSameOutletSouth(7) / nilaiDemandSouth(7)) * 100, 0) & " %"
113           Else
114               !K70 = 0 & " %"
115           End If
116           If nilaiDemandSouth(8) > 0 Then
117               !K80 = Round((nilaiSBSameOutletSouth(8) / nilaiDemandSouth(8)) * 100, 0) & " %"
118           Else
119               !K80 = 0 & " %"
120           End If
              
121           If nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
122               !tot16 = Round(((nilaiSBSameOutletSouth(0) + nilaiSBSameOutletSouth(1) + nilaiSBSameOutletSouth(2) + nilaiSBSameOutletSouth(3) + nilaiSBSameOutletSouth(4) + nilaiSBSameOutletSouth(5)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5))) * 100, 0) & " %"
123           Else
124               !tot16 = 0 & " %"
125           End If
              
126           If nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
127               !tot26 = Round(((nilaiSBSameOutletSouth(1) + nilaiSBSameOutletSouth(2) + nilaiSBSameOutletSouth(3) + nilaiSBSameOutletSouth(4) + nilaiSBSameOutletSouth(5)) / (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5))) * 100, 0) & " %"
128           Else
129               !tot26 = 0 & " %"
130           End If
               
131           If nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
132               !tot19 = Round(((nilaiSBNRMSouth(0) + nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5) + nilaiSBNRMSouth(6) + nilaiSBNRMSouth(7) + nilaiSBNRMSouth(8)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8))) * 100, 0) & " %"
133           Else
134               !tot19 = 0 & " %"
135           End If
               
136           If nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
137               !tot29 = Round(((nilaiSBSameOutletSouth(1) + nilaiSBSameOutletSouth(2) + nilaiSBSameOutletSouth(3) + nilaiSBSameOutletSouth(4) + nilaiSBSameOutletSouth(5) + nilaiSBSameOutletSouth(6) + nilaiSBSameOutletSouth(7) + nilaiSBSameOutletSouth(8)) / (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8))) * 100, 0) & " %"
138           Else
139               !tot29 = 0 & " %"
140           End If
               
141           !Urut = 5
142           !judul = "Retention Rate Same Outlet"
143           !model = "SOUTH AREA"
144           !north = 4
145           .Update
146       End With
              
147       With rsLaporan
148           .AddNew
149           If nilaiDemandPareto(0) > 0 Then
150               !K1 = Round((nilaiSBSameOutletPareto(0) / nilaiDemandPareto(0)) * 100, 0) & " %"
151           Else
152               !K1 = 0 & " %"
153           End If
154           If nilaiDemandPareto(1) > 0 Then
155               !K10 = Round((nilaiSBSameOutletPareto(1) / nilaiDemandPareto(1)) * 100, 0) & " %"
156           Else
157               !K10 = 0 & " %"
158           End If
159           If nilaiDemandPareto(2) > 0 Then
160               !K20 = Round((nilaiSBSameOutletPareto(2) / nilaiDemandPareto(2)) * 100, 0) & " %"
161           Else
162               !K20 = 0 & " %"
163           End If
164           If nilaiDemandPareto(3) > 0 Then
165               !K30 = Round((nilaiSBSameOutletPareto(3) / nilaiDemandPareto(3)) * 100, 0) & " %"
166           Else
167               !K30 = 0 & " %"
168           End If
169           If nilaiDemandPareto(4) > 0 Then
170               !K40 = Round((nilaiSBSameOutletPareto(4) / nilaiDemandPareto(4)) * 100, 0) & " %"
171           Else
172               !K40 = 0 & " %"
173           End If
174           If nilaiDemandPareto(5) > 0 Then
175               !K50 = Round((nilaiSBSameOutletPareto(5) / nilaiDemandPareto(5)) * 100, 0) & " %"
176           Else
177               !K50 = 0 & " %"
178           End If
179           If nilaiDemandPareto(6) > 0 Then
180               !K60 = Round((nilaiSBSameOutletPareto(6) / nilaiDemandPareto(6)) * 100, 0) & " %"
181           Else
182               !K60 = 0 & " %"
183           End If
184           If nilaiDemandPareto(7) > 0 Then
185               !K70 = Round((nilaiSBSameOutletPareto(7) / nilaiDemandPareto(7)) * 100, 0) & " %"
186           Else
187               !K70 = 0 & " %"
188           End If
189           If nilaiDemandPareto(8) > 0 Then
190               !K80 = Round((nilaiSBSameOutletPareto(8) / nilaiDemandPareto(8)) * 100, 0) & " %"
191           Else
192               !K80 = 0 & " %"
193           End If
              
194           If nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) > 0 Then
195               !tot16 = Round(((nilaiSBSameOutletPareto(0) + nilaiSBSameOutletPareto(1) + nilaiSBSameOutletPareto(2) + nilaiSBSameOutletPareto(3) + nilaiSBSameOutletPareto(4) + nilaiSBSameOutletPareto(5)) / (nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5))) * 100, 0) & " %"
196           Else
197               !tot16 = 0 & " %"
198           End If
              
199           If nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) > 0 Then
200               !tot26 = Round(((nilaiSBSameOutletPareto(1) + nilaiSBSameOutletPareto(2) + nilaiSBNRMPareto(3) + nilaiSBNRMPareto(4) + nilaiSBNRMPareto(5)) / (nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5))) * 100, 0) & " %"
201           Else
202               !tot26 = 0 & " %"
203           End If
               
204           If nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8) > 0 Then
205               !tot19 = Round(((nilaiSBSameOutletPareto(0) + nilaiSBSameOutletPareto(1) + nilaiSBSameOutletPareto(2) + nilaiSBSameOutletPareto(3) + nilaiSBSameOutletPareto(4) + nilaiSBSameOutletPareto(5) + nilaiSBSameOutletPareto(6) + nilaiSBSameOutletPareto(7) + nilaiSBSameOutletPareto(8)) / (nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8))) * 100, 0) & " %"
206           Else
207               !tot19 = 0 & " %"
208           End If
               
209           If nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8) > 0 Then
210               !tot29 = Round(((nilaiSBSameOutletPareto(1) + nilaiSBSameOutletPareto(2) + nilaiSBSameOutletPareto(3) + nilaiSBSameOutletPareto(4) + nilaiSBSameOutletPareto(5) + nilaiSBSameOutletPareto(6) + nilaiSBSameOutletPareto(7) + nilaiSBSameOutletPareto(8)) / (nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8))) * 100, 0) & " %"
211           Else
212               !tot29 = 0 & " %"
213           End If
               
214           !Urut = 5
215           !judul = "Retention Rate Same Outlet"
216           !model = "PARETO"
217           !north = 6
218           .Update
219       End With
          
220       With rsLaporan
221           .AddNew
222           If nilaiDemandNonPareto(0) > 0 Then
223               !K1 = Round((nilaiSBSameOutletNonPareto(0) / nilaiDemandNonPareto(0)) * 100, 0) & " %"
224           Else
225               !K1 = 0 & " %"
226           End If
227           If nilaiDemandNonPareto(1) > 0 Then
228               !K10 = Round((nilaiSBSameOutletNonPareto(1) / nilaiDemandNonPareto(1)) * 100, 0) & " %"
229           Else
230               !K10 = 0 & " %"
231           End If
232           If nilaiDemandNonPareto(2) > 0 Then
233               !K20 = Round((nilaiSBSameOutletNonPareto(2) / nilaiDemandNonPareto(2)) * 100, 0) & " %"
234           Else
235               !K20 = 0 & " %"
236           End If
237           If nilaiDemandNonPareto(3) > 0 Then
238               !K30 = Round((nilaiSBSameOutletNonPareto(3) / nilaiDemandNonPareto(3)) * 100, 0) & " %"
239           Else
240               !K30 = 0 & " %"
241           End If
242           If nilaiDemandNonPareto(4) > 0 Then
243               !K40 = Round((nilaiSBSameOutletNonPareto(4) / nilaiDemandNonPareto(4)) * 100, 0) & " %"
244           Else
245               !K40 = 0 & " %"
246           End If
247           If nilaiDemandNonPareto(5) > 0 Then
248               !K50 = Round((nilaiSBSameOutletNonPareto(5) / nilaiDemandNonPareto(5)) * 100, 0) & " %"
249           Else
250               !K50 = 0 & " %"
251           End If
252           If nilaiDemandNonPareto(6) > 0 Then
253               !K60 = Round((nilaiSBSameOutletNonPareto(6) / nilaiDemandNonPareto(6)) * 100, 0) & " %"
254           Else
255               !K60 = 0 & " %"
256           End If
257           If nilaiDemandNonPareto(7) > 0 Then
258               !K70 = Round((nilaiSBSameOutletNonPareto(7) / nilaiDemandNonPareto(7)) * 100, 0) & " %"
259           Else
260               !K70 = 0 & " %"
261           End If
262           If nilaiDemandNonPareto(8) > 0 Then
263               !K80 = Round((nilaiSBSameOutletNonPareto(8) / nilaiDemandNonPareto(8)) * 100, 0) & " %"
264           Else
265               !K80 = 0 & " %"
266           End If
              
267           If nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) > 0 Then
268               !tot16 = Round(((nilaiSBSameOutletNonPareto(0) + nilaiSBSameOutletNonPareto(1) + nilaiSBSameOutletNonPareto(2) + nilaiSBSameOutletNonPareto(3) + nilaiSBSameOutletNonPareto(4) + nilaiSBSameOutletNonPareto(5)) / (nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5))) * 100, 0) & " %"
269           Else
270               !tot16 = 0 & " %"
271           End If
              
272           If nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) > 0 Then
273               !tot26 = Round(((nilaiSBSameOutletNonPareto(1) + nilaiSBSameOutletNonPareto(2) + nilaiSBSameOutletNonPareto(3) + nilaiSBSameOutletNonPareto(4) + nilaiSBSameOutletNonPareto(5)) / (nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5))) * 100, 0) & " %"
274           Else
275               !tot26 = 0 & " %"
276           End If
               
277           If nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8) > 0 Then
278               !tot19 = Round(((nilaiSBSameOutletNonPareto(0) + nilaiSBSameOutletNonPareto(1) + nilaiSBSameOutletNonPareto(2) + nilaiSBSameOutletNonPareto(3) + nilaiSBSameOutletNonPareto(4) + nilaiSBSameOutletNonPareto(5) + nilaiSBSameOutletNonPareto(6) + nilaiSBSameOutletNonPareto(7) + nilaiSBSameOutletNonPareto(8)) / (nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8))) * 100, 0) & " %"
279           Else
280               !tot19 = 0 & " %"
281           End If
               
282           If nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8) > 0 Then
283               !tot29 = Round(((nilaiSBSameOutletNonPareto(1) + nilaiSBSameOutletNonPareto(2) + nilaiSBSameOutletNonPareto(3) + nilaiSBSameOutletNonPareto(4) + nilaiSBSameOutletNonPareto(5) + nilaiSBSameOutletNonPareto(6) + nilaiSBSameOutletNonPareto(7) + nilaiSBSameOutletNonPareto(8)) / (nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8))) * 100, 0) & " %"
284           Else
285               !tot29 = 0 & " %"
286           End If
               
287           !Urut = 5
288           !judul = "Retention Rate Same Outlet"
289           !model = "NON PARETO"
290           !north = 7
291           .Update
292       End With
          
          
293       With rsLaporan
294           .AddNew
295           If nilaiDemandSouth(0) + nilaiDemandNorth(0) > 0 Then
296               !K1 = Round((nilaiSBSameOutletNorth(0) + nilaiSBSameOutletSouth(0) / (nilaiDemandSouth(0) + nilaiDemandNorth(0))) * 100, 0) & " %"
297           Else
298               !K1 = 0 & " %"
299           End If
300           If nilaiDemandSouth(1) + nilaiDemandNorth(1) > 0 Then
301               !K10 = Round((nilaiSBSameOutletNorth(1) + nilaiSBSameOutletSouth(1) / (nilaiDemandSouth(1) + nilaiDemandNorth(1))) * 100, 0) & " %"
302           Else
303               !K10 = 0 & " %"
304           End If
305           If nilaiDemandSouth(2) + nilaiDemandNorth(2) > 0 Then
306               !K20 = Round((nilaiSBSameOutletNorth(2) + nilaiSBSameOutletSouth(2) / (nilaiDemandSouth(2) + nilaiDemandNorth(2))) * 100, 0) & " %"
307           Else
308               !K20 = 0 & " %"
309           End If
310           If nilaiDemandSouth(3) + nilaiDemandNorth(3) > 0 Then
311               !K30 = Round((nilaiSBSameOutletNorth(3) + nilaiSBSameOutletSouth(3) / (nilaiDemandSouth(3) + nilaiDemandNorth(3))) * 100, 0) & " %"
312           Else
313               !K30 = 0 & " %"
314           End If
315           If nilaiDemandSouth(4) + nilaiDemandNorth(4) > 0 Then
316               !K40 = Round((nilaiSBSameOutletNorth(4) + nilaiSBSameOutletSouth(4) / (nilaiDemandSouth(4) + nilaiDemandNorth(4))) * 100, 0) & " %"
317           Else
318               !K40 = 0 & " %"
319           End If
320           If nilaiDemandSouth(5) + nilaiDemandNorth(5) > 0 Then
321               !K50 = Round((nilaiSBSameOutletNorth(5) + nilaiSBSameOutletSouth(5) / (nilaiDemandSouth(5) + nilaiDemandNorth(0))) * 100, 0) & " %"
322           Else
323               !K50 = 0 & " %"
324           End If
325           If nilaiDemandSouth(6) + nilaiDemandNorth(6) > 0 Then
326               !K60 = Round((nilaiSBSameOutletNorth(6) + nilaiSBSameOutletSouth(6) / (nilaiDemandSouth(6) + nilaiDemandNorth(6))) * 100, 0) & " %"
327           Else
328               !K60 = 0 & " %"
329           End If
330           If nilaiDemandSouth(7) + nilaiDemandNorth(7) > 0 Then
331               !K70 = Round((nilaiSBSameOutletNorth(7) + nilaiSBSameOutletSouth(7) / (nilaiDemandSouth(7) + nilaiDemandNorth(7))) * 100, 0) & " %"
332           Else
333               !K70 = 0 & " %"
334           End If
335           If nilaiDemandSouth(8) + nilaiDemandNorth(8) > 0 Then
336               !K80 = Round((nilaiSBSameOutletNorth(8) + nilaiSBSameOutletSouth(8) / (nilaiDemandSouth(8) + nilaiDemandNorth(8))) * 100, 0) & " %"
337           Else
338               !K80 = 0 & " %"
339           End If
              
340           If (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
341               !tot16 = Round(((nilaiSBSameOutletNorth(0) + nilaiSBSameOutletNorth(1) + nilaiSBSameOutletNorth(2) + nilaiSBSameOutletNorth(3) + nilaiSBSameOutletNorth(4) + nilaiSBSameOutletNorth(5)) + (nilaiSBSameOutletSouth(0) + nilaiSBSameOutletSouth(1) + nilaiSBSameOutletSouth(2) + nilaiSBSameOutletSouth(3) + nilaiSBSameOutletSouth(4) + nilaiSBSameOutletSouth(5)) / ((nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5)) + (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)))) * 100, 0) & " %"
342           Else
343               !tot16 = 0 & " %"
344           End If
              
345           If (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
346               !tot26 = Round(((nilaiSBSameOutletSouth(1) + nilaiSBSameOutletSouth(2) + nilaiSBSameOutletSouth(3) + nilaiSBSameOutletSouth(4) + nilaiSBSameOutletSouth(5)) + (nilaiSBSameOutletNorth(1) + nilaiSBSameOutletNorth(2) + nilaiSBSameOutletNorth(3) + nilaiSBSameOutletNorth(4) + nilaiSBSameOutletNorth(5)) / ((nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5)))) * 100, 0) & " %"
347           Else
348               !tot26 = 0 & " %"
349           End If
               
350           If (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)) + nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
351               !tot19 = Round(((nilaiSBSameOutletNorth(0) + nilaiSBSameOutletNorth(1) + nilaiSBSameOutletNorth(2) + nilaiSBSameOutletNorth(3) + nilaiSBSameOutletNorth(4) + nilaiSBSameOutletNorth(5) + nilaiSBSameOutletNorth(6) + nilaiSBSameOutletNorth(7) + nilaiSBSameOutletNorth(8)) + (nilaiSBSameOutletSouth(0) + nilaiSBSameOutletSouth(1) + nilaiSBSameOutletSouth(2) + nilaiSBSameOutletSouth(3) + nilaiSBSameOutletSouth(4) + nilaiSBSameOutletSouth(5) + nilaiSBSameOutletSouth(6) + nilaiSBSameOutletSouth(7) + nilaiSBSameOutletSouth(8)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)) + (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
352           Else
353               !tot19 = 0 & " %"
354           End If
               
355           If (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8)) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
356               !tot29 = Round(((nilaiSBSameOutletSouth(1) + nilaiSBSameOutletSouth(2) + nilaiSBSameOutletSouth(3) + nilaiSBSameOutletSouth(4) + nilaiSBSameOutletSouth(5) + nilaiSBSameOutletSouth(6) + nilaiSBSameOutletSouth(7) + nilaiSBSameOutletSouth(8)) + (nilaiSBSameOutletNorth(1) + nilaiSBSameOutletNorth(2) + nilaiSBSameOutletNorth(3) + nilaiSBSameOutletNorth(4) + nilaiSBSameOutletNorth(5) + nilaiSBSameOutletNorth(6) + nilaiSBSameOutletNorth(7) + nilaiSBSameOutletNorth(8)) / ((nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8)) + (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)))) * 100, 0) & " %"
357           Else
358               !tot29 = 0 & " %"
359           End If
               
360           !Urut = 5
361           !judul = "Retention Rate Same Outlet"
362           !model = "NRM"
363           !north = 5
364           .Update
365       End With
          
          '' --- akhir urut 5 ---------------------
          
          
          '' -----------------------------------------
          '' urut 6
          '' -----------------------------------------
          
366       With rsLaporan
367           .AddNew
368           If nilaiDemandNorth(0) > 0 Then
369               !K1 = Round((nilaiSBNRMNorth(0) / nilaiDemandNorth(0)) * 100, 0) & " %"
370           Else
371               !K1 = 0 & " %"
372           End If
373           If nilaiDemandNorth(1) > 0 Then
374               !K10 = Round((nilaiSBNRMNorth(1) / nilaiDemandNorth(1)) * 100, 0) & " %"
375           Else
376               !K10 = 0 & " %"
377           End If
378           If nilaiDemandNorth(2) > 0 Then
379               !K20 = Round((nilaiSBNRMNorth(2) / nilaiDemandNorth(2)) * 100, 0) & " %"
380           Else
381               !K20 = 0 & " %"
382           End If
383           If nilaiDemandNorth(3) > 0 Then
384               !K30 = Round((nilaiSBNRMNorth(3) / nilaiDemandNorth(3)) * 100, 0) & " %"
385           Else
386               !K30 = 0 & " %"
387           End If
388           If nilaiDemandNorth(4) > 0 Then
389               !K40 = Round((nilaiSBNRMNorth(4) / nilaiDemandNorth(4)) * 100, 0) & " %"
390           Else
391               !K40 = 0 & " %"
392           End If
393           If nilaiDemandNorth(5) > 0 Then
394               !K50 = Round((nilaiSBNRMNorth(5) / nilaiDemandNorth(5)) * 100, 0) & " %"
395           Else
396               !K50 = 0 & " %"
397           End If
398           If nilaiDemandNorth(6) > 0 Then
399               !K60 = Round((nilaiSBNRMNorth(6) / nilaiDemandNorth(6)) * 100, 0) & " %"
400           Else
401               !K60 = 0 & " %"
402           End If
403           If nilaiDemandNorth(7) > 0 Then
404               !K70 = Round((nilaiSBNRMNorth(7) / nilaiDemandNorth(7)) * 100, 0) & " %"
405           Else
406               !K70 = 0 & " %"
407           End If
408           If nilaiDemandNorth(8) > 0 Then
409               !K80 = Round((nilaiSBNRMNorth(8) / nilaiDemandNorth(8)) * 100, 0) & " %"
410           Else
411               !K80 = 0 & " %"
412           End If
              
413           If nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) > 0 Then
414               !tot16 = Round(((nilaiSBNRMNorth(0) + nilaiSBNRMNorth(1) + nilaiSBNRMNorth(2) + nilaiSBNRMNorth(3) + nilaiSBNRMNorth(4) + nilaiSBNRMNorth(5)) / (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5))) * 100, 0) & " %"
415           Else
416               !tot16 = 0 & " %"
417           End If
              
418           If nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) > 0 Then
419               !tot26 = Round(((nilaiSBNRMNorth(1) + nilaiSBNRMNorth(2) + nilaiSBNRMNorth(3) + nilaiSBNRMNorth(4) + nilaiSBNRMNorth(5)) / (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5))) * 100, 0) & " %"
420           Else
421               !tot26 = 0 & " %"
422           End If
               
423           If nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8) > 0 Then
424               !tot19 = Round(((nilaiSBNRMNorth(0) + nilaiSBNRMNorth(1) + nilaiSBNRMNorth(2) + nilaiSBNRMNorth(3) + nilaiSBNRMNorth(4) + nilaiSBNRMNorth(5) + nilaiSBNRMNorth(6) + nilaiSBNRMNorth(7) + nilaiSBNRMNorth(8)) / (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
425           Else
426               !tot19 = 0 & " %"
427           End If
               
428           If nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8) > 0 Then
429               !tot29 = Round(((nilaiSBNRMNorth(1) + nilaiSBNRMNorth(2) + nilaiSBNRMNorth(3) + nilaiSBNRMNorth(4) + nilaiSBNRMNorth(5) + nilaiSBNRMNorth(6) + nilaiSBNRMNorth(7) + nilaiSBNRMNorth(8)) / (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
430           Else
431               !tot29 = 0 & " %"
432           End If
               
433           !Urut = 6
434           !judul = "Retention Rate Same Group NCM"
435           !model = "NORTH AREA"
436           !north = 2
437           .Update
438       End With
          
          
439       With rsLaporan
440           .AddNew
441           If nilaiDemandSouth(0) > 0 Then
442               !K1 = Round((nilaiSBNRMSouth(0) / nilaiDemandSouth(0)) * 100, 0) & " %"
443           Else
444               !K1 = 0 & " %"
445           End If
446           If nilaiDemandSouth(1) > 0 Then
447               !K10 = Round((nilaiSBNRMSouth(1) / nilaiDemandSouth(1)) * 100, 0) & " %"
448           Else
449               !K10 = 0 & " %"
450           End If
451           If nilaiDemandSouth(2) > 0 Then
452               !K20 = Round((nilaiSBNRMSouth(2) / nilaiDemandSouth(2)) * 100, 0) & " %"
453           Else
454               !K20 = 0 & " %"
455           End If
456           If nilaiDemandSouth(3) > 0 Then
457               !K30 = Round((nilaiSBNRMSouth(3) / nilaiDemandSouth(3)) * 100, 0) & " %"
458           Else
459               !K30 = 0 & " %"
460           End If
461           If nilaiDemandSouth(4) > 0 Then
462               !K40 = Round((nilaiSBNRMSouth(4) / nilaiDemandSouth(4)) * 100, 0) & " %"
463           Else
464               !K40 = 0 & " %"
465           End If
466           If nilaiDemandSouth(5) > 0 Then
467               !K50 = Round((nilaiSBNRMSouth(5) / nilaiDemandSouth(5)) * 100, 0) & " %"
468           Else
469               !K50 = 0 & " %"
470           End If
471           If nilaiDemandSouth(6) > 0 Then
472               !K60 = Round((nilaiSBNRMSouth(6) / nilaiDemandSouth(6)) * 100, 0) & " %"
473           Else
474               !K60 = 0 & " %"
475           End If
476           If nilaiDemandSouth(7) > 0 Then
477               !K70 = Round((nilaiSBNRMSouth(7) / nilaiDemandSouth(7)) * 100, 0) & " %"
478           Else
479               !K70 = 0 & " %"
480           End If
481           If nilaiDemandSouth(8) > 0 Then
482               !K80 = Round((nilaiSBNRMSouth(8) / nilaiDemandSouth(8)) * 100, 0) & " %"
483           Else
484               !K80 = 0 & " %"
485           End If
              
486           If nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
487               !tot16 = Round(((nilaiSBNRMSouth(0) + nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5))) * 100, 0) & " %"
488           Else
489               !tot16 = 0 & " %"
490           End If
              
491           If nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
492               !tot26 = Round(((nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5)) / (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5))) * 100, 0) & " %"
493           Else
494               !tot26 = 0 & " %"
495           End If
               
496           If nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
497               !tot19 = Round(((nilaiSBNRMSouth(0) + nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5) + nilaiSBNRMSouth(6) + nilaiSBNRMSouth(7) + nilaiSBNRMSouth(8)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8))) * 100, 0) & " %"
498           Else
499               !tot19 = 0 & " %"
500           End If
               
501           If nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
502               !tot29 = Round(((nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5) + nilaiSBNRMSouth(6) + nilaiSBNRMSouth(7) + nilaiSBNRMSouth(8)) / (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8))) * 100, 0) & " %"
503           Else
504               !tot29 = 0 & " %"
505           End If
               
506           !Urut = 6
507           !judul = "Retention Rate Same Group NCM"
508           !model = "SOUTH AREA"
509           !north = 4
510           .Update
511       End With
              
512       With rsLaporan
513           .AddNew
514           If nilaiDemandPareto(0) > 0 Then
515               !K1 = Round((nilaiSBNRMPareto(0) / nilaiDemandPareto(0)) * 100, 0) & " %"
516           Else
517               !K1 = 0 & " %"
518           End If
519           If nilaiDemandPareto(1) > 0 Then
520               !K10 = Round((nilaiSBNRMPareto(1) / nilaiDemandPareto(1)) * 100, 0) & " %"
521           Else
522               !K10 = 0 & " %"
523           End If
524           If nilaiDemandPareto(2) > 0 Then
525               !K20 = Round((nilaiSBNRMPareto(2) / nilaiDemandPareto(2)) * 100, 0) & " %"
526           Else
527               !K20 = 0 & " %"
528           End If
529           If nilaiDemandPareto(3) > 0 Then
530               !K30 = Round((nilaiSBNRMPareto(3) / nilaiDemandPareto(3)) * 100, 0) & " %"
531           Else
532               !K30 = 0 & " %"
533           End If
534           If nilaiDemandPareto(4) > 0 Then
535               !K40 = Round((nilaiSBNRMPareto(4) / nilaiDemandPareto(4)) * 100, 0) & " %"
536           Else
537               !K40 = 0 & " %"
538           End If
539           If nilaiDemandPareto(5) > 0 Then
540               !K50 = Round((nilaiSBNRMPareto(5) / nilaiDemandPareto(5)) * 100, 0) & " %"
541           Else
542               !K50 = 0 & " %"
543           End If
544           If nilaiDemandPareto(6) > 0 Then
545               !K60 = Round((nilaiSBNRMPareto(6) / nilaiDemandPareto(6)) * 100, 0) & " %"
546           Else
547               !K60 = 0 & " %"
548           End If
549           If nilaiDemandPareto(7) > 0 Then
550               !K70 = Round((nilaiSBNRMPareto(7) / nilaiDemandPareto(7)) * 100, 0) & " %"
551           Else
552               !K70 = 0 & " %"
553           End If
554           If nilaiDemandPareto(8) > 0 Then
555               !K80 = Round((nilaiSBNRMPareto(8) / nilaiDemandPareto(8)) * 100, 0) & " %"
556           Else
557               !K80 = 0 & " %"
558           End If
              
559           If nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) > 0 Then
560               !tot16 = Round(((nilaiSBNRMPareto(0) + nilaiSBNRMPareto(1) + nilaiSBNRMPareto(2) + nilaiSBNRMPareto(3) + nilaiSBNRMPareto(4) + nilaiSBNRMPareto(5)) / (nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5))) * 100, 0) & " %"
561           Else
562               !tot16 = 0 & " %"
563           End If
              
564           If nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) > 0 Then
565               !tot26 = Round(((nilaiSBNRMPareto(1) + nilaiSBNRMPareto(2) + nilaiSBNRMPareto(3) + nilaiSBNRMPareto(4) + nilaiSBNRMPareto(5)) / (nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5))) * 100, 0) & " %"
566           Else
567               !tot26 = 0 & " %"
568           End If
               
569           If nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8) > 0 Then
570               !tot19 = Round(((nilaiSBNRMPareto(0) + nilaiSBNRMPareto(1) + nilaiSBNRMPareto(2) + nilaiSBNRMPareto(3) + nilaiSBNRMPareto(4) + nilaiSBNRMPareto(5) + nilaiSBNRMPareto(6) + nilaiSBNRMPareto(7) + nilaiSBNRMPareto(8)) / (nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8))) * 100, 0) & " %"
571           Else
572               !tot19 = 0 & " %"
573           End If
               
574           If nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8) > 0 Then
575               !tot29 = Round(((nilaiSBNRMPareto(1) + nilaiSBNRMPareto(2) + nilaiSBNRMPareto(3) + nilaiSBNRMPareto(4) + nilaiSBNRMPareto(5) + nilaiSBNRMPareto(6) + nilaiSBNRMPareto(7) + nilaiSBNRMPareto(8)) / (nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8))) * 100, 0) & " %"
576           Else
577               !tot29 = 0 & " %"
578           End If
               
579           !Urut = 6
580           !judul = "Retention Rate Same Group NCM"
581           !model = "PARETO"
582           !north = 6
583           .Update
584       End With
          
585       With rsLaporan
586           .AddNew
587           If nilaiDemandNonPareto(0) > 0 Then
588               !K1 = Round((nilaiSBNRMNonPareto(0) / nilaiDemandNonPareto(0)) * 100, 0) & " %"
589           Else
590               !K1 = 0 & " %"
591           End If
592           If nilaiDemandNonPareto(1) > 0 Then
593               !K10 = Round((nilaiSBNRMNonPareto(1) / nilaiDemandNonPareto(1)) * 100, 0) & " %"
594           Else
595               !K10 = 0 & " %"
596           End If
597           If nilaiDemandNonPareto(2) > 0 Then
598               !K20 = Round((nilaiSBNRMNonPareto(2) / nilaiDemandNonPareto(2)) * 100, 0) & " %"
599           Else
600               !K20 = 0 & " %"
601           End If
602           If nilaiDemandNonPareto(3) > 0 Then
603               !K30 = Round((nilaiSBNRMNonPareto(3) / nilaiDemandNonPareto(3)) * 100, 0) & " %"
604           Else
605               !K30 = 0 & " %"
606           End If
607           If nilaiDemandNonPareto(4) > 0 Then
608               !K40 = Round((nilaiSBNRMNonPareto(4) / nilaiDemandNonPareto(4)) * 100, 0) & " %"
609           Else
610               !K40 = 0 & " %"
611           End If
612           If nilaiDemandNonPareto(5) > 0 Then
613               !K50 = Round((nilaiSBNRMNonPareto(5) / nilaiDemandNonPareto(5)) * 100, 0) & " %"
614           Else
615               !K50 = 0 & " %"
616           End If
617           If nilaiDemandNonPareto(6) > 0 Then
618               !K60 = Round((nilaiSBNRMNonPareto(6) / nilaiDemandNonPareto(6)) * 100, 0) & " %"
619           Else
620               !K60 = 0 & " %"
621           End If
622           If nilaiDemandNonPareto(7) > 0 Then
623               !K70 = Round((nilaiSBNRMNonPareto(7) / nilaiDemandNonPareto(7)) * 100, 0) & " %"
624           Else
625               !K70 = 0 & " %"
626           End If
627           If nilaiDemandNonPareto(8) > 0 Then
628               !K80 = Round((nilaiSBNRMNonPareto(8) / nilaiDemandNonPareto(8)) * 100, 0) & " %"
629           Else
630               !K80 = 0 & " %"
631           End If
              
632           If nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) > 0 Then
633               !tot16 = Round(((nilaiSBNRMNonPareto(0) + nilaiSBNRMNonPareto(1) + nilaiSBNRMNonPareto(2) + nilaiSBNRMNonPareto(3) + nilaiSBNRMNonPareto(4) + nilaiSBNRMNonPareto(5)) / (nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5))) * 100, 0) & " %"
634           Else
635               !tot16 = 0 & " %"
636           End If
              
637           If nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) > 0 Then
638               !tot26 = Round(((nilaiSBNRMNonPareto(1) + nilaiSBNRMNonPareto(2) + nilaiSBNRMNonPareto(3) + nilaiSBNRMNonPareto(4) + nilaiSBNRMNonPareto(5)) / (nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5))) * 100, 0) & " %"
639           Else
640               !tot26 = 0 & " %"
641           End If
               
642           If nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8) > 0 Then
643               !tot19 = Round(((nilaiSBNRMNonPareto(0) + nilaiSBNRMNonPareto(1) + nilaiSBNRMNonPareto(2) + nilaiSBNRMNonPareto(3) + nilaiSBNRMNonPareto(4) + nilaiSBNRMNonPareto(5) + nilaiSBNRMNonPareto(6) + nilaiSBNRMNonPareto(7) + nilaiSBNRMNonPareto(8)) / (nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8))) * 100, 0) & " %"
644           Else
645               !tot19 = 0 & " %"
646           End If
               
647           If nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8) > 0 Then
648               !tot29 = Round(((nilaiSBNRMNonPareto(1) + nilaiSBNRMNonPareto(2) + nilaiSBNRMNonPareto(3) + nilaiSBNRMNonPareto(4) + nilaiSBNRMNonPareto(5) + nilaiSBNRMNonPareto(6) + nilaiSBNRMNonPareto(7) + nilaiSBNRMNonPareto(8)) / (nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8))) * 100, 0) & " %"
649           Else
650               !tot29 = 0 & " %"
651           End If
               
652           !Urut = 6
653           !judul = "Retention Rate Same Group NCM"
654           !model = "NON PARETO"
655           !north = 7
656           .Update
657       End With
          
          
658       With rsLaporan
659           .AddNew
660           If nilaiDemandSouth(0) + nilaiDemandNorth(0) > 0 Then
661               !K1 = Round((nilaiSBNRMNorth(0) + nilaiSBNRMSouth(0) / (nilaiDemandSouth(0) + nilaiDemandNorth(0))) * 100, 0) & " %"
662           Else
663               !K1 = 0 & " %"
664           End If
665           If nilaiDemandSouth(1) + nilaiDemandNorth(1) > 0 Then
666               !K10 = Round((nilaiSBNRMNorth(1) + nilaiSBNRMSouth(1) / (nilaiDemandSouth(1) + nilaiDemandNorth(1))) * 100, 0) & " %"
667           Else
668               !K10 = 0 & " %"
669           End If
670           If nilaiDemandSouth(2) + nilaiDemandNorth(2) > 0 Then
671               !K20 = Round((nilaiSBNRMNorth(2) + nilaiSBNRMSouth(2) / (nilaiDemandSouth(2) + nilaiDemandNorth(2))) * 100, 0) & " %"
672           Else
673               !K20 = 0 & " %"
674           End If
675           If nilaiDemandSouth(3) + nilaiDemandNorth(3) > 0 Then
676               !K30 = Round((nilaiSBNRMNorth(3) + nilaiSBNRMSouth(3) / (nilaiDemandSouth(3) + nilaiDemandNorth(3))) * 100, 0) & " %"
677           Else
678               !K30 = 0 & " %"
679           End If
680           If nilaiDemandSouth(4) + nilaiDemandNorth(4) > 0 Then
681               !K40 = Round((nilaiSBNRMNorth(4) + nilaiSBNRMSouth(4) / (nilaiDemandSouth(4) + nilaiDemandNorth(4))) * 100, 0) & " %"
682           Else
683               !K40 = 0 & " %"
684           End If
685           If nilaiDemandSouth(5) + nilaiDemandNorth(5) > 0 Then
686               !K50 = Round((nilaiSBNRMNorth(5) + nilaiSBNRMSouth(5) / (nilaiDemandSouth(5) + nilaiDemandNorth(0))) * 100, 0) & " %"
687           Else
688               !K50 = 0 & " %"
689           End If
690           If nilaiDemandSouth(6) + nilaiDemandNorth(6) > 0 Then
691               !K60 = Round((nilaiSBNRMNorth(6) + nilaiSBNRMSouth(6) / (nilaiDemandSouth(6) + nilaiDemandNorth(6))) * 100, 0) & " %"
692           Else
693               !K60 = 0 & " %"
694           End If
695           If nilaiDemandSouth(7) + nilaiDemandNorth(7) > 0 Then
696               !K70 = Round((nilaiSBNRMNorth(7) + nilaiSBNRMSouth(7) / (nilaiDemandSouth(7) + nilaiDemandNorth(7))) * 100, 0) & " %"
697           Else
698               !K70 = 0 & " %"
699           End If
700           If nilaiDemandSouth(8) + nilaiDemandNorth(8) > 0 Then
701               !K80 = Round((nilaiSBNRMNorth(8) + nilaiSBNRMSouth(8) / (nilaiDemandSouth(8) + nilaiDemandNorth(8))) * 100, 0) & " %"
702           Else
703               !K80 = 0 & " %"
704           End If
              
705           If (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
706               !tot16 = Round(((nilaiSBNRMNorth(0) + nilaiSBNRMNorth(1) + nilaiSBNRMNorth(2) + nilaiSBNRMNorth(3) + nilaiSBNRMNorth(4) + nilaiSBNRMNorth(5)) + (nilaiSBNRMSouth(0) + nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5)) / ((nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5)) + (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)))) * 100, 0) & " %"
707           Else
708               !tot16 = 0 & " %"
709           End If
              
710           If (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
711               !tot26 = Round(((nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5)) + (nilaiSBNRMNorth(1) + nilaiSBNRMNorth(2) + nilaiSBNRMNorth(3) + nilaiSBNRMNorth(4) + nilaiSBNRMNorth(5)) / ((nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5)))) * 100, 0) & " %"
712           Else
713               !tot26 = 0 & " %"
714           End If
               
715           If (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)) + nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
716               !tot19 = Round(((nilaiSBNRMNorth(0) + nilaiSBNRMNorth(1) + nilaiSBNRMNorth(2) + nilaiSBNRMNorth(3) + nilaiSBNRMNorth(4) + nilaiSBNRMNorth(5) + nilaiSBNRMNorth(6) + nilaiSBNRMNorth(7) + nilaiSBNRMNorth(8)) + (nilaiSBNRMSouth(0) + nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5) + nilaiSBNRMSouth(6) + nilaiSBNRMSouth(7) + nilaiSBNRMSouth(8)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)) + (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
717           Else
718               !tot19 = 0 & " %"
719           End If
               
720           If (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8)) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
721               !tot29 = Round(((nilaiSBNRMSouth(1) + nilaiSBNRMSouth(2) + nilaiSBNRMSouth(3) + nilaiSBNRMSouth(4) + nilaiSBNRMSouth(5) + nilaiSBNRMSouth(6) + nilaiSBNRMSouth(7) + nilaiSBNRMSouth(8)) + (nilaiSBNRMNorth(1) + nilaiSBNRMNorth(2) + nilaiSBNRMNorth(3) + nilaiSBNRMNorth(4) + nilaiSBNRMNorth(5) + nilaiSBNRMNorth(6) + nilaiSBNRMNorth(7) + nilaiSBNRMNorth(8)) / ((nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8)) + (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)))) * 100, 0) & " %"
722           Else
723               !tot29 = 0 & " %"
724           End If
               
725           !Urut = 6
726           !judul = "Retention Rate Same Group NCM"
727           !model = "NRM"
728           !north = 5
729           .Update
730       End With
          
          '' --- akhir urut 6 ---------------------
          
          '' ---------------------------------------------------
          '' awal urut ke 7
          '' ---------------------------------------------------
          
731       With rsLaporan
732           .AddNew
733           If nilaiDemandNorth(0) > 0 Then
734               !K1 = Round((nilaiTotalNorth(0) / nilaiDemandNorth(0)) * 100, 0) & " %"
735           Else
736               !K1 = 0 & " %"
737           End If
738           If nilaiDemandNorth(1) > 0 Then
739               !K10 = Round((nilaiTotalNorth(1) / nilaiDemandNorth(1)) * 100, 0) & " %"
740           Else
741               !K10 = 0 & " %"
742           End If
743           If nilaiDemandNorth(2) > 0 Then
744               !K20 = Round((nilaiTotalNorth(2) / nilaiDemandNorth(2)) * 100, 0) & " %"
745           Else
746               !K20 = 0 & " %"
747           End If
748           If nilaiDemandNorth(3) > 0 Then
749               !K30 = Round((nilaiTotalNorth(3) / nilaiDemandNorth(3)) * 100, 0) & " %"
750           Else
751               !K30 = 0 & " %"
752           End If
753           If nilaiDemandNorth(4) > 0 Then
754               !K40 = Round((nilaiTotalNorth(4) / nilaiDemandNorth(4)) * 100, 0) & " %"
755           Else
756               !K40 = 0 & " %"
757           End If
758           If nilaiDemandNorth(5) > 0 Then
759               !K50 = Round((nilaiTotalNorth(5) / nilaiDemandNorth(5)) * 100, 0) & " %"
760           Else
761               !K50 = 0 & " %"
762           End If
763           If nilaiDemandNorth(6) > 0 Then
764               !K60 = Round((nilaiTotalNorth(6) / nilaiDemandNorth(6)) * 100, 0) & " %"
765           Else
766               !K60 = 0 & " %"
767           End If
768           If nilaiDemandNorth(7) > 0 Then
769               !K70 = Round((nilaiTotalNorth(7) / nilaiDemandNorth(7)) * 100, 0) & " %"
770           Else
771               !K70 = 0 & " %"
772           End If
773           If nilaiDemandNorth(8) > 0 Then
774               !K80 = Round((nilaiTotalNorth(8) / nilaiDemandNorth(8)) * 100, 0) & " %"
775           Else
776               !K80 = 0 & " %"
777           End If
              
778           If nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) > 0 Then
779               !tot16 = Round(((nilaiTotalNorth(0) + nilaiTotalNorth(1) + nilaiTotalNorth(2) + nilaiTotalNorth(3) + nilaiTotalNorth(4) + nilaiTotalNorth(5)) / (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5))) * 100, 0) & " %"
780           Else
781               !tot16 = 0 & " %"
782           End If
              
783           If nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) > 0 Then
784               !tot26 = Round(((nilaiTotalNorth(1) + nilaiTotalNorth(2) + nilaiTotalNorth(3) + nilaiTotalNorth(4) + nilaiTotalNorth(5)) / (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5))) * 100, 0) & " %"
785           Else
786               !tot26 = 0 & " %"
787           End If
               
788           If nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8) > 0 Then
789               !tot19 = Round(((nilaiTotalNorth(0) + nilaiTotalNorth(1) + nilaiTotalNorth(2) + nilaiTotalNorth(3) + nilaiTotalNorth(4) + nilaiTotalNorth(5) + nilaiTotalNorth(6) + nilaiTotalNorth(7) + nilaiTotalNorth(8)) / (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
790           Else
791               !tot19 = 0 & " %"
792           End If
               
793           If nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8) > 0 Then
794               !tot29 = Round(((nilaiTotalNorth(1) + nilaiTotalNorth(2) + nilaiTotalNorth(3) + nilaiTotalNorth(4) + nilaiTotalNorth(5) + nilaiTotalNorth(6) + nilaiTotalNorth(7) + nilaiTotalNorth(8)) / (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
795           Else
796               !tot29 = 0 & " %"
797           End If
               
798           !Urut = 7
799           !judul = "Retention Rate Total (NCM & Non NCM)"
800           !model = "NORTH AREA"
801           !north = 2
802           .Update
803       End With
          
804       With rsLaporan
805           .AddNew
806           If nilaiDemandSouth(0) > 0 Then
807               !K1 = Round((nilaiTotalSouth(0) / nilaiDemandSouth(0)) * 100, 0) & " %"
808           Else
809               !K1 = 0 & " %"
810           End If
811           If nilaiDemandSouth(1) > 0 Then
812               !K10 = Round((nilaiTotalSouth(1) / nilaiDemandSouth(1)) * 100, 0) & " %"
813           Else
814               !K10 = 0 & " %"
815           End If
816           If nilaiDemandSouth(2) > 0 Then
817               !K20 = Round((nilaiTotalSouth(2) / nilaiDemandSouth(2)) * 100, 0) & " %"
818           Else
819               !K20 = 0 & " %"
820           End If
821           If nilaiDemandSouth(3) > 0 Then
822               !K30 = Round((nilaiTotalSouth(3) / nilaiDemandSouth(3)) * 100, 0) & " %"
823           Else
824               !K30 = 0 & " %"
825           End If
826           If nilaiDemandSouth(4) > 0 Then
827               !K40 = Round((nilaiTotalSouth(4) / nilaiDemandSouth(4)) * 100, 0) & " %"
828           Else
829               !K40 = 0 & " %"
830           End If
831           If nilaiDemandSouth(5) > 0 Then
832               !K50 = Round((nilaiTotalSouth(5) / nilaiDemandSouth(5)) * 100, 0) & " %"
833           Else
834               !K50 = 0 & " %"
835           End If
836           If nilaiDemandSouth(6) > 0 Then
837               !K60 = Round((nilaiTotalSouth(6) / nilaiDemandSouth(6)) * 100, 0) & " %"
838           Else
839               !K60 = 0 & " %"
840           End If
841           If nilaiDemandSouth(7) > 0 Then
842               !K70 = Round((nilaiTotalSouth(7) / nilaiDemandSouth(7)) * 100, 0) & " %"
843           Else
844               !K70 = 0 & " %"
845           End If
846           If nilaiDemandSouth(8) > 0 Then
847               !K80 = Round((nilaiTotalSouth(8) / nilaiDemandSouth(8)) * 100, 0) & " %"
848           Else
849               !K80 = 0 & " %"
850           End If
              
851           If nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
852               !tot16 = Round(((nilaiTotalSouth(0) + nilaiTotalSouth(1) + nilaiTotalSouth(2) + nilaiTotalSouth(3) + nilaiTotalSouth(4) + nilaiTotalSouth(5)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5))) * 100, 0) & " %"
853           Else
854               !tot16 = 0 & " %"
855           End If
              
856           If nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
857               !tot26 = Round(((nilaiTotalSouth(1) + nilaiTotalSouth(2) + nilaiTotalSouth(3) + nilaiTotalSouth(4) + nilaiTotalSouth(5)) / (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5))) * 100, 0) & " %"
858           Else
859               !tot26 = 0 & " %"
860           End If
               
861           If nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
862               !tot19 = Round(((nilaiTotalSouth(0) + nilaiTotalSouth(1) + nilaiTotalSouth(2) + nilaiTotalSouth(3) + nilaiTotalSouth(4) + nilaiTotalSouth(5) + nilaiTotalSouth(6) + nilaiTotalSouth(7) + nilaiTotalSouth(8)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8))) * 100, 0) & " %"
863           Else
864               !tot19 = 0 & " %"
865           End If
               
866           If nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
867               !tot29 = Round(((nilaiTotalSouth(1) + nilaiTotalSouth(2) + nilaiTotalSouth(3) + nilaiTotalSouth(4) + nilaiTotalSouth(5) + nilaiTotalSouth(6) + nilaiTotalSouth(7) + nilaiTotalSouth(8)) / (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8))) * 100, 0) & " %"
868           Else
869               !tot29 = 0 & " %"
870           End If
               
871           !Urut = 7
872           !judul = "Retention Rate Total (NCM & Non NCM)"
873           !model = "SOUTH AREA"
874           !north = 4
875           .Update
876       End With
          
          
877       With rsLaporan
878           .AddNew
879           If nilaiDemandSouth(0) + nilaiDemandNorth(0) > 0 Then
880               !K1 = Round((nilaiTotalNorth(0) + nilaiTotalSouth(0) / (nilaiDemandSouth(0) + nilaiDemandNorth(0))) * 100, 0) & " %"
881           Else
882               !K1 = 0 & " %"
883           End If
884           If nilaiDemandSouth(1) + nilaiDemandNorth(1) > 0 Then
885               !K10 = Round((nilaiTotalNorth(1) + nilaiTotalSouth(1) / (nilaiDemandSouth(1) + nilaiDemandNorth(1))) * 100, 0) & " %"
886           Else
887               !K10 = 0 & " %"
888           End If
889           If nilaiDemandSouth(2) + nilaiDemandNorth(2) > 0 Then
890               !K20 = Round((nilaiTotalNorth(2) + nilaiTotalSouth(2) / (nilaiDemandSouth(2) + nilaiDemandNorth(2))) * 100, 0) & " %"
891           Else
892               !K20 = 0 & " %"
893           End If
894           If nilaiDemandSouth(3) + nilaiDemandNorth(3) > 0 Then
895               !K30 = Round((nilaiTotalNorth(3) + nilaiTotalSouth(3) / (nilaiDemandSouth(3) + nilaiDemandNorth(3))) * 100, 0) & " %"
896           Else
897               !K30 = 0 & " %"
898           End If
899           If nilaiDemandSouth(4) + nilaiDemandNorth(4) > 0 Then
900               !K40 = Round((nilaiTotalNorth(4) + nilaiTotalSouth(4) / (nilaiDemandSouth(4) + nilaiDemandNorth(4))) * 100, 0) & " %"
901           Else
902               !K40 = 0 & " %"
903           End If
904           If nilaiDemandSouth(5) + nilaiDemandNorth(5) > 0 Then
905               !K50 = Round((nilaiTotalNorth(5) + nilaiTotalSouth(5) / (nilaiDemandSouth(5) + nilaiDemandNorth(0))) * 100, 0) & " %"
906           Else
907               !K50 = 0 & " %"
908           End If
909           If nilaiDemandSouth(6) + nilaiDemandNorth(6) > 0 Then
910               !K60 = Round((nilaiTotalNorth(6) + nilaiTotalSouth(6) / (nilaiDemandSouth(6) + nilaiDemandNorth(6))) * 100, 0) & " %"
911           Else
912               !K60 = 0 & " %"
913           End If
914           If nilaiDemandSouth(7) + nilaiDemandNorth(7) > 0 Then
915               !K70 = Round((nilaiTotalNorth(7) + nilaiTotalSouth(7) / (nilaiDemandSouth(7) + nilaiDemandNorth(7))) * 100, 0) & " %"
916           Else
917               !K70 = 0 & " %"
918           End If
919           If nilaiDemandSouth(8) + nilaiDemandNorth(8) > 0 Then
920               !K80 = Round((nilaiTotalNorth(8) + nilaiTotalSouth(8) / (nilaiDemandSouth(8) + nilaiDemandNorth(8))) * 100, 0) & " %"
921           Else
922               !K80 = 0 & " %"
923           End If
              
924           If (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
925               !tot16 = Round(((nilaiTotalNorth(0) + nilaiTotalNorth(1) + nilaiTotalNorth(2) + nilaiTotalNorth(3) + nilaiTotalNorth(4) + nilaiTotalNorth(5)) + (nilaiTotalSouth(0) + nilaiTotalSouth(1) + nilaiTotalSouth(2) + nilaiTotalSouth(3) + nilaiTotalSouth(4) + nilaiTotalSouth(5)) / ((nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5)) + (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)))) * 100, 0) & " %"
926           Else
927               !tot16 = 0 & " %"
928           End If
              
929           If (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) > 0 Then
930               !tot26 = Round(((nilaiTotalSouth(1) + nilaiTotalSouth(2) + nilaiTotalSouth(3) + nilaiTotalSouth(4) + nilaiTotalSouth(5)) + (nilaiTotalNorth(1) + nilaiTotalNorth(2) + nilaiTotalNorth(3) + nilaiTotalNorth(4) + nilaiTotalNorth(5)) / ((nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5)) + (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5)))) * 100, 0) & " %"
931           Else
932               !tot26 = 0 & " %"
933           End If
               
934           If (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)) + nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
935               !tot19 = Round(((nilaiTotalNorth(0) + nilaiTotalNorth(1) + nilaiTotalNorth(2) + nilaiTotalNorth(3) + nilaiTotalNorth(4) + nilaiTotalNorth(5) + nilaiTotalNorth(6) + nilaiTotalNorth(7) + nilaiTotalNorth(8)) + (nilaiTotalSouth(0) + nilaiTotalSouth(1) + nilaiTotalSouth(2) + nilaiTotalSouth(3) + nilaiTotalSouth(4) + nilaiTotalSouth(5) + nilaiTotalSouth(6) + nilaiTotalSouth(7) + nilaiTotalSouth(8)) / (nilaiDemandSouth(0) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)) + (nilaiDemandNorth(0) + nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8))) * 100, 0) & " %"
936           Else
937               !tot19 = 0 & " %"
938           End If
               
939           If (nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8)) + nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8) > 0 Then
940               !tot29 = Round(((nilaiTotalSouth(1) + nilaiTotalSouth(2) + nilaiTotalSouth(3) + nilaiTotalSouth(4) + nilaiTotalSouth(5) + nilaiTotalSouth(6) + nilaiTotalSouth(7) + nilaiTotalSouth(8)) + (nilaiTotalNorth(1) + nilaiTotalNorth(2) + nilaiTotalNorth(3) + nilaiTotalNorth(4) + nilaiTotalNorth(5) + nilaiTotalNorth(6) + nilaiTotalNorth(7) + nilaiTotalNorth(8)) / ((nilaiDemandNorth(1) + nilaiDemandNorth(2) + nilaiDemandNorth(3) + nilaiDemandNorth(4) + nilaiDemandNorth(5) + nilaiDemandNorth(6) + nilaiDemandNorth(7) + nilaiDemandNorth(8)) + (nilaiDemandSouth(1) + nilaiDemandSouth(2) + nilaiDemandSouth(3) + nilaiDemandSouth(4) + nilaiDemandSouth(5) + nilaiDemandSouth(6) + nilaiDemandSouth(7) + nilaiDemandSouth(8)))) * 100, 0) & " %"
941           Else
942               !tot29 = 0 & " %"
943           End If
               
944           !Urut = 7
945           !judul = "Retention Rate Total (NCM & Non NCM)"
946           !model = "NRM"
947           !north = 5
948           .Update
949       End With
          
          
          
950       With rsLaporan
951           .AddNew
952           If nilaiDemandPareto(0) > 0 Then
953               !K1 = Round((nilaiTotalPareto(0) / nilaiDemandPareto(0)) * 100, 0) & " %"
954           Else
955               !K1 = 0 & " %"
956           End If
957           If nilaiDemandPareto(1) > 0 Then
958               !K10 = Round((nilaiTotalPareto(1) / nilaiDemandPareto(1)) * 100, 0) & " %"
959           Else
960               !K10 = 0 & " %"
961           End If
962           If nilaiDemandPareto(2) > 0 Then
963               !K20 = Round((nilaiTotalPareto(2) / nilaiDemandPareto(2)) * 100, 0) & " %"
964           Else
965               !K20 = 0 & " %"
966           End If
967           If nilaiDemandPareto(3) > 0 Then
968               !K30 = Round((nilaiTotalPareto(3) / nilaiDemandPareto(3)) * 100, 0) & " %"
969           Else
970               !K30 = 0 & " %"
971           End If
972           If nilaiDemandPareto(4) > 0 Then
973               !K40 = Round((nilaiTotalPareto(4) / nilaiDemandPareto(4)) * 100, 0) & " %"
974           Else
975               !K40 = 0 & " %"
976           End If
977           If nilaiDemandPareto(5) > 0 Then
978               !K50 = Round((nilaiTotalPareto(5) / nilaiDemandPareto(5)) * 100, 0) & " %"
979           Else
980               !K50 = 0 & " %"
981           End If
982           If nilaiDemandPareto(6) > 0 Then
983               !K60 = Round((nilaiTotalPareto(6) / nilaiDemandPareto(6)) * 100, 0) & " %"
984           Else
985               !K60 = 0 & " %"
986           End If
987           If nilaiDemandPareto(7) > 0 Then
988               !K70 = Round((nilaiTotalPareto(7) / nilaiDemandPareto(7)) * 100, 0) & " %"
989           Else
990               !K70 = 0 & " %"
991           End If
992           If nilaiDemandPareto(8) > 0 Then
993               !K80 = Round((nilaiTotalPareto(8) / nilaiDemandPareto(8)) * 100, 0) & " %"
994           Else
995               !K80 = 0 & " %"
996           End If
              
997           If nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) > 0 Then
998               !tot16 = Round(((nilaiTotalPareto(0) + nilaiTotalPareto(1) + nilaiTotalPareto(2) + nilaiTotalPareto(3) + nilaiTotalPareto(4) + nilaiTotalPareto(5)) / (nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5))) * 100, 0) & " %"
999           Else
1000              !tot16 = 0 & " %"
1001          End If
              
1002          If nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) > 0 Then
1003              !tot26 = Round(((nilaiTotalPareto(1) + nilaiTotalPareto(2) + nilaiTotalPareto(3) + nilaiTotalPareto(4) + nilaiTotalPareto(5)) / (nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5))) * 100, 0) & " %"
1004          Else
1005              !tot26 = 0 & " %"
1006          End If
               
1007          If nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8) > 0 Then
1008              !tot19 = Round(((nilaiTotalPareto(0) + nilaiTotalPareto(1) + nilaiTotalPareto(2) + nilaiTotalPareto(3) + nilaiTotalPareto(4) + nilaiTotalPareto(5) + nilaiTotalPareto(6) + nilaiTotalPareto(7) + nilaiTotalPareto(8)) / (nilaiDemandPareto(0) + nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8))) * 100, 0) & " %"
1009          Else
1010              !tot19 = 0 & " %"
1011          End If
               
1012          If nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8) > 0 Then
1013              !tot29 = Round(((nilaiTotalPareto(1) + nilaiTotalPareto(2) + nilaiTotalPareto(3) + nilaiTotalPareto(4) + nilaiTotalPareto(5) + nilaiTotalPareto(6) + nilaiTotalPareto(7) + nilaiTotalPareto(8)) / (nilaiDemandPareto(1) + nilaiDemandPareto(2) + nilaiDemandPareto(3) + nilaiDemandPareto(4) + nilaiDemandPareto(5) + nilaiDemandPareto(6) + nilaiDemandPareto(7) + nilaiDemandPareto(8))) * 100, 0) & " %"
1014          Else
1015              !tot29 = 0 & " %"
1016          End If
               
1017          !Urut = 7
1018          !judul = "Retention Rate Total (NCM & Non NCM)"
1019          !model = "PARETO"
1020          !north = 6
1021          .Update
1022      End With
          
          
1023      With rsLaporan
1024          .AddNew
1025          If nilaiDemandNonPareto(0) > 0 Then
1026              !K1 = Round((nilaiTotalNonPareto(0) / nilaiDemandNonPareto(0)) * 100, 0) & " %"
1027          Else
1028              !K1 = 0 & " %"
1029          End If
1030          If nilaiDemandNonPareto(1) > 0 Then
1031              !K10 = Round((nilaiTotalNonPareto(1) / nilaiDemandNonPareto(1)) * 100, 0) & " %"
1032          Else
1033              !K10 = 0 & " %"
1034          End If
1035          If nilaiDemandNonPareto(2) > 0 Then
1036              !K20 = Round((nilaiTotalNonPareto(2) / nilaiDemandNonPareto(2)) * 100, 0) & " %"
1037          Else
1038              !K20 = 0 & " %"
1039          End If
1040          If nilaiDemandNonPareto(3) > 0 Then
1041              !K30 = Round((nilaiTotalNonPareto(3) / nilaiDemandNonPareto(3)) * 100, 0) & " %"
1042          Else
1043              !K30 = 0 & " %"
1044          End If
1045          If nilaiDemandNonPareto(4) > 0 Then
1046              !K40 = Round((nilaiTotalNonPareto(4) / nilaiDemandNonPareto(4)) * 100, 0) & " %"
1047          Else
1048              !K40 = 0 & " %"
1049          End If
1050          If nilaiDemandNonPareto(5) > 0 Then
1051              !K50 = Round((nilaiTotalNonPareto(5) / nilaiDemandNonPareto(5)) * 100, 0) & " %"
1052          Else
1053              !K50 = 0 & " %"
1054          End If
1055          If nilaiDemandNonPareto(6) > 0 Then
1056              !K60 = Round((nilaiTotalNonPareto(6) / nilaiDemandNonPareto(6)) * 100, 0) & " %"
1057          Else
1058              !K60 = 0 & " %"
1059          End If
1060          If nilaiDemandNonPareto(7) > 0 Then
1061              !K70 = Round((nilaiTotalNonPareto(7) / nilaiDemandNonPareto(7)) * 100, 0) & " %"
1062          Else
1063              !K70 = 0 & " %"
1064          End If
1065          If nilaiDemandNonPareto(8) > 0 Then
1066              !K80 = Round((nilaiTotalNonPareto(8) / nilaiDemandNonPareto(8)) * 100, 0) & " %"
1067          Else
1068              !K80 = 0 & " %"
1069          End If
              
1070          If nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) > 0 Then
1071              !tot16 = Round(((nilaiTotalNonPareto(0) + nilaiTotalNonPareto(1) + nilaiTotalNonPareto(2) + nilaiTotalNonPareto(3) + nilaiTotalNonPareto(4) + nilaiTotalNonPareto(5)) / (nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5))) * 100, 0) & " %"
1072          Else
1073              !tot16 = 0 & " %"
1074          End If
              
1075          If nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) > 0 Then
1076              !tot26 = Round(((nilaiTotalNonPareto(1) + nilaiTotalNonPareto(2) + nilaiTotalNonPareto(3) + nilaiTotalNonPareto(4) + nilaiTotalNonPareto(5)) / (nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5))) * 100, 0) & " %"
1077          Else
1078              !tot26 = 0 & " %"
1079          End If
               
1080          If nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8) > 0 Then
1081              !tot19 = Round(((nilaiTotalNonPareto(0) + nilaiTotalNonPareto(1) + nilaiTotalNonPareto(2) + nilaiTotalNonPareto(3) + nilaiTotalNonPareto(4) + nilaiTotalNonPareto(5) + nilaiTotalNonPareto(6) + nilaiTotalNonPareto(7) + nilaiTotalNonPareto(8)) / (nilaiDemandNonPareto(0) + nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8))) * 100, 0) & " %"
1082          Else
1083              !tot19 = 0 & " %"
1084          End If
               
1085          If nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8) > 0 Then
1086              !tot29 = Round(((nilaiTotalNonPareto(1) + nilaiTotalNonPareto(2) + nilaiTotalNonPareto(3) + nilaiTotalNonPareto(4) + nilaiTotalNonPareto(5) + nilaiTotalNonPareto(6) + nilaiTotalNonPareto(7) + nilaiTotalNonPareto(8)) / (nilaiDemandNonPareto(1) + nilaiDemandNonPareto(2) + nilaiDemandNonPareto(3) + nilaiDemandNonPareto(4) + nilaiDemandNonPareto(5) + nilaiDemandNonPareto(6) + nilaiDemandNonPareto(7) + nilaiDemandNonPareto(8))) * 100, 0) & " %"
1087          Else
1088              !tot29 = 0 & " %"
1089          End If
               
1090          !Urut = 7
1091          !judul = "Retention Rate Total (NCM & Non NCM)"
1092          !model = "NON PARETO"
1093          !north = 7
1094          .Update
1095      End With
          
          '' akhir urut 7
          '' ---------------------------------------------------------------
          
End Sub

Private Sub addTotalRowDealer()
1         With rsLaporan
2             .AddNew
3             !K1 = nilaiTotal1(0)
4             !K10 = nilaiTotal1(1)
5             !K20 = nilaiTotal1(2)
6             !K30 = nilaiTotal1(3)
7             !K40 = nilaiTotal1(4)
8             !K50 = nilaiTotal1(5)
9             !K60 = nilaiTotal1(6)
10            !K70 = nilaiTotal1(7)
11            !K80 = nilaiTotal1(8)
              
12            !tot16 = nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)
13            !tot19 = nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)
14            !tot26 = nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)
15            !tot29 = nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)
                       
16            !judul = "SB Same Outlet"
17            !model = "Total"
18            !Urut = 1
19            .Update
20        End With
          
21        With rsLaporan
22            .AddNew
23            !K1 = nilaiTotal2(0)
24            !K10 = nilaiTotal2(1)
25            !K20 = nilaiTotal2(2)
26            !K30 = nilaiTotal2(3)
27            !K40 = nilaiTotal2(4)
28            !K50 = nilaiTotal2(5)
29            !K60 = nilaiTotal2(6)
30            !K70 = nilaiTotal2(7)
31            !K80 = nilaiTotal2(8)
              
32            !tot16 = nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)
33            !tot19 = nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)
34            !tot26 = nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)
35            !tot29 = nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)
                       
36            !judul = "Demand SB Same Outlet"
37            !model = "Total"
38            !Urut = 2
39            .Update
40        End With
          
41        With rsLaporan
42            .AddNew
43            If nilaiTotal2(0) > 0 Then
44                !K1 = Round(nilaiTotal1(0) * 100 / nilaiTotal2(0), 0) & " %"
45            Else
46                !K1 = "0 %"
47            End If
              
48            If nilaiTotal2(1) > 0 Then
49                !K10 = Round(nilaiTotal1(1) * 100 / nilaiTotal2(1), 0) & " %"
50            Else
51                !K10 = "0 %"
52            End If
                      
53            If nilaiTotal2(2) > 0 Then
54                !K20 = Round(nilaiTotal1(2) * 100 / nilaiTotal2(2), 0) & " %"
55            Else
56                !K20 = "0 %"
57            End If
              
58            If nilaiTotal2(3) > 0 Then
59                !K30 = Round(nilaiTotal1(3) * 100 / nilaiTotal2(3), 0) & " %"
60            Else
61                !K30 = "0 %"
62            End If
                       
63            If nilaiTotal2(4) > 0 Then
64                !K40 = Round(nilaiTotal1(4) * 100 / nilaiTotal2(4), 0) & " %"
65            Else
66                !K40 = "0 %"
67            End If

68            If nilaiTotal2(5) > 0 Then
69                !K50 = Round(nilaiTotal1(5) * 100 / nilaiTotal2(5), 0) & " %"
70            Else
71                !K50 = "0 %"
72            End If
              
73            If nilaiTotal2(6) > 0 Then
74                !K60 = Round(nilaiTotal1(6) * 100 / nilaiTotal2(6), 0) & " %"
75            Else
76                !K60 = "0 %"
77            End If
              
78            If nilaiTotal2(7) > 0 Then
79                !K70 = Round(nilaiTotal1(7) * 100 / nilaiTotal2(7), 0) & " %"
80            Else
81                !K70 = "0 %"
82            End If
              
83            If nilaiTotal2(8) > 0 Then
84                !K80 = Round(nilaiTotal1(8) * 100 / nilaiTotal2(8), 0) & " %"
85            Else
86                !K80 = "0 %"
87            End If
              
88            If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
89                !tot16 = Round(((nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
90            Else
91                !tot16 = "0 %"
92            End If
              
93            If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
94                !tot19 = Round((nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
95            Else
96                !tot19 = "0 %"
97            End If
              
98            If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
99                !tot26 = Round((nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
100           Else
101               !tot26 = "0 %"
102           End If
              
103           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
104               !tot29 = Round((nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
105           Else
106               !tot29 = "0 %"
107           End If
                     
108           !Urut = 3
109           !judul = "Retention Rate Same Outlate"
110           !model = "Total"
111           .Update
112       End With
          
113       With rsLaporan
114           .AddNew
115           !K1 = nilaiTotal4(0)
116           !K10 = nilaiTotal4(1)
117           !K20 = nilaiTotal4(2)
118           !K30 = nilaiTotal4(3)
119           !K40 = nilaiTotal4(4)
120           !K50 = nilaiTotal4(5)
121           !K60 = nilaiTotal4(6)
122           !K70 = nilaiTotal4(7)
123           !K80 = nilaiTotal4(8)
              
124           !tot16 = nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)
125           !tot19 = nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)
126           !tot26 = nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)
127           !tot29 = nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)
                   
128           !judul = "SB Same Group (NCM)"
129           !model = "Total"
130           !Urut = 4
131           .Update
132       End With
          
133       With rsLaporan
134           .AddNew
135           If nilaiTotal2(0) > 0 Then
136               !K1 = Round(nilaiTotal4(0) * 100 / nilaiTotal2(0), 0) & " %"
137           Else
138               !K1 = "0 %"
139           End If
              
140           If nilaiTotal2(1) > 0 Then
141               !K10 = Round(nilaiTotal4(1) * 100 / nilaiTotal2(1), 0) & " %"
142           Else
143               !K10 = "0 %"
144           End If
                      
145           If nilaiTotal2(2) > 0 Then
146               !K20 = Round(nilaiTotal4(2) * 100 / nilaiTotal2(2), 0) & " %"
147           Else
148               !K20 = "0 %"
149           End If
              
150           If nilaiTotal2(3) > 0 Then
151               !K30 = Round(nilaiTotal4(3) * 100 / nilaiTotal2(3), 0) & " %"
152           Else
153               !K30 = "0 %"
154           End If
                       
155           If nilaiTotal2(4) > 0 Then
156               !K40 = Round(nilaiTotal4(4) * 100 / nilaiTotal2(4), 0) & " %"
157           Else
158               !K40 = "0 %"
159           End If

160           If nilaiTotal2(5) > 0 Then
161               !K50 = Round(nilaiTotal4(5) * 100 / nilaiTotal2(5), 0) & " %"
162           Else
163               !K50 = "0 %"
164           End If
              
165           If nilaiTotal2(6) > 0 Then
166               !K60 = Round(nilaiTotal4(6) * 100 / nilaiTotal2(6), 0) & " %"
167           Else
168               !K60 = "0 %"
169           End If
              
170           If nilaiTotal2(7) > 0 Then
171               !K70 = Round(nilaiTotal4(7) * 100 / nilaiTotal2(7), 0) & " %"
172           Else
173               !K70 = "0 %"
174           End If
              
175           If nilaiTotal2(8) > 0 Then
176               !K80 = Round(nilaiTotal4(8) * 100 / nilaiTotal2(8), 0) & " %"
177           Else
178               !K80 = "0 %"
179           End If
              
180           If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
181               !tot16 = Round(((nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
182           Else
183               !tot16 = "0 %"
184           End If
              
185           If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
186               !tot19 = Round((nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
187           Else
188               !tot19 = "0 %"
189           End If
              
190           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
191               !tot26 = Round((nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
192           Else
193               !tot26 = "0 %"
194           End If
              
195           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
196               !tot29 = Round((nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
197           Else
198               !tot29 = "0 %"
199           End If
              
200           !Urut = 5
201           !judul = "Retention Rate Same Group (NCM)"
202           !model = "Total"
203           .Update
204       End With
          
205       With rsLaporan
206           .AddNew
207           !K1 = nilaiTotal6(0)
208           !K10 = nilaiTotal6(1)
209           !K20 = nilaiTotal6(2)
210           !K30 = nilaiTotal6(3)
211           !K40 = nilaiTotal6(4)
212           !K50 = nilaiTotal6(5)
213           !K60 = nilaiTotal6(6)
214           !K70 = nilaiTotal6(7)
215           !K80 = nilaiTotal6(8)
              
216           !tot16 = nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)
217           !tot19 = nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)
218           !tot26 = nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)
219           !tot29 = nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)
                   
220           !judul = "SB Total (NCM & Non NCM)"
221           !model = "Total"
222           !Urut = 6
223           .Update
224       End With
          
225       With rsLaporan
226           .AddNew
227           If nilaiTotal2(0) > 0 Then
228               !K1 = Round(nilaiTotal6(0) * 100 / nilaiTotal2(0), 0) & " %"
229           Else
230               !K1 = "0 %"
231           End If
              
232           If nilaiTotal2(1) > 0 Then
233               !K10 = Round(nilaiTotal6(1) * 100 / nilaiTotal2(1), 0) & " %"
234           Else
235               !K10 = "0 %"
236           End If
                      
237           If nilaiTotal2(2) > 0 Then
238               !K20 = Round(nilaiTotal6(2) * 100 / nilaiTotal2(2), 0) & " %"
239           Else
240               !K20 = "0 %"
241           End If
              
242           If nilaiTotal2(3) > 0 Then
243               !K30 = Round(nilaiTotal6(3) * 100 / nilaiTotal2(3), 0) & " %"
244           Else
245               !K30 = "0 %"
246           End If
                       
247           If nilaiTotal2(4) > 0 Then
248               !K40 = Round(nilaiTotal6(4) * 100 / nilaiTotal2(4), 0) & " %"
249           Else
250               !K40 = "0 %"
251           End If

252           If nilaiTotal2(5) > 0 Then
253               !K50 = Round(nilaiTotal6(5) * 100 / nilaiTotal2(5), 0) & " %"
254           Else
255               !K50 = "0 %"
256           End If
              
257           If nilaiTotal2(6) > 0 Then
258               !K60 = Round(nilaiTotal6(6) * 100 / nilaiTotal2(6), 0) & " %"
259           Else
260               !K60 = "0 %"
261           End If
              
262           If nilaiTotal2(7) > 0 Then
263               !K70 = Round(nilaiTotal6(7) * 100 / nilaiTotal2(7), 0) & " %"
264           Else
265               !K70 = "0 %"
266           End If
              
267           If nilaiTotal2(8) > 0 Then
268               !K80 = Round(nilaiTotal6(8) * 100 / nilaiTotal2(8), 0) & " %"
269           Else
270               !K80 = "0 %"
271           End If
              
272           If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
273               !tot16 = Round(((nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
274           Else
275               !tot16 = "0 %"
276           End If
              
277           If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
278               !tot19 = Round((nilaiTotal6(0) + nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
279           Else
280               !tot19 = "0 %"
281           End If
              
282           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
283               !tot26 = Round((nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
284           Else
285               !tot26 = "0 %"
286           End If
              
287           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
288               !tot29 = Round((nilaiTotal6(1) + nilaiTotal6(2) + nilaiTotal6(3) + nilaiTotal6(4) + nilaiTotal6(5) + nilaiTotal6(6) + nilaiTotal6(7) + nilaiTotal6(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
289           Else
290               !tot29 = "0 %"
291           End If
292           !Urut = 7
293           !judul = "Retention Rate Total (NCM & Non NCM)"
294           !model = "Total"
295           .Update
296       End With
End Sub


Private Sub addTotalRow()
1         With rsLaporan
2             .AddNew
3             !K1 = nilaiTotal1(0)
4             !K10 = nilaiTotal1(1)
5             !K20 = nilaiTotal1(2)
6             !K30 = nilaiTotal1(3)
7             !K40 = nilaiTotal1(4)
8             !K50 = nilaiTotal1(5)
9             !K60 = nilaiTotal1(6)
10            !K70 = nilaiTotal1(7)
11            !K80 = nilaiTotal1(8)
              
12            !tot16 = nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)
13            !tot19 = nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)
14            !tot26 = nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)
15            !tot29 = nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)
                       
16            !judul = "SB Same Group (NCM)"
17            !model = "Total"
18            !Urut = 1
19            .Update
20        End With
          
21        With rsLaporan
22            .AddNew
23            !K1 = nilaiTotal2(0)
24            !K10 = nilaiTotal2(1)
25            !K20 = nilaiTotal2(2)
26            !K30 = nilaiTotal2(3)
27            !K40 = nilaiTotal2(4)
28            !K50 = nilaiTotal2(5)
29            !K60 = nilaiTotal2(6)
30            !K70 = nilaiTotal2(7)
31            !K80 = nilaiTotal2(8)
              
32            !tot16 = nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)
33            !tot19 = nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)
34            !tot26 = nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)
35            !tot29 = nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)
                       
36            !judul = "Demand SB Same Group (NCM)"
37            !model = "Total"
38            !Urut = 2
39            .Update
40        End With
          
41        With rsLaporan
42            .AddNew
43            If nilaiTotal2(0) > 0 Then
44                !K1 = Round(nilaiTotal1(0) * 100 / nilaiTotal2(0), 0) & " %"
45            Else
46                !K1 = "0 %"
47            End If
              
48            If nilaiTotal2(1) > 0 Then
49                !K10 = Round(nilaiTotal1(1) * 100 / nilaiTotal2(1), 0) & " %"
50            Else
51                !K10 = "0 %"
52            End If
                      
53            If nilaiTotal2(2) > 0 Then
54                !K20 = Round(nilaiTotal1(2) * 100 / nilaiTotal2(2), 0) & " %"
55            Else
56                !K20 = "0 %"
57            End If
              
58            If nilaiTotal2(3) > 0 Then
59                !K30 = Round(nilaiTotal1(3) * 100 / nilaiTotal2(3), 0) & " %"
60            Else
61                !K30 = "0 %"
62            End If
                       
63            If nilaiTotal2(4) > 0 Then
64                !K40 = Round(nilaiTotal1(4) * 100 / nilaiTotal2(4), 0) & " %"
65            Else
66                !K40 = "0 %"
67            End If

68            If nilaiTotal2(5) > 0 Then
69                !K50 = Round(nilaiTotal1(5) * 100 / nilaiTotal2(5), 0) & " %"
70            Else
71                !K50 = "0 %"
72            End If
              
73            If nilaiTotal2(6) > 0 Then
74                !K60 = Round(nilaiTotal1(6) * 100 / nilaiTotal2(6), 0) & " %"
75            Else
76                !K60 = "0 %"
77            End If
              
78            If nilaiTotal2(7) > 0 Then
79                !K70 = Round(nilaiTotal1(7) * 100 / nilaiTotal2(7), 0) & " %"
80            Else
81                !K70 = "0 %"
82            End If
              
83            If nilaiTotal2(8) > 0 Then
84                !K80 = Round(nilaiTotal1(8) * 100 / nilaiTotal2(8), 0) & " %"
85            Else
86                !K80 = "0 %"
87            End If
              
88            If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
89                !tot16 = Round(((nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
90            Else
91                !tot16 = "0 %"
92            End If
              
93            If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
94                !tot19 = Round((nilaiTotal1(0) + nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
95            Else
96                !tot19 = "0 %"
97            End If
              
98            If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
99                !tot26 = Round((nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
100           Else
101               !tot26 = "0 %"
102           End If
              
103           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
104               !tot29 = Round((nilaiTotal1(1) + nilaiTotal1(2) + nilaiTotal1(3) + nilaiTotal1(4) + nilaiTotal1(5) + nilaiTotal1(6) + nilaiTotal1(7) + nilaiTotal1(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
105           Else
106               !tot29 = "0 %"
107           End If
                     
108           !Urut = 3
109           !judul = "Retention Rate Same Group (NCM)"
110           !model = "Total"
111           .Update
112       End With
           
113       With rsLaporan
114           .AddNew
115           !K1 = nilaiTotal4(0)
116           !K10 = nilaiTotal4(1)
117           !K20 = nilaiTotal4(2)
118           !K30 = nilaiTotal4(3)
119           !K40 = nilaiTotal4(4)
120           !K50 = nilaiTotal4(5)
121           !K60 = nilaiTotal4(6)
122           !K70 = nilaiTotal4(7)
123           !K80 = nilaiTotal4(8)
              
124           !tot16 = nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)
125           !tot19 = nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)
126           !tot26 = nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)
127           !tot29 = nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)
                   
128           !judul = "SB Total (NCM & Non NCM)"
129           !model = "Total"
130           !Urut = 4
131           .Update
132       End With
          
133       With rsLaporan
134           .AddNew
135           If nilaiTotal2(0) > 0 Then
136               !K1 = Round(nilaiTotal4(0) * 100 / nilaiTotal2(0), 0) & " %"
137           Else
138               !K1 = "0 %"
139           End If
              
140           If nilaiTotal2(1) > 0 Then
141               !K10 = Round(nilaiTotal4(1) * 100 / nilaiTotal2(1), 0) & " %"
142           Else
143               !K10 = "0 %"
144           End If
                      
145           If nilaiTotal2(2) > 0 Then
146               !K20 = Round(nilaiTotal4(2) * 100 / nilaiTotal2(2), 0) & " %"
147           Else
148               !K20 = "0 %"
149           End If
              
150           If nilaiTotal2(3) > 0 Then
151               !K30 = Round(nilaiTotal4(3) * 100 / nilaiTotal2(3), 0) & " %"
152           Else
153               !K30 = "0 %"
154           End If
                       
155           If nilaiTotal2(4) > 0 Then
156               !K40 = Round(nilaiTotal4(4) * 100 / nilaiTotal2(4), 0) & " %"
157           Else
158               !K40 = "0 %"
159           End If

160           If nilaiTotal2(5) > 0 Then
161               !K50 = Round(nilaiTotal4(5) * 100 / nilaiTotal2(5), 0) & " %"
162           Else
163               !K50 = "0 %"
164           End If
              
165           If nilaiTotal2(6) > 0 Then
166               !K60 = Round(nilaiTotal4(6) * 100 / nilaiTotal2(6), 0) & " %"
167           Else
168               !K60 = "0 %"
169           End If
              
170           If nilaiTotal2(7) > 0 Then
171               !K70 = Round(nilaiTotal4(7) * 100 / nilaiTotal2(7), 0) & " %"
172           Else
173               !K70 = "0 %"
174           End If
              
175           If nilaiTotal2(8) > 0 Then
176               !K80 = Round(nilaiTotal4(8) * 100 / nilaiTotal2(8), 0) & " %"
177           Else
178               !K80 = "0 %"
179           End If
              
180           If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
181               !tot16 = Round(((nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)) * 100) / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
182           Else
183               !tot16 = "0 %"
184           End If
              
185           If nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
186               !tot19 = Round((nilaiTotal4(0) + nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)) * 100 / (nilaiTotal2(0) + nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
187           Else
188               !tot19 = "0 %"
189           End If
              
190           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) > 0 Then
191               !tot26 = Round((nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5)), 0) & " %"
192           Else
193               !tot26 = "0 %"
194           End If
              
195           If nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8) > 0 Then
196               !tot29 = Round((nilaiTotal4(1) + nilaiTotal4(2) + nilaiTotal4(3) + nilaiTotal4(4) + nilaiTotal4(5) + nilaiTotal4(6) + nilaiTotal4(7) + nilaiTotal4(8)) * 100 / (nilaiTotal2(1) + nilaiTotal2(2) + nilaiTotal2(3) + nilaiTotal2(4) + nilaiTotal2(5) + nilaiTotal2(6) + nilaiTotal2(7) + nilaiTotal2(8)), 0) & " %"
197           Else
198               !tot29 = "0 %"
199           End If
200           !Urut = 5
201           !judul = "Retention Rate Total (NCM & Non NCM)"
202           !model = "Total"
203           .Update
204       End With
End Sub
 

Private Sub GridDealer_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        If rsDealer.State Then
            rsDealer.MoveFirst
            Do While Not rsDealer.EOF
                rsDealer!terlibat = IIf(rsDealer!terlibat = True, False, True)
                rsDealer.MoveNext
            Loop
            rsDealer.MoveFirst
        End If
    End If
End Sub
