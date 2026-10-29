VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmValidasiTWC 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Validasi Warranty"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   12525
   Begin VB.CommandButton cmdThn 
      Caption         =   ".."
      Height          =   375
      Left            =   -120
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   255
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
      Height          =   270
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   10080
      Width           =   1050
   End
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
      Height          =   270
      Left            =   5160
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   55
      Top             =   10080
      Width           =   1395
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   2175
      Left            =   4200
      TabIndex        =   54
      Top             =   10440
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
   Begin VB.PictureBox frmTidakSetuju 
      BackColor       =   &H00E0E0E0&
      Height          =   2175
      Left            =   3240
      ScaleHeight     =   2115
      ScaleWidth      =   5595
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox Text1 
         Height          =   855
         Index           =   5
         Left            =   120
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   49
         Top             =   480
         Width           =   5355
      End
      Begin ucXPButton.XPButton cmdOK 
         Height          =   495
         Left            =   1560
         TabIndex        =   50
         Top             =   1440
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&O"
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
         MICON           =   "frmValidasiTwc.frx":0000
         PICN            =   "frmValidasiTwc.frx":001C
         PICH            =   "frmValidasiTwc.frx":04BD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdCancel 
         Height          =   495
         Left            =   2760
         TabIndex        =   51
         Top             =   1440
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&c"
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
         MICON           =   "frmValidasiTwc.frx":09D7
         PICN            =   "frmValidasiTwc.frx":09F3
         PICH            =   "frmValidasiTwc.frx":0F37
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: Alasan Tidak Disetujui  ::."
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
         Left            =   1290
         TabIndex        =   53
         Top             =   240
         Width           =   3045
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   5355
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   5520
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   5520
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Label Ket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: Alasan Pembatalan  ::."
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
         Left            =   1440
         TabIndex        =   52
         Top             =   165
         Width           =   2745
      End
   End
   Begin VB.Frame frmeDetail 
      BackColor       =   &H00FFFFFF&
      Height          =   9615
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Visible         =   0   'False
      Width           =   10575
      Begin VB.CheckBox chPWC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "PWC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9240
         TabIndex        =   58
         Top             =   720
         Width           =   1215
      End
      Begin TrueOleDBGrid70.TDBGrid partGridFill 
         Height          =   2025
         Left            =   120
         TabIndex        =   57
         Top             =   6120
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   3572
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
         Columns(3).Caption=   "INFO"
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
         Splits(0)._ColumnProps(29)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Index           =   0
         Left            =   120
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   47
         Top             =   1680
         Width           =   5025
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Index           =   1
         Left            =   5280
         MaxLength       =   500
         MultiLine       =   -1  'True
         TabIndex        =   46
         Top             =   1680
         Width           =   5175
      End
      Begin VB.CommandButton TblBatal 
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
         Height          =   450
         Left            =   120
         Picture         =   "frmValidasiTwc.frx":152E
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   8880
         Visible         =   0   'False
         Width           =   1850
      End
      Begin VB.CommandButton tblSetuju 
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
         Height          =   450
         Left            =   120
         Picture         =   "frmValidasiTwc.frx":1C8B
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   8280
         Width           =   1850
      End
      Begin VB.CommandButton TblTdkSetuju 
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
         Height          =   450
         Left            =   2040
         Picture         =   "frmValidasiTwc.frx":23BB
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   8280
         Width           =   1850
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   23
         Top             =   195
         Width           =   1275
         _ExtentX        =   2249
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   1320
         TabIndex        =   24
         Top             =   480
         Width           =   3435
         _ExtentX        =   6059
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   5880
         TabIndex        =   25
         Top             =   765
         Width           =   1875
         _ExtentX        =   3307
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   1320
         TabIndex        =   26
         Top             =   765
         Width           =   2955
         _ExtentX        =   5212
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   3360
         TabIndex        =   27
         Top             =   195
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   1080
         TabIndex        =   28
         Top             =   5760
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
         Left            =   4020
         TabIndex        =   29
         Top             =   5760
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   9
         Left            =   4125
         TabIndex        =   30
         Top             =   5760
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
      Begin MSMask.MaskEdBox text 
         Height          =   345
         Index           =   7
         Left            =   8400
         TabIndex        =   31
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
      Begin TrueOleDBGrid70.TDBGrid opennoGrid 
         Height          =   4560
         Left            =   7635
         TabIndex        =   32
         Top             =   3600
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   8043
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Openno"
         Columns(0).DataField=   "OLI"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
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
         DeadAreaBackColor=   12577007
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid gridJobOrder 
         Height          =   2280
         Left            =   120
         TabIndex        =   33
         Top             =   3240
         Width           =   7380
         _ExtentX        =   13018
         _ExtentY        =   4022
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Openno"
         Columns(0).DataField=   "OLI"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   ""
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
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=2725"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=2646"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
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
         DeadAreaBackColor=   12577007
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=58,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=55,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=56,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=57,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
      Begin TrueOleDBGrid70.TDBGrid partGrid 
         Height          =   2055
         Left            =   120
         TabIndex        =   34
         Top             =   6120
         Visible         =   0   'False
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   3625
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
         Columns(2).Caption=   "NAMA BY PIC"
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
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(18)=   "Column(3)._ColStyle=8196"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=1244"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1164"
         Splits(0)._ColumnProps(23)=   "Column(4)._ColStyle=8196"
         Splits(0)._ColumnProps(24)=   "Column(4).Order=5"
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
      Begin ucXPButton.XPButton TblTutup 
         Height          =   450
         Left            =   9240
         TabIndex        =   35
         Top             =   8280
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   794
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
         MICON           =   "frmValidasiTwc.frx":2BE5
         PICN            =   "frmValidasiTwc.frx":2C01
         PICH            =   "frmValidasiTwc.frx":30D5
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton tblSimpan 
         Height          =   450
         Left            =   8040
         TabIndex        =   36
         Top             =   8280
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   794
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
         MICON           =   "frmValidasiTwc.frx":3620
         PICN            =   "frmValidasiTwc.frx":363C
         PICH            =   "frmValidasiTwc.frx":3B4F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton tblAdd 
         Height          =   330
         Left            =   9960
         TabIndex        =   37
         Top             =   3240
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmValidasiTwc.frx":40F1
         PICN            =   "frmValidasiTwc.frx":410D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   10
         Left            =   5880
         TabIndex        =   60
         Top             =   195
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   11
         Left            =   5040
         TabIndex        =   62
         Top             =   480
         Width           =   2715
         _ExtentX        =   4789
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
      Begin VB.Shape Shape1 
         Height          =   255
         Left            =   7700
         Top             =   195
         Width           =   55
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode SSC"
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
         Left            =   5070
         TabIndex        =   61
         Top             =   240
         Width           =   690
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   5280
         Picture         =   "frmValidasiTwc.frx":4454
         Top             =   1200
         Width           =   12600
      End
      Begin VB.Image Image5 
         Height          =   390
         Left            =   120
         Picture         =   "frmValidasiTwc.frx":5949
         Top             =   1200
         Width           =   4380
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   120
         X2              =   10440
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Prediagnosis"
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
         Left            =   0
         TabIndex        =   45
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label4 
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
         Left            =   600
         TabIndex        =   44
         Top             =   525
         Width           =   690
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   5175
         TabIndex        =   43
         Top             =   810
         Width           =   585
      End
      Begin VB.Label Label6 
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
         Left            =   360
         TabIndex        =   42
         Top             =   810
         Width           =   840
      End
      Begin VB.Label Label7 
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
         Left            =   2715
         TabIndex        =   41
         Top             =   210
         Width           =   510
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   120
         X2              =   10440
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         BorderWidth     =   3
         X1              =   120
         X2              =   7560
         Y1              =   5640
         Y2              =   5640
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
         Left            =   240
         TabIndex        =   40
         Top             =   5805
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Open No"
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
         Left            =   7680
         TabIndex        =   39
         Top             =   3285
         Width           =   630
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RETURN JOB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   8520
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   10440
         Y1              =   8235
         Y2              =   8235
      End
   End
   Begin VB.CommandButton cmdRepo 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Report"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      Picture         =   "frmValidasiTwc.frx":68A7
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   600
      Width           =   810
   End
   Begin VB.CommandButton tblSearch 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cari"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      Picture         =   "frmValidasiTwc.frx":8DD8
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   600
      Width           =   810
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9360
      Width           =   2130
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   6720
      TabIndex        =   4
      Top             =   480
      Width           =   3975
      Begin VB.CheckBox chView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tunggu Validasi TL"
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
         Index           =   6
         Left            =   120
         TabIndex        =   59
         Top             =   840
         Width           =   2055
      End
      Begin VB.CheckBox chView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sudah di Invoice"
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
         Index           =   5
         Left            =   2280
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox chView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sudah jadi WO"
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
         Index           =   4
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
      Begin VB.CheckBox chView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tidak disetujui"
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
         Index           =   3
         Left            =   2280
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.CheckBox chView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sudah di Validasi"
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
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox chView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Belum di Validasi"
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
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chView 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Prediagnosis di Batalkan"
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
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   2175
      End
   End
   Begin TrueOleDBGrid70.TDBGrid DiagGrid 
      Height          =   7695
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   13573
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
      Columns(4).NumberFormat=   "Standard"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6).NumberFormat=   "General Number"
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
      Columns(9).NumberFormat=   "Standard"
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
      Columns.Count   =   13
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=13"
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
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=8196"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=8196"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
      Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
      Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=8196"
      Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(36)=   "Column(7).Width=2725"
      Splits(0)._ColumnProps(37)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(38)=   "Column(7)._WidthInPix=2646"
      Splits(0)._ColumnProps(39)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(40)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(41)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(42)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(43)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(44)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(45)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(47)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(48)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(49)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(50)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(51)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(52)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(53)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(54)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(55)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(56)=   "Column(12).Width=2725"
      Splits(0)._ColumnProps(57)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(12)._WidthInPix=2646"
      Splits(0)._ColumnProps(59)=   "Column(12).Order=13"
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
      Caption         =   "Prediagnosis"
      MultipleLines   =   0
      CellTipsWidth   =   0
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
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13,.locked=-1"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=66,.parent=13,.locked=-1"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=63,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=64,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=65,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13,.locked=-1"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13,.locked=-1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.locked=-1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13,.locked=-1"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13,.locked=-1"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=82,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=79,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=80,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=81,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=74,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=71,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=72,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=73,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=13,.locked=0"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=29,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=30,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=31,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=86,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=83,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=84,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=85,.parent=17"
      _StyleDefs(88)  =   "Named:id=33:Normal"
      _StyleDefs(89)  =   ":id=33,.parent=0"
      _StyleDefs(90)  =   "Named:id=34:Heading"
      _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(92)  =   ":id=34,.wraptext=-1"
      _StyleDefs(93)  =   "Named:id=35:Footing"
      _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(95)  =   "Named:id=36:Selected"
      _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=37:Caption"
      _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(99)  =   "Named:id=38:HighlightRow"
      _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(101) =   "Named:id=39:EvenRow"
      _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(103) =   "Named:id=40:OddRow"
      _StyleDefs(104) =   ":id=40,.parent=33"
      _StyleDefs(105) =   "Named:id=41:RecordSelector"
      _StyleDefs(106) =   ":id=41,.parent=34"
      _StyleDefs(107) =   "Named:id=42:FilterBar"
      _StyleDefs(108) =   ":id=42,.parent=33"
   End
   Begin MSComCtl2.DTPicker tanggal1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   357695489
      CurrentDate     =   39859
   End
   Begin MSComCtl2.DTPicker tanggal2 
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      Format          =   357695489
      CurrentDate     =   39859
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   1170
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   503
      _Version        =   393216
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
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   270
      Left            =   1635
      TabIndex        =   11
      Top             =   1170
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   1
      Left            =   1755
      TabIndex        =   12
      Top             =   1170
      Width           =   2955
      _ExtentX        =   5212
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
   Begin AIFCmp1.asxBars asxBars2 
      Height          =   375
      Left            =   0
      Top             =   11160
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   661
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   120
      Top             =   0
      Width           =   12735
      _ExtentX        =   22463
      _ExtentY        =   661
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S/d"
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
      Left            =   2640
      TabIndex        =   15
      Top             =   690
      Width           =   240
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
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
      Left            =   120
      TabIndex        =   14
      Top             =   690
      Width           =   570
   End
   Begin VB.Label Label3 
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
      Left            =   495
      TabIndex        =   13
      Top             =   1215
      Width           =   195
   End
End
Attribute VB_Name = "frmValidasiTWC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsGrid As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns

Dim rsJob As New ADODB.Recordset
Dim rsOpenno As New ADODB.Recordset
Dim rsGridPart As New ADODB.Recordset
Dim Report As Object
'- tambahan
Dim rsPermintaan As New ADODB.Recordset
Dim rsKeluhan As New ADODB.Recordset
'---------

Dim sFilter As String
Dim sFilterReport As String
Dim hNamaPart As String
Dim hSatPart As String
Dim nStok As Integer
Dim openBaris As Integer
Private isBySA As Boolean
Private NoFak As String
Private vNamaTabel As String

Private Sub createTblJob(stat As Boolean)
    
    If stat Then
        If rsJob.State Then rsJob.Close
        Set rsJob = Nothing
        With rsJob
            .Fields.Append "GROUP", adChar, 5
            .Fields.Append "JOB", adChar, 50
            .Fields.Append "RATE", adDouble
            .Open
        End With
        
        Set GridJobOrder.DataSource = rsJob
        Call headColumns
    End If
End Sub

Private Sub headColumns()
    
    With GridJobOrder
        .Columns(0).Caption = "GROUP"
        .Columns(0).Width = 800
        .Columns(0).DataField = "GROUP"
        .Columns(0).Locked = True
        .Columns(1).Caption = "JOB ORDER"
        .Columns(1).Width = 4000
        .Columns(1).DataField = "JOB"
        .Columns(1).Locked = True
        .Columns(2).Caption = "RATE"
        .Columns(2).Width = 800
        .Columns(2).DataField = "RATE"
        .Columns(2).Locked = True
    End With
End Sub

Private Sub hapusForm(clearAll As Boolean)
    If clearAll Then
        Text(0) = ""
        Text(1) = ""
    End If
    For i = 2 To 9
        Select Case i
            Case 9
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    Set rsJob = Nothing
    Set rsOpenno = Nothing
    Set rsGridPart = Nothing
    TblBatal.Enabled = True
    tblSetuju.Enabled = True
    TblTdkSetuju.Enabled = True
    chPWC.value = False
End Sub

Private Sub chview_report()
 sFilter = ""
    sFilterReport = ""
    '' 20190701REV1
    If Log.lockSCC = 1 And Log.newLockSSC = 1 Then
        For i = 0 To 6
            If chView(i).value = Checked Then
                Select Case i
                    Case 0
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " or ", "") & "deliv = 'BTL'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " or ", "") & "dta.deliv = 'BTL'"
                    Case 1
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = '' AND FLAGTL = 0"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = '' AND dta.FLAGTL = 0"
                    Case 2
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = 'OKE'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = 'OKE'"
                    Case 3
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = 'NOK'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = 'NOK'"
                    Case 4
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "NO_WO <> ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.NO_WO <> ''"
                    Case 5
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "NOINV <> ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.NOINV <> ''"
                    Case 6
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = '' AND FLAGTL = 1"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = '' AND dta.FLAGTL = 1"
                End Select
            End If
        Next i
    Else
        For i = 0 To 5
            If chView(i).value = Checked Then
                Select Case i
                    Case 0
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " or ", "") & "deliv = 'BTL'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " or ", "") & "dta.deliv = 'BTL'"
                    Case 1
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = ''"
                    Case 2
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = 'OKE'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = 'OKE'"
                    Case 3
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = 'NOK'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = 'NOK'"
                    Case 4
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "NO_WO <> ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.NO_WO <> ''"
                    Case 5
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "NOINV <> ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.NOINV <> ''"
                End Select
            End If
        Next i
    End If
    
    If sFilter <> "" Then
      sFilter = "(" & sFilter & ")"
    End If
    If sFilterReport <> "" Then
      sFilterReport = "(" & sFilterReport & ")"
    End If
End Sub

Private Sub chView_Click(Index As Integer)
    
    sFilter = ""
    sFilterReport = ""
    '' 20190701REV1
    If Log.lockSCC = 1 And Log.newLockSSC = 1 Then
        For i = 0 To 6
            If chView(i).value = Checked Then
                Select Case i
                    Case 0
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " or ", "") & "deliv = 'BTL'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " or ", "") & "dta.deliv = 'BTL'"
                    Case 1
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = '' AND FLAGTL = 0"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = ''"
                    Case 2
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = 'OKE'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = 'OKE'"
                    Case 3
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = 'NOK'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = 'NOK'"
                    Case 4
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "NO_WO <> ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.NO_WO <> ''"
                    Case 5
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "NOINV <> ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.NOINV <> ''"
                    Case 6
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = '' AND FLAGTL = 1"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = ''"
                End Select
            End If
        Next i
    Else
        For i = 0 To 5
            If chView(i).value = Checked Then
                Select Case i
                    Case 0
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " or ", "") & "deliv = 'BTL'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " or ", "") & "dta.deliv = 'BTL'"
                    Case 1
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = ''"
                    Case 2
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = 'OKE'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = 'OKE'"
                    Case 3
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "VALIDTWC = 'NOK'"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.VALIDTWC = 'NOK'"
                    Case 4
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "NO_WO <> ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.NO_WO <> ''"
                    Case 5
                        sFilter = sFilter + IIf(Trim(sFilter) <> "", " OR ", "") & "NOINV <> ''"
                        sFilterReport = sFilterReport + IIf(Trim(sFilterReport) <> "", " OR ", "") & "dta.NOINV <> ''"
                End Select
            End If
        Next i
    End If
    If sFilter <> "" Then
      sFilter = "(" & sFilter & ")"
    End If
    If sFilterReport <> "" Then
      sFilterReport = "(" & sFilterReport & ")"
    End If
        Call refreshGrid(tanggal1, tanggal2, Text(0), sFilter)
    
End Sub

Private Sub cmdCancel_Click()
    frmTidakSetuju.Visible = False
End Sub

Private Sub cmdOK_Click()
    ''20190711 | tidak semua warranty ada SSC
      '    If NulSID(text(10)) = "" Then
      '        MsgBox "Id Warranty harus diisi.", vbCritical
      '        Exit Sub
      '    Else
      '        Set rsSearch = objO.FetchData("SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH (NOLOCK) WHERE id_WARRANTY = '" & NulSID(text(10)) & "' and (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)")
      '        If rsSearch.EOF Then
      '            MsgBox "Id Warranty tidak ditemukan.", vbCritical
      '            Exit Sub
      '        End If
      '    End If
1         If Len(Trim(Text1(5))) > 5 Then
2             If Not opennoGrid.EOF Then
3                 opennoGrid.MoveFirst
4                 Call objO.UpdateData("DELETE DIAGAPPROVAL WHERE NO_DIAG = '" & Trim(Text(2)) & "'")
5                 Do While Not rsOpenno.EOF
6                     Call objO.UpdateData("INSERT INTO DIAGAPPROVAL VALUES ('" & Trim(Text(2)) & "','" & Trim(Text(6)) & "','" & Trim(rsOpenno("OPENNO")) & "','" & Trim(gSA) & "', '" & Format(isnow, "MM/DD/YYYY HH:NN") & "','" & Trim(Text1(5)) & "','" & NoFak & "','" & IIf(chPWC.value, 1, 0) & "')")
7                     rsOpenno.MoveNext
8                 Loop
9             End If
      '        If rsOpenno.State <> adStateClosed Then
      '            rsOpenno.MoveFirst
      '            If Not rsOpenno.EOF Then
      '                Call objO.UpdateData("DELETE DIAGAPPROVAL WHERE NO_DIAG = '" & Trim(text(2)) & "'")
      '                Do While Not rsOpenno.EOF
      '                    Call objO.UpdateData("INSERT INTO DIAGAPPROVAL VALUES ('" & Trim(text(2)) & "','" & Trim(text(6)) & "','" & Trim(rsOpenno("OPENNO")) & "','" & Trim(gSA) & "', '" & Format(isnow, "MM/DD/YYYY HH:NN") & "','" & Trim(Text1(5)) & "','" & nofak & "')")
      '                    rsOpenno.MoveNext
      '                Loop
      '            End If
      '        End If
10            Call objO.UpdateData("UPDATE DIAGDTA SET VALIDTWC = 'NOK' WHERE NO_DIAG = '" & Trim(Text(2)) & "'")
11            TblBatal.Enabled = False
12            tblSetuju.Enabled = False
13            frmTidakSetuju.Visible = False
14        Else
15            MsgBox "Alasan tidak lengkap. Harap melengkapi Alasan tidak menyetujui prediagnosis", vbCritical
16        End If
End Sub

Private Sub cmdRepo_Click()
'    MsgBox "Under Construction !!! ", vbCritical
    
    Call chview_report
    Call IsiTempRepo
End Sub

Private Function createTemp(vNamaTabel As String) As Boolean
Dim rsHlp As ADODB.Recordset
    createTemp = False
    Set rsHlp = objO.FetchData("select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ")
    If Not rsHlp.EOF Then
        createTemp = True
    Else
        Call BuatTableValidasiTWC(vNamaTabel)
        createTemp = True
    End If
End Function

Private Sub IsiTempRepo()
Dim hlp As String
Dim hlp2 As String
Dim hlp3 As String
Dim hlp4 As String
Dim hlp5 As String
    
    'ambil data
    'Set rsTemp = objO.FetchData("SELECT dta.no_diag,dta.no_polisi,dta.mdl + ' ' + dta.fmdl as mdl,dta.no_chassis, dta.[date],dp.tglapprove,dtl.cmpl1 + ' , ' + dtl.cmpl2 + ' , ' + dtl.cmpl3 + ' , ' + dtl.cmpl4 + ' , ' + dtl.cmpl5 as cmpl, dta.no_wo, dp.alasan_tdksetuju FROM DIAGDTA dta LEFT OUTER JOIN DIAGAPPROVAL dp ON dp.no_diag = dta.no_diag LEFT OUTER JOIN DIAGDTL dtl ON dtl.no_diag = dta.no_diag WHERE dta.[DATE] BETWEEN '" & Format(tanggal1, "MM/DD/YYYY") & "' AND '" & Format(tanggal2, "MM/DD/YYYY") & "' " & IIf(Trim(sFilterReport) <> "", " AND " & sFilterReport, ""))
    Set rsTemp = objO.FetchData("SELECT dta.no_diag,dta.no_polisi,dta.mdl + ' ' + dta.fmdl as mdl,dta.no_chassis, dta.[date],dp.tglapprove,dtl.cmpl1 + CHAR(13) + CHAR(10)  + dtl.cmpl2 + CHAR(13) + CHAR(10)  + dtl.cmpl3 + CHAR(13) + CHAR(10)  + dtl.cmpl4 + CHAR(13) + CHAR(10)  + dtl.cmpl5 as cmpl, dta.no_wo, dp.alasan_tdksetuju FROM DIAGDTA dta WITH(NOLOCK) LEFT OUTER JOIN DIAGAPPROVAL dp WITH(NOLOCK) ON dp.no_diag = dta.no_diag LEFT OUTER JOIN DIAGDTL dtl WITH(NOLOCK) ON dtl.no_diag = dta.no_diag WHERE dta.[DATE] BETWEEN '" & Format(tanggal1, "MM/DD/YYYY") & "' AND '" & Format(tanggal2, "MM/DD/YYYY") & "' " & IIf(Trim(sFilterReport) <> "", " AND " & sFilterReport, ""))
    
    If Not rsTemp.EOF Then
        Call BuatTableValidasiTWC(vNamaTabel)
        Do While Not rsTemp.EOF
            'lakukan pengisian data
            If createTemp(vNamaTabel) Then
                Call objO.InsertData("INSERT INTO [TEMPDB]..[" & vNamaTabel & "] VALUES ('" & rsTemp("no_diag") & "', '" & rsTemp("no_polisi") & "', '" & rsTemp("mdl") & "', " _
                             & "'" & rsTemp("no_chassis") & "', '" & Format(rsTemp("DATE"), "MM/DD/YYYY") & "', '" & Format(rsTemp("tglapprove"), "MM/DD/YYYY") & "', '" & Trim(rsTemp("cmpl")) & "','" & rsTemp("no_wo") & "', '" & IIf(rsTemp("alasan_tdksetuju") <> "", "Ya", "Tidak") & "', '" & rsTemp("alasan_tdksetuju") & "','')")
            End If
            rsTemp.MoveNext
        Loop
        Call Cetak(vNamaTabel)
    Else
        MsgBox "Tidak ada data !!"
        Exit Sub
    End If
End Sub

Private Sub Cetak(vNamaTabel As String)
Dim rsRepo As Object
Screen.MousePointer = vbHourglass
    
    Set Report = New RptRepoValidasiTWSC
    If txt <> "" Then
        txt = " WHERE NO_WO IN (SELECT DISTINCT NO_WO FROM [TEMPDB]..[" & vNamaTabel & "] WHERE " & txt & ")"
    End If
    sSQL = "SELECT * FROM [TEMPDB]..[" & vNamaTabel & "]" & txt & " ORDER BY NO_DIAG DESC"
    Set rsRepo = objO.FetchData(sSQL)
    Report.database.SetDataSource rsRepo
    Report.txtPerusahaan.SetText gPerush
    Report.txtAlamat.SetText gAlamat
  '  Report.txtShift.SetText IIf(opLap(0), "", "SHIFT : " & Trim(text(16)) & " ~ Jam " & Trim(text(17)) & " - " & Trim(text(18)))
    Report.txtTanggal.SetText "TANGGAL : " & Format(tanggal1, "DD-MM-YYYY") & " s/d " & Format(tanggal2, "DD-MM-YYYY")
    
    With CRViewer1
        .ReportSource = Report
        .ViewReport
        .Zoom 100
        .Visible = True
    End With
    cmdTutup.Visible = True
    cmdSettingPrinter.Visible = True
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSettingPrinter_Click()
    Report.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdThn_Click()
    If Trim(Text(2)) <> "" Then
        sSQL = "UPDATE DIAGDTA SET KODETIPE = '" & NulSID(Text(10)) & "' WHERE NO_DIAG = '" & Trim(Text(2)) & "'"
        Call objO.InsertData(sSQL)
        
    Else
        MsgBox "No Prediagnose kosong", vbInformation
    End If
End Sub

Private Sub cmdTutup_Click()
    sSQL = "if exists (select * from [tempdb]..sysobjects where id = object_id(N'[tempdb]..[" & vNamaTabel & "]')  ) " _
         & " drop table [tempdb]..[" & vNamaTabel & "]"
    Call objO.DeleteData(sSQL)
    Set Report = Nothing
    
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DiagGrid_DblClick()
      Dim rsSearch As ADODB.Recordset
      Dim tmp As String
      Dim tmp2 As String
      Dim tmp3 As String

1         If rsGrid.State <> adStateClosed Then
2             If Not rsGrid.EOF Then
3                 Call hapusForm(False)
4                 TblBatal.Visible = False
5                 tblSetuju.Visible = True
6                 TblTdkSetuju.Visible = True
7                 lblStatus.Visible = False
8                 Text(2) = Trim(rsGrid.Fields(2))
9                 Text(3) = Trim(rsGrid.Fields(7))
10                Text(4) = Trim(rsGrid.Fields(11))
11                Text(5) = Trim(rsGrid.Fields(12))
12                Text(6) = Trim(rsGrid.Fields(5))
13                If Trim(rsGrid.Fields(1)) = "BTL" Then
14                    TblBatal.Visible = False
15                    tblSetuju.Visible = False
16                    TblTdkSetuju.Visible = False
17                    lblStatus.Caption = "Dibatalkan"
18                    lblStatus.Visible = True
19                End If
20                Select Case Trim(rsGrid.Fields(0))
                      Case "OKE"
21                        tblSetuju.Visible = False
22                        lblStatus.Caption = "Disetujui"
23                        lblStatus.Visible = True
24                    Case "NOK"
25                        TblTdkSetuju.Visible = False
26                        lblStatus.Caption = "Ditolak"
27                        lblStatus.Visible = True
28                End Select
                  
29                If rsJob.State Then rsJob.Close
30                Set rsJob = Nothing
31                Set GridJobOrder.DataSource = rsJob
32                Call headColumns
33                If rsGridPart.State Then rsGridPart.Close
34                Set rsGridPart = Nothing
35                Set partGrid.DataSource = rsGridPart
36                Call HeadColumn
37                If rsOpenno.State Then rsOpenno.Close
38                Set rsOpenno = Nothing
39                Set opennoGrid.DataSource = rsOpenno
40                With opennoGrid
41                    .Columns(0).Caption = "Open No"
42                    .Columns(0).DataField = "openNO"
43                    .Columns(0).Width = 1300
44                End With
                  
                  '' AMBIL DIAGORDER
45                Set rsSearch = objO.FetchData("SELECT * FROM diagDta WITH(NOLOCK) where no_diag = '" & Trim(rsGrid.Fields(2)) & "'")
46                If Not rsSearch.EOF Then
47                    Text(10) = NulSID(rsSearch!KODETIPE)
                      
48                    For i = 1 To 7
49                        tmp = "REP" & i
50                        tmp2 = "RTE" & i
51                        tmp3 = "KLPKERJA" & i
52                        Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
53                        If Trim(rsSearch(tmp)) <> "" Then
54                            With rsJob
55                                .AddNew
56                                !Group = rsSearch(tmp3)
57                                !job = rsSearch(tmp)
58                                !Rate = rsSearch(tmp2)
59                                .Update
60                            End With
61                        End If
62                        Set GridJobOrder.DataSource = rsJob
63                        Call headColumns
64                    Next
65                End If
66                If NulSID(Text(10)) <> "" Then
67                  Set rsSearch = objO.FetchData("SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH (NOLOCK) WHERE id_WARRANTY = '" & NulSID(Text(10)) & "' and (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)")
68                  If Not rsSearch.EOF Then
69                       Text(11).Text = NulSID(rsSearch!warranty)
70                  End If
71                End If
                  '' AMBIL DIAGADD
72                Set rsSearch = objO.FetchData("SELECT * FROM diagAdd2 WITH(NOLOCK) where no_diag = '" & Trim(rsGrid.Fields(2)) & "'")
73                If Not rsSearch.EOF Then
74                    Do While Not rsSearch.EOF
75                        Call createTblJob(IIf(rsJob.State = adStateClosed, True, False))
76                        With rsJob
77                            .AddNew
78                            !Group = rsSearch("KLPKERJA")
79                            !job = rsSearch("REP")
80                            !Rate = rsSearch("RTE")
81                            .Update
82                        End With
83                        Set GridJobOrder.DataSource = rsJob
84                        Call headColumns
85                        rsSearch.MoveNext
86                    Loop
87                End If
                  
                  '' AMBIL DIAGPART
88                Set rsSearch = objO.FetchData("SELECT * FROM diagPART WITH(NOLOCK) where no_diag = '" & Trim(rsGrid.Fields(2)) & "'")
89                If rsGridPart.State = adStateOpen Then rsGridPart.Close
                  
90                If Not rsSearch.EOF Then
91                    Set partGrid.DataSource = Nothing
92                    Do While Not rsSearch.EOF
93                        If rsGridPart.State = adStateClosed Then createTblPart
94                        With rsGridPart
95                            .AddNew
96                            !KodeItem = rsSearch("KODE_ITEM")
97                            !NAMAITEM = rsSearch("NAMA_ITEM")
98                            !SATITEM = rsSearch("SAT_ITEM")
99                            !qty = CDbl(rsSearch("QTY"))
100                           If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
101                               !namabySA = Trim(Nul(rsSearch("NAMABYSA")))
102                           Else
103                               !namabySA = ""
104                           End If
                              
105                           .Update
106                       End With
                          
                          
107                       rsSearch.MoveNext
108                   Loop
109                   Set partGrid.DataSource = rsGridPart
110                   Set partGridFill.DataSource = rsGridPart
111                   partGridFill.MoveFirst
112                   partGrid.MoveFirst
113                   Call HeadColumn
114               Else
115                   If rsGridPart.State Then rsGridPart.Close
116                   If rsGridPart.State = adStateClosed Then createTblPart
                      
117                   Set partGrid.DataSource = rsGridPart
118                   Set partGridFill.DataSource = rsGridPart
      '                partGridFill.MoveFirst
      '                PartGrid.MoveFirst
119                   Call HeadColumn
120               End If
                  
                  '' AMBIL DIAGOPENNO
121               Set rsSearch = objO.FetchData("SELECT DISTINCT OPENNO FROM DIAGAPPROVAL WITH(NOLOCK) where no_diag = '" & Trim(rsGrid.Fields(2)) & "'")
122               If Not rsSearch.EOF Then
123                   Do While Not rsSearch.EOF
124                       If rsOpenno.State = adStateClosed Then createTblOpenNo
125                       With rsOpenno
126                           .AddNew
127                           !OPENNO = rsSearch("OPENNO")
128                           .Update
129                           openBaris = openBaris + 1
130                       End With
131                       Set opennoGrid.DataSource = rsOpenno
132                       With opennoGrid
133                           .Columns(0).Caption = "Open No"
134                           .Columns(0).DataField = "openNO"
135                           .Columns(0).Width = 1300
136                       End With
137                       rsSearch.MoveNext
138                   Loop
139               End If
                  
140               GridJobOrder.Enabled = True
141               partGrid.Enabled = True
142               opennoGrid.Enabled = True
143               If lblStatus.Visible Then
144                   GridJobOrder.Enabled = False
145                   partGrid.Enabled = False
146                   opennoGrid.Enabled = False
147               End If
148               frmeDetail.Visible = True
149               Image6.Width = 5770
150               NoFak = Trim(rsGrid.Fields(13))
151               Call getDataKeluhan(Trim(rsGrid.Fields(2)))
152               Call getDataPermintaan(Trim(rsGrid.Fields(2)))
153           End If
154       End If
End Sub

Private Sub DiagGrid_FilterChange()
   On Error GoTo errHandler
    Set cols = DiagGrid.Columns
    Dim C As Integer
    
    C = DiagGrid.col
    DiagGrid.HoldFields
    rsGrid.filter = getFilter()
    If rsGrid.EOF Then
        For Each col In DiagGrid.Columns
            col.FilterText = ""
        Next col
        rsGrid.filter = ""
    End If
    DiagGrid.col = C
    DiagGrid.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
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


Private Sub DiagGrid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If frmeDetail.Visible Then
        Call DiagGrid_DblClick
    End If
End Sub

Private Sub Form_Load()
    Call hapusForm(True)
    sFilter = ""
    tanggal1 = Format(isnow, "dd/mm/yyyy")
    tanggal2 = Format(isnow, "dd/mm/yyyy")
    '' 20190701REV1
    If Log.lockSCC = 1 And Log.newLockSSC = 1 Then
        For i = 0 To 6
            chView(i).value = Checked
        Next
    Else
        chView(6).Visible = False
        For i = 0 To 5
            chView(i).value = Checked
        Next
    End If
    
    Call refreshGrid(tanggal1, tanggal2, Text(0), sFilter)
    'Call getDataPermintaan(text(2))
    vNamaTabel = "ReportValidasiTWC" & DriveSerialNumber("c:")
    Call cmdTutup_Click
    If NulSID(gTipeUser) = "ADMIN" Then
        cmdThn.Visible = True
    Else
        cmdThn.Visible = False
    End If
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = asxBars1.Top + asxBars1.Height '0
    CRViewer1.Left = 0 'frmValidasiTWC.Width 'asxPowerBanner1.Left + asxPowerBanner1.Width '0
    If ScaleHeight > 100 Then
        CRViewer1.Height = ScaleHeight - asxBars1.Height - 100 'frmValidasiTWC.Height - asxBars2.Height
    End If
    If frmValidasiTWC.Width > 200 Then
        CRViewer1.Width = frmValidasiTWC.Width - 200 'ScaleWidth - asxPowerBanner1.Width - 100
    End If
    asxBars1.Width = CRViewer1.Width
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
    cmdTutup.Top = CRViewer1.Top + 400
    cmdTutup.Left = CRViewer1.Left + 50
    cmdSettingPrinter.Top = cmdTutup.Top
    cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
    'asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub opennoGrid_DblClick()
Dim nCari As String

    If rsOpenno.State <> adStateClosed Then
        If Not rsOpenno.EOF Then
            nCari = Trim(opennoGrid.Columns(0))
            Text(7) = Trim(nCari)
            rsOpenno.MoveFirst
            rsOpenno.Find "OPENNO = '" & Trim(nCari) & "'", adSearchForward, 1
            If Not rsOpenno.BOF Then
                With rsOpenno
                    .Delete
                    openBaris = openBaris - 1
                End With
            End If
        End If
    End If
    Set opennoGrid.DataSource = rsOpenno
End Sub

Private Sub partGrid_Click()
MsgBox partGrid.Columns(0)

End Sub

Private Sub partGrid_DblClick()
Dim nCari As String

    If rsGridPart.RecordCount > 0 Then
        If Not rsGridPart.EOF Then
            If Trim(partGrid.Columns(0)) <> "" Then
                'nCari = Trim(partGrid.Columns(0))
                'text(8) = Trim(nCari)
                'text(9) = CDbl(partGrid.Columns(4))
                'rsGridPart.MoveFirst
                'rsGridPart.Find "KODEITEM = '" & Trim(nCari) & "'", adSearchForward, 1
                'If Not rsGridPart.EOF Then
                    With rsGridPart
                        .Delete
                        isBySA = False
                    End With
                'End If
            Else
                nCari = Trim(partGrid.Columns(2))
                Text(8) = Trim(nCari)
                Text(9) = CDbl(partGrid.Columns(4))
                rsGridPart.MoveFirst
                rsGridPart.Find "NAMABYSA = '" & Trim(nCari) & "'", adSearchForward, 1
                If Not rsGridPart.EOF Then
                    With rsGridPart
                        .Delete
                        isBySA = True
                    End With
                End If
            End If
        End If
        rsGridPart.MoveFirst
    End If
    
    Set partGrid.DataSource = rsGridPart
    Call HeadColumn
End Sub

Private Sub partGridFill_DblClick()
    Dim nCari As String
If rsGridPart.State <> adStateClosed Then
    If rsGridPart.RecordCount > 0 Then
        If Not rsGridPart.EOF Then
            If Trim(partGridFill.Columns(0)) <> "" Then
                'nCari = Trim(partGrid.Columns(0))
                'text(8) = Trim(nCari)
                'text(9) = CDbl(partGrid.Columns(4))
                'rsGridPart.MoveFirst
                'rsGridPart.Find "KODEITEM = '" & Trim(nCari) & "'", adSearchForward, 1
                'If Not rsGridPart.EOF Then
                    With rsGridPart
                        .Delete
                        isBySA = False
                    End With
                'End If
            Else
                nCari = Trim(partGridFill.Columns(2))
                Text(8) = Trim(nCari)
                Text(9) = CDbl(partGridFill.Columns(4))
                rsGridPart.MoveFirst
                rsGridPart.Find "NAMABYSA = '" & Trim(nCari) & "'", adSearchForward, 1
                If Not rsGridPart.EOF Then
                    With rsGridPart
                        .Delete
                        isBySA = True
                    End With
                End If
            End If
        End If
        rsGridPart.MoveFirst
    End If
End If
    Set partGridFill.DataSource = rsGridPart
End Sub

Private Sub TblAdd_Click()
    If Text(7) <> "" Then
        Call isiOpenNo(Trim(Text(7)))
        Text(7) = ""
        Text(7).SetFocus
    End If
End Sub

Private Sub tblBatal_Click()
    If MsgBox("Akan membatalkan PreDiagnosis ?", vbYesNo, "Pembatalan Prediagnosis") = vbYes Then
        Call objO.UpdateData("UPDATE DIAGDTA SET DELIV = 'BTL' WHERE NO_DIAG = '" & Trim(Text(2)) & "'")
        Call objO.UpdateData("UPDATE STALL SET FLAG = 'X' WHERE KETERANGAN like '% |" & Trim(Text(2)) & "| %'")
        tblSetuju.Enabled = False
        TblTdkSetuju.Enabled = False
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSearch_Click()
    Call refreshGrid(tanggal1, tanggal2, Text(0), sFilter)
End Sub

Private Sub tblSetuju_Click()
      Dim rsSearch As ADODB.Recordset
      Dim boleh As Boolean
      Dim pesan As String

1         boleh = True
2         pesan = ""
3         If rsOpenno.State = adStateClosed Then
4             pesan = pesan + "Openno. "
5             boleh = False
6         End If
7         If partGridFill.EOF Then
8             pesan = pesan + "Part Order. "
9             boleh = False
10        End If
      ''20190711 | tidak semua warranty ada SSC
      '    If NulSID(text(10)) = "" Then
      '        pesan = pesan + "Id Warranty. "
      '        boleh = False
      '    Else
      '        Set rsSearch = objO.FetchData("SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH (NOLOCK) WHERE id_WARRANTY = '" & NulSID(text(10)) & "' and (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)")
      '        If rsSearch.EOF Then
      '            pesan = pesan + "Id Warranty. "
      '            boleh = False
      '        End If
      '    End If
11        If boleh Then
12            If MsgBox("Prediagnosis ini disetujui ?", vbYesNo + vbDefaultButton1 + vbQuestion, "Persetujuan") = vbYes Then
13                erTranAll = False
14                onTR = True
15                Call objHT.OpenTransactionBengkel
                              
                  '' SIMPAN OPENNO
16                If rsOpenno.State <> adStateClosed Then
17                    rsOpenno.MoveFirst
18                    If Not rsOpenno.EOF Then
19                        Call objHT.DeleteData_BengkelOnly("DELETE DIAGAPPROVAL WHERE NO_DIAG = '" & Trim(Text(2)) & "'")
20                        Do While Not rsOpenno.EOF
21                            Call objHT.InsertData_BengkelOnly("INSERT INTO DIAGAPPROVAL VALUES ('" & Trim(Text(2)) & "','" & Trim(Text(6)) & "','" & Trim(rsOpenno("OPENNO")) & "','" & Trim(gSA) & "', '" & Format(isnow, "MM/DD/YYYY HH:NN") & "','','','" & IIf(chPWC.value, 1, 0) & "')")
22                            rsOpenno.MoveNext
23                        Loop
24                    End If
25                End If
                  
                  '' SIMPAN PART
26                If rsGridPart.State <> adStateClosed Then
27                    rsGridPart.MoveFirst
28                    If Not rsGridPart.EOF Then
29                        Call objHT.DeleteData_BengkelOnly("DELETE DIAGPART WHERE NO_DIAG = '" & Trim(Text(2)) & "'")
30                        Do While Not rsGridPart.EOF
31                            If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
32                              Call objHT.InsertData_BengkelOnly("INSERT INTO DIAGPART VALUES ('" & Trim(Text(2)) & "','" & Trim(Text(6)) & "','" & Trim(Text(5)) & "','','" & Trim(rsGridPart("KODEITEM")) & "','" & Trim(rsGridPart("NAMAITEM")) & "','" & Trim(rsGridPart("SATITEM")) & "','" & CDbl(rsGridPart("QTY")) & "','0','0','',0,'',0,'" & Trim(rsGridPart("NAMABYSA")) & "')")
33                            Else
34                              Call objHT.InsertData_BengkelOnly("INSERT INTO DIAGPART VALUES ('" & Trim(Text(2)) & "','" & Trim(Text(6)) & "','" & Trim(Text(5)) & "','','" & Trim(rsGridPart("KODEITEM")) & "','" & Trim(rsGridPart("NAMAITEM")) & "','" & Trim(rsGridPart("SATITEM")) & "','" & CDbl(rsGridPart("QTY")) & "','0','0','',0,'',0,'')")
35                            End If
36                            rsGridPart.MoveNext
37                        Loop
38                    End If
39                End If
                  
                  ''Call objHT.UpdateData_BengkelOnly("UPDATE DIAGDTA SET VALIDTWC = 'OKE' WHERE NO_DIAG = '" & Trim(text(2)) & "'")
                  ''20190705 penambahan kodetipe id warranty
                  
40                Call objHT.UpdateData_BengkelOnly("UPDATE DIAGDTA SET KODETIPE = '" & NulSID(Text(10)) & "',VALIDTWC = 'OKE' WHERE NO_DIAG = '" & Trim(Text(2)) & "'")
41                If erTranAll = False Then
42                    Call objHT.CloseTransactionBengkel
43                    onTR = False
44                    Call refreshGrid(tanggal1, tanggal2, Text(0), sFilter)
45                    TblBatal.Enabled = False
46                    tblSetuju.Enabled = False
47                    frmeDetail.Visible = False
48                Else
49                    Call objHT.CloseTransactionBengkel_err
50                    onTR = False
51                    MsgBox "Transaksi tidak dapat disimpan !!", vbCritical, "Penyimpanan"
52                End If
                  
53            End If
54        Else
55            MsgBox "Lengkapi field berikut : " & vbCr & pesan, vbExclamation, "Kelengkapan"
56        End If
End Sub


Private Sub TblTdkSetuju_Click()
    Dim boleh As Boolean
    Dim pesan As String

    boleh = True
    pesan = ""
'    If rsOpenno.State = adStateClosed Then
'        pesan = pesan + "Openno. "
'        boleh = False
'    End If
    
    If opennoGrid.EOF Then
    'If rsOpenno.RecordCount > 0 Then
        pesan = pesan + "Openno. harus diisi" & vbCrLf
        boleh = False
    End If
    
'    If rsGridPart.State = adStateClosed Then
'        pesan = pesan + "Part Order. "
'        boleh = False
'    End If
'
'    If Not partGridFill.EOF Then
'    'If  rsGridPart.RecordCount > 0 Then
'        pesan = pesan + "Part Order. "
'        boleh = False
'    End If
'
    If boleh Then
        If MsgBox("Prediagnosis ini tidak disetujui ?", vbYesNo, "Persetujuan") = vbYes Then
            Text1(5) = ""
            frmTidakSetuju.Visible = True
    '        Call objO.UpdateData("UPDATE DIAGDTA SET VALIDTWC = 'NOK' WHERE NO_DIAG = '" & Trim(text(2)) & "'")
    '        TblBatal.Enabled = False
    '        tblSetuju.Enabled = False
        End If
    Else
        MsgBox "Lengkapi field berikut : " & vbCr & pesan, vbExclamation, "Kelengkapan"
    End If
    
End Sub

Private Sub tblTutup_Click()
    frmeDetail.Visible = False
End Sub

Private Sub getDataPermintaan(noDiag As String)
    Set rsPermintaan = objO.FetchData("SELECT req1,req2,req3,req4,req5 FROM diagdta WITH(NOLOCK) WHERE no_diag = '" & Trim(noDiag) & "'")
    If Not rsPermintaan.EOF Then
        Text1(0) = rsPermintaan.Fields(0) & vbCrLf & rsPermintaan.Fields(1) & vbCrLf & rsPermintaan.Fields(2) & vbCrLf & rsPermintaan.Fields(3) & vbCrLf & rsPermintaan.Fields(4)
'        text(23) = rsPermintaan.Fields(0)
'        text(24) = rsPermintaan.Fields(1)
'        text(25) = rsPermintaan.Fields(2)
'        text(26) = rsPermintaan.Fields(3)
'        text(27) = rsPermintaan.Fields(4)
    End If
End Sub

Private Sub getDataKeluhan(noDiag As String)
    Set rsKeluhan = objO.FetchData("SELECT cmpl1,cmpl2,cmpl3,cmpl4,cmpl5 FROM diagdtl WITH(NOLOCK) WHERE no_diag = '" & Trim(noDiag) & "'")
    If Not rsKeluhan.EOF Then
        Text1(1) = rsKeluhan.Fields(0) & vbCrLf & rsKeluhan.Fields(1) & vbCrLf & rsKeluhan.Fields(2) & vbCrLf & rsKeluhan.Fields(3) & vbCrLf & rsKeluhan.Fields(4)
'        Text1(0) = rsKeluhan.Fields(0)
'        Text1(1) = rsKeluhan.Fields(1)
'        Text1(2) = rsKeluhan.Fields(2)
'        Text1(3) = rsKeluhan.Fields(3)
'        Text1(4) = rsKeluhan.Fields(4)
    End If
End Sub

Private Sub refreshGrid(Tgl1 As String, Tgl2 As String, sa As String, filter As String)
    DiagGrid.DataSource = Nothing
    Set rsGrid = objO.FetchData("SELECT VALIDTWC, DELIV, NO_DIAG, DATE, FRONT, NO_WO, NOINV, CUSTOMER,CUST_ADRES, KOTA, TELPON, NO_POLISI, NO_CHASSIS,NOFAK FROM DIAGDTA WITH(NOLOCK) WHERE " & IIf(sa <> "", "FRONT = '" & sa & "' AND ", "") & " DATE BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "' " & IIf(Trim(filter) <> "", " AND " & filter, "") & "")
    If Not rsGrid.EOF Then
        Set DiagGrid.DataSource = rsGrid
        Call headGrid
    End If
End Sub

Private Sub headGrid()
    With DiagGrid
        .Columns(0).Caption = "Persetujuan"
        .Columns(0).DataField = "VALIDTWC"
        .Columns(0).Width = 1000
        .Columns(0).Locked = True
        .Columns(1).Caption = "Status"
        .Columns(1).DataField = "DELIV"
        .Columns(1).Width = 1000
        .Columns(1).Locked = True
        .Columns(2).Caption = "No Diag"
        .Columns(2).DataField = "NO_DIAG"
        .Columns(2).Width = 1000
        .Columns(2).Locked = True
        .Columns(3).Caption = "Tanggal"
        .Columns(3).DataField = "DATE"
        .Columns(3).Width = 1200
        .Columns(3).Locked = True
        .Columns(4).Caption = "SA"
        .Columns(4).DataField = "FRONT"
        .Columns(4).Width = 800
        .Columns(4).Locked = True
        .Columns(5).Caption = "No. WO"
        .Columns(5).DataField = "NO_WO"
        .Columns(5).Width = 1000
        .Columns(5).Locked = True
        
        .Columns(6).Caption = "No Inv"
        .Columns(6).DataField = "NOINV"
        .Columns(6).Width = 1000
        .Columns(6).Locked = True
        
        .Columns(7).Caption = "Customer"
        .Columns(7).DataField = "CUSTOMER"
        .Columns(7).Width = 4000
        .Columns(7).Locked = True
        .Columns(8).Caption = "Alamat"
        .Columns(8).DataField = "CUST_ADRES"
        .Columns(8).Width = 4500
        .Columns(8).Locked = True
        .Columns(9).Caption = "Kota"
        .Columns(9).DataField = "KOTA"
        .Columns(9).Width = 1000
        .Columns(9).Locked = True
        .Columns(10).Caption = "Telpon"
        .Columns(10).DataField = "TELPON"
        .Columns(10).Width = 1000
        .Columns(10).Locked = True
        .Columns(11).Caption = "No Polisi"
        .Columns(11).DataField = "NO_POLISI"
        .Columns(11).Width = 1000
        .Columns(11).Locked = True
        .Columns(12).Caption = "No Rangka"
        .Columns(12).DataField = "NO_CHASSIS"
        .Columns(12).Width = 2000
        .Columns(12).Locked = True
    End With
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
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 8
                gMainKeyCari = Text(Index)
                gSearchOption = "PART"
                frmCariAcess.show (vbModal)
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "SA"
                frmCari.show (vbModal)
            Case 10
                gMainKeyCari = Text(Index)
                gSearchOption = "WARRANTY"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    Select Case Index
        Case 9
            KeyAscii = CekNumerik(KeyAscii)
            simpan = False
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsTemp As New ADODB.Recordset
    
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_PETUGAS WITH(NOLOCK) WHERE TIPE = 'SA' AND FS_KD_PETUGAS = '" & Text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SA"
                    frmCari.show (vbModal)
                Else
                    Text(Index + 1) = rsSearch("FS_NM_PETUGAS")
                End If
            Else
                Text(Index + 1) = ""
            End If
        Case 10
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT ID_WARRANTY, WARRANTY, KETERANGAN, TGLMULAI, TGLSELESAI FROM [" & dbReplikasiHO & "]..HO_WARRANTY WITH (NOLOCK) WHERE id_WARRANTY = '" & NulSID(Text(Index)) & "' and (DATEDIFF(day, TGLMULAI, TGLSELESAI) = 0 OR '" & Format(isnow, "MM/DD/YYYY") & "' BETWEEN TGLMULAI AND TGLSELESAI)")
                If rsSearch.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "WARRANTY"
                    frmCari.show (vbModal)
                Else
                    Text(11).Text = NulSID(rsSearch!warranty)
                End If
            End If
        Case 7
            If Text(Index) <> "" Then
                Call isiOpenNo(Trim(Text(7)))
                openBaris = openBaris + 1
                Text(7) = ""
                Text(7).SetFocus
            End If
        Case 8
            If Text(Index) <> "" Then
                    isBySA = False
                    Set oRs = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH(NOLOCK) INNER JOIN intrn B WITH(NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & Text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
                     
                    If oRs.EOF Then
                        If Log.isFill = 1 And Log.startAppRate <= Format(isnow, "dd/mm/yyyy") Then
                        If MsgBox("Inputkan Part, Nama By PIC ?", vbYesNo + vbDefaultButton1, "Nama Part By PIC") = vbYes Then
                            isBySA = True
                            hNamaPart = Text(8)
                            hSatPart = ""
                        End If
                Else
                  gMainKeyCari = Text(Index)
                        Text(Index) = ""
                        gSearchOption = "PART"
                        frmCariAcess.show (vbModal)
                        nStok = gStok
                        hNamaPart = gNamaPart
                        hSatPart = gSatPart
                        hHrgPart = gHargaJual
                        If rsTemp.State Then rsTemp.Close
                        Set rsTemp = objO.FetchData("SELECT Top 1 Tipe FROM T_OLI WITH(NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "' AND TIPE = 'O'")
                        If Not rsTemp.EOF Then
                            hOliPart = True
                        Else
                            hOliPart = False
                        End If
                End If
                        
                        
'                        If MsgBox("Bukan Part Stok Gudang ?", vbYesNo + vbDefaultButton1, "Bukan Stok") = vbYes Then
'                            If Trim(hNamaPart) = "" Or Trim(hSatPart) = "" Then
'                                hNamaPart = InputBox("Masukkan Nama Part : ", "Nama Part")
'                                hSatPart = InputBox("Masukkan Satuan Part : ", "Satuan Part")
'                            End If
'                        End If
                    Else
                        hNamaPart = oRs("NAMABARANG")
                        nStok = oRs("QTYSTOCK")
                        hHrgPart = oRs("HARGAJUALA")
                        hSatPart = oRs("KODESATUAN")
                        Set rsTemp = objO.FetchData("SELECT * FROM T_OLI WITH(NOLOCK) WHERE FS_KD_OLI = '" & Left(Text(Index), 1) & "' AND TIPE = 'O'")
                        If Not rsTemp.EOF Then
                            hOliPart = True
                        Else
                            hOliPart = False
                        End If
                    End If
                
            End If
        Case 9
            If Trim(Text(Index)) = "" Then
                Text(Index) = 0
                Sendkeys ("+{TAB}")
                Exit Sub
            Else
                If hNamaPart <> "" And Trim(Text(8)) <> "" Then
                    If Text(Index) > 0 Then
                        If rsGridPart.State = adStateClosed Then createTblPart
                        If rsGridPart.State <> adStateClosed Then
                            notallow = False
                            Do While Not rsGridPart.EOF
                                If Trim(rsGridPart("KODEITEM")) = Trim(Text(8)) Then
                                    notallow = True
                                End If
                                rsGridPart.MoveNext
                            Loop
                            
1                            If Not notallow Then
                            
                                    'ISI RS PART
                                    Call isiPart(Text(8), hNamaPart, hSatPart, CDbl(Text(Index)), Text(8))
                                    Text(8) = ""
                                    Text(Index) = 0
                            Else
                                Text(8) = ""
                                Text(9) = 0
                            End If
                        End If
                    End If
                Else
                    Text(8).SetFocus
                    Sendkeys ("+{TAB}")
                    Exit Sub
                End If
            End If
    End Select
End Sub

Public Sub isiPart(iKdItem As String, iNmItem As String, iSatItem As String, iQty As Integer, iNamaBySA As String)
    
    If rsGridPart.State = adStateClosed Then createTblPart
    
    If Not rsGridPart.BOF Then rsGridPart.MoveFirst
    rsGridPart.Find "KODEITEM = '" & Trim(iKdItem) & "'", adSearchForward, 1
    If rsGridPart.EOF Then
        rsGridPart.AddNew
        rsGridPart!KodeItem = IIf(isBySA = False, iKdItem, "")
        rsGridPart!NAMAITEM = IIf(isBySA = False, iNmItem, "")
        rsGridPart!SATITEM = iSatItem
        rsGridPart!qty = CDbl(iQty)
        rsGridPart!namabySA = IIf(isBySA = False, "", iNamaBySA)
        rsGridPart.Update
    End If
    Set partGridFill.DataSource = rsGridPart
    Call HeadColumn
End Sub

Public Sub isiOpenNo(iOpenNo As String)
    
    If rsOpenno.State = adStateClosed Then createTblOpenNo
    If Not rsOpenno.BOF Then rsOpenno.MoveFirst
    rsOpenno.Find "openNo = '" & Trim(iOpenNo) & "'", adSearchForward, 1
    If rsOpenno.EOF Then
        rsOpenno.AddNew
        rsOpenno!OPENNO = iOpenNo
        rsOpenno.Update
    End If
    Set opennoGrid.DataSource = rsOpenno
    With opennoGrid
        .Columns(0).Caption = "Open No"
        .Columns(0).DataField = "openNO"
        .Columns(0).Width = 1300
    End With
End Sub

Function createTblPart() As Boolean
1         On Error GoTo vb_error
          
2         If rsGridPart.State Then rsGridPart.Close
3         Set rsGridPart = Nothing
4         With rsGridPart
5             .Fields.Append "KODEITEM", adVarChar, 30
6             .Fields.Append "NAMAITEM", adVarChar, 200
7             .Fields.Append "SATITEM", adVarChar, 30
8             .Fields.Append "QTY", adInteger
9             .Fields.Append "NAMABYSA", adVarChar, 100
10            .Open
11        End With
          'If rsGridPart.State Then createTblPart = True
12        Exit Function
          
vb_error:
13        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Function createTblOpenNo() As Boolean
1         On Error GoTo vb_error
          
2         If rsOpenno.State Then rsOpenno.Close
3         Set rsOpenno = Nothing
4         With rsOpenno
5             .Fields.Append "OPENNO", adVarChar, 30
6             .Open
7             openBaris = 1
8         End With
9         If rsOpenno.State Then createTblOpenNo = True
10        Exit Function
          
vb_error:
11        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub HeadColumn()
    partGridFill.Columns(0).Caption = "KODE PART"
    partGridFill.Columns(0).DataField = "KODEITEM"
    partGridFill.Columns(0).Width = 1300
    partGridFill.Columns(1).Caption = "NAMA PART"
    partGridFill.Columns(1).DataField = "NAMAITEM"
    partGridFill.Columns(1).Width = 2500
    partGridFill.Columns(2).Caption = "NAMA BY PIC"
    partGridFill.Columns(2).DataField = "NAMABYSA"
    partGridFill.Columns(2).Width = 2500
    partGridFill.Columns(3).Caption = "SATUAN PART"
    partGridFill.Columns(3).DataField = "SATITEM"
    partGridFill.Columns(3).Width = 800
    partGridFill.Columns(4).Caption = "QTY"
    partGridFill.Columns(4).DataField = "QTY"
    partGridFill.Columns(4).Width = 600
'    partGrid.Columns(5).Caption = "OLI"
'    partGrid.Columns(5).DataField = "OLI"
'    partGrid.Columns(5).Visible = False
End Sub


