VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOrderListCuci 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Daftar Cuci"
   ClientHeight    =   9240
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12750
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9240
   ScaleWidth      =   12750
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   11760
      Top             =   3600
   End
   Begin ucXPButton.XPButton cmdKeluar 
      Height          =   495
      Left            =   10440
      TabIndex        =   1
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
      MICON           =   "frmOrderListCuci.frx":0000
      PICN            =   "frmOrderListCuci.frx":001C
      PICH            =   "frmOrderListCuci.frx":04F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   8493
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14737632
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SELESAI CUCI"
      TabPicture(0)   =   "frmOrderListCuci.frx":0A3B
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblHari"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblJumlahSelesaiCuci"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblUpdateSelesaiCuci"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DtArrAkhir"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DtArrAwal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "listSelesaiCuci"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdRefreshSelesaiCuci"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "ANTRI CUCI"
      TabPicture(1)   =   "frmOrderListCuci.frx":0A57
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblJumlahAC"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblHariAC"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblUpdateAC"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "asxPowerButton1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "listAntriCuci"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "framePembatalanWashing"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "PERMINTAAN CUCI"
      TabPicture(2)   =   "frmOrderListCuci.frx":0A73
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblJumlahPC"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblHariPC"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lblUpdatePC"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "btnBatal"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "listPermintaan"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "SELESAI PENGERINGAN"
      TabPicture(3)   =   "frmOrderListCuci.frx":0A8F
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "lblJumlahSP"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "lblHariSP"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "lblUpdateSP"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "listSelesaiPengeringan"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "ANTRI PENGERINGAN"
      TabPicture(4)   =   "frmOrderListCuci.frx":0AAB
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblHariAP"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "lblJumlahAP"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label7"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "lblUpdateAP"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "listAntriPengering"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).ControlCount=   5
      Begin ucXPButton.XPButton cmdRefreshSelesaiCuci 
         Height          =   375
         Left            =   8400
         TabIndex        =   42
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Refresh"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         MICON           =   "frmOrderListCuci.frx":0AC7
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame framePembatalanWashing 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Pembatalan Washing"
         Height          =   2295
         Left            =   -71040
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   3615
         Begin VB.TextBox txtAlasanBtlWashing 
            Height          =   1455
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   38
            Top             =   240
            Width           =   3375
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00DCFCFB&
            Caption         =   "Simpan"
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
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1800
            Width           =   1725
         End
      End
      Begin TrueOleDBGrid70.TDBGrid listSelesaiCuci 
         Height          =   3000
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5292
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
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid listSelesaiPengeringan 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   7
         Top             =   1080
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5292
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
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid listAntriCuci 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   13
         Top             =   1080
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5292
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
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid listAntriPengering 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   18
         Top             =   1320
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5292
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
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid listPermintaan 
         Height          =   3000
         Left            =   -74880
         TabIndex        =   23
         Top             =   1080
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5292
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
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
      End
      Begin AIFCmp1.asxPowerButton btnBatal 
         Height          =   495
         Left            =   -65400
         TabIndex        =   28
         Top             =   4200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BackColor       =   255
         Caption         =   "BATAL CUCI"
         Enabled         =   0   'False
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
         Left            =   -65400
         TabIndex        =   35
         Top             =   4200
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   873
         BackColor       =   255
         Caption         =   "BATAL CUCI"
         Enabled         =   0   'False
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
      Begin TDBDate6Ctl.TDBDate DtArrAwal 
         Height          =   330
         Left            =   4800
         TabIndex        =   39
         Top             =   360
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmOrderListCuci.frx":0AE3
         Caption         =   "frmOrderListCuci.frx":0C0F
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOrderListCuci.frx":0C7B
         Keys            =   "frmOrderListCuci.frx":0C99
         Spin            =   "frmOrderListCuci.frx":0CF7
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
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
      Begin TDBDate6Ctl.TDBDate DtArrAkhir 
         Height          =   330
         Left            =   6840
         TabIndex        =   40
         Top             =   360
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   582
         Calendar        =   "frmOrderListCuci.frx":0D1F
         Caption         =   "frmOrderListCuci.frx":0E4B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmOrderListCuci.frx":0EB7
         Keys            =   "frmOrderListCuci.frx":0ED5
         Spin            =   "frmOrderListCuci.frx":0F33
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   0
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
      Begin VB.Label Label9 
         Caption         =   "S / D"
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
         Left            =   6360
         TabIndex        =   41
         Top             =   405
         Width           =   495
      End
      Begin VB.Label lblUpdatePC 
         Caption         =   "Label1"
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
         Left            =   -72360
         TabIndex        =   24
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblHariPC 
         Caption         =   "Label1"
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
         Left            =   -74880
         TabIndex        =   27
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblJumlahPC 
         Caption         =   "JUMLAH DATA = 999"
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
         Left            =   -65520
         TabIndex        =   26
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "PERMINTAAN CUCI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -70680
         TabIndex        =   25
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblUpdateAP 
         Caption         =   "Label1"
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
         Left            =   -72360
         TabIndex        =   19
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label7 
         Caption         =   "ANTRI PENGERING"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -70680
         TabIndex        =   22
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label lblJumlahAP 
         Caption         =   "JUMLAH DATA = 999"
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
         Left            =   -65520
         TabIndex        =   21
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblHariAP 
         Caption         =   "Label1"
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
         Left            =   -74880
         TabIndex        =   20
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label lblUpdateAC 
         Caption         =   "Label1"
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
         Left            =   -72360
         TabIndex        =   14
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblHariAC 
         Caption         =   "Label1"
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
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblJumlahAC 
         Caption         =   "JUMLAH DATA = 999"
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
         Left            =   -65520
         TabIndex        =   16
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "ANTRI CUCI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -69960
         TabIndex        =   15
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblUpdateSP 
         Caption         =   "Label1"
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
         Left            =   -72360
         TabIndex        =   12
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblUpdateSelesaiCuci 
         Caption         =   "Label1"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblHariSP 
         Caption         =   "Label1"
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
         Left            =   -74880
         TabIndex        =   10
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label lblJumlahSP 
         Caption         =   "JUMLAH DATA = 999"
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
         Left            =   -65520
         TabIndex        =   9
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "SELESAI PENGERINGAN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   -70680
         TabIndex        =   8
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "SELESAI CUCI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5040
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblJumlahSelesaiCuci 
         Caption         =   "JUMLAH DATA = 999"
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
         Left            =   9480
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lblHari 
         Caption         =   "Label1"
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
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
   End
   Begin TrueOleDBGrid70.TDBGrid listProsesCuci 
      Height          =   2400
      Left            =   120
      TabIndex        =   29
      Top             =   840
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4233
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
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin TrueOleDBGrid70.TDBGrid listProsesPengering 
      Height          =   2400
      Left            =   6000
      TabIndex        =   31
      Top             =   840
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4233
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
      Columns.Count   =   8
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=8"
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
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
      _StyleDefs(68)  =   "Named:id=33:Normal"
      _StyleDefs(69)  =   ":id=33,.parent=0"
      _StyleDefs(70)  =   "Named:id=34:Heading"
      _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(72)  =   ":id=34,.wraptext=-1"
      _StyleDefs(73)  =   "Named:id=35:Footing"
      _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(75)  =   "Named:id=36:Selected"
      _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(77)  =   "Named:id=37:Caption"
      _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(79)  =   "Named:id=38:HighlightRow"
      _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(81)  =   "Named:id=39:EvenRow"
      _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(83)  =   "Named:id=40:OddRow"
      _StyleDefs(84)  =   ":id=40,.parent=33"
      _StyleDefs(85)  =   "Named:id=41:RecordSelector"
      _StyleDefs(86)  =   ":id=41,.parent=34"
      _StyleDefs(87)  =   "Named:id=42:FilterBar"
      _StyleDefs(88)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH DATA = 999"
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
      Left            =   9960
      TabIndex        =   34
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label lblJumlahProsesCuci 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH DATA = 999"
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
      Left            =   3000
      TabIndex        =   33
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DAFTAR UNIT PROSES PENGERING"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   32
      Top             =   600
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DAFTAR UNIT PROSES CUCI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   30
      Top             =   600
      Width           =   2655
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   11640
      X2              =   11040
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   11640
      X2              =   10800
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   11640
      X2              =   10680
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
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   11640
      X2              =   120
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   11640
      X2              =   120
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Daftar Cuci"
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
      Left            =   4800
      TabIndex        =   0
      Top             =   120
      Width           =   2775
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
      Width           =   11535
   End
End
Attribute VB_Name = "frmOrderListCuci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsSelesaiCuci As ADODB.Recordset
Private rsTempSelesaiCuci As New ADODB.Recordset

Private rsAntriCuci As ADODB.Recordset
Private rsTempAntriCuci As New ADODB.Recordset

Private rsPermintaanCuci As ADODB.Recordset
Private rsTempPermintaanCuci As New ADODB.Recordset

Private rsSelesaiPengeringan As ADODB.Recordset
Private rsTempSelesaiPengeringan As New ADODB.Recordset

Private rsAntriPengeringan As ADODB.Recordset
Private rsTempAntriPengeringCuci As New ADODB.Recordset

Private rsTemp As ADODB.Recordset

Dim rsStallCuci As ADODB.Recordset
Dim rsTemStallCuci As New ADODB.Recordset

Dim rsStallPengering As ADODB.Recordset
Dim rsTemStallPengering As New ADODB.Recordset

Private strSql As String
Private isPengering As Boolean

Private Function hitungEstWaktu(mulai As String, selesai As String) As String
    Dim jMulai As Integer
    Dim mMulai As Integer
    Dim jSelesai As Integer
    Dim mSelesai As Integer
    Dim total As Integer
    
    jMulai = CInt(Left(mulai, 2))
    mMulai = CInt(Mid(mulai, 4, 2))
    
    jSelesai = CInt(Left(selesai, 2))
    mSelesai = CInt(Mid(selesai, 4, 2))
    
    total = (jMulai * 60) + (mMulai) + (jSelesai * 60) + (mSelesai)
    
    If total > 60 Then
        jSelesai = CInt(Val(total / 60))
        mSelesai = CInt(total Mod 60)
        
        hitungEstWaktu = Format(jSelesai, "0#") & ":" & Format(mSelesai, "0#")
        
    Else
        hitungEstWaktu = "00:" & Format(Abs(CInt(total Mod 60)), "0#")
    End If
    
End Function

Private Function hitungWaktu(waktuCuci As String, menitProses As Integer) As String
    Dim jamW As Integer
    Dim menitW As Integer
    Dim totalWaktu As Double
    jamW = Left(waktuCuci, 2)
    menitW = Mid(waktuCuci, 4, 2)
    
    totalWaktu = (jamW * 60) + menitW
    totalWaktu = totalWaktu - menitProses
    If totalWaktu > 60 Then
        jamW = CInt(Val(totalWaktu / 60))
        menitW = CInt(totalWaktu Mod 60)
        
        hitungWaktu = Format(jamW, "0#") & ":" & Format(menitW, "0#")
    Else
        
        hitungWaktu = "00:" & Format(Abs(CInt(totalWaktu Mod 60)), "0#")
    End If
End Function

Private Function createSelesaiPengeringan() As Boolean
1         On Error GoTo vb_error
          
2         If rsTempSelesaiPengeringan.State Then
3         rsTempSelesaiPengeringan.Close
4         End If
5         Set rsTempSelesaiPengeringan = Nothing
6         With rsTempSelesaiPengeringan
7             .Fields.Append "NoPolisi", adVarChar, 25
8             .Fields.Append "SA", adVarChar, 5
9             .Fields.Append "NoWo", adVarChar, 15
10            .Fields.Append "OTD", adVarChar, 25
11            .Fields.Append "MulaiCuci", adVarChar, 25
12            .Fields.Append "Finish", adVarChar, 25
13            .Fields.Append "NoStall", adVarChar, 25
14            .Fields.Append "Mech", adVarChar, 25
15            .Fields.Append "Waktu", adVarChar, 25
16            .Fields.Append "EstFinish", adVarChar, 25
17            .Open
18        End With
19        If rsTempSelesaiPengeringan.State Then createSelesaiPengeringan = True
20        Exit Function
          
vb_error:
21        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createAntriPengeringCuci() As Boolean
1         On Error GoTo vb_error
          
2         If rsTempAntriPengeringCuci.State Then rsTempAntriPengeringCuci.Close
3         Set rsTempAntriPengeringCuci = Nothing
4         With rsTempAntriPengeringCuci
5             .Fields.Append "NoPolisi", adVarChar, 25
6             .Fields.Append "SA", adVarChar, 5
7             .Fields.Append "NoWo", adVarChar, 15
8             .Fields.Append "OTD", adVarChar, 25
9             .Fields.Append "MulaiCuci", adVarChar, 25
10            .Fields.Append "Finish", adVarChar, 25
11            .Fields.Append "NoStall", adVarChar, 25
12            .Fields.Append "Mech", adVarChar, 25
13            .Fields.Append "tipe", adVarChar, 25
14            .Fields.Append "SisaWaktu", adVarChar, 25
15            .Open
16        End With
17        If rsTempAntriPengeringCuci.State Then createAntriPengeringCuci = True
18        Exit Function
          
vb_error:
19        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createStallCuci() As Boolean
1         On Error GoTo vb_error
          
2         If rsTemStallCuci.State Then rsTemStallCuci.Close
3         Set rsTemStallCuci = Nothing
4         With rsTemStallCuci
5             .Fields.Append "NoPolisi", adVarChar, 25
6             .Fields.Append "SA", adVarChar, 5
7             .Fields.Append "NoWo", adVarChar, 15
8             .Fields.Append "OTD", adVarChar, 25
9             .Fields.Append "MulaiCuci", adVarChar, 25
10            .Fields.Append "Waktu", adVarChar, 25
11            .Fields.Append "EstFinish", adVarChar, 25
12            .Fields.Append "Mech", adVarChar, 25
13            .Fields.Append "NamaStall", adVarChar, 50
14            .Open
15        End With
16        If rsTemStallCuci.State Then createStallCuci = True
17        Exit Function
          
vb_error:
18        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createStallPengering() As Boolean
1         On Error GoTo vb_error
          
2         If rsTemStallPengering.State Then rsTemStallPengering.Close
3         Set rsTemStallPengering = Nothing
4         With rsTemStallPengering
5             .Fields.Append "NoPolisi", adVarChar, 25
6             .Fields.Append "SA", adVarChar, 5
7             .Fields.Append "NoWo", adVarChar, 15
8             .Fields.Append "OTD", adVarChar, 25
9             .Fields.Append "MulaiCuci", adVarChar, 25
10            .Fields.Append "Waktu", adVarChar, 25
11            .Fields.Append "EstFinish", adVarChar, 25
12            .Fields.Append "Mech", adVarChar, 25
13            .Fields.Append "NamaStall", adVarChar, 50
14            .Open
15        End With
16        If rsTemStallPengering.State Then createStallPengering = True
17        Exit Function
          
vb_error:
18        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createAntriCuci() As Boolean
    On Error GoTo vb_error
    
    If rsTempAntriCuci.State Then rsTempAntriCuci.Close
    Set rsTempAntriCuci = Nothing
    With rsTempAntriCuci
        .Fields.Append "NoPolisi", adVarChar, 25
        .Fields.Append "SA", adVarChar, 5
        .Fields.Append "NoWo", adVarChar, 15
        .Fields.Append "OTD", adVarChar, 25
        .Fields.Append "MulaiCuci", adVarChar, 25
        .Fields.Append "Finish", adVarChar, 25
        .Fields.Append "NoStall", adVarChar, 25
        .Fields.Append "Mech", adVarChar, 25
        .Fields.Append "tipe", adVarChar, 25
        .Fields.Append "SisaWaktu", adVarChar, 25
        .Open
    End With
    If rsTempAntriCuci.State Then createAntriCuci = True
    Exit Function
    
vb_error:
    MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createPermintaanCuci() As Boolean
1         On Error GoTo vb_error
          
2         If rsTempPermintaanCuci.State Then rsTempPermintaanCuci.Close
3         Set rsTempPermintaanCuci = Nothing
4         With rsTempPermintaanCuci
5             .Fields.Append "NoPolisi", adVarChar, 25
6             .Fields.Append "SA", adVarChar, 5
7             .Fields.Append "NoWo", adVarChar, 15
8             .Fields.Append "OTD", adVarChar, 25
9             .Fields.Append "Finish", adVarChar, 25
10            .Fields.Append "NoStall", adVarChar, 25
11            .Fields.Append "Posisi", adVarChar, 50
12            .Fields.Append "tipe", adVarChar, 25
13            .Open
14        End With
15        If rsTempPermintaanCuci.State Then createPermintaanCuci = True
16        Exit Function
          
vb_error:
17        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createSelesaiCuci() As Boolean
1         On Error GoTo vb_error
          
2         If rsTempSelesaiCuci.State Then rsTempSelesaiCuci.Close
3         Set rsTempSelesaiCuci = Nothing
4         With rsTempSelesaiCuci
5             .Fields.Append "NoPolisi", adVarChar, 25
6             .Fields.Append "SA", adVarChar, 5
7             .Fields.Append "NoWo", adVarChar, 15
8             .Fields.Append "OTD", adVarChar, 25
9             .Fields.Append "MulaiCuci", adVarChar, 25
10            .Fields.Append "Finish", adVarChar, 25
11            .Fields.Append "NoStall", adVarChar, 25
12            .Fields.Append "Mech", adVarChar, 25
13            .Fields.Append "PosisiAkhir", adVarChar, 50
14            .Open
15        End With
16        If rsTempSelesaiCuci.State Then createSelesaiCuci = True
17        Exit Function
          
vb_error:
18        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function convertMenit(menit As Double) As String
    Dim jam As Double
    Dim menitM As Double
    If menit > 60 Then
        jam = CDbl(Val(menit / 60))
        menitM = CDbl(menit Mod 60)
        
        convertMenit = Format(jam, "0#") & ":" & Format(menitM, "0#")
    Else
        convertMenit = "00:" & Format(menit, "0#")
    End If
End Function

Private Sub refreshGridAntriPengeringCuci()
Set rsAntriPengeringan = objO.FetchData("SELECT convert(varchar(5),getdate(),108) as waktuSekarang,pds,j.[code],j.kd_k,j.sbe,KLPKERJA1,KLPKERJA2,KLPKERJA3,KLPKERJA4,KLPKERJA5,KLPKERJA6,KLPKERJA7,j.pds,J.NO_POLISI AS NOPOLISI,A.*,J.FRONT,L.TIMESERAH,DATEDIFF(MINUTE, CONVERT(VARCHAR(5),getdate(),108),timeserah) as sisa FROM AntriCuci A WITH (NOLOCK) INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = A.NOORDER INNER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_WO = A.NOORDER WHERE tipestall = 'pengering' and A.FLAG = '' ORDER BY L.TIMESERAH")
    Dim sisaWaktu As String
    Dim tipe As String
    Dim waktuSekarang As String
    If createAntriPengeringCuci Then
        If rsAntriPengeringan.RecordCount > 0 Then
            rsAntriPengeringan.MoveFirst
            Do While Not rsAntriPengeringan.EOF
                waktuSekarang = rsAntriPengeringan("waktuSekarang")
                
                If Trim(rsAntriPengeringan("SISA")) >= 0 Then
                    sisaWaktu = convertMenit(CDbl(Trim(rsAntriPengeringan("SISA"))))
                Else
                    sisaWaktu = "-" & convertMenit(CDbl(Mid(Trim(rsAntriPengeringan("SISA")), 2, Len(Trim(rsAntriPengeringan("SISA")) - 1))))
                End If
                If CInt(rsAntriPengeringan("CODE")) = 1 Then
                    'em
                    tipe = "EM"
                Else
                    'np
                    If CInt(rsAntriPengeringan("KD_K")) = 1 Then
                        tipe = "NP"
                    Else
                        'sbi
                        If rsAntriPengeringan("KLPKERJA1") = "SBI" Or rsAntriPengeringan("KLPKERJA2") = "SBI" Or rsAntriPengeringan("KLPKERJA3") = "SBI" Or rsAntriPengeringan("KLPKERJA4") = "SBI" Or rsAntriPengeringan("KLPKERJA5") = "SBI" Or rsAntriPengeringan("KLPKERJA6") = "SBI" Or rsAntriPengeringan("KLPKERJA7") = "SBI" Then
                            tipe = "SBI"
                        Else
                            'pds
                            If rsAntriPengeringan("PDS") = 1 Then
                                tipe = "PDS"
                            Else
                                'sbe
                                If rsAntriPengeringan("KLPKERJA1") = "SBE" Or rsAntriPengeringan("KLPKERJA2") = "SBE" Or rsAntriPengeringan("KLPKERJA3") = "SBE" Or rsAntriPengeringan("KLPKERJA4") = "SBE" Or rsAntriPengeringan("KLPKERJA5") = "SBE" Or rsAntriPengeringan("KLPKERJA6") = "SBE" Or rsAntriPengeringan("KLPKERJA7") = "SBE" Then
                                    tipe = "SBE"
                                Else
                                    'gr
                                    tipe = "GRP"
                                End If
                            End If
                        End If
                    End If
                End If
                With rsTempAntriPengeringCuci
                    .AddNew
                    !NoPolisi = Trim(rsAntriPengeringan("NoPolisi"))
                    !sa = Trim(rsAntriPengeringan("FRONT"))
                    !noWO = Trim(rsAntriPengeringan("noorder"))
                    !OTD = Trim(rsAntriPengeringan("TIMESERAH"))
                    !tipe = tipe
                    !sisaWaktu = sisaWaktu
                    !Mech = Trim(rsAntriPengeringan("MECH"))
                    .Update
                    .MoveNext
                End With
                rsAntriPengeringan.MoveNext
                j = j + 1
            Loop
        End If
    End If
    listAntriPengering.DataSource = rsTempAntriPengeringCuci
    Call headColumnPengeringCuci
    lblUpdateAP.Caption = waktuSekarang
    lblJumlahAP.Caption = "JUMLAH DATA = " & rsTempAntriPengeringCuci.RecordCount
    lblHariAP.Caption = WeekdayName(IIf(Weekday(isnow) = 1, 7, Weekday(isnow) - 1)) & ", " & Format(isnow, "DD - MMMM - YYYY")
End Sub

Private Sub refreshGridAntriCuci()
Set rsAntriCuci = objO.FetchData("SELECT convert(varchar(5),getdate(),108) as waktuSekarang,pds,j.[code],j.kd_k,j.sbe,KLPKERJA1,KLPKERJA2,KLPKERJA3,KLPKERJA4,KLPKERJA5,KLPKERJA6,KLPKERJA7,j.pds,J.NO_POLISI AS NOPOLISI,A.*,J.FRONT,L.TIMESERAH,DATEDIFF(MINUTE, getdate(), dateserah + timeserah) as sisa,CONVERT(varchar(20),a.tanggal,108) as waktuAntri FROM AntriCuci A WITH (NOLOCK) INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = A.NOORDER INNER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_WO = A.NOORDER WHERE a.tipestall = 'Cuci' and A.FLAG = '' ORDER BY L.TIMESERAH")
    Dim sisaWaktu As String
    Dim tipe As String
    Dim waktuSekarang As String
    If createAntriCuci Then
        If rsAntriCuci.RecordCount > 0 Then
            rsAntriCuci.MoveFirst
            Do While Not rsAntriCuci.EOF
                waktuSekarang = rsAntriCuci("waktuSekarang")
                
                If Trim(rsAntriCuci("SISA")) >= 0 Then
                    sisaWaktu = convertMenit(CDbl(Trim(rsAntriCuci("SISA"))))
                Else
                    sisaWaktu = "-" & convertMenit(CDbl(Mid(Trim(rsAntriCuci("SISA")), 2, Len(Trim(rsAntriCuci("SISA")) - 1))))
                End If
                If CInt(rsAntriCuci("CODE")) = 1 Then
                    'em
                    tipe = "EM"
                Else
                    'np
                    If CInt(rsAntriCuci("KD_K")) = 1 Then
                        tipe = "NP"
                    Else
                        'sbi
                        If rsAntriCuci("KLPKERJA1") = "SBI" Or rsAntriCuci("KLPKERJA2") = "SBI" Or rsAntriCuci("KLPKERJA3") = "SBI" Or rsAntriCuci("KLPKERJA4") = "SBI" Or rsAntriCuci("KLPKERJA5") = "SBI" Or rsAntriCuci("KLPKERJA6") = "SBI" Or rsAntriCuci("KLPKERJA7") = "SBI" Then
                            tipe = "SBI"
                        Else
                            'pds
                            If rsAntriCuci("PDS") = 1 Then
                                tipe = "PDS"
                            Else
                                'sbe
                                If rsAntriCuci("KLPKERJA1") = "SBE" Or rsAntriCuci("KLPKERJA2") = "SBE" Or rsAntriCuci("KLPKERJA3") = "SBE" Or rsAntriCuci("KLPKERJA4") = "SBE" Or rsAntriCuci("KLPKERJA5") = "SBE" Or rsAntriCuci("KLPKERJA6") = "SBE" Or rsAntriCuci("KLPKERJA7") = "SBE" Then
                                    tipe = "SBE"
                                Else
                                    'gr
                                    tipe = "GRP"
                                End If
                            End If
                        End If
                    End If
                End If
                With rsTempAntriCuci
                    .AddNew
                    !NoPolisi = Trim(rsAntriCuci("NoPolisi"))
                    !sa = Trim(rsAntriCuci("FRONT"))
                    !noWO = Trim(rsAntriCuci("noorder"))
                    !OTD = Trim(rsAntriCuci("TIMESERAH"))
                    !tipe = tipe
                    !sisaWaktu = sisaWaktu
                    !Mech = Trim(rsAntriCuci("MECH"))
                    !MulaiCuci = Trim(rsAntriCuci("waktuAntri"))
                    .Update
                    .MoveNext
                End With
                rsAntriCuci.MoveNext
                j = j + 1
            Loop
        End If
    End If
    listAntriCuci.DataSource = rsTempAntriCuci
    Call headColumnAntriCuci
    lblUpdateAC.Caption = waktuSekarang
    lblJumlahAC.Caption = "JUMLAH DATA = " & rsTempAntriCuci.RecordCount
    lblHariAC.Caption = WeekdayName(IIf(Weekday(isnow) = 1, 7, Weekday(isnow) - 1)) & ", " & Format(isnow, "DD - MMMM - YYYY")
End Sub

Private Sub headColumnAntriCuci()
    With listAntriCuci
        .Columns(0).Caption = "No Polisi"
        .Columns(0).DataField = "NoPolisi"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "SA"
        .Columns(1).DataField = "SA"
        .Columns(1).Width = 700
        .Columns(2).Caption = "No Wo"
        .Columns(2).DataField = "NoWo"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "OTD"
        .Columns(3).DataField = "OTD"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "R / Type"
        .Columns(4).DataField = "tipe"
        .Columns(4).Width = 1500
        .Columns(5).Caption = "Sisa Waktu (HH:MM)"
        .Columns(5).DataField = "SisaWaktu"
        .Columns(5).Width = 1000
        .Columns(6).Caption = "Waktu Antri"
        .Columns(6).DataField = "mulaiCuci"
        .Columns(6).Width = 1000
        .Columns(6).Visible = True
    End With
End Sub

Private Sub headColumnPengeringCuci()
    With listAntriPengering
        .Columns(0).Caption = "No Polisi"
        .Columns(0).DataField = "NoPolisi"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "SA"
        .Columns(1).DataField = "SA"
        .Columns(1).Width = 700
        .Columns(2).Caption = "No Wo"
        .Columns(2).DataField = "NoWo"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "OTD"
        .Columns(3).DataField = "OTD"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "R / Type"
        .Columns(4).DataField = "tipe"
        .Columns(4).Width = 1500
        .Columns(5).Caption = "Sisa Waktu (HH:MM)"
        .Columns(5).DataField = "SisaWaktu"
        .Columns(5).Width = 1000
        .Columns(6).Caption = "MECH"
        .Columns(6).DataField = "MECH"
        .Columns(6).Width = 1000
        .Columns(6).Visible = False
    End With
End Sub

Private Sub refreshGridProsesPengering()
    Dim waktuCuci As String
    Dim waktu As String
    Dim waktuSekarang As String
    Dim estWaktu As String
    Dim selisihWaktu As Integer
    
    'waktuCuci = "00:30"
    
    
    noOrder = ""
    waktu = ""
    Set rsStallPengering = objO.FetchData("SELECT datediff(minute,CONVERT(VARCHAR(5),getdate(),108),l.timeserah) as selisih, CONVERT(VARCHAR(5),getdate(),108) as waktuSekarang, datediff(minute,start,CONVERT(VARCHAR,getdate(),108)) as proses, T.*,J.NO_POLISI,J.FRONT,L.TIMESERAH,Q.NAMASTALL,Q.LTCUCI FROM TempProgressCuci t WITH (NOLOCK) INNER JOIN QTY_STALLWASHING Q WITH (NOLOCK) ON Q.NOSTALL = T.NOSTALL AND ISDEL = '0' AND AKTIF = '1' INNER JOIN JOBDTA J WITH (NOLOCK) ON T.NOORDER = J.NO_WO INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NOORDER WHERE  FLAG = '' AND Q.TIPESTALL = 'Pengering'")
    Dim nopol As String
    If createStallPengering Then
        If rsStallPengering.RecordCount > 0 Then
            rsStallPengering.MoveFirst
            Do While Not rsStallPengering.EOF
                waktuCuci = Trim(rsStallPengering("LTCUCI"))
                nopol = Trim(rsStallPengering("NO_POLISI"))
                waktu = hitungWaktu(waktuCuci, rsStallPengering("proses"))
                waktuSekarang = rsStallPengering("waktuSekarang")
                estWaktu = hitungEstWaktu(rsStallPengering("START"), waktuCuci)
                selisihWaktu = rsStallPengering("selisih")
                noOrder = Trim(rsStallPengering("NOORDER"))
                With rsTemStallPengering
                    .AddNew
                    !NoPolisi = Trim(rsStallPengering("NO_POLISI"))
                    !sa = Trim(rsStallPengering("FRONT"))
                    !noWO = Trim(rsStallPengering("NOORDER"))
                    !OTD = Trim(rsStallPengering("TIMESERAH"))
                    !MulaiCuci = Trim(rsStallPengering("START"))
                    !waktu = waktu
                    !EstFinish = estWaktu
                    !Mech = Trim(rsStallPengering("MECH"))
                    !NAMASTALL = Trim(rsStallPengering("NAMASTALL"))
                    .Update
                    .MoveNext
                End With
                rsStallPengering.MoveNext
                j = j + 1
            Loop
        End If
    End If
    listProsesPengering.DataSource = rsTemStallPengering
    Call headColumnsProsesPengering
    Label8.Caption = "JUMLAH DATA = " & rsTemStallPengering.RecordCount
End Sub

Private Sub refreshGridProsesCuci()
    Dim waktuCuci As String
    Dim waktu As String
    Dim waktuSekarang As String
    Dim estWaktu As String
    Dim selisihWaktu As Integer
    
    'waktuCuci = "00:30"
    
    
    noOrder = ""
    waktu = ""
    Set rsStallCuci = objO.FetchData("SELECT datediff(minute,CONVERT(VARCHAR(5),getdate(),108),l.timeserah) as selisih, CONVERT(VARCHAR(5),getdate(),108) as waktuSekarang, datediff(minute,start,CONVERT(VARCHAR,getdate(),108)) as proses, T.*,J.NO_POLISI,J.FRONT,L.TIMESERAH,Q.NAMASTALL,Q.LTCUCI FROM TempProgressCuci t WITH (NOLOCK) INNER JOIN JOBDTA J WITH (NOLOCK) ON T.NOORDER = J.NO_WO INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NOORDER INNER JOIN QTY_STALLWASHING Q WITH (NOLOCK) ON Q.NOSTALL = T.NOSTALL WHERE  FLAG = '' AND Q.TIPESTALL = 'Cuci'")
    Dim nopol As String
    If createStallCuci Then
        If rsStallCuci.RecordCount > 0 Then
            rsStallCuci.MoveFirst
            Do While Not rsStallCuci.EOF
                waktuCuci = Trim(rsStallCuci("LTCUCI"))
                nopol = Trim(rsStallCuci("NO_POLISI"))
                waktu = hitungWaktu(waktuCuci, rsStallCuci("proses"))
                waktuSekarang = rsStallCuci("waktuSekarang")
                estWaktu = hitungEstWaktu(rsStallCuci("START"), waktuCuci)
                selisihWaktu = rsStallCuci("selisih")
                noOrder = Trim(rsStallCuci("NOORDER"))
                With rsTemStallCuci
                    .AddNew
                    !NoPolisi = Trim(rsStallCuci("NO_POLISI"))
                    !sa = Trim(rsStallCuci("FRONT"))
                    !noWO = Trim(rsStallCuci("NOORDER"))
                    !OTD = Trim(rsStallCuci("TIMESERAH"))
                    !MulaiCuci = Trim(rsStallCuci("START"))
                    !waktu = waktu
                    !EstFinish = estWaktu
                    !Mech = Trim(rsStallCuci("MECH"))
                    !NAMASTALL = Trim(rsStallCuci("NAMASTALL"))
                    .Update
                    .MoveNext
                End With
                rsStallCuci.MoveNext
                j = j + 1
            Loop
        End If
    End If
    listProsesCuci.DataSource = rsTemStallCuci
    Call headColumnsProsesCuci
    lblJumlahProsesCuci.Caption = "JUMLAH DATA = " & rsTemStallCuci.RecordCount
End Sub

Private Sub refreshGridSelesaiCuci()
    Dim flag As String
    Dim waktuCuci As String
    Dim waktu As String
    Dim waktuSekarang As String
    Dim estWaktu As String
    Dim selisihWaktu As Integer
    
    'waktuCuci = "00:30"
'    If SSTab1.TabVisible(3) = True Then
'        flag = "D"
'    Else
        flag = "F"
'    End If
    
    strSql = "SELECT datediff(minute,CONVERT(VARCHAR(5),getdate(),108),l.timeserah) as selisih, " & _
             "CONVERT(VARCHAR(5),getdate(),108) as waktuSekarang,datediff(minute,start,CONVERT(VARCHAR,getdate(),108)) as proses, " & _
             "T.*,J.NO_POLISI,J.FRONT,L.TIMESERAH,a.posisiAkhir " & _
             "FROM TempProgressCuci t WITH (NOLOCK) " & _
             "INNER JOIN ANTRICUCI a WITH (NOLOCK) ON a.NOORDER = T.NOORDER " & _
             "INNER JOIN STALL S WITH (NOLOCK) ON substring(S.keterangan,charindex(' |',S.keterangan)+2,charindex('| ',S.keterangan)-charindex(' |',S.keterangan)-2) = T.NOORDER " & _
             "INNER JOIN JOBDTA J WITH (NOLOCK) ON T.NOORDER = J.NO_WO " & _
             "INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NOORDER " & _
             "WHERE T.tanggal >= '" & Format(DtArrAwal.value, "YYYY-MM-DD") & " 00:00:00' and T.tanggal <= '" & Format(DtArrAkhir.value, "YYYY-MM-DD") & " 23:59:59' and T.FLAG = '" & flag & "'"
    Set rsSelesaiCuci = objO.FetchData(strSql)
    
    If createSelesaiCuci Then
        If rsSelesaiCuci.RecordCount > 0 Then
            rsSelesaiCuci.MoveFirst
            Do While Not rsSelesaiCuci.EOF
                
'                waktu = hitungWaktu(waktuCuci, rsSelesaiCuci("proses"))
                waktuSekarang = rsSelesaiCuci("waktuSekarang")
'                estWaktu = hitungEstWaktu(rsSelesaiCuci("START"), waktuCuci)
                selisihWaktu = rsSelesaiCuci("selisih")
                noOrder = Trim(rsSelesaiCuci("NOORDER"))
                With rsTempSelesaiCuci
                    .AddNew
                    !NoPolisi = Trim(rsSelesaiCuci("NO_POLISI"))
                    !sa = Trim(rsSelesaiCuci("FRONT"))
                    !noWO = Trim(rsSelesaiCuci("NOORDER"))
                    !OTD = Trim(rsSelesaiCuci("TIMESERAH"))
                    !MulaiCuci = Trim(rsSelesaiCuci("START"))
                    !finish = Trim(rsSelesaiCuci("Finish"))
                    !noStall = Trim(rsSelesaiCuci("NoStall"))
                    !Mech = Trim(rsSelesaiCuci("MECH"))
                    !posisiAkhir = Trim(rsSelesaiCuci!posisiAkhir)
                    .Update
                    .MoveNext
                End With
                rsSelesaiCuci.MoveNext
            Loop
        End If
    End If
    
    listSelesaiCuci.DataSource = rsTempSelesaiCuci
    Call headColumns
    lblUpdateSelesaiCuci.Caption = waktuSekarang
    lblJumlahSelesaiCuci.Caption = "JUMLAH DATA = " & rsTempSelesaiCuci.RecordCount
    lblHari.Caption = WeekdayName(IIf(Weekday(isnow) = 1, 7, Weekday(isnow) - 1)) & ", " & Format(isnow, "DD - MMMM - YYYY")
End Sub

Private Sub refreshGridSelesaiPengeringan()
    Dim flag As String
    Dim waktuCuci As String
    Dim waktu As String
    Dim waktuSekarang As String
    Dim estWaktu As String
    Dim selisihWaktu As Integer
    
   ' waktuCuci = "00:30"
    If SSTab1.TabVisible(3) = True Then
        flag = "F"
    End If
    
    strSql = "SELECT datediff(minute,CONVERT(VARCHAR(5),getdate(),108),l.timeserah) as selisih, " & _
             "CONVERT(VARCHAR(5),getdate(),108) as waktuSekarang,datediff(minute,start,CONVERT(VARCHAR,getdate(),108)) as proses, " & _
             "T.*,J.NO_POLISI,J.FRONT,L.TIMESERAH,Q.LTCUCI " & _
             "FROM TempProgressCuci t WITH (NOLOCK) " & _
             "INNER JOIN QTY_STALLWASHING Q WITH (NOLOCK) ON Q.NOSTALL = T.NOSTALL AND ISDEL = '0' AND AKTIF = '1' " & _
             "INNER JOIN STALL S WITH (NOLOCK) ON substring(S.keterangan,charindex(' |',S.keterangan)+2,charindex('| ',S.keterangan)-charindex(' |',S.keterangan)-2) = T.NOORDER " & _
             "INNER JOIN JOBDTA J WITH (NOLOCK) ON T.NOORDER = J.NO_WO " & _
             "INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NOORDER " & _
             "WHERE T.tanggal >= '" & Format(isnow, "YYYY-MM-DD") & " 00:00:00' and T.tanggal <= '" & Format(isnow, "YYYY-MM-DD") & " 23:59:59' and T.FLAG = '" & flag & "'"
    Set rsSelesaiPengeringan = objO.FetchData(strSql)
    
    If createSelesaiPengeringan Then
        If rsSelesaiPengeringan.RecordCount > 0 Then
            rsSelesaiPengeringan.MoveFirst
            Do While Not rsSelesaiPengeringan.EOF
                waktuCuci = Trim(rsSelesaiPengeringan("LTCUCI"))
                waktu = hitungWaktu(waktuCuci, rsSelesaiPengeringan("proses"))
                waktuSekarang = rsSelesaiPengeringan("waktuSekarang")
                estWaktu = hitungEstWaktu(rsSelesaiPengeringan("START"), waktuCuci)
                selisihWaktu = rsSelesaiPengeringan("selisih")
                noOrder = Trim(rsSelesaiPengeringan("NOORDER"))
                With rsTempSelesaiPengeringan
                    .AddNew
                    !NoPolisi = Trim(rsSelesaiPengeringan("NO_POLISI"))
                    !sa = Trim(rsSelesaiPengeringan("FRONT"))
                    !noWO = Trim(rsSelesaiPengeringan("NOORDER"))
                    !OTD = Trim(rsSelesaiPengeringan("TIMESERAH"))
                    !MulaiCuci = Trim(rsSelesaiPengeringan("START"))
                    !waktu = waktu
                    !EstFinish = estWaktu
                    !Mech = Trim(rsSelesaiPengeringan("MECH"))
                    !finish = Trim(rsSelesaiPengeringan("FINISH"))
                    !noStall = Trim(rsSelesaiPengeringan("NoStall"))
                    .Update
                    .MoveNext
                End With
                rsSelesaiPengeringan.MoveNext
            Loop
        End If
    End If
    
    listSelesaiPengeringan.DataSource = rsTempSelesaiPengeringan
    Call headColumnsSP
    lblUpdateSP.Caption = waktuSekarang
    lblJumlahSP.Caption = "JUMLAH DATA = " & rsTempSelesaiPengeringan.RecordCount
    lblHariSP.Caption = WeekdayName(IIf(Weekday(isnow) = 1, 7, Weekday(isnow) - 1)) & ", " & Format(isnow, "DD - MMMM - YYYY")
End Sub

Private Sub refreshGridPermintaanCuci()
    Dim flag As String
    Dim waktuCuci As String
    Dim waktu As String
    Dim waktuSekarang As String
    Dim estWaktu As String
    Dim selisihWaktu As Integer
    Dim tipe As String
    Dim Posisi As String
    
'    waktuCuci = "00:30"
    
    strSql = "select a.no_polisi,a.front,a.no_wo,l.timeserah,convert(varchar(5),getdate(),108) as waktuSekarang,pds,a.[code],a.kd_k,a.sbe,KLPKERJA1,KLPKERJA2,KLPKERJA3,KLPKERJA4,KLPKERJA5,KLPKERJA6,KLPKERJA7,a.pds,s.flag " & _
             "from jobdta a " & _
             "inner join jobdtl l on l.no_wo = a.no_wo " & _
             "inner join stall s on substring(S.keterangan,charindex(' |',S.keterangan)+2,charindex('| ',S.keterangan)-charindex(' |',S.keterangan)-2) = a.no_wo " & _
             "where s.flag in ('N','S','P','H') and l.washing = '1' "
    Set rsPermintaanCuci = objO.FetchData(strSql)
    
    If createPermintaanCuci Then
        If rsPermintaanCuci.RecordCount > 0 Then
            rsPermintaanCuci.MoveFirst
            Do While Not rsPermintaanCuci.EOF
                waktuSekarang = rsPermintaanCuci("waktuSekarang")
                If CInt(IIf(rsPermintaanCuci("CODE") = "", 0, rsPermintaanCuci("CODE"))) = 1 Then
                    'em
                    tipe = "EM"
                Else
                    'np
                    If CInt(IIf(rsPermintaanCuci("KD_K") = "", 0, rsPermintaanCuci("KD_K"))) = 1 Then
                        tipe = "NP"
                    Else
                        'sbi
                        If rsPermintaanCuci("KLPKERJA1") = "SBI" Or rsPermintaanCuci("KLPKERJA2") = "SBI" Or rsPermintaanCuci("KLPKERJA3") = "SBI" Or rsPermintaanCuci("KLPKERJA4") = "SBI" Or rsPermintaanCuci("KLPKERJA5") = "SBI" Or rsPermintaanCuci("KLPKERJA6") = "SBI" Or rsPermintaanCuci("KLPKERJA7") = "SBI" Then
                            tipe = "SBI"
                        Else
                            'pds
                            If rsPermintaanCuci("PDS") = 1 Then
                                tipe = "PDS"
                            Else
                                'sbe
                                If rsPermintaanCuci("KLPKERJA1") = "SBE" Or rsPermintaanCuci("KLPKERJA2") = "SBE" Or rsPermintaanCuci("KLPKERJA3") = "SBE" Or rsPermintaanCuci("KLPKERJA4") = "SBE" Or rsPermintaanCuci("KLPKERJA5") = "SBE" Or rsPermintaanCuci("KLPKERJA6") = "SBE" Or rsPermintaanCuci("KLPKERJA7") = "SBE" Then
                                    tipe = "SBE"
                                Else
                                    'gr
                                    tipe = "GRP"
                                End If
                            End If
                        End If
                    End If
                End If
                Select Case Trim(rsPermintaanCuci("FLAG"))
                    Case ""
                        Posisi = "Unit belum dibuat WO masih booking waiting "
                    Case "N"
                        Posisi = "Unit belum dikerjakan "
                    Case "S"
                        Posisi = "Unit sedang dikerjakan"
                    Case "P"
                        Posisi = "Unit sedang dihentikan sementara"
                    Case "H"
                        Posisi = "Unit menunggu proses Final Inspection"
                    Case "Q"
                        Posisi = "Unit dalam proses Final Inspection"
                    Case "G"
                        Posisi = "Unit menunggu proses Washing"
                    Case "A"
                        Posisi = "Unit sedang proses Washing"
                    Case "F"
                        Posisi = "Unit telah selesai dikerjakan"
                End Select
                
                waktuSekarang = rsPermintaanCuci("waktuSekarang")
                'estWaktu = hitungEstWaktu(rsPermintaanCuci("START"), waktuCuci)
                'selisihWaktu = rsPermintaanCuci("selisih")
                noOrder = Trim(rsPermintaanCuci("NO_WO"))
                With rsTempPermintaanCuci
                    .AddNew
                    !NoPolisi = Trim(rsPermintaanCuci("NO_POLISI"))
                    !sa = Trim(rsPermintaanCuci("FRONT"))
                    !noWO = Trim(rsPermintaanCuci("NO_WO"))
                    !OTD = Trim(rsPermintaanCuci("TIMESERAH"))
                    !tipe = tipe
                    !Posisi = Posisi
                    .Update
                    .MoveNext
                End With
                rsPermintaanCuci.MoveNext
            Loop
        End If
    End If
    
    listPermintaan.DataSource = rsTempPermintaanCuci
    Call headColumnsPermintaan
    lblUpdatePC.Caption = waktuSekarang
    lblJumlahPC.Caption = "JUMLAH DATA = " & rsTempPermintaanCuci.RecordCount
    lblHariPC.Caption = WeekdayName(IIf(Weekday(isnow) = 1, 7, Weekday(isnow) - 1)) & ", " & Format(isnow, "DD - MMMM - YYYY")
End Sub

Private Sub headColumnsProsesPengering()
    With listProsesPengering
        .Columns(0).Caption = "Nama Stall"
        .Columns(0).DataField = "NamaStall"
        .Columns(0).Width = 1500
        .Columns(1).Caption = "No Polisi"
        .Columns(1).DataField = "NoPolisi"
        .Columns(1).Width = 1000
        .Columns(2).Caption = "SA"
        .Columns(2).DataField = "SA"
        .Columns(2).Width = 700
        .Columns(3).Caption = "No Wo"
        .Columns(3).DataField = "NoWo"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "OTD"
        .Columns(4).DataField = "OTD"
        .Columns(4).Width = 1000
        .Columns(5).Caption = "Mulai Cuci"
        .Columns(5).DataField = "MulaiCuci"
        .Columns(5).Width = 1000
        .Columns(6).Caption = "Waktu"
        .Columns(6).DataField = "Waktu"
        .Columns(6).Width = 1000
        .Columns(7).Caption = "Est Selesai Cuci"
        .Columns(7).DataField = "EstFinish"
        .Columns(7).Width = 1000
    End With
End Sub

Private Sub headColumnsProsesCuci()
    With listProsesCuci
        .Columns(0).Caption = "Nama Stall"
        .Columns(0).DataField = "NamaStall"
        .Columns(0).Width = 1500
        .Columns(1).Caption = "No Polisi"
        .Columns(1).DataField = "NoPolisi"
        .Columns(1).Width = 1000
        .Columns(2).Caption = "SA"
        .Columns(2).DataField = "SA"
        .Columns(2).Width = 700
        .Columns(3).Caption = "No Wo"
        .Columns(3).DataField = "NoWo"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "OTD"
        .Columns(4).DataField = "OTD"
        .Columns(4).Width = 1000
        .Columns(5).Caption = "Mulai Cuci"
        .Columns(5).DataField = "MulaiCuci"
        .Columns(5).Width = 1000
        .Columns(6).Caption = "Waktu"
        .Columns(6).DataField = "Waktu"
        .Columns(6).Width = 1000
        .Columns(7).Caption = "Est Selesai Cuci"
        .Columns(7).DataField = "EstFinish"
        .Columns(7).Width = 1000
    End With
    
    
End Sub

Private Sub headColumns()
    With listSelesaiCuci
        .Columns(0).Caption = "No Polisi"
        .Columns(0).DataField = "NoPolisi"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "SA"
        .Columns(1).DataField = "SA"
        .Columns(1).Width = 700
        .Columns(2).Caption = "No Wo"
        .Columns(2).DataField = "NoWo"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "OTD"
        .Columns(3).DataField = "OTD"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "Mulai Cuci"
        .Columns(4).DataField = "MulaiCuci"
        .Columns(4).Width = 1000
        .Columns(5).Caption = "Selesai Cuci"
        .Columns(5).DataField = "Finish"
        .Columns(5).Width = 1000
        .Columns(6).Caption = "No Stall Cuci"
        .Columns(6).DataField = "NoStall"
        .Columns(6).Width = 1000
        .Columns(7).Caption = "Posisi Akhir"
        .Columns(7).DataField = "posisiAkhir"
        .Columns(7).Width = 1500
        '.Columns(7).Visible = False
    End With
End Sub

Private Sub headColumnsPermintaan()
    With listPermintaan
        .Columns(0).Caption = "No Polisi"
        .Columns(0).DataField = "NoPolisi"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "SA"
        .Columns(1).DataField = "SA"
        .Columns(1).Width = 700
        .Columns(2).Caption = "No Wo"
        .Columns(2).DataField = "NoWo"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "OTD"
        .Columns(3).DataField = "OTD"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "R / Type"
        .Columns(4).DataField = "tipe"
        .Columns(4).Width = 1500
        .Columns(5).Caption = "Posisi Terakhir"
        .Columns(5).DataField = "Posisi"
        .Columns(5).Width = 3000
        .Columns(6).Caption = "MECH"
        .Columns(6).DataField = "MECH"
        .Columns(6).Width = 1000
        .Columns(6).Visible = False
    End With
End Sub


Private Sub headColumnsSP()
    With listSelesaiPengeringan
        .Columns(0).Caption = "No Polisi"
        .Columns(0).DataField = "NoPolisi"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "SA"
        .Columns(1).DataField = "SA"
        .Columns(1).Width = 700
        .Columns(2).Caption = "No Wo"
        .Columns(2).DataField = "NoWo"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "OTD"
        .Columns(3).DataField = "OTD"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "Mulai Cuci"
        .Columns(4).DataField = "MulaiCuci"
        .Columns(4).Width = 1000
        .Columns(5).Caption = "Selesai Cuci"
        .Columns(5).DataField = "Finish"
        .Columns(5).Width = 1000
        .Columns(6).Caption = "No Stall Cuci"
        .Columns(6).DataField = "NoStall"
        .Columns(6).Width = 1000
        .Columns(7).Visible = False
    End With
End Sub

Private Sub headSelesaiPengeringanColumns()
    With listSelesaiCuci
        .Columns(0).Caption = "No Polisi"
        .Columns(0).DataField = "NoPolisi"
        .Columns(0).Width = 1000
        .Columns(1).Caption = "SA"
        .Columns(1).DataField = "SA"
        .Columns(1).Width = 700
        .Columns(2).Caption = "No Wo"
        .Columns(2).DataField = "NoWo"
        .Columns(2).Width = 1000
        .Columns(3).Caption = "OTD"
        .Columns(3).DataField = "OTD"
        .Columns(3).Width = 1000
        .Columns(4).Caption = "Mulai Cuci"
        .Columns(4).DataField = "MulaiCuci"
        .Columns(4).Width = 1000
        .Columns(5).Caption = "Waktu"
        .Columns(5).DataField = "Waktu"
        .Columns(5).Width = 1000
        .Columns(6).Caption = "Est Selesai Cuci"
        .Columns(6).DataField = "EstFinish"
        .Columns(6).Width = 1000
    End With
End Sub

Private Sub asxPowerButton1_Click()
   If Log.lockWashing = 1 Then
      If MsgBox("Apakah ingin membatalkan antri cuci dari No Polisi " & listAntriCuci.Columns("nopolisi").value & " , No WO " & listAntriCuci.Columns("nowo").value & " ?" & vbCrLf & "Jika Yes maka pilihan anda akan direport ke HO!", vbYesNo, "") = vbYes Then
         framePembatalanWashing.Visible = True
      End If
   Else
      If MsgBox("Apakah ingin membatalkan antri cuci dari No Polisi " & listAntriCuci.Columns("nopolisi").value & " , No WO " & listAntriCuci.Columns("nowo").value & " ?", vbYesNo, "Konfirmasi Pembatalan Cuci") = vbYes Then
          framePembatalanWashing.Visible = True
      End If
   End If
    
End Sub

Private Sub btnBatal_Click()
    If MsgBox("Apakah ingin membatalkan proses cuci dari No Polisi " & listPermintaan.Columns("nopolisi").value & " , No WO " & listPermintaan.Columns("nowo").value & " ?", vbYesNo, "Konfirmasi Pembatalan Cuci") = vbYes Then
        objO.UpdateData ("UPDATE JOBDTL SET WASHING = '0' WHERE no_wo = '" & listPermintaan.Columns("nowo").value & "'")
        Call refreshGridPermintaanCuci
    End If
End Sub

Private Sub CmdKeluar_Click()
Unload Me
End Sub

Private Sub cmdRefreshSelesaiCuci_Click()
    Call refreshGridSelesaiCuci
    
End Sub

Private Sub Command2_Click()
    If Trim(txtAlasanBtlWashing.text) <> "" Then
        If Len(txtAlasanBtlWashing.text) < 5 Then
            MsgBox "Alasan Pembatalan Washing minimal 5 karakter!", vbCritical, "Informasi Pembatalan Washing"
        Else
            On Error GoTo vbError
            erTranAll = False
               If Trim(listAntriCuci.Columns("nowo").value) = "" Then
                  MsgBox "No Wo kosong tidak dapat melanjutkan proses Finish", vbInformation
                  Exit Sub
               End If
            objHT.OpenTransactionBengkel
            
            Call objHT.InsertData_Bengkel("INSERT INTO WASHING VALUES ('" & Trim(listAntriCuci.Columns("nowo").value) & "',GETDATE(),'" & gSA & "','" & txtAlasanBtlWashing.text & "')")
            Call objHT.InsertData_Bengkel("UPDATE ANTRICUCI SET FLAG = 'CANCEL' WHERE NOORDER = '" & Trim(listAntriCuci.Columns("nowo").value) & "'")
            Call objHT.InsertData_Bengkel("UPDATE JOBDTL SET WASHING = 0 WHERE NO_wo = '" & Trim(listAntriCuci.Columns("nowo").value) & "'")
            Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'F' WHERE KETERANGAN LIKE '%" & Trim(listAntriCuci.Columns("nowo").value) & "%'")
            
            If erTranAll = False Then
                objHT.CloseTransactionBengkel
            End If
            MsgBox "Proses Pembatalan Washing berhasil dilakukan", vbInformation, "Informasi Pembatalan Washing"
            txtAlasanBtlWashing.text = ""
            framePembatalanWashing.Visible = False
            Call refreshGridAntriCuci
            Exit Sub
vbError:
            objHT.CloseTransactionBengkel_err
            MsgBox "Proses Pembatalan Washing gagal dilakukan", vbCritical
            Exit Sub
        End If
    Else
        MsgBox "Alasan Pembatalan Washing harus diisi!", vbCritical, "Informasi Pembatalan Washing"
    End If
End Sub

Private Sub Form_Load()
    DtArrAkhir.value = isnow
    DtArrAwal.value = isnow
    Set rsTemp = objO.FetchData("SELECT * FROM QTY_StallWashing WITH (NOLOCK) WHERE TIPESTALL = 'PENGERING'  AND ISDEL = '0'  and aktif = '1'")
    If rsTemp.RecordCount > 0 Then
        'SSTab1.Tabs = 3
        isPengering = True
        refreshGridProsesPengering
    Else
        isPengering = False
        listProsesCuci.Width = 11500
        listProsesPengering.Visible = False
        Label6.Visible = False
        Label8.Visible = False
        SSTab1.TabVisible(3) = False
        SSTab1.TabVisible(4) = False
    End If
    refreshGridSelesaiCuci
    refreshGridProsesCuci
End Sub

Private Sub listAntriCuci_Click()
    If listAntriCuci.Columns("NoWo").value <> "" Then
        asxPowerButton1.Enabled = True
    End If
End Sub

Private Sub listPermintaan_Click()
    If listPermintaan.Columns("NoWo").value <> "" Then
        btnBatal.Enabled = True
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        Call refreshGridSelesaiCuci
    ElseIf SSTab1.Tab = 1 Then
        Call refreshGridAntriCuci
    ElseIf SSTab1.Tab = 2 Then
        Call refreshGridPermintaanCuci
    ElseIf SSTab1.Tab = 3 Then
        Call refreshGridSelesaiPengeringan
    ElseIf SSTab1.Tab = 4 Then
        Call refreshGridAntriPengeringCuci
    End If
End Sub

Private Sub Timer1_Timer()
    If SSTab1.Tab = 0 Then
        Call refreshGridSelesaiCuci
    ElseIf SSTab1.Tab = 1 Then
        Call refreshGridAntriCuci
    ElseIf SSTab1.Tab = 2 Then
        Call refreshGridPermintaanCuci
    ElseIf SSTab1.Tab = 3 Then
        Call refreshGridSelesaiPengeringan
    ElseIf SSTab1.Tab = 4 Then
        Call refreshGridAntriPengeringCuci
    End If
    
    If isPengering = True Then
        Call refreshGridProsesPengering
    End If
    Call refreshGridProsesCuci
End Sub
