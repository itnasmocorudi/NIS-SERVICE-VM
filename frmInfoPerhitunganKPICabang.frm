VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmInfoPerhitunganKPICabang 
   Caption         =   ".: Perhitungan KPI Cabang :."
   ClientHeight    =   6375
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   12930
   ScaleWidth      =   21360
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10455
      Left            =   0
      Picture         =   "frmInfoPerhitunganKPICabang.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   15
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8535
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   12615
      Begin VB.TextBox Text 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1695
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3120
         MaxLength       =   60
         TabIndex        =   16
         Top             =   960
         Width           =   3615
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   10920
         TabIndex        =   9
         Top             =   7440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInfoPerhitunganKPICabang.frx":8C4A
         PICN            =   "frmInfoPerhitunganKPICabang.frx":8C66
         PICH            =   "frmInfoPerhitunganKPICabang.frx":F599
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   5280
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1635
         Width           =   1455
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   5
         Top             =   1635
         Width           =   1815
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   2
         Top             =   960
         Width           =   1455
      End
      Begin TrueOleDBGrid70.TDBGrid GridSupplier 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   2040
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9340
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode KPI"
         Columns(0).DataField=   "KodeKPI"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KPI"
         Columns(1).DataField=   "KPI"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Target"
         Columns(2).DataField=   "Target"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Actual"
         Columns(3).DataField=   "Actual"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Achievement"
         Columns(4).DataField=   "Achievement"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "SetAchv"
         Columns(5).DataField=   "SetAchv"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   6
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=6"
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
         Splits(0)._ColumnProps(24)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(25)=   "Column(5).Order=6"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
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
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFF80&"
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
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
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
         _StyleDefs(74)  =   ":id=39,.parent=33,.bgcolor=&H4080FF&"
         _StyleDefs(75)  =   "Named:id=40:OddRow"
         _StyleDefs(76)  =   ":id=40,.parent=33"
         _StyleDefs(77)  =   "Named:id=41:RecordSelector"
         _StyleDefs(78)  =   ":id=41,.parent=34"
         _StyleDefs(79)  =   "Named:id=42:FilterBar"
         _StyleDefs(80)  =   ":id=42,.parent=33"
         _StyleDefs(81)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(82)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(83)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(84)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   495
         Left            =   4920
         TabIndex        =   10
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInfoPerhitunganKPICabang.frx":16250
         PICN            =   "frmInfoPerhitunganKPICabang.frx":1626C
         PICH            =   "frmInfoPerhitunganKPICabang.frx":1CA0F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdEdit 
         Height          =   495
         Left            =   7320
         TabIndex        =   11
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInfoPerhitunganKPICabang.frx":23538
         PICN            =   "frmInfoPerhitunganKPICabang.frx":23554
         PICH            =   "frmInfoPerhitunganKPICabang.frx":29AB6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdBatal 
         Height          =   495
         Left            =   9720
         TabIndex        =   12
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInfoPerhitunganKPICabang.frx":302D9
         PICN            =   "frmInfoPerhitunganKPICabang.frx":302F5
         PICH            =   "frmInfoPerhitunganKPICabang.frx":368EE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   8520
         TabIndex        =   13
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInfoPerhitunganKPICabang.frx":3D19D
         PICN            =   "frmInfoPerhitunganKPICabang.frx":3D1B9
         PICH            =   "frmInfoPerhitunganKPICabang.frx":439BC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapus 
         Height          =   495
         Left            =   6120
         TabIndex        =   14
         Top             =   7440
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
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
         MICON           =   "frmInfoPerhitunganKPICabang.frx":4A562
         PICN            =   "frmInfoPerhitunganKPICabang.frx":4A57E
         PICH            =   "frmInfoPerhitunganKPICabang.frx":510F1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBDate6Ctl.TDBDate Periode 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   300
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calendar        =   "frmInfoPerhitunganKPICabang.frx":57E8B
         Caption         =   "frmInfoPerhitunganKPICabang.frx":57FB7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmInfoPerhitunganKPICabang.frx":58023
         Keys            =   "frmInfoPerhitunganKPICabang.frx":58041
         Spin            =   "frmInfoPerhitunganKPICabang.frx":5809F
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm-yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/yyyy"
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
         Text            =   "11/2024"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   45611
         CenturyMode     =   0
      End
      Begin ucXPButton.XPButton cmdViewKPI 
         Height          =   495
         Left            =   2520
         TabIndex        =   19
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInfoPerhitunganKPICabang.frx":580C7
         PICN            =   "frmInfoPerhitunganKPICabang.frx":580E3
         PICH            =   "frmInfoPerhitunganKPICabang.frx":5C144
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdUpdate 
         Height          =   495
         Left            =   6840
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmInfoPerhitunganKPICabang.frx":60433
         PICN            =   "frmInfoPerhitunganKPICabang.frx":6044F
         PICH            =   "frmInfoPerhitunganKPICabang.frx":632D6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Aktual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   20
         Top             =   1680
         Width           =   555
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   12600
         X2              =   120
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "KPI"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1350
         TabIndex        =   17
         Top             =   1005
         Width           =   285
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
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
         Left            =   315
         TabIndex        =   4
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Target"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1065
         TabIndex        =   6
         Top             =   1365
         Width           =   570
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Achievement"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4095
         TabIndex        =   8
         Top             =   1680
         Width           =   1125
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
End
Attribute VB_Name = "frmInfoPerhitunganKPICabang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmInfoPerhitunganKPICabang
' DateTime  : 19/09/2018
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------


Dim rsGrid      As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset
Dim cols        As TrueOleDBGrid70.Columns
Dim Urut        As String
Dim isTambah    As Boolean
Dim isBegin     As Boolean


Private Sub ber_sih()
    For i = 0 To 4
        Text(i) = ""
    Next
    cmdUpdate.Enabled = True
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
'    For i = 1 To 4
'        Text(i).Locked = Not Flg
'        If Flg Then
'            Text(i).BackColor = vbWhite
'        Else
'            Text(i).BackColor = &HC0E0FF
'        End If
'    Next
    cmdSimpan.Enabled = Flg
    cmdBatal.Enabled = Flg
    
   ' cmdTambah.Enabled = Not Flg
    cmdEdit.Enabled = Not Flg
    cmdHapus.Enabled = Not Flg
End Sub

Private Sub chAllDealer_Click()
     
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

Private Sub chModel_Click()
    If chModel.value = Checked Then
        frmeModel.Visible = False
    ElseIf chModel.value = Unchecked Then
        frmeModel.Visible = True
        Call isiModel
    End If
End Sub

Private Sub cmdBatal_Click()
    Call buka_Tutup(False)
    Call refreshGrid
End Sub

Private Sub cmdEdit_Click()
1         isTambah = False
2         Call buka_Tutup(True)
3         If NulSID(GridSupplier.Columns(8).value) = "NR 00" Then
4             For i = 1 To 8
5                 Text(i).Locked = True
6                 Text(i).BackColor = &HC0E0FF
7             Next
8             Text(10).Locked = True
9             Text(10).BackColor = &HC0E0FF
10        ElseIf NulSID(GridSupplier.Columns(8).value) <> NulSID(gKdCbg) Then
11            For i = 1 To 8
12                Text(i).Locked = True
13                Text(i).BackColor = &HC0E0FF
14            Next
15            Text(10).Locked = True
16            Text(10).BackColor = &HC0E0FF
17        End If
End Sub

Private Sub cmdHapus_Click()
1        On Error GoTo cmdHapus_Click_Error
2         isBegin = False
3         If NulSID(GridSupplier.Columns(8).value) = "NR 00" Then
4             MsgBox "Master supplier HO tidak dapat di hapus!", vbInformation
5             Exit Sub
6         ElseIf NulSID(GridSupplier.Columns(8).value) <> NulSID(gKdCbg) Then
7             MsgBox "Master supplier Dealer lain tidak dapat di hapus!", vbInformation
8             Exit Sub
9         End If
          
10        Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE ispakai = '1' and KDSUPP = '" & NulSID(Text(0)) & "'")
11        If Not rsTemp.EOF Then
12            MsgBox "Sudah Ada Transaksi Work Order untuk sublet ini" & vbCrLf & "Kode Supplier " & NulSID(Text(0)) & " tidak dapat dihapus", vbInformation
13            Exit Sub
14        End If
15        If MsgBox("Apakah yakin menghapus data kode supplier " & NulSID(Text(0)) & " ?", vbYesNo, "Konfirmasi Hapus Data") = vbYes Then
16            Set rsTemp = objO.FetchDataHO("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..mstsublet WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND KDSUPP = '" & NulSID(Text(0)) & "'")
17            If Not rsTemp.EOF Then
18                MsgBox "Masih ada sublet untuk supplier ini." & vbCrLf & "Kode Sublet " & NulSID(rsTemp!KDSUBLET) & " masih aktif tidak dapat menghapus supplier", vbInformation
19                Exit Sub
20            End If
21            Call objHT.OpenTransactionBengkel
22            isBegin = True
23            erTranAll = True
24            sSQL = "UPDATE [" & dbReplikasiHO & "]..MSTSUPPLIER SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = GETDATE() ,STATUS = 0 WHERE KDSUPP = '" & NulSID(Text(0)) & "'"
25            Call objHT.InsertData_Bengkel(sSQL)
26            sSQL = "UPDATE T_SUPPLIER SET FS_LOG_DEL = '" & gSA & "',FD_TGL_DEL = GETDATE() ,STATUS = 0 WHERE KDSUPP = '" & NulSID(Text(0)) & "'"
27            Call objHT.InsertData_Bengkel(sSQL)
              
28            If erTranAll Then
29                Call objHT.CloseTransactionBengkel
30                MsgBox "Data berhasil dihapus", vbInformation
31                Call refreshGrid
32                Call buka_Tutup(False)
33            Else
34                Call objHT.CloseTransactionBengkel_err
35                MsgBox "Gagal hapus data", vbInformation
36                Exit Sub
37            End If
38        End If

39       On Error GoTo 0
40       Exit Sub

cmdHapus_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
41        varErlError = Erl
42        varErrNumber = err.Number
43        If isBegin Then
44            Call objHT.CloseTransactionBengkel_err
45        End If
46        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmInfoPerhitunganKPICabang")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
1        On Error GoTo cmdSimpan_Click_Error
          Dim boleh           As Boolean
          Dim pesan           As String
          Dim rsTemp          As ADODB.Recordset
          Dim namaSupplier    As String
2         boleh = True
3         isBegin = False
4         If NulSID(Text(1)) = "" Then
5             pesan = pesan & "Nama supplier harus diisi." & vbCrLf
6             boleh = False
7         Else
8             namaSupplier = Replace(NulSID(Text(1)), " ", "%")
              ''Set rsTemp = objO.FetchData("select KDSUPP,SUPPLIER,KDDEALER from [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE SUPPLIER LIKE '%" & namaSupplier & "%' AND KDSUPP <> '" & NulSID(Text(0)) & "'")
9             Set rsTemp = objO.FetchData("select KDSUPP,SUPPLIER,KDDEALER from [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE SUPPLIER = '" & NulSID(Text(1)) & "' AND KDSUPP <> '" & NulSID(Text(0)) & "' AND (SUBSTRING(KDSUPP,1,4) = 'NR00' OR SUBSTRING(KDSUPP,1,4) = '" & Replace(gKdCbg, " ", "") & "') ")
10            If Not rsTemp.EOF Then
11                MsgBox "Nama supplier mirip : " & NulSID(rsTemp!supplier) & " Dealer " & NulSID(rsTemp!kdDealer), vbInformation
12                pesan = pesan & " Nama supplier kembar" & vbCrLf
13                boleh = False
                   
14            End If
15        End If
16        If NulSID(Text(2)) = "" Then
17            pesan = pesan & "Alamat 1 supplier harus diisi." & vbCrLf
18            boleh = False
19        End If
          
20        If chPKP.value = vbChecked Then
              ''penambahan master supplier dgn memasukkan no npwp sbg kuncian ( khusus supplier yang ditandain PKP)
21            If NulSID(Text(10).Text) = "" Then
22                pesan = pesan & "No NPWP harus diisi." & vbCrLf
23                boleh = False
24            Else
25                If ckAfiliasi.value = vbUnchecked Then
26                    Set rsTemp = objO.FetchData("select KDSUPP,SUPPLIER,KDDEALER,NPWP from [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE NPWP = '" & NulSID(Text(10)) & "' AND KDSUPP <> '" & NulSID(Text(0)) & "' AND (SUBSTRING(KDSUPP,1,4) = 'NR00' OR SUBSTRING(KDSUPP,1,4) = '" & Replace(gKdCbg, " ", "") & "')")
27                    If Not rsTemp.EOF Then
28                        MsgBox "No NPWP : " & NulSID(rsTemp!npwp) & " telah digunakan Kode Supplier " & NulSID(rsTemp!KDSUPP), vbInformation
29                        pesan = pesan & " No NPWP kembar" & vbCrLf
30                        boleh = False
31                    End If
32                End If
33            End If
34        Else
35            If NulSID(Text(10).Text) <> "" Then
36                If ckAfiliasi.value = vbUnchecked Then
37                    Set rsTemp = objO.FetchData("select KDSUPP,SUPPLIER,KDDEALER,NPWP from [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE NPWP = '" & NulSID(Text(10)) & "' AND KDSUPP <> '" & NulSID(Text(0)) & "'")
38                    If Not rsTemp.EOF Then
39                        MsgBox "No NPWP : " & NulSID(rsTemp!npwp) & " telah digunakan Kode Supplier " & NulSID(rsTemp!KDSUPP), vbInformation
40                        pesan = pesan & " No NPWP kembar" & vbCrLf
41                        boleh = False
42                    End If
43                End If
44            End If
45        End If
          
46        If NulSID(Text(4)) = "" Then
47            pesan = pesan & "Kota supplier harus diisi." & vbCrLf
48            boleh = False
49        End If
          
50        If NulSID(Text(5)) = "" Then
51            pesan = pesan & "Telepon supplier harus diisi." & vbCrLf
52            boleh = False
53        End If
          
54        If NulSID(Text(6)) = "" Then
55            pesan = pesan & "Contact Person supplier harus diisi."
56            boleh = False
57        End If
          
58        If NulSID(Text(9)) = "" Then
59            If isTambah Then
60                pesan = pesan & "Kodelgn harus diisi." & vbCrLf
61                boleh = False
62            Else
63                If MsgBox("Apakah ingin mengosongkan kodelgn supplier?", vbYesNo) = vbNo Then
                   
64                    pesan = pesan & "Kodelgn harus diisi." & vbCrLf
65                    boleh = False
66                End If
67            End If
              
68        Else
69            sSQL = "SELECT KODELGN FROM APMST WITH (NOLOCK) WHERE KODELGN = '" & NulSID(Text(9)) & "' AND NONAKTIFFLAG = '0'"
70            If objO.FetchData_String_Acc(sSQL) = "" Then
71                pesan = pesan & "Kodelgn " & NulSID(Text(9)) & " tidak ditemukan." & vbCrLf
72                boleh = False
73            Else
74                sSQL = "SELECT KDSUPP FROM  T_SUPPLIER WITH (NOLOCK) WHERE KODEAP = '" & NulSID(Text(9)) & "' AND KDSUPP <> '" & NulSID(Text(0)) & "'"
75                If objO.FetchData_String(sSQL) <> "" Then
76                    pesan = pesan & "Kodelgn " & NulSID(Text(9)) & " digunakan kode supplier lain." & vbCrLf
77                    boleh = False
78                End If
79            End If
80        End If
          
81        If boleh Then
82            Call ClearsSql
        
83            gsSql1(0) = "KDSUPP"
84            gsSql2(0) = NulSID(Text(0))
85            gsSql1(1) = "SUPPLIER"
86            gsSql2(1) = NulSID(Text(1))
87            gsSql1(2) = "ALAMAT1"
88            gsSql2(2) = NulSID(Text(2))
89            gsSql1(3) = "ALAMAT2"
90            gsSql2(3) = NulSID(Text(3))
91            gsSql1(4) = "Kota"
92            gsSql2(4) = NulSID(Text(4))
93            gsSql1(5) = "TELPON"
94            gsSql2(5) = NulSID(Text(5))
95            gsSql1(6) = "CP"
96            gsSql2(6) = NulSID(Text(6))
97            gsSql1(7) = "kdDealer"
98            gsSql2(7) = IIf(isTambah, NulSID(gKdCbg), NulSID(rsGrid!kdDealer))
99            gsSql1(8) = "CABANG"
100           gsSql2(8) = IIf(isTambah, NulSID(gDb), NulSID(rsGrid!cabang)) 'NulSID(gDb)
101           gsSql1(9) = "STATUS"
102           gsSql2(9) = "1"
103           gsSql1(10) = "KETERANGAN"
104           gsSql2(10) = NulSID(Text(7))
105           If isTambah = True Then
106               gsSql1(11) = "TGLENTRY"
107               gsSql2(11) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
108               gsSql1(12) = "USERENTRY"
109               gsSql2(12) = NulSID(gSA)
110           Else
111               gsSql1(11) = "TGLUPDATE"
112               gsSql2(11) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
113               gsSql1(12) = "USERUPDATE"
114               gsSql2(12) = NulSID(gSA)
115           End If
              
116           gsSql1(13) = "PKP"
117           gsSql2(13) = IIf(chPKP, "1", "0")
118           gsSql1(14) = "NONJURNAL"
119           gsSql2(14) = IIf(chJurnal, "1", "0")
120           gsSql1(15) = "ISPAKAI"
121           gsSql2(15) = "1"
122           gsSql1(16) = "NPWP"
123           gsSql2(16) = NulSID(Text(10))
124           gsSql1(17) = "FS_LOG_DEL"
125           gsSql2(17) = ""
126           gsSql1(18) = "FD_TGL_DEL"
127           gsSql2(18) = "1900-01-01"
128           gsSql1(19) = "aflflag"
129           gsSql2(19) = IIf(ckAfiliasi, "1", "0")
              ''karena default value d replbengkel hilang
130           If isTambah = True Then
131             gsSql1(20) = "TGLUPDATE"
132             gsSql2(20) = "1900-01-01"
133             gsSql1(21) = "USERUPDATE"
134             gsSql2(21) = ""
135           End If
136           If isTambah = True Then
137               sSQL = GensSql2("[" & dbReplikasiHO & "]..MSTSUPPLIER", False)
138           Else
139               sSQL = GensSql2("[" & dbReplikasiHO & "]..MSTSUPPLIER", True) & " WHERE KDSUPP = '" & NulSID(Text(0)) & "' "
140           End If
141           Call objHT.OpenTransactionBengkel
142           isBegin = True
143           erTranAll = False
              
144           Call objHT.InsertData_Bengkel(sSQL)
              
145           Set rsTemp = objHT.FetchDataBengkel("SELECT * FROM T_SUPPLIER WITH (NOLOCK) WHERE KDSUPP = '" & NulSID(Text(0)) & "'")
146           isTambah = False
147           If rsTemp.EOF Then isTambah = True
              
148           Call ClearsSql
        
149           gsSql1(0) = "KDSUPP"
150           gsSql2(0) = NulSID(Text(0))
151           gsSql1(1) = "SUPPLIER"
152           gsSql2(1) = NulSID(Text(1))
153           gsSql1(2) = "ALAMAT1"
154           gsSql2(2) = NulSID(Text(2))
155           gsSql1(3) = "ALAMAT2"
156           gsSql2(3) = NulSID(Text(3))
157           gsSql1(4) = "Kota"
158           gsSql2(4) = NulSID(Text(4))
159           gsSql1(5) = "TELPON"
160           gsSql2(5) = NulSID(Text(5))
161           gsSql1(6) = "CP"
162           gsSql2(6) = NulSID(Text(6))
163           gsSql1(7) = "kdDealer"
164           gsSql2(7) = IIf(isTambah, NulSID(gKdCbg), NulSID(rsGrid!kdDealer)) 'NulSID(gKdCbg)
165           gsSql1(8) = "CABANG"
166           gsSql2(8) = IIf(isTambah, NulSID(gDb), NulSID(rsGrid!cabang)) 'NulSID(gDb)
167           gsSql1(9) = "STATUS"
168           gsSql2(9) = "1"
169           gsSql1(10) = "KETERANGAN"
170           gsSql2(10) = NulSID(Text(7))
171           If isTambah = True Then
172               gsSql1(11) = "TGLENTRY"
173               gsSql2(11) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
174               gsSql1(12) = "USERENTRY"
175               gsSql2(12) = NulSID(gSA)
176           Else
177               gsSql1(11) = "TGLUPDATE"
178               gsSql2(11) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
179               gsSql1(12) = "USERUPDATE"
180               gsSql2(12) = NulSID(gSA)
181           End If
              
182           gsSql1(13) = "PKP"
183           gsSql2(13) = IIf(chPKP, "1", "0")
184           gsSql1(14) = "NONJURNAL"
185           gsSql2(14) = IIf(chJurnal, "1", "0")
186           gsSql1(15) = "KODEAP"
187           gsSql2(15) = NulSID(Text(9))
188           gsSql1(16) = "NPWP"
189           gsSql2(16) = NulSID(Text(10))
190           gsSql1(17) = "aflflag"
191           gsSql2(17) = IIf(ckAfiliasi, "1", "0")
              
              
192           If isTambah = True Then
193               sSQL = GensSql2("T_SUPPLIER", False)
194           Else
195               sSQL = GensSql2("T_SUPPLIER", True) & " WHERE KDSUPP = '" & NulSID(Text(0)) & "' "
196           End If
197           Call objHT.InsertData_Bengkel(sSQL)
              
198           If erTranAll = False Then
199               Call objHT.CloseTransactionBengkel
200               MsgBox "Data berhasil disimpan", vbInformation
201               Call refreshGrid
202               Call buka_Tutup(False)
203           Else
204               If isBegin Then
205                   Call objHT.CloseTransactionBengkel_err
206               End If
207               MsgBox "Data gagal disimpan.", vbInformation
208               Exit Sub
209           End If
210       Else
211           MsgBox "Lengkapi data berikut :" & vbCrLf & pesan, vbCritical, "KELENGKAPAN"
212       End If
213      On Error GoTo 0
214      Exit Sub
cmdSimpan_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
215       varErlError = Erl
216       varErrNumber = err.Number
217       If isBegin = True Then
218           Call objHT.CloseTransactionBengkel_err
219       End If
220       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmInfoPerhitunganKPICabang")
End Sub

Private Sub cmdTambah_Click()
    isTambah = True
    Call ber_sih
    Call buka_Tutup(True)
    Call isiId
End Sub

Private Sub cmdTutup_Click()
    frmeModel.Visible = False
End Sub

Private Sub cmdTutupDealer_Click()

    frmeDealer.Visible = False
End Sub

Private Sub cmdUpdate_Click()
     Dim target, aktual, ach As Double
    target = k2t(Text(2))
    aktual = k2t(Text(3))
    If target > 0 Then
        ach = Round((aktual / target) * 100, 2)
    Else
        ach = 0
    End If
    Call objO.InsertData("UPDATE [INSENTIF].dbo.grKPIResultCbg SET TargetKPICbg = '" & k2t(target) & "',ActKpiCbg = '" & k2t(aktual) & "',AchvKpiCbg = '" & k2t(ach) & "' where KdKPI = '" & NulSID(Text(0)) & "' and kategori='0' and Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_DEALER = '" & NulSID(gKdCbg) & "'")
    Call cmdViewKPI_Click
End Sub

Private Sub cmdViewKPI_Click()
    Call ber_sih
    Call buka_Tutup(False)
    Call refreshGrid
End Sub

Private Sub Form_Load()
    Periode.value = Format(Date, "mm-yyyy")
    Call ber_sih
    Call buka_Tutup(False)
    Call refreshGrid
End Sub


Private Sub GridSupplier_FilterChange()
    On Error GoTo errHandler
    Set cols = GridSupplier.Columns
    Dim C As Integer
    
    C = GridSupplier.col
    GridSupplier.HoldFields
    rsGrid.filter = getFilter()
    GridSupplier.col = C
    GridSupplier.EditActive = True
    Exit Sub
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub GridSupplier_HeadClick(ByVal ColIndex As Integer)
   On Error GoTo GridSupplier_HeadClick_Error

    Urut = IIf(Urut = "DESC", "ASC", "DESC")
    rsGrid.Sort = GridSupplier.Columns(ColIndex).DataField & " " & Urut
    GridSupplier.SetFocus
    Sendkeys "{Right " & Val(ColIndex) & "}"
 
   Exit Sub

GridSupplier_HeadClick_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure GridSupplier_HeadClick of Form frmInfoPerhitunganKPICabang")
 End Sub

Private Sub gridSupplier_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF Then
        Text(0).Text = NulSID(rsGrid!kodeKPI)
        Text(1).Text = NulSID(rsGrid!KPI)
        Text(2).Text = NulSID(rsGrid!target)
        Text(3).Text = NulSID(rsGrid!actual)
        Text(4).Text = NulSID(rsGrid!Achievement)
        If rsGrid!setAchv = 1 Then
            Text(2).Enabled = True
            Text(3).Enabled = True
        Else
            Text(2).Enabled = False
            Text(3).Enabled = False
        End If
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
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
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    'KeyCode = UCase(KeyCode)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 9
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLGAP"
                frmCariAcess.show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) ORDER BY RIGHT(KDSUPP,10) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KDSUPP"), 4))
        noUrut = noUrut + 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    End If
End Sub

Private Sub isiModel()
    Set rsTemp = objO.FetchData("SELECT DISTINCT a.MODEL, a.mdl, a.[desc], ISNULL(id_program, '') as PRO FROM [" & dbReplikasiHO & "]..model a WITH (NOLOCK) LEFT JOIN ho_PROmdl b WITH (NOLOCK) ON a.MDL = b.MDL AND B.ID_PROGRAM = '" & Trim(Text(0)) & "' ORDER BY a.MODEL, a.mdl")
    If Not rsTemp.EOF Then
        Call createTbl
        Do While Not rsTemp.EOF
            rsGridMdl.AddNew
            rsGridMdl!grpMDL = Trim(rsTemp("MODEL"))
            rsGridMdl!kodeMDL = Trim(rsTemp("MDL"))
            rsGridMdl!NAMAMDL = Trim(rsTemp("DESC"))
            rsGridMdl!ikut = IIf(rsTemp("PRO") = "", False, True)
            rsGridMdl.Update
            rsTemp.MoveNext
        Loop

        rsGridMdl.MoveFirst
        GridModel.Columns(0).Width = 1200
        GridModel.Columns(1).Width = 2000
        GridModel.Columns(2).Width = 4000
        GridModel.Columns(3).Alignment = dbgCenter
        GridModel.Columns(3).Width = 800
    End If
End Sub

Private Sub createTbl()
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    With rsGridMdl
        .Fields.Append "GRPMDL", adVarChar, 15
        .Fields.Append "KODEMDL", adVarChar, 20
        .Fields.Append "NAMAMDL", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridModel.DataSource = rsGridMdl
    GridModel.AlternatingRowStyle = True
End Sub

Private Sub GridModel_FilterChange()
    On Error GoTo errHandler
    Set cols = GridModel.Columns
    Dim C As Integer
    
    C = GridModel.col
    GridModel.HoldFields
    rsGridMdl.filter = getFilter()
    GridModel.col = C
    GridModel.EditActive = True
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


Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
    Set rsTmp = objO.FetchDataHO("SELECT fs_kd_dealer, fs_nm_dealer, ISNULL(ID_PROGRAM,'') AS KODE FROM t_dealer a WITH (NOLOCK) " _
                            & "LEFT JOIN HO_PROCBG b ON a.fs_kd_dealer = b.kodedealer and A.FS_LOG_DEL = '' AND b.ID_PROGRAM = '" & Text(0) & "' WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
    If Not rsTmp.EOF Then
        Call createRS
        Do While Not rsTmp.EOF
            rsDealer.AddNew
            rsDealer!KodeDealer = rsTmp("FS_KD_DEALER")
            rsDealer!namaDealer = rsTmp("FS_NM_DEALER")
            If rsTmp("KODE") <> "" Then
                rsDealer!terlibat = True
            Else
                rsDealer!terlibat = False
            End If
            rsDealer.Update
            rsTmp.MoveNext
        Loop
        
        GridDealer.DataSource = rsDealer
        GridDealer.EvenRowStyle.BackColor = vbYellow
        GridDealer.OddRowStyle.BackColor = vbWhite
        GridDealer.AlternatingRowStyle = True
    End If
End Sub

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 3500
    GridDealer.Columns(2).Width = 1000
    GridDealer.Columns(2).Alignment = dbgCenter
End Sub

Private Sub createRS()
    
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    With rsDealer
        .Fields.Append "KODEDEALER", adVarChar, 15
        .Fields.Append "NAMADEALER", adVarChar, 25
        .Fields.Append "TERLIBAT", adBoolean
        .Open
    End With
    
    Set GridDealer.DataSource = rsDealer
End Sub

Private Sub refreshGrid()
    GridSupplier.ClearFields

    ''sSQL = "SELECT KDSUPP,SUPPLIER,ALAMAT1,ALAMAT2,KOTA,TELPON,CP,KETERANGAN,KDDEALER,PKP,NONJURNAL FROM [" & dbReplikasiHO & "]..MSTSUPPLIER WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY KDDEALER "
    
    sSQL = "Select KdKPI as [KodeKPI], NamaKPI as KPI, TargetKPICbg as Target, ActKpiCbg as Actual, AchvKpiCbg as Achievement,SetAchv from [INSENTIF].dbo.grKPIResultCbg where kategori='0' and Periode = '" & Format(Periode.value, "MM-YYYY") & "' and FS_KD_DEALER = '" & NulSID(gKdCbg) & "'"
    Set rsGrid = objO.FetchData(sSQL)
    Set GridSupplier.DataSource = rsGrid
End Sub
 
Private Sub XPButton1_Click()
 
End Sub

