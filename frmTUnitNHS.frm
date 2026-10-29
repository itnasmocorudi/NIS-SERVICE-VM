VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmTUnitNHS 
   Caption         =   ".: MASTER UNIT NHS  :."
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9195
   ScaleWidth      =   16320
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Master Unit NHS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   360
      TabIndex        =   10
      Top             =   360
      Width           =   12975
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   2880
         TabIndex        =   22
         Top             =   3360
         Visible         =   0   'False
         Width           =   7215
         Begin VB.TextBox txtAlasan 
            Height          =   975
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   23
            Top             =   480
            Width           =   6975
         End
         Begin ucXPButton.XPButton cmdSimpanPerubahan 
            Height          =   495
            Left            =   4680
            TabIndex        =   25
            Top             =   1560
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
            MICON           =   "frmTUnitNHS.frx":0000
            PICN            =   "frmTUnitNHS.frx":001C
            PICH            =   "frmTUnitNHS.frx":681F
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton cmdBatalKonfirmasi 
            Height          =   495
            Left            =   5880
            TabIndex        =   26
            Top             =   1560
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
            MICON           =   "frmTUnitNHS.frx":D3C5
            PICN            =   "frmTUnitNHS.frx":D3E1
            PICH            =   "frmTUnitNHS.frx":139DA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Alasan"
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
            TabIndex        =   24
            Top             =   240
            Width           =   585
         End
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
         Index           =   2
         Left            =   3600
         MaxLength       =   60
         TabIndex        =   21
         Top             =   550
         Width           =   2535
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
         TabIndex        =   9
         Top             =   225
         Width           =   1815
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
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   550
         Width           =   1815
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
         TabIndex        =   1
         Top             =   840
         Width           =   4455
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
         Index           =   4
         Left            =   7800
         MaxLength       =   60
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   4455
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
         Index           =   5
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1450
         Width           =   3135
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   10920
         TabIndex        =   11
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
         MICON           =   "frmTUnitNHS.frx":1A289
         PICN            =   "frmTUnitNHS.frx":1A2A5
         PICH            =   "frmTUnitNHS.frx":20BD8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid gridNHS 
         Height          =   5535
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   9763
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode Unit"
         Columns(0).DataField=   "KodeUnit"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "No Polisi"
         Columns(1).DataField=   "NoPol"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No Rangka"
         Columns(2).DataField=   "NoRangka"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "No Mesin"
         Columns(3).DataField=   "NoMesin"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Model"
         Columns(4).DataField=   "Model"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Ket Model"
         Columns(5).DataField=   "KetModel"
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H93C6FD&"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
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
         TabIndex        =   13
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
         MICON           =   "frmTUnitNHS.frx":2788F
         PICN            =   "frmTUnitNHS.frx":278AB
         PICH            =   "frmTUnitNHS.frx":2E04E
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
         Left            =   6120
         TabIndex        =   14
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
         MICON           =   "frmTUnitNHS.frx":34B77
         PICN            =   "frmTUnitNHS.frx":34B93
         PICH            =   "frmTUnitNHS.frx":3B0F5
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
         TabIndex        =   8
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
         MICON           =   "frmTUnitNHS.frx":41918
         PICN            =   "frmTUnitNHS.frx":41934
         PICH            =   "frmTUnitNHS.frx":47F2D
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
         TabIndex        =   7
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
         MICON           =   "frmTUnitNHS.frx":4E7DC
         PICN            =   "frmTUnitNHS.frx":4E7F8
         PICH            =   "frmTUnitNHS.frx":54FFB
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
         Left            =   7350
         TabIndex        =   15
         Top             =   7440
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
         MICON           =   "frmTUnitNHS.frx":5BBA1
         PICN            =   "frmTUnitNHS.frx":5BBBD
         PICH            =   "frmTUnitNHS.frx":62730
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox nopol 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   3
         Top             =   1123
         Width           =   495
         _ExtentX        =   873
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox nopol 
         Height          =   285
         Index           =   1
         Left            =   2250
         TabIndex        =   4
         Top             =   1125
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox nopol 
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   5
         Top             =   1125
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
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
         Left            =   930
         TabIndex        =   20
         Top             =   1200
         Width           =   690
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   7920
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3500
         Top             =   550
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka"
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
         Left            =   720
         TabIndex        =   19
         Top             =   930
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Kendaraan"
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
         Left            =   150
         TabIndex        =   18
         Top             =   645
         Width           =   1470
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kode NHS"
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
         TabIndex        =   17
         Top             =   255
         Width           =   1305
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Mesin"
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
         Left            =   840
         TabIndex        =   16
         Top             =   1485
         Width           =   750
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12180
      _ExtentX        =   21484
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmTUnitNHS.frx":694CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmTUnitNHS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmTUnitNHS
' DateTime  : 05/01/2021 15:13
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
Dim isHapus     As Boolean
Dim isBegin     As Boolean


Private Sub ber_sih()
    For i = 0 To 5
        Text(i) = ""
    Next
    nopol(0).Mask = "CC"
    nopol(0).Text = ""
    nopol(1).Mask = "9999"
    nopol(1) = ""
    nopol(2).Mask = "CCC"
    nopol(2) = ""
    isTambah = True
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
    For i = 1 To 5
        Text(i).Locked = Not Flg
        If Flg And i <> 2 Then
            Text(i).BackColor = vbWhite
            Text(i).Locked = False
        Else
            Text(i).BackColor = &HC0E0FF
            Text(i).Locked = True
                    
        End If
    Next
    If Flg Then
        nopol(0).BackColor = vbWhite
        nopol(0).Enabled = True
        nopol(1).BackColor = vbWhite
        nopol(1).Enabled = True
        nopol(2).BackColor = vbWhite
        nopol(2).Enabled = True
    Else
        nopol(0).BackColor = &HC0E0FF
        nopol(0).Enabled = False
        nopol(1).BackColor = &HC0E0FF
        nopol(1).Enabled = False
        nopol(2).BackColor = &HC0E0FF
        nopol(2).Enabled = False
    End If
    cmdSimpan.Enabled = Flg
    cmdBatal.Enabled = Flg
    
    'chAllDealer.Enabled = Flg
   ' chAllDealer.Enabled = False
   ' chModel.Enabled = Flg
    
    cmdTambah.Enabled = Not Flg
    cmdEdit.Enabled = Not Flg
    cmdHapus.Enabled = Not Flg
End Sub
 

Private Sub cmdBatal_Click()
    Call buka_Tutup(False)
    gridNHS.Enabled = True
    Call refreshGrid
End Sub

Private Sub cmdBatalKonfirmasi_Click()
    If MsgBox("Apakah yakin membatalkan proses?", vbYesNo) = vbYes Then
        Frame2.Visible = False
        Call ber_sih
        Call buka_Tutup(False)
        Call refreshGrid
    End If
End Sub

Private Sub cmdEdit_Click()
     
    Me.MousePointer = vbHourglass
    isTambah = False
    isHapus = False
    Call buka_Tutup(True)
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdHapus_Click()
        
   On Error GoTo cmdHapus_Click_Error
   
    If MsgBox("Apakah yakin menghapus data kode NHS " & NulSID(Text(0)) & " ?", vbYesNo, "Konfirmasi Hapus Data") = vbYes Then
        isHapus = True
        txtAlasan.Text = ""
        Frame2.Visible = True
    End If

   On Error GoTo 0
   Exit Sub

cmdHapus_Click_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmTUnitNHS")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
          
1        On Error GoTo cmdSimpan_Click_Error
          Dim isBegin     As Boolean
          Dim boleh       As Boolean
          Dim pesan       As String
          Dim allModel    As Byte
          Dim allDealer   As Byte
          Dim KodeUnit    As String
          
          
2         boleh = True
3         isBegin = False
          
4         If NulSID(Text(1)) = "" Then
5             pesan = pesan & " - Kode Model harus diisi." & vbCrLf
6             boleh = False
7         Else
8             Set rsTemp = objO.FetchData("select MDL, [DESC], MODEL, FS_GRP_MDL from [" & dbReplikasiHO & "]..model WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND MDL = '" & NulSID(Text(1)) & "'")
9             If rsTemp.EOF Then
10                pesan = pesan & " - Kode Model tidak terdaftar." & vbCrLf
11                boleh = False
12            End If
13        End If
          
14        If NulSID(Text(3)) = "" Then
15            pesan = pesan & " - No Rangka harus diisi." & vbCrLf
16            boleh = False
17        Else
18            If Len(NulSID(Text(3))) < 10 Then
19                pesan = pesan & " - No Rangka minimal 10 karakter." & vbCrLf
20                boleh = False
21            Else
22                Set rsTemp = objO.FetchData("SELECT * FROM MstUnitNHS WITH (NOLOCK) WHERE KodeUnit <> '" & NulSID(Text(0)) & "' AND NORANGKA = '" & NulSID(Text(3).Text) & "' and kodecabang = '" & gDb & "'")
23                If Not rsTemp.EOF Then
24                    rsTemp.MoveFirst
25                    MsgBox "Sudah ada norangka " & NulSID(Text(3).Text) & " dengan kode NHS " & NulSID(rsTemp!KodeUnit) & ", vbInformation"
26                    Exit Sub
27                Else
28                    KodeUnit = "N" & Right(NulSID(Text(3)), 7)
29                End If
30            End If
31        End If
          
32        If NulSID(Text(5)) = "" Then
33            pesan = pesan & " - No Engine harus diisi." & vbCrLf
34            boleh = False
35        End If
          
          
36        If NulSID(nopol(1)) = "" Then
37            pesan = pesan & " - No Polisi harus diisi." & vbCrLf
38            boleh = False
39        End If
          
40        If boleh Then
41            If isTambah = False Then
42                Frame2.Visible = True
43                txtAlasan.Text = ""
44                Call buka_Tutup(False)
45                cmdTambah.Enabled = False
46                cmdEdit.Enabled = False
47                cmdHapus.Enabled = False
48                Exit Sub
49            End If
50            Call objHT.OpenTransactionBengkel
51            isBegin = True
              
52            Call ClearsSql
        
53            gsSql1(0) = "KodeUnit"
54            gsSql2(0) = NulSID(KodeUnit)
55            gsSql1(1) = "NoRangka"
56            gsSql2(1) = NulSID(Text(3))
57            gsSql1(2) = "NoMesin"
58            gsSql2(2) = NulSID(Text(5))
59            gsSql1(3) = "Model"
60            gsSql2(3) = NulSID(Text(1))
61            gsSql1(4) = "ketModel"
62            gsSql2(4) = NulSID(Text(2))
63            gsSql1(5) = "NoPol"
64            gsSql2(5) = NulSID(nopol(0)) & IIf(NulSID(nopol(0)) <> "", "-", "") & NulSID(nopol(1)) & IIf(NulSID(nopol(2)) <> "", "-", "") & NulSID(nopol(2))
                  
65            If isTambah Then
66                gsSql1(6) = "TGLENTRY"
67                gsSql2(6) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
68                gsSql1(7) = "USERENTRY"
69                gsSql2(7) = gSA
70                gsSql1(8) = "kodecabang"
71                gsSql2(8) = gDb
72            End If

73            If isTambah = True Then
74                sSQL = GensSql2("MstUnitNHS", False)
75            Else
76                sSQL = GensSql2("MstUnitNHS", True) & " WHERE KodeUnit = '" & NulSID(Text(0)) & "' "
77            End If
78            erTranAll = False
79            Call objHT.InsertData_Bengkel(sSQL)
              
80            gridNHS.Enabled = True
              
81            If erTranAll = False Then
82                Call objHT.CloseTransactionBengkel
83                MsgBox "Berhasil menyimpan data.", vbInformation
84                Call ber_sih
85                Call buka_Tutup(False)
86                Call refreshGrid
87            Else
errorRollback:
88                Call objHT.CloseTransactionBengkel_err
89                MsgBox "Gagal menyimpan data.", vbInformation
90                Exit Sub
91            End If
92        Else
93            MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
94        End If

95       On Error GoTo 0
96       Exit Sub

cmdSimpan_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
97        varErlError = Erl
98        varErrNumber = err.Number
99        If isBegin = True Then
100           Call objHT.CloseTransactionBengkel_err
101       End If
102       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmTUnitNHS")
End Sub

Private Sub cmdSimpanPerubahan_Click()
          
1        On Error GoTo cmdSimpanPerubahan_Click_Error
          Dim isBegin     As Boolean
          Dim boleh       As Boolean
          Dim pesan       As String
          Dim KodeUnit    As String
          Dim rsCek       As ADODB.Recordset
          
2         boleh = True
3         isBegin = False
          
4         If NulSID(txtAlasan.Text) = "" Then
5             boleh = False
6             pesan = pesan & " - Alasan harus diisi." & vbCrLf
7         ElseIf Len(NulSID(txtAlasan.Text)) < 5 Then
8             boleh = False
9             pesan = pesan & " - Alasan minimal 5 karakter." & vbCrLf
10        End If
          
11        If isHapus Then
12            sSQL = "SELECT * FROM PlottingPetugasNHS WITH (NOLOCK) WHERE KODEUNIT = '" & NulSID(Text(0)) & "' "
13            Set rsCek = objO.FetchData(sSQL)
14            If Not rsCek.EOF Then
15                boleh = False
16                pesan = pesan & " - Kode Unit ini masih diplotingkan ke petugas " & NulSID(rsCek!FS_KD_PETUGAS) & vbCrLf
17            End If
18        End If
19        If boleh Then
20            Call objHT.OpenTransactionBengkel
21            isBegin = True
              
              
22            pesan = ""
23            If isHapus = True Then
24                sSQL = "INSERT INTO LogMstUnitNHS (KodeUnit,nopol,noRangka,NoMesin,model,KetModel,UserEntry,TglEntry,kodeCabang,UserDel,TglDel,AlasanDel)"
25                sSQL = sSQL & " SELECT KodeUnit,nopol,noRangka,NoMesin,model,KetModel,UserEntry,TglEntry,kodeCabang,'" & gSA & "',getdate(),'" & NulSID(txtAlasan.Text) & "' FROM MstUnitNHS WITH (NOLOCK) WHERE KodeUnit = '" & NulSID(Text(0)) & "'"
26                Call objHT.InsertData_Bengkel(sSQL)
                  
27                sSQL = "DELETE FROM MstUnitNHS WHERE KodeUnit = '" & NulSID(Text(0)) & "' "
28                Call objHT.InsertData_Bengkel(sSQL)
29                pesan = "Data berhasil dihapus."
30            Else
31                sSQL = "INSERT INTO LogMstUnitNHS (KodeUnit,nopol,noRangka,NoMesin,model,KetModel,UserEntry,TglEntry,kodeCabang,UserEdit,TglEdit,AlasanEdit)"
32                sSQL = sSQL & " SELECT KodeUnit,nopol,noRangka,NoMesin,model,KetModel,UserEntry,TglEntry,kodeCabang,'" & gSA & "',getdate(),'" & NulSID(txtAlasan.Text) & "' FROM MstUnitNHS WITH (NOLOCK) WHERE KodeUnit = '" & NulSID(Text(0)) & "'"
33                Call objHT.InsertData_Bengkel(sSQL)
                  
34                sSQL = "UPDATE MstUnitNHS SET nopol = '" & NulSID(nopol(0)) & IIf(NulSID(Text(0)) <> "", "-", "") & NulSID(nopol(1)) & IIf(NulSID(nopol(2)) <> "", "-", "") & NulSID(nopol(2)) & "',noRangka = '" & NulSID(Text(3)) & "',NoMesin = '" & NulSID(Text(5)) & "',model = '" & NulSID(Text(1)) & "',KetModel = '" & NulSID(Text(2)) & "' WHERE KodeUnit = '" & NulSID(Text(0)) & "' "
35                Call objHT.InsertData_Bengkel(sSQL)
36                pesan = "Data berhasil disimpan."
37            End If
              
38            Call objHT.CloseTransactionBengkel
39            gridNHS.Enabled = True
40            MsgBox pesan, vbInformation
41            Frame2.Visible = False
42            Call ber_sih
43            Call buka_Tutup(False)
44            Call refreshGrid
45        Else
46            MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
47        End If
48       On Error GoTo 0
49       Exit Sub

cmdSimpanPerubahan_Click_Error:
          
50        If isBegin Then
51            Call objHT.CloseTransactionBengkel_err
52        End If
53        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpanPerubahan_Click of Form frmTUnitNHS"


End Sub

Private Sub cmdTambah_Click()
    isTambah = True
    Call ber_sih
    ''Call isiId
    Call buka_Tutup(True)
    gridNHS.Enabled = False
End Sub
 

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False)
    Call refreshGrid
End Sub

 
Private Sub gridnhs_FilterChange()
On Error GoTo errHandler
    Set cols = gridNHS.Columns
    Dim C As Integer
    
    C = gridNHS.col
    gridNHS.HoldFields
    rsGrid.filter = getFilter()
    gridNHS.col = C
    gridNHS.EditActive = True
    Exit Sub
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub gridnhs_HeadClick(ByVal ColIndex As Integer)

1        On Error GoTo gridnhs_HeadClick_Error
2         Urut = IIf(Urut = "DESC", "ASC", "DESC")
3         rsGrid.Sort = gridNHS.Columns(ColIndex).DataField & " " & Urut
4         gridNHS.SetFocus
5         Sendkeys "{Right " & Val(ColIndex) & "}"
          
6        On Error GoTo 0
7        Exit Sub

gridnhs_HeadClick_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
8         varErlError = Erl
9         varErrNumber = err.Number
10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure gridnhs_HeadClick of Form frmTUnitNHS")
End Sub

Private Sub gridnhs_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF Then
        Dim nPol As String
        Dim nDpn1 As String
        Dim nTgh1 As String
        Dim nBlk1 As String
        nPol = NulSID(rsGrid("NoPol"))
        nDpn1 = ""
        nTgh1 = ""
        nBlk1 = ""
        Call pecahNoPol(Trim(nPol), nDpn1, nTgh1, nBlk1)
        nopol(0) = Trim(nDpn1)
        nopol(1) = Trim(nTgh1)
        nopol(2) = Trim(nBlk1)
        Text(0).Text = NulSID(rsGrid!KodeUnit)
        Text(1).Text = NulSID(rsGrid!model)
        Text(2).Text = NulSID(rsGrid!KetModel)
        Text(3).Text = NulSID(rsGrid!noRangka)
        Text(5).Text = NulSID(rsGrid!noMesin)
      
    End If
End Sub


Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
    Set rsTmp = objO.FetchData("SELECT A.fs_kd_dealer, A.fs_nm_dealer, ISNULL(b.KDSUBLET,'') AS KODE FROM [" & dbReplikasiHO & "]..t_dealer a WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETDEALER b ON a.fs_kd_dealer = b.FS_KD_DEALER and A.FS_LOG_DEL = '' AND b.KDSUBLET = '" & NulSID(Text(0)) & "' WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
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


Private Sub isiModel()
    Set rsTemp = objO.FetchData("SELECT DISTINCT a.MODEL, a.mdl, a.[desc], ISNULL(kdsublet, '') as PRO FROM [" & dbReplikasiHO & "]..model a WITH (NOLOCK) LEFT JOIN [" & dbReplikasiHO & "]..MSTSUBLETMDL b WITH (NOLOCK) ON a.MDL = b.MDL AND B.kdsublet = '" & NulSID(Text(0)) & "' ORDER BY a.MODEL, a.mdl")
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

Private Sub refreshGrid()
    ''gridNHS.ClearFields
    sSQL = "SELECT KodeUnit,nopol,noRangka,NoMesin,model,KetModel FROM MstUnitNHS WITH (NOLOCK) WHERE kodecabang = '" & gDb & "'"
    Set rsGrid = objO.FetchData(sSQL)
    Set gridNHS.DataSource = rsGrid
End Sub


Private Sub nopol_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            'Text(Index).Mask = ""
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            If Index = 1 Then
                Set rsTemp = objO.FetchData("select MDL, [DESC], MODEL, FS_GRP_MDL from [" & dbReplikasiHO & "]..model WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND MDL = '" & NulSID(Text(1)) & "'")
                If rsTemp.EOF Then
                    MsgBox " - Kode model tidak terdaftar.", vbInformation
                    Text(1).Text = ""
                Else
                    Text(2).Text = NulSID(rsTemp!desc)
                End If
            End If
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
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchData("SELECT TOP 1 * FROM [" & dbReplikasiHO & "]..MSTSUBLET WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY RIGHT(KDSUBLET,10) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KDSUBLET"), 4))
        noUrut = noUrut + 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
    Else
        noUrut = 1
        Text(0) = Replace(gKdCbg, " ", "") & Right(year(isnow), 2) & Format(Month(isnow), "00") & Format(noUrut, "00000#")
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
    
    If KeyAscii = 13 And cmdSimpan.Enabled = True Then
        Select Case Index
            Case 1
                Set rsTemp = objO.FetchData("select MDL, [DESC], MODEL, FS_GRP_MDL from [" & dbReplikasiHO & "]..model WITH (NOLOCK) WHERE FS_LOG_DEL = '' AND MDL = '" & NulSID(Text(1)) & "'")
                
                If rsTemp.EOF Then
                    MsgBox " - Kode model tidak terdaftar.", vbInformation
                    Text(1).Text = ""
                    Text(2).Text = ""
                Else
                    Text(1).Text = NulSID(rsTemp!MDL)
                    Text(2).Text = NulSID(rsTemp!desc)
                End If
        End Select
    End If
End Sub
 
Private Sub Text_LostFocus(Index As Integer)
    If Index = 1 Then
        Call text_KeyPress(1, 13)
    End If
End Sub

Private Function getFilter() As String
Dim tmp As String
Dim n As Integer
   Dim col As TrueOleDBGrid70.Columns
   
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
 
