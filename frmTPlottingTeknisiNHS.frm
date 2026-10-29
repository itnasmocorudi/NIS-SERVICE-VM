VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmTPlottingTeknisiNHS 
   Caption         =   ".: PLOTTING TEKNISI NHS  :."
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
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
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plotting Teknisi NHS"
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
      TabIndex        =   1
      Top             =   360
      Width           =   12975
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
         Left            =   10320
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   21
         Top             =   240
         Visible         =   0   'False
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
         Left            =   7680
         MaxLength       =   50
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
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
         Index           =   2
         Left            =   9600
         MaxLength       =   60
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Height          =   2175
         Left            =   2880
         TabIndex        =   9
         Top             =   3360
         Visible         =   0   'False
         Width           =   7215
         Begin VB.TextBox txtAlasan 
            Height          =   975
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   10
            Top             =   480
            Width           =   6975
         End
         Begin ucXPButton.XPButton cmdSimpanPerubahan 
            Height          =   495
            Left            =   4680
            TabIndex        =   12
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
            MICON           =   "frmTPlottingTeknisiNHS.frx":0000
            PICN            =   "frmTPlottingTeknisiNHS.frx":001C
            PICH            =   "frmTPlottingTeknisiNHS.frx":681F
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
            TabIndex        =   22
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
            MICON           =   "frmTPlottingTeknisiNHS.frx":D3C5
            PICN            =   "frmTPlottingTeknisiNHS.frx":D3E1
            PICH            =   "frmTPlottingTeknisiNHS.frx":139DA
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lblParam 
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   18
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblParam 
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblParam 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   1560
            Visible         =   0   'False
            Width           =   1335
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
            TabIndex        =   11
            Top             =   240
            Width           =   585
         End
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   10920
         TabIndex        =   2
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
         MICON           =   "frmTPlottingTeknisiNHS.frx":1A289
         PICN            =   "frmTPlottingTeknisiNHS.frx":1A2A5
         PICH            =   "frmTPlottingTeknisiNHS.frx":20BD8
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
         Height          =   6375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   11245
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
         Columns(2).Caption=   "Ket Model"
         Columns(2).DataField=   "KetModel"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Petugas"
         Columns(3).DataField=   "Petugas"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "nik"
         Columns(4).DataField=   "nik"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "idplot"
         Columns(5).DataField=   "idplot"
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
         Splits(0)._ColumnProps(20)=   "Column(4).Visible=0"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(22)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(5).Visible=0"
         Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
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
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=62,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=59,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=60,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=61,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
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
         TabIndex        =   4
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
         MICON           =   "frmTPlottingTeknisiNHS.frx":2788F
         PICN            =   "frmTPlottingTeknisiNHS.frx":278AB
         PICH            =   "frmTPlottingTeknisiNHS.frx":2E04E
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
         TabIndex        =   5
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
         MICON           =   "frmTPlottingTeknisiNHS.frx":34B77
         PICN            =   "frmTPlottingTeknisiNHS.frx":34B93
         PICH            =   "frmTPlottingTeknisiNHS.frx":3B0F5
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
         TabIndex        =   0
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
         MICON           =   "frmTPlottingTeknisiNHS.frx":41918
         PICN            =   "frmTPlottingTeknisiNHS.frx":41934
         PICH            =   "frmTPlottingTeknisiNHS.frx":47F2D
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
         TabIndex        =   6
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
         MICON           =   "frmTPlottingTeknisiNHS.frx":4E7DC
         PICN            =   "frmTPlottingTeknisiNHS.frx":4E7F8
         PICH            =   "frmTPlottingTeknisiNHS.frx":5536B
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DcKodeNhs 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Tag             =   "Kode Supplier TAM"
         Top             =   240
         Width           =   3090
         BevelType       =   0
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   12648447
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5450
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSAdodcLib.Adodc AdoAs 
         Height          =   330
         Index           =   0
         Left            =   9000
         Top             =   240
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc AdoAs 
         Height          =   330
         Index           =   1
         Left            =   8640
         Top             =   120
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DCPetugas 
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Tag             =   "Kode Supplier TAM"
         Top             =   600
         Width           =   3090
         BevelType       =   0
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   12648447
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5450
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   8520
         TabIndex        =   15
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
         MICON           =   "frmTPlottingTeknisiNHS.frx":5C105
         PICN            =   "frmTPlottingTeknisiNHS.frx":5C121
         PICH            =   "frmTPlottingTeknisiNHS.frx":62924
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Petugas NHS"
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
         Left            =   555
         TabIndex        =   8
         Top             =   645
         Width           =   1065
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
         TabIndex        =   7
         Top             =   255
         Width           =   1305
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
      Picture         =   "frmTPlottingTeknisiNHS.frx":694CA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmTPlottingTeknisiNHS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmTPlottingTeknisiNHS
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
    AdoAs(0).ConnectionString = SQLCONBengkel
    AdoAs(0).RecordSource = "SELECT KodeUnit + ' - ' + NoPol AS UNIT,NoPol,KetModel,KodeUnit FROM mstunitnhs WITH (NOLOCK)"
    DcKodeNhs.DataSourceList = AdoAs(0)
    DcKodeNhs.DataFieldList = "UNIT"
    DcKodeNhs.Columns(0).Width = 3200
    DcKodeNhs.Columns(1).Width = 1500
    DcKodeNhs.Columns(2).Width = 4000
    DcKodeNhs.Columns(0).Caption = "Unit"
    DcKodeNhs.Columns(1).Caption = "No Polisi"
    DcKodeNhs.Columns(2).Caption = "Keterangan"
    DcKodeNhs.Columns(3).Visible = False
    
    AdoAs(1).ConnectionString = SQLCONBengkel
    AdoAs(1).RecordSource = "SELECT FS_NM_PETUGAS + ' | ' + nik as petugas,FS_KD_PETUGAS,nik FROM T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA' and ISNULL(nik,'') <> ''"
    DCPetugas.DataSourceList = AdoAs(1)
    DCPetugas.DataFieldList = "petugas"
    DCPetugas.Columns(0).Width = 3000
    DCPetugas.Columns(1).Width = 1500
    DCPetugas.Columns(0).Caption = "Petugas"
    DCPetugas.Columns(1).Caption = "Kode Petugas"
    DCPetugas.Columns(2).Visible = False
    isTambah = True
    DcKodeNhs.value = ""
    DCPetugas.value = ""
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
    DcKodeNhs.Enabled = Flg
    DCPetugas.Enabled = Flg
    
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
    Call ber_sih
    gridNHS.Enabled = True
    Call refreshGrid
End Sub

Private Sub cmdBatalKonfirmasi_Click()
    If MsgBox("Apakah yakin membatalkan proses?", vbYesNo) = vbYes Then
        gridNHS.Enabled = True
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
   
    If MsgBox("Apakah yakin menghapus data kode NHS " & NulSID(rsGrid!KodeUnit) & " ?", vbYesNo, "Konfirmasi Hapus Data") = vbYes Then
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
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmTPlottingTeknisiNHS")
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
          Dim nik         As String
          Dim FS_KD_PETUGAS   As String
              
2         boleh = True
3         isBegin = False
          
4         If NulSID(DcKodeNhs.value) = "" Then
5             pesan = pesan & " - Kode Unit harus dipilih." & vbCrLf
6             boleh = False
7         Else
8             sSQL = "SELECT KodeUnit + ' - ' + NoPol AS UNIT,KetModel,KodeUnit FROM mstunitnhs WITH (NOLOCK) WHERE KodeUnit + ' - ' + NoPol  = '" & NulSID(DcKodeNhs.value) & "'"
9             Set rsTemp = objO.FetchData(sSQL)
10            If rsTemp.EOF Then
11                pesan = pesan & " - Kode Unit tidak terdaftar." & vbCrLf
12                boleh = False
13            Else
14                KodeUnit = NulSID(rsTemp!KodeUnit)
15                lblParam(0).Caption = NulSID(KodeUnit)
16            End If
17        End If
          
18        If NulSID(DCPetugas.value) = "" Then
19            pesan = pesan & " - Kode Petugas harus dipilih." & vbCrLf
20            boleh = False
21        Else
22            sSQL = "SELECT FS_NM_PETUGAS + ' | ' + nik as petugas,FS_KD_PETUGAS,nik FROM T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA' and ISNULL(nik,'') <> '' and FS_NM_PETUGAS + ' | ' + nik  = '" & NulSID(DCPetugas.value) & "'"
23            Set rsTemp = objO.FetchData(sSQL)
24            If rsTemp.EOF Then
25                pesan = pesan & " - Kode Petugas tidak terdaftar." & vbCrLf
26                boleh = False
27            Else
28                FS_KD_PETUGAS = NulSID(rsTemp!FS_KD_PETUGAS)
29                nik = NulSID(rsTemp!nik)
30                lblParam(1).Caption = NulSID(FS_KD_PETUGAS)
31                lblParam(2).Caption = NulSID(nik)
32            End If
33            If isTambah = False Then
34                sSQL = "SELECT kodeunit FROM PlottingPetugasNHS WITH (NOLOCK) WHERE nik  = '" & NulSID(nik) & "' and kodeunit <> '" & NulSID(KodeUnit) & "'"
35            Else
36                sSQL = "SELECT kodeunit FROM PlottingPetugasNHS WITH (NOLOCK) WHERE nik  = '" & NulSID(nik) & "'"
37            End If
38            Set rsTemp = objO.FetchData(sSQL)
39            If Not rsTemp.EOF Then
40                pesan = pesan & " - Kode Petugas sudah terdaftar untuk kode unit " & NulSID(rsTemp!KodeUnit) & "." & vbCrLf
41                boleh = False
42            End If
43        End If
           
44        If boleh Then
45            If isTambah = False Then
46                Frame2.Visible = True
47                txtAlasan.Text = ""
48                Call buka_Tutup(False)
49                cmdTambah.Enabled = False
50                cmdEdit.Enabled = False
51                cmdHapus.Enabled = False
52                Exit Sub
53            End If
54            Call objHT.OpenTransactionBengkel
55            isBegin = True

56            Call ClearsSql

57            gsSql1(0) = "KodeUnit"
58            gsSql2(0) = NulSID(KodeUnit)
59            gsSql1(1) = "NIK"
60            gsSql2(1) = NulSID(nik)
61            gsSql1(2) = "FS_KD_PETUGAS"
62            gsSql2(2) = NulSID(FS_KD_PETUGAS)
63            If isTambah Then
64                gsSql1(3) = "TGLENTRY"
65                gsSql2(3) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
66                gsSql1(4) = "USERENTRY"
67                gsSql2(4) = gSA
68                gsSql1(5) = "kodecabang"
69                gsSql2(5) = gDb
70            End If

71            If isTambah = True Then
72                sSQL = GensSql2("PlottingPetugasNHS", False)
73            Else
74                sSQL = GensSql2("PlottingPetugasNHS", True) & " WHERE KodeUnit = '" & NulSID(KodeUnit) & "' "
75            End If
76            erTranAll = False
77            Call objHT.InsertData_Bengkel(sSQL)
              
78            gridNHS.Enabled = True

79            If erTranAll = False Then
80                Call objHT.CloseTransactionBengkel
81                MsgBox "Berhasil menyimpan data.", vbInformation
82                Call ber_sih
83                Call buka_Tutup(False)
84                Call refreshGrid
85            Else
errorRollback:
86                Call objHT.CloseTransactionBengkel_err
87                MsgBox "Gagal menyimpan data.", vbInformation
88                Exit Sub
89            End If
90        Else
91            MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
92        End If

93       On Error GoTo 0
94       Exit Sub

cmdSimpan_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
95        varErlError = Erl
96        varErrNumber = err.Number
97        If isBegin = True Then
98            Call objHT.CloseTransactionBengkel_err
99        End If
100       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmTPlottingTeknisiNHS")
End Sub

Private Sub cmdSimpanPerubahan_Click()
          
1        On Error GoTo cmdSimpanPerubahan_Click_Error
          Dim isBegin     As Boolean
          Dim boleh       As Boolean
          Dim pesan       As String
          Dim rsCek       As ADODB.Recordset
          
          Dim KodeUnit    As String
          Dim nik         As String
          Dim FS_KD_PETUGAS   As String
          
2         boleh = True
3         isBegin = False
          
4         If NulSID(txtAlasan.Text) = "" Then
5             boleh = False
6             pesan = pesan & " - Alasan harus diisi." & vbCrLf
7         ElseIf Len(NulSID(txtAlasan.Text)) < 5 Then
8             boleh = False
9             pesan = pesan & " - Alasan minimal 5 karakter." & vbCrLf
10        End If
           
11        If boleh Then
12            Call objHT.OpenTransactionBengkel
13            isBegin = True
              
              
14            pesan = ""
15            If isHapus = True Then
16                sSQL = "INSERT INTO LogPlottingPetugasNHS (KodeUnit,NIK,FS_KD_PETUGAS,UserEntry,TglEntry,KodeCabang,IdPlot,UserDel,TglDel,AlasanDel)"
17                sSQL = " SELECT KodeUnit,NIK,FS_KD_PETUGAS,UserEntry,TglEntry,KodeCabang,IdPlot,'" & gSA & "',getdate(),'" & NulSID(txtAlasan.Text) & "' FROM PlottingPetugasNHS WITH (NOLOCK) WHERE idplot = '" & NulSID(rsGrid!idplot) & "'"
18                Call objHT.InsertData_Bengkel(sSQL)
                  
19                sSQL = "DELETE FROM PlottingPetugasNHS WHERE idplot = '" & NulSID(rsGrid!idplot) & "' "
20                Call objHT.InsertData_Bengkel(sSQL)
21                pesan = "Data berhasil dihapus."
22            Else
23                sSQL = "INSERT INTO LogPlottingPetugasNHS (KodeUnit,NIK,FS_KD_PETUGAS,UserEntry,TglEntry,KodeCabang,IdPlot,UserEdit,TglEdit,AlasanEdit)"
24                sSQL = " SELECT KodeUnit,NIK,FS_KD_PETUGAS,UserEntry,TglEntry,KodeCabang,IdPlot,'" & gSA & "',getdate(),'" & NulSID(txtAlasan.Text) & "' FROM PlottingPetugasNHS WITH (NOLOCK) WHERE idplot = '" & NulSID(rsGrid!idplot) & "'"
25                Call objHT.InsertData_Bengkel(sSQL)
                  
26                sSQL = "UPDATE PlottingPetugasNHS SET KodeUnit = '" & NulSID(lblParam(0).Caption) & "' ,NIK = '" & NulSID(lblParam(2).Caption) & "',FS_KD_PETUGAS = '" & NulSID(lblParam(1).Caption) & "' WHERE idplot = '" & NulSID(rsGrid!idplot) & "' "
27                Call objHT.InsertData_Bengkel(sSQL)
28                pesan = "Data berhasil disimpan."
29            End If
              
30            Call objHT.CloseTransactionBengkel
31            gridNHS.Enabled = True
32            MsgBox pesan, vbInformation
33            Frame2.Visible = False
34            Call ber_sih
35            Call buka_Tutup(False)
36            Call refreshGrid
37        Else
38            MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
39        End If
40       On Error GoTo 0
41       Exit Sub

cmdSimpanPerubahan_Click_Error:
          
42        If isBegin Then
43            Call objHT.CloseTransactionBengkel_err
44        End If
45        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpanPerubahan_Click of Form frmTPlottingTeknisiNHS"


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
10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure gridnhs_HeadClick of Form frmTPlottingTeknisiNHS")
End Sub

Private Sub gridnhs_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
   On Error GoTo gridnhs_RowColChange_Error
    If Not rsGrid.EOF Then
        sSQL = "SELECT KodeUnit + ' - ' + NoPol AS UNIT,KetModel,KodeUnit FROM mstunitnhs WITH (NOLOCK) where kodeunit = '" & NulSID(rsGrid!KodeUnit) & "'"
        DcKodeNhs.value = objO.FetchData_String(sSQL)
        sSQL = "SELECT FS_NM_PETUGAS + ' | ' + nik as petugas,FS_KD_PETUGAS,nik FROM T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'SA' and ISNULL(nik,'') = '" & NulSID(rsGrid!nik) & "'"
        DCPetugas.value = objO.FetchData_String(sSQL)
    End If
   On Error GoTo 0
   Exit Sub
gridnhs_RowColChange_Error:
End Sub
   
Private Sub refreshGrid()
    sSQL = "SELECT a.kodeunit, b.NoPol,b.KetModel,c.FS_NM_PETUGAS as petugas,c.nik,a.idplot "
    sSQL = sSQL & " FROM PlottingPetugasNHS a with (nolock)"
    sSQL = sSQL & " inner join MstUnitNHS b with (nolock) on a.kodeunit = b.KodeUnit"
    sSQL = sSQL & " left join T_PETUGAS c with (nolock) on a.fs_kd_petugas = c.FS_KD_PETUGAS"
    sSQL = sSQL & " where a.kodecabang = '" & gDb & "'"
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
 
Private Sub XPButton1_Click()

End Sub
