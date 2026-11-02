VERSION 5.00
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtprofil 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Profil"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   18960
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20220
   WindowState     =   2  'Maximized
   Begin VB.Frame FrameAR 
      BackColor       =   &H80000016&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9255
      Left            =   12120
      TabIndex        =   146
      Top             =   0
      Visible         =   0   'False
      Width           =   11400
      Begin VB.Frame frmPilihKodeLgn 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frame16"
         Height          =   4455
         Left            =   120
         TabIndex        =   303
         Top             =   600
         Visible         =   0   'False
         Width           =   11175
         Begin VB.CommandButton TblSimpan 
            BackColor       =   &H00FF8080&
            Caption         =   "PILIH"
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
            Index           =   1
            Left            =   9360
            MaskColor       =   &H00FF8080&
            Style           =   1  'Graphical
            TabIndex        =   304
            Top             =   3960
            Width           =   1575
         End
         Begin TrueOleDBGrid70.TDBGrid gridKodelgn 
            Height          =   3255
            Left            =   120
            TabIndex        =   305
            Top             =   600
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   5741
            _LayoutType     =   4
            _RowHeight      =   25
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   4
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Pilih"
            Columns(0).DataField=   "pilih"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "KodeLgn"
            Columns(1).DataField=   "KodeLgn"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "NamaLgn"
            Columns(2).DataField=   "NamaLgn"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Alamat"
            Columns(3).DataField=   "Alamat"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Npwp"
            Columns(4).DataField=   "Npwp"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "nom"
            Columns(5).DataField=   "nom"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   6
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14933984
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=6"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
            Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
            Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
            Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(9)=   "Column(2).Width=4154"
            Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=4048"
            Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(13)=   "Column(3).Width=6800"
            Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=6694"
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
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   0
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
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
            _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=862,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=859,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=860,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=861,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=870,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=867,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=868,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=869,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=878,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=875,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=876,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=877,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=886,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=883,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=884,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=885,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=894,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=891,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=892,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=893,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
            _StyleDefs(60)  =   "Named:id=33:Normal"
            _StyleDefs(61)  =   ":id=33,.parent=0"
            _StyleDefs(62)  =   "Named:id=34:Heading"
            _StyleDefs(63)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(64)  =   ":id=34,.wraptext=-1"
            _StyleDefs(65)  =   "Named:id=35:Footing"
            _StyleDefs(66)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(67)  =   "Named:id=36:Selected"
            _StyleDefs(68)  =   ":id=36,.parent=33,.bgcolor=&H80FF&,.fgcolor=&H8000000E&"
            _StyleDefs(69)  =   "Named:id=37:Caption"
            _StyleDefs(70)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(71)  =   "Named:id=38:HighlightRow"
            _StyleDefs(72)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=39:EvenRow"
            _StyleDefs(74)  =   ":id=39,.parent=33"
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
            _StyleDefs(85)  =   "Named:id=87:SudahDitagih"
            _StyleDefs(86)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(87)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(88)  =   ":id=87,.fontname=Tahoma"
            _StyleDefs(89)  =   "Named:id=514:Kuning"
            _StyleDefs(90)  =   ":id=514,.parent=87,.bgcolor=&HFFFF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
            _StyleDefs(91)  =   ":id=514,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(92)  =   ":id=514,.fontname=Tahoma"
            _StyleDefs(93)  =   "Named:id=515:Biru"
            _StyleDefs(94)  =   ":id=515,.parent=514,.bgcolor=&HFF8080&"
            _StyleDefs(95)  =   "Named:id=516:Merah"
            _StyleDefs(96)  =   ":id=516,.parent=515,.bgcolor=&H88A0F7&,.fgcolor=&H0&"
            _StyleDefs(97)  =   "Named:id=517:Hijau"
            _StyleDefs(98)  =   ":id=517,.parent=516,.bgcolor=&HFF00&"
            _StyleDefs(99)  =   "Named:id=518:Putih"
            _StyleDefs(100) =   ":id=518,.parent=517,.bgcolor=&HFFFFFF&"
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PILIH KODE LANGGANAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   308
            Top             =   120
            Width           =   10935
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "KPI RETENTION"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   450
            TabIndex        =   307
            Top             =   495
            Width           =   9975
         End
         Begin VB.Shape Shape15 
            BackColor       =   &H00404040&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00404040&
            Height          =   495
            Left            =   15
            Shape           =   4  'Rounded Rectangle
            Top             =   15
            Width           =   11115
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "PILIH KODE LANGGANAN"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   306
            Top             =   120
            Width           =   8055
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4215
         Left            =   120
         TabIndex        =   147
         Top             =   4940
         Width           =   5535
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   300
            Top             =   570
            Width           =   4005
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            Index           =   100
            Left            =   1320
            MaxLength       =   25
            TabIndex        =   152
            Top             =   120
            Width           =   2250
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   109
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   151
            Top             =   1560
            Width           =   2205
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   110
            Left            =   1320
            MaxLength       =   25
            TabIndex        =   150
            Top             =   1890
            Width           =   2205
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   107
            Left            =   1320
            MaxLength       =   40
            TabIndex        =   149
            Top             =   900
            Width           =   4005
         End
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   320
            IMEMode         =   3  'DISABLE
            Index           =   108
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   148
            Top             =   1230
            Width           =   4005
         End
         Begin TDBNumber6Ctl.TDBNumber txtRp 
            Height          =   315
            Index           =   1
            Left            =   1920
            TabIndex        =   153
            Top             =   3060
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   556
            Calculator      =   "frmtprofil.frx":0000
            Caption         =   "frmtprofil.frx":0020
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmtprofil.frx":008C
            Keys            =   "frmtprofil.frx":00AA
            Spin            =   "frmtprofil.frx":00F4
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
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   "."
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1397882885
            MinValueVT      =   1112342533
         End
         Begin TDBNumber6Ctl.TDBNumber txtRp 
            Height          =   315
            Index           =   2
            Left            =   1920
            TabIndex        =   154
            Top             =   3390
            Width           =   1965
            _Version        =   65536
            _ExtentX        =   3466
            _ExtentY        =   556
            Calculator      =   "frmtprofil.frx":011C
            Caption         =   "frmtprofil.frx":013C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "frmtprofil.frx":01A8
            Keys            =   "frmtprofil.frx":01C6
            Spin            =   "frmtprofil.frx":0210
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
            MinValue        =   0
            MousePointer    =   0
            MoveOnLRKey     =   0
            NegativeColor   =   255
            OLEDragMode     =   0
            OLEDropMode     =   0
            ReadOnly        =   0
            Separator       =   "."
            ShowContextMenu =   -1
            ValueVT         =   5
            Value           =   0
            MaxValueVT      =   1397882885
            MinValueVT      =   1112342533
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID Cust"
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
            Index           =   6
            Left            =   120
            TabIndex        =   302
            Top             =   183
            Width           =   705
         End
         Begin VB.Label Label8 
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
            Index           =   5
            Left            =   1200
            TabIndex        =   301
            Top             =   633
            Width           =   60
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   315
            Index           =   1
            Left            =   3600
            Top             =   120
            Width           =   60
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
            Index           =   4
            Left            =   1200
            TabIndex        =   172
            Top             =   183
            Width           =   60
         End
         Begin VB.Line Line1 
            BorderColor     =   &H000080FF&
            X1              =   5400
            X2              =   0
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rp."
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
            Left            =   1560
            TabIndex        =   171
            Top             =   3090
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Rp."
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
            Left            =   1560
            TabIndex        =   170
            Top             =   3420
            Width           =   300
         End
         Begin VB.Label Label5 
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
            Index           =   9
            Left            =   120
            TabIndex        =   169
            Top             =   3090
            Width           =   1050
         End
         Begin VB.Label Label8 
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
            Index           =   8
            Left            =   1320
            TabIndex        =   168
            Top             =   3120
            Width           =   60
         End
         Begin VB.Label Label8 
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
            Index           =   6
            Left            =   1320
            TabIndex        =   167
            Top             =   3420
            Width           =   60
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plafond"
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
            Index           =   7
            Left            =   120
            TabIndex        =   166
            Top             =   3420
            Width           =   735
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Lgn "
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
            Left            =   120
            TabIndex        =   165
            Top             =   633
            Width           =   945
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kota"
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
            Left            =   120
            TabIndex        =   164
            Top             =   1623
            Width           =   435
         End
         Begin VB.Label Label3 
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
            Left            =   1200
            TabIndex        =   163
            Top             =   1623
            Width           =   60
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
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
            Left            =   120
            TabIndex        =   162
            Top             =   1293
            Width           =   690
         End
         Begin VB.Label Label8 
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
            Left            =   1200
            TabIndex        =   161
            Top             =   1293
            Width           =   60
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telp"
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
            Left            =   120
            TabIndex        =   160
            Top             =   1920
            Width           =   420
         End
         Begin VB.Label Label8 
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
            Left            =   1200
            TabIndex        =   159
            Top             =   963
            Width           =   60
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama"
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
            Left            =   120
            TabIndex        =   158
            Top             =   963
            Width           =   555
         End
         Begin VB.Label Label68 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".:: INFORMASI KREDIT ::."
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
            Left            =   1350
            TabIndex        =   157
            Top             =   2580
            Width           =   2655
         End
         Begin VB.Label Label3 
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
            Left            =   1200
            TabIndex        =   156
            Top             =   1920
            Width           =   60
         End
         Begin VB.Label nAR 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "89%"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   21.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   525
            Left            =   3960
            TabIndex        =   155
            Top             =   3120
            Width           =   1185
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   2520
            Width           =   5295
         End
      End
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   4425
         Left            =   120
         TabIndex        =   173
         Top             =   120
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7805
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "KodeLgn"
         Columns(0).DataField=   "KodeLgn"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Current AR"
         Columns(1).DataField=   "CurrentAR"
         Columns(1).NumberFormat=   "Standard"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Plafond AR"
         Columns(2).DataField=   "PlatformAR"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Valid"
         Columns(3).DataField=   "Valid"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "User"
         Columns(4).DataField=   "UserLog"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Tanggal"
         Columns(5).DataField=   "TglLog"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "NoWO"
         Columns(6).DataField=   "NoWO"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Keterangan"
         Columns(7).DataField=   "Ket"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2302"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2196"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=3175"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3069"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3175"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3069"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(13)=   "Column(3).Width=1323"
         Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=1217"
         Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(17)=   "Column(4).Width=1773"
         Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=1667"
         Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(21)=   "Column(5).Width=4763"
         Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=4657"
         Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(25)=   "Column(6).Width=2117"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2011"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=26458"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=26353"
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         BorderStyle     =   0
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
         DirectionAfterEnter=   2
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
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
      Begin ucXPButton.XPButton XPButton1 
         Height          =   495
         Left            =   10080
         TabIndex        =   193
         Top             =   8160
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
         MICON           =   "frmtprofil.frx":0238
         PICN            =   "frmtprofil.frx":0254
         PICH            =   "frmtprofil.frx":0728
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H80000016&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   5640
         TabIndex        =   174
         Top             =   5520
         Width           =   5535
         Begin VB.TextBox txtKet 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   175
            Top             =   720
            Width           =   5025
         End
         Begin ucXPButton.XPButton cmdUp2 
            Height          =   495
            Left            =   1560
            TabIndex        =   176
            Top             =   1560
            Width           =   2145
            _ExtentX        =   3784
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   "&u"
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
            MICON           =   "frmtprofil.frx":0C73
            PICN            =   "frmtprofil.frx":0C8F
            PICH            =   "frmtprofil.frx":14F4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Keterangan :"
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
            Left            =   120
            TabIndex        =   145
            Top             =   480
            Width           =   1245
         End
         Begin VB.Label Label69 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".:: VALIDASI ::."
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
            Left            =   1920
            TabIndex        =   177
            Top             =   60
            Width           =   1635
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   375
            Left            =   0
            Shape           =   4  'Rounded Rectangle
            Top             =   0
            Width           =   5295
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   120
         X2              =   11400
         Y1              =   4920
         Y2              =   4920
      End
   End
   Begin VB.Frame frmSimplify 
      Caption         =   "Frame8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   74440
      TabIndex        =   275
      Top             =   240
      Width           =   11415
      Begin TrueOleDBGrid70.TDBGrid gridSA2 
         Height          =   2865
         Left            =   6360
         TabIndex        =   299
         Top             =   600
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5054
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode"
         Columns(0).DataField=   "fs_kd_petugas"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama"
         Columns(1).DataField=   "fs_nm_petugas"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "HP"
         Columns(2).DataField=   "nohp"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1508"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1402"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=8196"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         BorderStyle     =   0
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
         DirectionAfterEnter=   2
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.locked=-1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
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
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   78
         Left            =   1710
         TabIndex        =   276
         Top             =   600
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   79
         Left            =   1710
         TabIndex        =   277
         Top             =   840
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   80
         Left            =   1710
         TabIndex        =   278
         Top             =   1200
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   81
         Left            =   1710
         TabIndex        =   279
         Top             =   1440
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   82
         Left            =   1710
         TabIndex        =   280
         Top             =   1800
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   83
         Left            =   1710
         TabIndex        =   281
         Top             =   2040
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   84
         Left            =   1710
         TabIndex        =   282
         Top             =   2400
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   85
         Left            =   1710
         TabIndex        =   283
         Top             =   2640
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   86
         Left            =   1710
         TabIndex        =   284
         Top             =   3000
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin ucXPButton.XPButton cmdSimpanSimplify 
         Height          =   495
         Left            =   8760
         TabIndex        =   285
         Top             =   3720
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
         MICON           =   "frmtprofil.frx":1D59
         PICN            =   "frmtprofil.frx":1D75
         PICH            =   "frmtprofil.frx":2288
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton XPButton3 
         Height          =   495
         Left            =   10080
         TabIndex        =   286
         Top             =   3720
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
         MICON           =   "frmtprofil.frx":282A
         PICN            =   "frmtprofil.frx":2846
         PICH            =   "frmtprofil.frx":2D1A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   87
         Left            =   1710
         TabIndex        =   287
         Top             =   3240
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin VB.Label Label95 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   ".:: SETTING DATA SIMPLIFY ::."
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
         TabIndex        =   298
         Top             =   120
         Width           =   11295
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H0080C0FF&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   11415
      End
      Begin VB.Label Label96 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   645
         TabIndex        =   297
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label97 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Bank 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         TabIndex        =   296
         Top             =   900
         Width           =   765
      End
      Begin VB.Label Label98 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   645
         TabIndex        =   295
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label99 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Bank 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         TabIndex        =   294
         Top             =   1500
         Width           =   765
      End
      Begin VB.Label Label101 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Bank 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   645
         TabIndex        =   293
         Top             =   1860
         Width           =   975
      End
      Begin VB.Label Label102 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Bank 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   855
         TabIndex        =   292
         Top             =   2100
         Width           =   765
      End
      Begin VB.Label Label103 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jam layanan 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   291
         Top             =   2460
         Width           =   1020
      End
      Begin VB.Label Label104 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jam Layanan 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   540
         TabIndex        =   290
         Top             =   2700
         Width           =   1080
      End
      Begin VB.Label Label105 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fax Perusahaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   465
         TabIndex        =   289
         Top             =   3060
         Width           =   1155
      End
      Begin VB.Label Label106 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Email Perusahaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   288
         Top             =   3300
         Width           =   1275
      End
   End
   Begin VB.Frame frameHansaRenkei 
      Caption         =   "Frame8"
      Height          =   1335
      Left            =   2520
      TabIndex        =   266
      Top             =   8640
      Width           =   5295
      Begin TDBNumber6Ctl.TDBNumber txtReminderSA 
         Height          =   375
         Left            =   1800
         TabIndex        =   272
         Top             =   480
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   661
         Calculator      =   "frmtprofil.frx":3265
         Caption         =   "frmtprofil.frx":3285
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmtprofil.frx":32EA
         Keys            =   "frmtprofil.frx":3308
         Spin            =   "frmtprofil.frx":3352
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   -99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin TDBNumber6Ctl.TDBNumber txtReminderMRA 
         Height          =   375
         Left            =   1800
         TabIndex        =   273
         Top             =   840
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   661
         Calculator      =   "frmtprofil.frx":337A
         Caption         =   "frmtprofil.frx":339A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmtprofil.frx":33FF
         Keys            =   "frmtprofil.frx":341D
         Spin            =   "frmtprofil.frx":3467
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   16777215
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   -99
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   0
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin ucXPButton.XPButton cmdSimpanHansaRenkei 
         Height          =   495
         Left            =   2520
         TabIndex        =   274
         Top             =   720
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
         MICON           =   "frmtprofil.frx":348F
         PICN            =   "frmtprofil.frx":34AB
         PICH            =   "frmtprofil.frx":39BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reminder MRA"
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
         Index           =   5
         Left            =   120
         TabIndex        =   271
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label8 
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
         Index           =   4
         Left            =   1680
         TabIndex        =   270
         Top             =   960
         Width           =   60
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reminder SA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   269
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label8 
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
         Left            =   1680
         TabIndex        =   268
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label68 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: SETTING REMINDER HANSA RENKEI ::."
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
         Left            =   15
         TabIndex        =   267
         Top             =   75
         Width           =   5265
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   5295
      End
   End
   Begin ucXPButton.XPButton cmdHansaRenkei 
      Height          =   495
      Left            =   2520
      TabIndex        =   265
      Top             =   8640
      Visible         =   0   'False
      Width           =   2300
      _ExtentX        =   4048
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
      MICON           =   "frmtprofil.frx":3F60
      PICN            =   "frmtprofil.frx":3F7C
      PICH            =   "frmtprofil.frx":AE9D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame frameMasterSetting 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   23600
      TabIndex        =   250
      Top             =   240
      Width           =   6375
      Begin MSMask.MaskEdBox Text 
         Height          =   315
         Index           =   75
         Left            =   1680
         TabIndex        =   251
         Top             =   555
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin ucXPButton.XPButton cmdSimpanSettingCR 
         Height          =   495
         Left            =   1680
         TabIndex        =   252
         Top             =   1920
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
         MICON           =   "frmtprofil.frx":120C5
         PICN            =   "frmtprofil.frx":120E1
         PICH            =   "frmtprofil.frx":125F4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   315
         Index           =   76
         Left            =   1680
         TabIndex        =   256
         Top             =   960
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   315
         Index           =   77
         Left            =   1680
         TabIndex        =   261
         Top             =   1320
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin VB.Label Label100 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hari Sebelum JT"
         Height          =   240
         Index           =   7
         Left            =   2280
         TabIndex        =   262
         Top             =   1365
         Width           =   1410
      End
      Begin VB.Label Label100 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hari Sesudah Service Terakhir"
         Height          =   240
         Index           =   3
         Left            =   2280
         TabIndex        =   258
         Top             =   1005
         Width           =   2610
      End
      Begin VB.Label Label100 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reminder Lubing"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   257
         Top             =   1005
         Width           =   1440
      End
      Begin VB.Label Label100 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hari Sebelum JT"
         Height          =   240
         Index           =   1
         Left            =   2280
         TabIndex        =   255
         Top             =   600
         Width           =   1410
      End
      Begin VB.Label Ket 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: MASTER SETTING CR ::."
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
         Left            =   1440
         TabIndex        =   254
         Top             =   195
         Width           =   2745
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label Label100 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reminder CR 7"
         Height          =   240
         Index           =   0
         Left            =   240
         TabIndex        =   253
         Top             =   600
         Width           =   1290
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   9240
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000080FF&
         X1              =   0
         X2              =   6360
         Y1              =   2520
         Y2              =   2520
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10455
      Left            =   0
      Picture         =   "frmtprofil.frx":12B96
      ScaleHeight     =   10455
      ScaleWidth      =   495
      TabIndex        =   192
      Top             =   -120
      Width           =   495
   End
   Begin VB.Frame Frame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8220
      Left            =   480
      TabIndex        =   47
      Top             =   60
      Width           =   11340
      Begin ucXPButton.XPButton cmdSettingSimplify 
         Height          =   495
         Left            =   8560
         TabIndex        =   263
         Top             =   240
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtprofil.frx":14972
         PICN            =   "frmtprofil.frx":1498E
         PICH            =   "frmtprofil.frx":1BCAC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Frame SetServer 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5415
         Left            =   12000
         TabIndex        =   83
         Top             =   480
         Visible         =   0   'False
         Width           =   9495
         Begin VB.Frame setServerEnkrip 
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame8"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   120
            TabIndex        =   217
            Top             =   240
            Visible         =   0   'False
            Width           =   9255
         End
         Begin VB.TextBox txtPassSPKHO 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   7380
            PasswordChar    =   "*"
            TabIndex        =   97
            Top             =   1305
            Width           =   1815
         End
         Begin VB.TextBox txtPassHO 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   1140
            PasswordChar    =   "*"
            TabIndex        =   89
            Top             =   1305
            Width           =   1815
         End
         Begin VB.TextBox txtpassSPK 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   4260
            PasswordChar    =   "*"
            TabIndex        =   93
            Top             =   1305
            Width           =   1815
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   34
            Left            =   4260
            TabIndex        =   91
            Top             =   720
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   35
            Left            =   4260
            TabIndex        =   92
            Top             =   1005
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   36
            Left            =   4260
            TabIndex        =   94
            Top             =   1605
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   37
            Left            =   1140
            TabIndex        =   87
            Top             =   720
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   38
            Left            =   1140
            TabIndex        =   88
            Top             =   1005
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   39
            Left            =   1140
            TabIndex        =   90
            Top             =   1605
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   31
            Left            =   7380
            TabIndex        =   95
            Top             =   720
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   32
            Left            =   7380
            TabIndex        =   96
            Top             =   1005
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   33
            Left            =   7380
            TabIndex        =   98
            Top             =   1605
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   40
            Left            =   1080
            TabIndex        =   99
            Top             =   2640
            Width           =   5220
            _ExtentX        =   9208
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   41
            Left            =   1080
            TabIndex        =   100
            Top             =   2930
            Width           =   5220
            _ExtentX        =   9208
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   42
            Left            =   1080
            TabIndex        =   101
            Top             =   3240
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   43
            Left            =   1080
            TabIndex        =   102
            Top             =   3530
            Width           =   5220
            _ExtentX        =   9208
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Mask            =   "CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   44
            Left            =   1080
            TabIndex        =   103
            Top             =   3840
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
            Height          =   270
            Index           =   45
            Left            =   1080
            TabIndex        =   104
            Top             =   4200
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox MaskEdBox1 
            Height          =   270
            Left            =   2720
            TabIndex        =   123
            Top             =   4200
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin ucXPButton.XPButton cmdSave 
            Height          =   495
            Left            =   4200
            TabIndex        =   208
            Top             =   4680
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
            MICON           =   "frmtprofil.frx":23895
            PICN            =   "frmtprofil.frx":238B1
            PICH            =   "frmtprofil.frx":23DC4
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Line Line8 
            BorderColor     =   &H000080FF&
            X1              =   120
            X2              =   9360
            Y1              =   5280
            Y2              =   5280
         End
         Begin VB.Line Line7 
            BorderColor     =   &H000080FF&
            X1              =   120
            X2              =   9360
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Label Ket 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".:: DATA HO ::."
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
            Left            =   2550
            TabIndex        =   207
            Top             =   2220
            Width           =   1545
         End
         Begin VB.Label Ket 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".:: SPK HO SERVER  ::."
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
            Left            =   6735
            TabIndex        =   206
            Top             =   300
            Width           =   2325
         End
         Begin VB.Label Ket 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".:: SPK SERVER  ::."
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
            Left            =   3810
            TabIndex        =   205
            Top             =   300
            Width           =   1935
         End
         Begin VB.Label Ket 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   ".:: HO SERVER  ::."
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
            Left            =   720
            TabIndex        =   204
            Top             =   300
            Width           =   1875
         End
         Begin VB.Label Label54 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Kode Plg"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   330
            TabIndex        =   122
            Top             =   4245
            Width           =   645
         End
         Begin VB.Label Label53 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "NPWP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   480
            TabIndex        =   121
            Top             =   3885
            Width           =   495
         End
         Begin VB.Label Label52 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telp."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   615
            TabIndex        =   120
            Top             =   3568
            Width           =   360
         End
         Begin VB.Label Label51 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Kota"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   645
            TabIndex        =   119
            Top             =   3278
            Width           =   330
         End
         Begin VB.Label Label50 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Alamat"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   495
            TabIndex        =   118
            Top             =   2968
            Width           =   480
         End
         Begin VB.Label Label49 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Perusahaan"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   117
            Top             =   2685
            Width           =   855
         End
         Begin VB.Label Label44 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6480
            TabIndex        =   116
            Top             =   1635
            Width           =   690
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6480
            TabIndex        =   115
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6840
            TabIndex        =   114
            Top             =   1035
            Width           =   330
         End
         Begin VB.Label Label33 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Server"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   6705
            TabIndex        =   113
            Top             =   765
            Width           =   465
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   112
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label47 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   111
            Top             =   1635
            Width           =   690
         End
         Begin VB.Label Label45 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   110
            Top             =   1035
            Width           =   330
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Server"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   465
            TabIndex        =   109
            Top             =   765
            Width           =   465
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Database"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   108
            Top             =   1635
            Width           =   690
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3360
            TabIndex        =   107
            Top             =   1350
            Width           =   690
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "User"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3720
            TabIndex        =   106
            Top             =   1035
            Width           =   330
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Server"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   3585
            TabIndex        =   105
            Top             =   765
            Width           =   465
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   3015
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   375
            Index           =   2
            Left            =   3240
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   3015
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   375
            Index           =   3
            Left            =   6360
            Shape           =   4  'Rounded Rectangle
            Top             =   240
            Width           =   3015
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00FFFFFF&
            Height          =   375
            Index           =   4
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   2160
            Width           =   6255
         End
      End
      Begin VB.CommandButton cmdMasterSettingCR 
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
         Height          =   495
         Left            =   6720
         Picture         =   "frmtprofil.frx":24366
         Style           =   1  'Graphical
         TabIndex        =   249
         Top             =   240
         Width           =   1815
      End
      Begin VB.Frame FrDisc 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   9150
         TabIndex        =   198
         Top             =   4830
         Visible         =   0   'False
         Width           =   2175
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   60
            Left            =   1440
            TabIndex        =   199
            Top             =   120
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   61
            Left            =   1440
            TabIndex        =   200
            Top             =   405
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin ucXPButton.XPButton cmdSaveDisc 
            Height          =   495
            Left            =   360
            TabIndex        =   203
            Top             =   840
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   873
            BTYPE           =   6
            TX              =   ""
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
            MICON           =   "frmtprofil.frx":2ADAC
            PICN            =   "frmtprofil.frx":2ADC8
            PICH            =   "frmtprofil.frx":2B2DB
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Line Line6 
            BorderColor     =   &H000080FF&
            X1              =   120
            X2              =   2040
            Y1              =   705
            Y2              =   705
         End
         Begin VB.Label Label75 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Disc Non NRM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   202
            Top             =   450
            Width           =   1080
         End
         Begin VB.Label Label74 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Disc NRM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   201
            Top             =   150
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdDiscNRM 
         Caption         =   "DISC NRM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10350
         TabIndex        =   197
         Top             =   4575
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   7560
         TabIndex        =   178
         Top             =   1920
         Width           =   3255
         Begin VB.CommandButton cmdAturThs 
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
            Height          =   440
            Left            =   600
            Picture         =   "frmtprofil.frx":2B87D
            Style           =   1  'Graphical
            TabIndex        =   189
            Top             =   3000
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.CommandButton cmdAturStall 
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
            Height          =   440
            Left            =   600
            Picture         =   "frmtprofil.frx":2C07C
            Style           =   1  'Graphical
            TabIndex        =   190
            Top             =   2580
            Width           =   2055
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   57
            Left            =   2400
            TabIndex        =   179
            Top             =   420
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   58
            Left            =   2400
            TabIndex        =   180
            Top             =   720
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   56
            Left            =   2400
            TabIndex        =   185
            Top             =   120
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   59
            Left            =   2400
            TabIndex        =   186
            Top             =   1020
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   27
            Left            =   2400
            TabIndex        =   187
            Top             =   1320
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin VB.CommandButton cmdShift 
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
            Height          =   440
            Left            =   600
            Picture         =   "frmtprofil.frx":2C8A6
            Style           =   1  'Graphical
            TabIndex        =   191
            Top             =   2160
            Width           =   2055
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   67
            Left            =   2400
            TabIndex        =   225
            Top             =   1600
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin VB.Label Label81 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Jml Service Point"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   226
            Top             =   1605
            Width           =   1215
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000080FF&
            X1              =   120
            X2              =   3120
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label Label30 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Jml Unit THS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   188
            Top             =   1350
            Width           =   930
         End
         Begin VB.Label Label73 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Stall EM"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   184
            Top             =   150
            Width           =   1125
         End
         Begin VB.Label Label72 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Stall Produktif"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   183
            Top             =   750
            Width           =   1515
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Jumlah Stall NP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   182
            Top             =   450
            Width           =   1110
         End
         Begin VB.Label Label70 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Jumnlah Stall Non Produktif"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   181
            Top             =   1050
            Width           =   1950
         End
      End
      Begin VB.CheckBox chkCash 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pengaturan Cash / Kredit manual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7665
         TabIndex        =   137
         Top             =   960
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.CommandButton cmdServer 
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
         Height          =   495
         Left            =   4440
         Picture         =   "frmtprofil.frx":2CF2E
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox ChkAAS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Connect dengan ACC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7680
         TabIndex        =   35
         Top             =   1680
         Visible         =   0   'False
         Width           =   1980
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00BFE8EA&
         Caption         =   "Cetak Retur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   11400
         TabIndex        =   79
         Top             =   2040
         Visible         =   0   'False
         Width           =   3960
         Begin VB.OptionButton Option8 
            BackColor       =   &H00BFE8EA&
            Caption         =   "Invoice Retur Kecil (tanpa form)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   31
            Top             =   510
            Width           =   3435
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00BFE8EA&
            Caption         =   "From Invoice Retur Besar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   30
            Top             =   240
            Width           =   2220
         End
      End
      Begin VB.CheckBox ChkIERC 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pembuatan Faktur Pajak Auto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7665
         TabIndex        =   34
         Top             =   1200
         Visible         =   0   'False
         Width           =   2940
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00BFE8EA&
         Caption         =   "Preview Order && Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   11400
         TabIndex        =   75
         Top             =   2925
         Visible         =   0   'False
         Width           =   3960
         Begin VB.OptionButton Option6 
            BackColor       =   &H00BFE8EA&
            Caption         =   "Setelah simpan langsung cetak"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   33
            Top             =   510
            Width           =   2910
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00BFE8EA&
            Caption         =   "Setelah simpan melalui preview"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   32
            Top             =   240
            Width           =   2895
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00BFE8EA&
         Caption         =   "Cetak Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   11400
         TabIndex        =   74
         Top             =   1155
         Visible         =   0   'False
         Width           =   3960
         Begin VB.OptionButton Option4 
            BackColor       =   &H00BFE8EA&
            Caption         =   "From Invoice Besar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   28
            Top             =   240
            Width           =   2220
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00BFE8EA&
            Caption         =   "Invoice Kecil (tanpa form)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   29
            Top             =   510
            Width           =   2220
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00BFE8EA&
         Caption         =   "Cetak Order"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   11400
         TabIndex        =   73
         Top             =   270
         Visible         =   0   'False
         Width           =   3960
         Begin VB.OptionButton Option2 
            BackColor       =   &H00BFE8EA&
            Caption         =   "Order Kecil (tanpa form)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   27
            Top             =   510
            Width           =   2220
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00BFE8EA&
            Caption         =   "From Order Besar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   165
            TabIndex        =   26
            Top             =   240
            Width           =   2220
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   2190
         TabIndex        =   0
         Top             =   285
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   2190
         TabIndex        =   1
         Top             =   585
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   7
         Left            =   2190
         TabIndex        =   2
         Top             =   960
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   8
         Left            =   2190
         TabIndex        =   3
         Top             =   1245
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   13395
         TabIndex        =   36
         Top             =   6435
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   3
         Left            =   13395
         TabIndex        =   37
         Top             =   6735
         Visible         =   0   'False
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   4
         Left            =   12510
         TabIndex        =   38
         Top             =   3885
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   5
         Left            =   12510
         TabIndex        =   39
         Top             =   4185
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   6
         Left            =   12510
         TabIndex        =   40
         Top             =   4485
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   9
         Left            =   2190
         TabIndex        =   9
         Top             =   3465
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   10
         Left            =   2190
         TabIndex        =   10
         Top             =   3765
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   11
         Left            =   2190
         TabIndex        =   13
         Top             =   4365
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   12
         Left            =   13395
         TabIndex        =   46
         Top             =   7035
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   13
         Left            =   13395
         TabIndex        =   41
         Top             =   6150
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   14
         Left            =   2145
         TabIndex        =   11
         Top             =   8385
         Visible         =   0   'False
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   15
         Left            =   13380
         TabIndex        =   42
         Top             =   4905
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   16
         Left            =   2190
         TabIndex        =   14
         Top             =   4665
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   17
         Left            =   2190
         TabIndex        =   15
         Top             =   4965
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   18
         Left            =   13395
         TabIndex        =   43
         Top             =   5205
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   19
         Left            =   2190
         TabIndex        =   5
         Top             =   1905
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   20
         Left            =   2190
         TabIndex        =   7
         Top             =   2865
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   21
         Left            =   2190
         TabIndex        =   8
         Top             =   3165
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   22
         Left            =   2190
         TabIndex        =   6
         Top             =   2205
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   23
         Left            =   13395
         TabIndex        =   44
         Top             =   5535
         Visible         =   0   'False
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   24
         Left            =   13395
         TabIndex        =   45
         Top             =   5835
         Visible         =   0   'False
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   25
         Left            =   2190
         TabIndex        =   16
         Top             =   5265
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   26
         Left            =   6150
         TabIndex        =   17
         Top             =   6405
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   28
         Left            =   2190
         TabIndex        =   12
         Top             =   4065
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   29
         Left            =   2190
         TabIndex        =   4
         Top             =   1560
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   3120
         TabIndex        =   84
         Top             =   1560
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
         Index           =   30
         Left            =   3285
         TabIndex        =   85
         Top             =   1560
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Height          =   270
         Index           =   46
         Left            =   2190
         TabIndex        =   18
         Top             =   5865
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
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
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   47
         Left            =   10110
         TabIndex        =   25
         Top             =   7125
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
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
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   48
         Left            =   2190
         TabIndex        =   20
         Top             =   6165
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">99999"
         PromptChar      =   ">"
      End
      Begin VB.CheckBox chkSpart 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Connect dengan Sparepart"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7665
         TabIndex        =   124
         Top             =   1440
         Visible         =   0   'False
         Width           =   2340
      End
      Begin MSComCtl2.DTPicker TglPkp 
         Height          =   255
         Left            =   5280
         TabIndex        =   131
         Top             =   3180
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   98762753
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   49
         Left            =   2190
         TabIndex        =   22
         Top             =   6765
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   35
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   50
         Left            =   2190
         TabIndex        =   23
         Top             =   7080
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
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
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   51
         Left            =   2880
         TabIndex        =   134
         Top             =   7080
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">99999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   52
         Left            =   2190
         TabIndex        =   24
         Top             =   5565
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">999"
         PromptChar      =   ">"
      End
      Begin AIFCmp1.asxPowerButton cmdDisconnect 
         Height          =   480
         Left            =   4560
         TabIndex        =   136
         Top             =   7080
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   847
         BackColor       =   255
         BorderStyle     =   5
         Caption         =   "&Disconnect HO"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotTracking     =   -1  'True
         Appearance      =   1
      End
      Begin VB.CheckBox chQueue 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Connect dengan Antrian"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7665
         TabIndex        =   138
         Top             =   720
         Width           =   2340
      End
      Begin MSMask.MaskEdBox Text 
         DataField       =   "m "
         Height          =   270
         Index           =   53
         Left            =   3960
         TabIndex        =   19
         Top             =   5880
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
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
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   54
         Left            =   2190
         TabIndex        =   21
         Top             =   6465
         Width           =   2940
         _ExtentX        =   5186
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   5160
         TabIndex        =   141
         Top             =   6465
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   55
         Left            =   2190
         TabIndex        =   143
         Top             =   2520
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   57
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   62
         Left            =   6150
         TabIndex        =   209
         Top             =   5830
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
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
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   63
         Left            =   2190
         TabIndex        =   212
         Top             =   7440
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   476
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   64
         Left            =   5655
         TabIndex        =   214
         Top             =   5275
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   476
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
         Mask            =   ">&&"
         PromptChar      =   ">"
      End
      Begin VB.Frame frameDMS 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Otorisasi Disc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   7560
         TabIndex        =   218
         Top             =   5040
         Visible         =   0   'False
         Width           =   3255
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   65
            Left            =   1680
            TabIndex        =   219
            Top             =   240
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   66
            Left            =   1680
            TabIndex        =   222
            Top             =   600
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Mask            =   ">99999999"
            PromptChar      =   ">"
         End
         Begin VB.Label Label80 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rp."
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   1335
            TabIndex        =   224
            Top             =   660
            Width           =   315
         End
         Begin VB.Label Label79 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Plafon DP Part"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   240
            TabIndex        =   223
            Top             =   660
            Width           =   1050
         End
         Begin VB.Label Label78 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2280
            TabIndex        =   221
            Top             =   240
            Width           =   210
         End
         Begin VB.Label Label77 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Persentase DP Part"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   220
            Top             =   300
            Width           =   1395
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   68
         Left            =   2190
         TabIndex        =   227
         Top             =   7755
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Left            =   6600
         TabIndex        =   228
         Top             =   7755
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   69
         Left            =   5655
         TabIndex        =   230
         Top             =   5570
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   476
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
         Mask            =   ">&&"
         PromptChar      =   ">"
      End
      Begin VB.Frame frameDiscKacab 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Otoritas Disc"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   7020
         TabIndex        =   233
         Top             =   6165
         Visible         =   0   'False
         Width           =   4230
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   70
            Left            =   1350
            TabIndex        =   234
            Top             =   225
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   71
            Left            =   1350
            TabIndex        =   238
            Top             =   555
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   73
            Left            =   3345
            TabIndex        =   241
            Top             =   555
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            DataField       =   "m "
            Height          =   270
            Index           =   72
            Left            =   3345
            TabIndex        =   243
            Top             =   210
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   476
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
            Mask            =   ">99"
            PromptChar      =   ">"
         End
         Begin VB.Label Label90 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3915
            TabIndex        =   245
            Top             =   240
            Width           =   210
         End
         Begin VB.Label Label91 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Disc Part"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2595
            TabIndex        =   244
            Top             =   300
            Width           =   645
         End
         Begin VB.Label Label92 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3975
            TabIndex        =   242
            Top             =   555
            Width           =   210
         End
         Begin VB.Label Label89 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Disc Part Kacab"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   2145
            TabIndex        =   240
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Label85 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1830
            TabIndex        =   239
            Top             =   555
            Width           =   210
         End
         Begin VB.Label Label88 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Disc Jasa"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   135
            TabIndex        =   237
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label87 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "%"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1875
            TabIndex        =   236
            Top             =   255
            Width           =   210
         End
         Begin VB.Label Label86 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Disc Jasa Kacab"
            DataField       =   "m "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   105
            TabIndex        =   235
            Top             =   570
            Width           =   1200
         End
      End
      Begin MSMask.MaskEdBox Text 
         DataField       =   "m "
         Height          =   270
         Index           =   74
         Left            =   6150
         TabIndex        =   246
         Top             =   6120
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
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
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   88
         Left            =   10110
         TabIndex        =   309
         Top             =   7440
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   476
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
         Mask            =   ">99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   89
         Left            =   10095
         TabIndex        =   312
         Top             =   7800
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   476
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999"
         PromptChar      =   ">"
      End
      Begin VB.Label Label110 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Max Hari Booking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8715
         TabIndex        =   314
         Top             =   7845
         Width           =   1260
      End
      Begin VB.Label Label109 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " Hari"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   10860
         TabIndex        =   313
         Top             =   7830
         Width           =   330
      End
      Begin VB.Label Label108 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disc Max Sublet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8820
         TabIndex        =   311
         Top             =   7515
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label107 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10800
         TabIndex        =   310
         Top             =   7455
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label94 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         DataField       =   "m "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6800
         TabIndex        =   248
         Top             =   6150
         Width           =   210
      End
      Begin VB.Label Label93 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disc Part"
         DataField       =   "m "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5385
         TabIndex        =   247
         Top             =   6210
         Width           =   645
      End
      Begin VB.Label Label84 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Prosentase Follow Up :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3915
         TabIndex        =   232
         Top             =   5595
         Width           =   1635
      End
      Begin VB.Label Label83 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   231
         Top             =   5595
         Width           =   120
      End
      Begin VB.Label Label82 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Format Sms Billing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   795
         TabIndex        =   229
         Top             =   7800
         Width           =   1275
      End
      Begin VB.Label Label76 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Menit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6240
         TabIndex        =   216
         Top             =   5300
         Width           =   390
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estimasi Invoice :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4320
         TabIndex        =   215
         Top             =   5313
         Width           =   1230
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nomor Counter Booking"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   405
         TabIndex        =   213
         Top             =   7440
         Width           =   1695
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disc Max KBG"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4980
         TabIndex        =   211
         Top             =   5910
         Width           =   1035
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6800
         TabIndex        =   210
         Top             =   5895
         Width           =   210
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000080FF&
         Height          =   8175
         Left            =   0
         Top             =   0
         Width           =   11295
      End
      Begin VB.Label Label67 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hotline"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1605
         TabIndex        =   144
         Top             =   2580
         Width           =   495
      End
      Begin VB.Label Label66 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Estimasi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1170
         TabIndex        =   142
         Top             =   6510
         Width           =   930
      End
      Begin VB.Label Label65 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "MOU Limit"
         DataField       =   "m "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   140
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         DataField       =   "m "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4560
         TabIndex        =   139
         Top             =   5880
         Width           =   210
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Limit Unit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1395
         TabIndex        =   135
         Top             =   5595
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ukuran JPCB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1140
         TabIndex        =   133
         Top             =   7125
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Folder Backup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1065
         TabIndex        =   132
         Top             =   6795
         Width           =   1035
      End
      Begin VB.Label Label60 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Menit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3120
         TabIndex        =   130
         Top             =   6210
         Width           =   390
      End
      Begin VB.Label Label59 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Batas No Show"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   990
         TabIndex        =   129
         Top             =   6210
         Width           =   1110
      End
      Begin VB.Label Label58 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2790
         TabIndex        =   128
         Top             =   5880
         Width           =   210
      End
      Begin VB.Label Label57 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10800
         TabIndex        =   127
         Top             =   7140
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Upping Sublet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8970
         TabIndex        =   126
         Top             =   7200
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1125
         TabIndex        =   125
         Top             =   5895
         Width           =   975
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1515
         TabIndex        =   86
         Top             =   1605
         Width           =   555
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Administration Head"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   690
         TabIndex        =   81
         Top             =   4095
         Width           =   1410
      End
      Begin VB.Label Label29 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Stall"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   80
         Top             =   6435
         Visible         =   0   'False
         Width           =   840
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Retur"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1455
         TabIndex        =   78
         Top             =   5295
         Width           =   645
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Direktori Database AAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11550
         TabIndex        =   77
         Top             =   5910
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Direktori System AAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11730
         TabIndex        =   76
         Top             =   5610
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Telp"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1785
         TabIndex        =   72
         Top             =   2265
         Width           =   315
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. PKP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1485
         TabIndex        =   71
         Top             =   3225
         Width           =   615
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. NPWP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1305
         TabIndex        =   70
         Top             =   2925
         Width           =   795
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Perusahaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   69
         Top             =   1965
         Width           =   1380
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "detik"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   14265
         TabIndex        =   68
         Top             =   5250
         Width           =   1065
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Refresh Info All Repair"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11640
         TabIndex        =   67
         Top             =   5265
         Width           =   1590
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Acc. No."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1470
         TabIndex        =   66
         Top             =   4995
         Width           =   630
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bank"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1725
         TabIndex        =   65
         Top             =   4695
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Perkalian Return Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11655
         TabIndex        =   64
         Top             =   4980
         Width           =   1575
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Section Manager"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   63
         Top             =   8400
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "(Hari)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   15000
         TabIndex        =   62
         Top             =   7080
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat Printer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12255
         TabIndex        =   61
         Top             =   6225
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Batas Kredit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   12375
         TabIndex        =   60
         Top             =   7110
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1800
         TabIndex        =   59
         Top             =   4395
         Width           =   330
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kepala Cabang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1005
         TabIndex        =   58
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kepala Bengkel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   975
         TabIndex        =   57
         Top             =   3540
         Width           =   1125
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rate IRC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   11730
         TabIndex        =   56
         Top             =   4560
         Width           =   660
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rate PAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11685
         TabIndex        =   55
         Top             =   4260
         Width           =   705
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rate COM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   11640
         TabIndex        =   54
         Top             =   3960
         Width           =   750
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Invoice F. Pajak"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   11820
         TabIndex        =   53
         Top             =   6810
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Invoice INT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   12135
         TabIndex        =   52
         Top             =   6510
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Divisi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1695
         TabIndex        =   51
         Top             =   1305
         Width           =   375
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Perusahaan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   750
         TabIndex        =   50
         Top             =   1005
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No Invoice"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1290
         TabIndex        =   49
         Top             =   660
         Width           =   780
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No WO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1530
         TabIndex        =   48
         Top             =   345
         Width           =   540
      End
   End
   Begin ucXPButton.XPButton cmdSimpan 
      Height          =   495
      Left            =   9240
      TabIndex        =   195
      Top             =   8640
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
      MICON           =   "frmtprofil.frx":2D91C
      PICN            =   "frmtprofil.frx":2D938
      PICH            =   "frmtprofil.frx":2DE4B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdExit 
      Height          =   495
      Left            =   10560
      TabIndex        =   196
      Top             =   8640
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
      MICON           =   "frmtprofil.frx":2E3ED
      PICN            =   "frmtprofil.frx":2E409
      PICH            =   "frmtprofil.frx":2E8DD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ucXPButton.XPButton cmdValAR 
      Height          =   495
      Left            =   600
      TabIndex        =   194
      Top             =   8640
      Width           =   1875
      _ExtentX        =   3307
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
      MICON           =   "frmtprofil.frx":2EE28
      PICN            =   "frmtprofil.frx":2EE44
      PICH            =   "frmtprofil.frx":2F4C6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblKeteranganScoring 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   264
      Top             =   9360
      Width           =   5295
   End
   Begin VB.Label Label100 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Hari Sebelum JT"
      Height          =   240
      Index           =   5
      Left            =   2040
      TabIndex        =   260
      Top             =   45
      Width           =   1410
   End
   Begin VB.Label Label100 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Reminder CR 7"
      Height          =   240
      Index           =   4
      Left            =   0
      TabIndex        =   259
      Top             =   45
      Width           =   1290
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000080FF&
      X1              =   480
      X2              =   11760
      Y1              =   8520
      Y2              =   8520
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000080FF&
      X1              =   480
      X2              =   11760
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   5040
      Top             =   8760
      Width           =   495
   End
End
Attribute VB_Name = "frmtprofil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private support As New clsFormDealer
Private nkode As String
Private isOver As Boolean

Private cEM As Double
Private cNP As Double
Private cPro As Double
Private cNPro As Double
Private cTHS As Double
Private cSP As Double
Private rsSA As ADODB.Recordset
Dim rsKodeLgn   As New ADODB.Recordset

Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim xHasil As String
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC"
                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
            Case Else
                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
            End Select
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "PPN", "DISC"
                xHasil = xHasil + gsSql2(i)
            Case Else
                xHasil = xHasil + "'" + gsSql2(i) + "'"
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql = xHasil
End Function

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub ChkIERC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        ChkAAS.Visible = Not ChkAAS.Visible
    End If
End Sub

Private Sub cmdAturStall_Click()
    Me.Tag = "Stall"
    frmAturStall.show
End Sub

Private Sub cmdAturThs_Click()
    frmAturStall.Tag = "THS"
    frmAturStall.show
End Sub

Private Sub cmdDiscNRM_Click()
    If FrDisc.Visible = False Then
        FrDisc.Visible = True
    Else
        FrDisc.Visible = False
    End If
End Sub

Private Sub cmdDisconnect_Click()
    Call objO.UpdateData("UPDATE T_PROFIL SET KONEKSIHO = 'FAILED'")
    gStatKoneksi = False
    Log.LogHOConn = 0
    MsgBox "Koneksi ke HO di putus !"
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHansaRenkei_Click()
    frameHansaRenkei.Visible = True
    Dim rsHansaRenkei As ADODB.Recordset
    Set rsHansaRenkei = objO.FetchData("select * from settingHansaRenkei with (nolock)")
    If Not rsHansaRenkei.EOF Then
        txtReminderMRA.value = CInt(rsHansaRenkei!remindermra)
        txtReminderSA.value = CInt(rsHansaRenkei!remindersa)
    Else
        txtReminderMRA.value = 0
        txtReminderSA.value = 7
    End If
End Sub

Private Sub cmdMasterSettingCR_Click()
    frameMasterSetting.Visible = True
    frameMasterSetting.Left = 3600
    Dim rsMasterSetting As ADODB.Recordset
    
    Set rsMasterSetting = objO.FetchData("SELECT remindcr,leadremindlub,remindlub FROM MASTERREMINDERCR WITH (NOLOCK) ")
    If Not rsMasterSetting.EOF Then
        Text(75) = NulSID(rsMasterSetting!remindCR)
        Text(76) = NulSID(rsMasterSetting!leadRemindLub)
        Text(77) = NulSID(rsMasterSetting!remindLub)
    Else
        Text(75) = 0
        Text(76) = 0
        Text(77) = 0
    End If
End Sub

Private Sub cmdSave_Click()
Dim rsSearch As ADODB.Recordset

    ClearsSql
    If Not newEnkrip Then
        gsSql1(0) = "FS_HO_KOMP"
        gsSql2(0) = Trim(Replace(Replace(Text(37), ">", ""), " ", ""))
        gsSql1(1) = "FS_HO_USER"
        gsSql2(1) = Trim(Text(38))
        gsSql1(2) = "FS_HO_PASS"
        gsSql2(2) = Trim(txtPassHO)
        gsSql1(3) = "FS_HO_DB"
        gsSql2(3) = Trim(Text(39))
        
        gsSql1(4) = "FS_SPK_KOMP"
        gsSql2(4) = Trim(Replace(Replace(Text(34), ">", ""), " ", ""))
        gsSql1(5) = "FS_SPK_USER"
        gsSql2(5) = Trim(Text(35))
        gsSql1(6) = "FS_SPK_PASS"
        gsSql2(6) = Trim(txtpassSPK)
        gsSql1(7) = "FS_SPK_DB"
        gsSql2(7) = Trim(Text(36))
        
        gsSql1(8) = "FS_SPKHO_KOMP"
        gsSql2(8) = Trim(Replace(Replace(Text(31), ">", ""), " ", ""))
        gsSql1(9) = "FS_SPKHO_USER"
        gsSql2(9) = Trim(Text(32))
        gsSql1(10) = "FS_SPKHO_PASS"
        gsSql2(10) = Trim(txtPassSPKHO)
        gsSql1(11) = "FS_SPKHO_DB"
        gsSql2(11) = Trim(Text(33))
    
        gsSql1(12) = "FN_HO"
        gsSql2(12) = Trim(Text(40))
        gsSql1(13) = "FN_HOADRES"
        gsSql2(13) = Trim(Text(41))
        gsSql1(14) = "FN_HOKOTA"
        gsSql2(14) = Trim(Text(42))
        gsSql1(15) = "FN_HOTELP"
        gsSql2(15) = Trim(Text(43))
        gsSql1(16) = "FN_HONPWP"
        gsSql2(16) = Trim(Text(44))
        gsSql1(17) = "FN_HOKDPLG"
        gsSql2(17) = Trim(Text(45))
        gsSql1(18) = "TGLPKP"
        gsSql2(18) = Format(TglPkp, "MM/DD/YYYY")
    
    Else
        gsSql1(0) = "FN_HO"
        gsSql2(0) = Trim(Text(40))
        gsSql1(1) = "FN_HOADRES"
        gsSql2(1) = Trim(Text(41))
        gsSql1(2) = "FN_HOKOTA"
        gsSql2(2) = Trim(Text(42))
        gsSql1(3) = "FN_HOTELP"
        gsSql2(3) = Trim(Text(43))
        gsSql1(4) = "FN_HONPWP"
        gsSql2(4) = Trim(Text(44))
        gsSql1(5) = "FN_HOKDPLG"
        gsSql2(5) = Trim(Text(45))
        gsSql1(6) = "TGLPKP"
        gsSql2(6) = Format(TglPkp, "MM/DD/YYYY")
        
    End If
    
    
    
    
    Set rsSearch = support.listFieldInTable("T_PROFIL")
    If Not rsSearch.BOF Then
        sSQL = GensSql("T_PROFIL", True)
    Else
        sSQL = GensSql("T_PROFIL", False)
    End If
    Call objO.UpdateData(sSQL)
    
    MsgBox "Data Server disimpan "
    SetServer.Visible = False
End Sub

Private Sub cmdSaveDisc_Click()
    If Text(60).Text = "" Then Text(60).Text = 0
    If Text(61).Text = "" Then Text(61).Text = 0
    Call objSt.Update_DiscNRM(Text(60), Text(61))
    nDiscNRM = Val(Text(60))
    nDiscNON = Val(Text(61))
    MsgBox "Data telah disimpan !, Silahkan Keluar Dari NIS Service Untuk Mengaktifkan Perubahan Diskon (Login Ulang)"
End Sub

Private Sub cmdServer_Click()
    SetServer.Top = 720
    SetServer.Left = 720
    SetServer.Visible = True
    If newEnkrip Then
        setServerEnkrip.Visible = True
    End If
End Sub

Private Sub cmdSettingSimplify_Click()
    Dim rsSimplify As ADODB.Recordset
    Set rsSimplify = objO.FetchData("SELECT FS_FAX,FS_EMAIL,bankname1,bankname2,bankname3,bankno1,bankno2,bankno3,jamlayanan1,jamlayanan2 FROM T_PROFIL WITH (NOLOCK)")
    If Not rsSimplify.EOF Then
        Text(78) = Trim(rsSimplify!bankname1)
        Text(79) = Trim(rsSimplify!bankno1)
        
        Text(80) = Trim(rsSimplify!bankname2)
        Text(81) = Trim(rsSimplify!bankno2)
        
        Text(82) = Trim(rsSimplify!bankname3)
        Text(83) = Trim(rsSimplify!bankno3)
                
        Text(84) = Trim(rsSimplify!jamlayanan1)
        Text(85) = Trim(rsSimplify!jamlayanan2)
                
        Text(86) = Trim(rsSimplify!FS_FAX)
        Text(87) = Trim(rsSimplify!FS_EMAIL)
         
    End If
    Set rsSA = objO.FetchData("SELECT FS_KD_PETUGAS,FS_NM_PETUGAS,NOHP FROM T_petugas WITH (NOLOCK) WHERE tipe = 'SA'")
    gridSA2.DataSource = rsSA
    frmSimplify.Left = 4440
End Sub

Private Sub cmdShift_Click()
    frmTShift.show
End Sub

Private Sub cmdSimpan_Click()
    Dim cChange As Boolean
   On Error GoTo cmdSimpan_Click_Error

    Text(26) = Val(Text(56)) + Val(Text(57)) + Val(Text(58)) + Val(Text(59))
    cChange = False
    If (Val(Text(56)) <> cEM) Or (Val(Text(57)) <> cNP) Or (Val(Text(58)) <> cPro) Or (Val(Text(59)) <> cNPro Or (Val(Text(27)) <> cTHS) Or (Val(Text(67)) <> cSP)) Then
        cChange = True
        cEM = Text(56)
        cNP = Text(57)
        cPro = Text(58)
        cNPro = Text(59)
        cTHS = Text(27)
        cSP = Text(67)
    End If
    
    Dim rsSearch As ADODB.Recordset
    Dim Invoice As Byte
    Dim Order As Byte
    Dim Preview As Byte
    Dim Retur As Byte
    
    Call ClearsSql
    gsSql1(0) = "FN_NO_WO"
    gsSql2(0) = Val(Text(0))
    gsSql1(1) = "FN_RATE_IRC"
    gsSql2(1) = Val(Text(6))
    gsSql1(2) = "FS_PRINTER"
    gsSql2(2) = Text(13)
    gsSql1(3) = "FS_NM_PERUSH"
    gsSql2(3) = Text(7)
    gsSql1(4) = "FS_CABANG"
    gsSql2(4) = Text(8)
    gsSql1(5) = "FS_KA_BENGKEL"
    gsSql2(5) = Text(9)
    gsSql1(6) = "FS_KA_CABANG"
    gsSql2(6) = Text(10)
    gsSql1(7) = "FS_KOTA"
    gsSql2(7) = Text(11)
    gsSql1(8) = "FS_KALI_RET_RATE"
    gsSql2(8) = Val(Text(15))
    gsSql1(9) = "FN_NO_INV"
    gsSql2(9) = Val(Text(1))
    gsSql1(10) = "FN_RATE_COM"
    gsSql2(10) = Val(Text(4))
    gsSql1(11) = "FS_SEC_MGR"
    gsSql2(11) = Text(14)
    gsSql1(12) = "FS_BANK"
    gsSql2(12) = Text(16)
    gsSql1(13) = "FS_ACC"
    gsSql2(13) = Text(17)
    gsSql1(14) = "FN_REFRESH_INFO"
    gsSql2(14) = 0 'text(18)
    gsSql1(15) = "FS_ALM_PERUSH"
    gsSql2(15) = Text(19)
    gsSql1(16) = "FS_NPWP"
    gsSql2(16) = Text(20)
    gsSql1(17) = "FS_PKP"
    gsSql2(17) = Text(21)
    gsSql1(18) = "FS_TELP"
    gsSql2(18) = Text(22)
    gsSql1(19) = "FN_RATE_PAS"
    gsSql2(19) = Val(Text(5))
    
    
    
    If Option1.value = True Then
        Order = 0
    Else
        Order = 1
    End If
    If Option4.value = True Then
        Invoice = 0
    Else
        Invoice = 1
    End If
    If Option5.value = True Then
        Preview = 0
    Else
        Preview = 1
    End If
    If Option7.value = True Then
        Retur = 1
    Else
        Retur = 0
    End If
    gsSql1(20) = "FN_ORDER"
    gsSql2(20) = Order
    gsSql1(21) = "FN_INVOICE"
    gsSql2(21) = Invoice
    gsSql1(22) = "FN_PREVIEW"
    gsSql2(22) = Preview
    gsSql1(23) = "FN_IERC"
    gsSql2(23) = ChkIERC.value
    gsSql1(24) = "FS_SYS_AAS"
    gsSql2(24) = Text(23)
    gsSql1(25) = "FS_DB_AAS"
    gsSql2(25) = Text(24)
    gsSql1(26) = "FN_NO_RET"
    gsSql2(26) = Val(Text(25))
    gsSql1(27) = "FN_RETUR"
    gsSql2(27) = Retur
    gsSql1(28) = "FN_AAS"
    gsSql2(28) = ChkAAS.value
    gsSql1(29) = "FN_JML_STALL"
    gsSql2(29) = Val(Text(26))
    gsSql1(30) = "FN_JML_THS"
    gsSql2(30) = Val(Text(27))
    gsSql1(31) = "FS_KD_DEALER"
    gsSql2(31) = Trim(Text(29))
    gsSql1(32) = "FS_NM_DEALER"
    gsSql2(32) = Trim(Text(30))
    gsSql1(33) = "FN_PART"
    gsSql2(33) = IIf(chkSpart = Checked, 1, 0)
    gsSql1(34) = "FS_KA_ADH"
    gsSql2(34) = Trim(Text(28))
    gsSql1(35) = "DISCMAX"
    gsSql2(35) = CDbl(Text(46))
    gsSql1(36) = "UPSUBLET"
    gsSql2(36) = CDbl(Text(47))
    gsSql1(37) = "NOSHOW"
    gsSql2(37) = CDbl(Text(48))
    gsSql1(38) = "TGLPKP"
    gsSql2(38) = Format(TglPkp, "MM/DD/YYYY")
    gsSql1(39) = "FOLDERBACKUP"
    gsSql2(39) = Trim(Text(49))
    gsSql1(40) = "NOFONT"
    gsSql2(40) = CDbl(Text(50))
    gsSql1(41) = "NOHEIGHT"
    gsSql2(41) = CDbl(Text(51))
    gsSql1(42) = "LIMITDEALER"
    gsSql2(42) = CDbl(Text(52))
    gsSql1(43) = "FN_CASH"
    gsSql2(43) = IIf(chkCash = Checked, 1, 0)
    gsSql1(44) = "FN_QUEUE"
    gsSql2(44) = IIf(chQueue = Checked, 1, 0)
    gsSql1(45) = "ESTIMASIDEALER"
    gsSql2(45) = Trim(Text(54))
    gsSql1(46) = "LIMITMOU"
    gsSql2(46) = CDbl(Text(53))
    gsSql1(47) = "FS_HOTLINE"
    gsSql2(47) = Trim(Text(55))
    gsSql1(48) = "DISCMAXKBG"
    gsSql2(48) = CDbl(Text(62))
    gsSql1(49) = "FN_NO_BOOK"
    gsSql2(49) = Trim(Text(63))
    
    gsSql1(50) = "ESTINV"
    gsSql2(50) = Trim(Text(64))
    
    gsSql1(51) = "PERSENTASEDP"
    gsSql2(51) = Trim(Text(65))
    gsSql1(52) = "PLAFONDP"
    gsSql2(52) = Trim(Text(66))
    gsSql1(53) = "smsBilling"
    gsSql2(53) = Trim(Text(68))
    gsSql1(54) = "persenFollow"
    gsSql2(54) = CInt(Text(69))
    gsSql1(55) = "maxHariBooking"
    gsSql2(55) = CInt(Text(69))
    
    
    If Log.standarisasiSublet = 1 Then
        gsSql1(56) = "discmaxsublet"
        gsSql2(56) = CInt(Text(88))
    End If
    
    Set rsSearch = support.listFieldInTable("T_PROFIL")
    If Not rsSearch.BOF Then
        sSQL = GensSql("T_PROFIL", True)
    Else
        sSQL = GensSql("T_PROFIL", False)
    End If
    Call objO.UpdateData(sSQL)
    
    If Log.discKacab = 1 Then
        sSQL = "UPDATE t_profil SET discMax = " & CDbl(Text(70)) & " ,discPartMax = " & CDbl(Text(72)) & ",discMaxKbg = " & CDbl(Text(71)) & ",discPartKacabMax =  " & CDbl(Text(73))
        Call objO.UpdateData(sSQL)
    End If
        
    If ChkAAS.value = 1 Then
        gConnecAAS = True
    Else
        gConnecAAS = False
    End If
    
    '- sintak asli -
'    sSQL = "INSERT INTO LogProfilDealer (FN_NO_WO , FN_NO_INV, FS_PRINTER, FS_NM_PERUSH, FS_ALM_PERUSH, FS_NPWP, FS_PKP, FS_TELP, FS_CABANG, FS_KA_BENGKEL, FS_KA_CABANG, FS_KA_ADH, FS_KOTA, FN_BTS_KREDIT, FN_INV_EXT, " & _
'            "FN_INV_INT , FN_FAK_PAJAK, FN_RATE_COM, FN_RATE_PAS, FN_RATE_IRC, FS_SEC_MGR, FS_KALI_RET_RATE, FS_BANK, FS_ACC, FN_REFRESH_INFO, FN_ORDER, FN_INVOICE, FN_PREVIEW, FN_IERC,FS_SYS_AAS, FS_DB_AAS, FS_KD_FAKTUR, FS_GENERASI, FS_FISKAL, FN_NO_RET, FN_RETUR, FN_AAS, FS_GUDANG, FN_MAN_PART, FD_MAN_AW, FD_MAN_AK, FS_NIS_KOMP, FS_NIS_USER, FS_NIS_PASS, " & _
'            "FS_NIS_DB, FN_JML_STALL, FN_JML_THS, FS_HO_KOMP, FS_HO_USER, FS_HO_PASS, FS_HO_DB, FS_KD_DEALER, FS_NM_DEALER, FS_SPK_KOMP, FS_SPK_USER, FS_SPK_PASS, FS_SPK_DB, FS_SPKHO_KOMP,FS_SPKHO_USER, FS_SPKHO_PASS, FS_SPKHO_DB, LIMITKACAB, HUTANGKACAB, BAYARKACAB, LIMITKABENG, HUTANGKABENG, BAYARKABENG, FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP, " & _
'            "FN_HOKDPLG, FN_PART, DISCMAX, UPSUBLET, FS_NISSP_KOMP, FS_NISSP_USER, FS_NISSP_PASS, FS_NISSP_DB, NOSHOW, FS_NISSP_DB2, TglPkp, FN_HOPKP, KONEKSIHO, FOLDERBACKUP , NOFONT, NOHEIGHT, LIMITDEALER,USERENTRY, CLIENTNAME, TGLENTRY, FS_HOTLINE) SELECT FN_NO_WO , FN_NO_INV, FS_PRINTER, FS_NM_PERUSH, FS_ALM_PERUSH, FS_NPWP, FS_PKP, FS_TELP, FS_CABANG, FS_KA_BENGKEL, FS_KA_CABANG, FS_KA_ADH, FS_KOTA, FN_BTS_KREDIT, FN_INV_EXT, " & _
'            "FN_INV_INT , FN_FAK_PAJAK, FN_RATE_COM, FN_RATE_PAS, FN_RATE_IRC, FS_SEC_MGR, FS_KALI_RET_RATE, FS_BANK, FS_ACC, FN_REFRESH_INFO, FN_ORDER, FN_INVOICE, FN_PREVIEW, FN_IERC,FS_SYS_AAS, FS_DB_AAS, FS_KD_FAKTUR, FS_GENERASI, FS_FISKAL, FN_NO_RET, FN_RETUR, FN_AAS, FS_GUDANG, FN_MAN_PART, CONVERT( VARCHAR(10), FD_MAN_AW,101), CONVERT( VARCHAR(10), FD_MAN_AK,101), FS_NIS_KOMP, FS_NIS_USER, FS_NIS_PASS, " & _
'            "FS_NIS_DB, FN_JML_STALL, FN_JML_THS, FS_HO_KOMP, FS_HO_USER, FS_HO_PASS, FS_HO_DB, FS_KD_DEALER, FS_NM_DEALER, FS_SPK_KOMP, FS_SPK_USER, FS_SPK_PASS, FS_SPK_DB, FS_SPKHO_KOMP,FS_SPKHO_USER, FS_SPKHO_PASS, FS_SPKHO_DB, LIMITKACAB, HUTANGKACAB, BAYARKACAB, LIMITKABENG, HUTANGKABENG, BAYARKABENG, FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP, " & _
'            "FN_HOKDPLG, FN_PART, DISCMAX, UPSUBLET, FS_NISSP_KOMP, FS_NISSP_USER, FS_NISSP_PASS, FS_NISSP_DB, NOSHOW, FS_NISSP_DB2, CONVERT( VARCHAR(10), TGLPKP,105), FN_HOPKP, KONEKSIHO, FOLDERBACKUP , NOFONT, NOHEIGHT, LIMITDEALER, '" & gSA & "','" & ComptName & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_HOTLINE FROM " & gDb & "..T_PROFIL"
            
    '- sintak tambah fn_no_book
    sSQL = "INSERT INTO LogProfilDealer (FN_NO_WO , FN_NO_INV, FS_PRINTER, FS_NM_PERUSH, FS_ALM_PERUSH, FS_NPWP, FS_PKP, FS_TELP, FS_CABANG, FS_KA_BENGKEL, FS_KA_CABANG, FS_KA_ADH, FS_KOTA, FN_BTS_KREDIT, FN_INV_EXT, " & _
            "FN_INV_INT , FN_FAK_PAJAK, FN_RATE_COM, FN_RATE_PAS, FN_RATE_IRC, FS_SEC_MGR, FS_KALI_RET_RATE, FS_BANK, FS_ACC, FN_REFRESH_INFO, FN_ORDER, FN_INVOICE, FN_PREVIEW, FN_IERC,FS_SYS_AAS, FS_DB_AAS, FS_KD_FAKTUR, FS_GENERASI, FS_FISKAL, FN_NO_RET, FN_RETUR, FN_AAS, FS_GUDANG, FN_MAN_PART, FD_MAN_AW, FD_MAN_AK, FS_NIS_KOMP, FS_NIS_USER, FS_NIS_PASS, " & _
            "FS_NIS_DB, FN_JML_STALL, FN_JML_THS, FS_HO_KOMP, FS_HO_USER, FS_HO_PASS, FS_HO_DB, FS_KD_DEALER, FS_NM_DEALER, FS_SPK_KOMP, FS_SPK_USER, FS_SPK_PASS, FS_SPK_DB, FS_SPKHO_KOMP,FS_SPKHO_USER, FS_SPKHO_PASS, FS_SPKHO_DB, LIMITKACAB, HUTANGKACAB, BAYARKACAB, LIMITKABENG, HUTANGKABENG, BAYARKABENG, FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP, " & _
            "FN_HOKDPLG, FN_PART, DISCMAX, UPSUBLET, FS_NISSP_KOMP, FS_NISSP_USER, FS_NISSP_PASS, FS_NISSP_DB, NOSHOW, FS_NISSP_DB2, TglPkp, FN_HOPKP, KONEKSIHO, FOLDERBACKUP , NOFONT, NOHEIGHT, LIMITDEALER,USERENTRY, CLIENTNAME, TGLENTRY, FS_HOTLINE, FN_NO_BOOK,ESTINV) SELECT FN_NO_WO , FN_NO_INV, FS_PRINTER, FS_NM_PERUSH, FS_ALM_PERUSH, FS_NPWP, FS_PKP, FS_TELP, FS_CABANG, FS_KA_BENGKEL, FS_KA_CABANG, FS_KA_ADH, FS_KOTA, FN_BTS_KREDIT, FN_INV_EXT, " & _
            "FN_INV_INT , FN_FAK_PAJAK, FN_RATE_COM, FN_RATE_PAS, FN_RATE_IRC, FS_SEC_MGR, FS_KALI_RET_RATE, FS_BANK, FS_ACC, FN_REFRESH_INFO, FN_ORDER, FN_INVOICE, FN_PREVIEW, FN_IERC,FS_SYS_AAS, FS_DB_AAS, FS_KD_FAKTUR, FS_GENERASI, FS_FISKAL, FN_NO_RET, FN_RETUR, FN_AAS, FS_GUDANG, FN_MAN_PART, CONVERT( VARCHAR(10), FD_MAN_AW,101), CONVERT( VARCHAR(10), FD_MAN_AK,101), FS_NIS_KOMP, FS_NIS_USER, FS_NIS_PASS, " & _
            "FS_NIS_DB, FN_JML_STALL, FN_JML_THS, FS_HO_KOMP, FS_HO_USER, FS_HO_PASS, FS_HO_DB, FS_KD_DEALER, FS_NM_DEALER, FS_SPK_KOMP, FS_SPK_USER, FS_SPK_PASS, FS_SPK_DB, FS_SPKHO_KOMP,FS_SPKHO_USER, FS_SPKHO_PASS, FS_SPKHO_DB, LIMITKACAB, HUTANGKACAB, BAYARKACAB, LIMITKABENG, HUTANGKABENG, BAYARKABENG, FN_HO, FN_HOADRES, FN_HOKOTA, FN_HOTELP, FN_HONPWP, " & _
            "FN_HOKDPLG, FN_PART, DISCMAX, UPSUBLET, FS_NISSP_KOMP, FS_NISSP_USER, FS_NISSP_PASS, FS_NISSP_DB, NOSHOW, FS_NISSP_DB2, CONVERT( VARCHAR(10), TGLPKP,101), FN_HOPKP, KONEKSIHO, FOLDERBACKUP , NOFONT, NOHEIGHT, LIMITDEALER, '" & gSA & "','" & ComptName & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_HOTLINE, FN_NO_BOOK,ESTINV FROM " & gDb & "..T_PROFIL WITH(NOLOCK) "
    
    
    Call objO.InsertData_Profil(sSQL)
    
    If cChange = True Then
        Call objSt.AddNew_QtyStall(cEM, cNP, cPro, cNPro, cTHS, cSP)
        Call objSt.AddNew_QtyStall_Log(cEM, cNP, cPro, cNPro, cTHS, cSP)
    End If
    
    MsgBox "Data telah disimpan, Setelah melakukan perubahan keluar dari program bengkel, lalu masuk program bengkel kembali untuk perubahan yang terjadi !"

   On Error GoTo 0
   Exit Sub

cmdSimpan_Click_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmtprofil")
End Sub

Private Sub cmdSimpanHansaRenkei_Click()
    Dim rsHansaRenkei As ADODB.Recordset
    Set rsHansaRenkei = objO.FetchData("SELECT * FROM settingHansaRenkei WITH (NOLOCK)")
    If Not rsHansaRenkei.EOF Then
        Call objO.InsertData("UPDATE settingHansaRenkei SET reminderSA = '" & CInt(txtReminderSA) & "',reminderMRA = '" & CInt(txtReminderMRA) & "'")
    Else
        Call objO.InsertData("INSERT INTO settingHansaRenkei VALUES ('" & CInt(txtReminderSA) & "','" & CInt(txtReminderMRA) & "')")
    End If
    MsgBox "Data setting hansa renkei berhasil disimpan", vbInformation
    frameHansaRenkei.Visible = False
    
End Sub

Private Sub cmdSimpanSettingCR_Click()
          Dim boleh   As Boolean
          Dim pesan   As String
          Dim begin   As Boolean
          Dim rsData  As ADODB.Recordset
1        On Error GoTo cmdSimpanSettingCR_Click_Error
2         begin = False
3         boleh = True
4         pesan = ""
          
5         If CDbl(NulSID(Text(75))) = 0 Or CDbl(NulSID(Text(75))) > 14 Then
6             boleh = False
7             pesan = " Reminder CR 7 harus dalam range 1 sampe 14 hari." & vbCrLf
8         End If
          
9         If CDbl(NulSID(Text(76))) = 0 Or CDbl(NulSID(Text(76))) > 180 Then
10            boleh = False
11            pesan = " Reminder Lubing sesudah service harus dalam range 1 sampe 180 hari." & vbCrLf
12        End If
          
13        If CDbl(NulSID(Text(77))) = 0 Or CDbl(NulSID(Text(77))) > 14 Then
14            boleh = False
15            pesan = " Reminder Lubing dari JT harus dalam range 1 sampe 14 hari." & vbCrLf
16        End If
          
17        If boleh Then
18            Call objHT.OpenTransactionBengkel
19            begin = True
              
20            Set rsData = objHT.FetchDataBengkel("SELECT remindcr,leadremindlub,remindlub FROM MASTERREMINDERCR WITH (NOLOCK)")
21            If Not rsData.EOF Then
22                Call objHT.InsertData_Bengkel("UPDATE MASTERREMINDERCR SET remindcr = '" & NulSID(Text(75)) & "',leadremindlub = '" & NulSID(Text(76)) & "',remindlub = '" & NulSID(Text(77)) & "'")
23            Else
24                Call objHT.InsertData_Bengkel("INSERT INTO MASTERREMINDERCR (remindcr,leadremindlub,remindlub) VALUES ('" & NulSID(Text(75)) & "','" & NulSID(Text(76)) & "','" & NulSID(Text(77)) & "')")
25            End If
              
26            Call objHT.CloseTransactionBengkel
27            MsgBox "Data berhasil disimpan", vbInformation
28            frameMasterSetting.Left = 23600
29        Else
30            MsgBox pesan, vbInformation
31            Exit Sub
32        End If

33       On Error GoTo 0
34       Exit Sub

cmdSimpanSettingCR_Click_Error:
35        If begin = False Then
36            objHT.CloseTransactionBengkel_err
37        End If
38        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSimpanSettingCR_Click of Form frmtprofil"
End Sub

Private Sub cmdSimpanSimplify_Click()
   On Error GoTo cmdSimpanSimplify_Click_Error
    
    If Not rsSA.EOF Then
        rsSA.MoveFirst
        
        Do While Not rsSA.EOF
            sSQL = "UPDATE T_PETUGAS SET NOHP = '" & NulSID(gridSA2.Columns(2).value) & "' WHERE FS_KD_PETUGAS = '" & NulSID(rsSA!FS_KD_PETUGAS) & "'"
            ''20200928 || req ms adit update insert delete t_petugas ke repliksi tanpa flag hansarenkeibp
            ''----------------------------------------------------------------------------------------------
            If Log.RPLHistori = 1 Then
                sSQL = "UPDATE [" & dbReplikasiHO & "]..T_PETUGAS SET NOHP = '" & NulSID(gridSA2.Columns(2).value) & "' WHERE FS_KD_PETUGAS = '" & NulSID(rsSA!FS_KD_PETUGAS) & "' and dealer = '" & gDb & "'"
            End If
            Call objO.InsertData(sSQL)
            rsSA.MoveNext
        Loop
    End If
    Call ClearsSql
    gsSql1(0) = "bankname1"
    gsSql2(0) = Trim(Text(78))
    gsSql1(1) = "bankno1"
    gsSql2(1) = Trim(Text(79))
    
    gsSql1(2) = "bankname2"
    gsSql2(2) = Trim(Text(80))
    gsSql1(3) = "bankno2"
    gsSql2(3) = Trim(Text(81))
    
    gsSql1(4) = "bankname3"
    gsSql2(4) = Trim(Text(82))
    gsSql1(5) = "bankno3"
    gsSql2(5) = Trim(Text(83))
    
    gsSql1(6) = "jamlayanan1"
    gsSql2(6) = Trim(Text(84))
    gsSql1(7) = "jamlayanan2"
    gsSql2(7) = Trim(Text(85))
    
    gsSql1(8) = "FS_FAX"
    gsSql2(8) = Trim(Text(86))
    gsSql1(9) = "FS_EMAIL"
    gsSql2(9) = Trim(Text(87))
    
    sSQL = GensSql("T_PROFIL", True)
    Call objO.InsertData(sSQL)
    MsgBox "Data berhasil disimpan.", vbInformation
    frmSimplify.Left = 444000
   On Error GoTo 0
   Exit Sub

cmdSimpanSimplify_Click_Error:
    '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
    varErlError = Erl
    varErrNumber = err.Number
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpanSimplify_Click of Form frmtprofil")
    
End Sub

Private Sub cmdUp2_Click()
    If UCase(gTipeUser) = "KACAB" Or UCase(gTipeUser) = "ADMIN" Then
        If txtKet.Text = "" Then
            MsgBox "Keterangan Otorisasi harus diisi !"
            Exit Sub
        End If
        If nkode <> "" Then
            If (Round((txtRp(1) / IIf(txtRp(2) = 0, 1, txtRp(2))) * 100, 2)) > 100 Or isOver = True Then
                If MsgBox("Set Otoritas AR pada " & txt(107).Text & " ?", vbQuestion + vbYesNo) = vbYes Then
                    Call objO.AddNew_Valid_AR(nkode, txtRp(1), txtRp(2), txtKet.Text)
                    MsgBox "Otoritas AR telah di`set !"
                    nkode = ""
                    txt(100).Text = ""
                    Call ViewOtoritas
                End If
            Else
                MsgBox "Otoritas hanya dapat diproses pada :" & Chr(13) & "A.  Current AR > 100% Nilai Plafond" & Chr(13) & "B.  AR Overdue                                   "
            End If
        End If
    Else
        MsgBox "Otoritas hanya dapat diproses oleh Kepala Cabang !", vbInformation + vbOKOnly
    End If
End Sub

Private Sub cmdValAR_Click()
    If Me.cmdSimpan.Visible = True Then
        Me.cmdSimpan.Visible = False
    Else
        Me.cmdSimpan.Visible = True
    End If
    If FrameAR.Visible = True Then
        FrameAR.Visible = False
    Else
        FrameAR.Visible = True
        FrameAR.Left = 480
        Call ViewOtoritas
    End If
End Sub

Private Sub Form_Load()
      Dim rsSearch As ADODB.Recordset
          
1         If Log.singleIDScoring = 1 Then
2             lblKeteranganScoring.Caption = "Scoring Single ID Aktif"
3         End If
          
4         Call hapusForm
5         If Log.newCetak = 1 Then
6             cmdSettingSimplify.Visible = True
7         Else
8             cmdSettingSimplify.Visible = False
9         End If
          
10        Set rsSearch = support.listFieldInTable("T_PROFIL")
11        If Not rsSearch.BOF Then
12            Text(25) = Format(rsSearch("FN_NO_RET"), "00000000")
13            Text(0) = Format(rsSearch("FN_NO_WO"), "0000000")
14            Text(6) = rsSearch("FN_RATE_IRC")
15            Text(13) = rsSearch("FS_PRINTER")
16            Text(7) = rsSearch("FS_NM_PERUSH")
17            Text(8) = rsSearch("FS_CABANG")
18            Text(9) = rsSearch("FS_KA_BENGKEL")
19            Text(10) = rsSearch("FS_KA_CABANG")
20            Text(11) = rsSearch("FS_KOTA")
21            Text(1) = Format(rsSearch("FN_NO_INV"), "00000000")
22            Text(4) = rsSearch("FN_RATE_COM")
23            Text(5) = rsSearch("FN_RATE_PAS")
24            Text(14) = rsSearch("FS_SEC_MGR")
25            Text(15) = rsSearch("FS_KALI_RET_RATE")
26            Text(16) = rsSearch("FS_BANK")
27            Text(17) = rsSearch("FS_ACC")
28            Text(18) = rsSearch("FN_REFRESH_INFO")
29            Text(19) = rsSearch("FS_ALM_PERUSH")
30            Text(20) = rsSearch("FS_NPWP")
31            Text(21) = rsSearch("FS_PKP")
32            Text(22) = rsSearch("FS_TELP")
33            Text(55) = rsSearch("FS_HOTLINE")
              
34            If rsSearch("FN_ORDER") = False Then
35                Option1.value = True
36            Else
37                Option2.value = True
38            End If
39            If rsSearch("FN_INVOICE") = False Then
40                Option4.value = True
41            Else
42                Option3.value = True
43            End If
44            If rsSearch("FN_PREVIEW") = False Then
45                Option5.value = True
46            Else
47                Option6.value = True
48            End If
49            If rsSearch("FN_IERC") = True Then
50                ChkIERC.value = 1
51            Else
52                ChkIERC.value = 0
53            End If
54            If rsSearch("FN_RETUR") = True Then
55                Option7.value = True
56            Else
57                Option8.value = True
58            End If
59            Text(23) = rsSearch("FS_SYS_AAS")
60            Text(24) = rsSearch("FS_DB_AAS")
61            Text(26) = Nul(rsSearch("FN_JML_STALL"), 0)
62            Text(27) = Nul(rsSearch("FN_JML_THS"), 0)
63            If rsSearch("FN_AAS") = True Then
64                ChkAAS.value = 1
65            Else
66                ChkAAS.value = 0
67            End If
68            chkSpart.value = IIf(rsSearch("FN_PART"), Checked, Unchecked)
69            chkCash.value = IIf(rsSearch("FN_CASH"), Checked, Unchecked)
70            chQueue.value = IIf(rsSearch("FN_QUEUE") = "1", Checked, Unchecked)
71            Text(29) = rsSearch("FS_KD_DEALER")
72            Text(30) = rsSearch("FS_NM_DEALER")
73            Text(28) = rsSearch("FS_KA_ADH")
74            Text(46) = rsSearch("DISCMAX")
75            Text(62) = rsSearch("DISCMAXKBG")
76            Text(47) = rsSearch("UPSUBLET")
77            Text(53) = rsSearch("LIMITMOU")
78            Text(54) = rsSearch("ESTIMASIDEALER")
79            Text(64) = CInt(rsSearch("ESTINV"))
              
80            Text(65) = CInt(rsSearch("PERSENTASEDP"))
81            Text(66) = CDbl(rsSearch("PLAFONDP"))
82            Text(68) = Trim(rsSearch("SMSBILLING"))
83            Text(69) = CInt(rsSearch("persenFollow"))
              
              ' DATA SERVER
84            Text(37) = rsSearch("FS_HO_KOMP")
85            Text(38) = rsSearch("FS_HO_USER")
86            txtPassHO = rsSearch("FS_HO_PASS")
87            Text(39) = rsSearch("FS_HO_DB")
          
88            Text(34) = rsSearch("FS_SPK_KOMP")
89            Text(35) = rsSearch("FS_SPK_USER")
90            txtpassSPK = rsSearch("FS_SPK_PASS")
91            Text(36) = rsSearch("FS_SPK_DB")
              
92            Text(31) = rsSearch("FS_SPKHO_KOMP")
93            Text(32) = rsSearch("FS_SPKHO_USER")
94            txtPassSPKHO = rsSearch("FS_SPKHO_PASS")
95            Text(33) = rsSearch("FS_SPKHO_DB")
              
96            Text(40) = rsSearch("FN_HO")
97            Text(41) = rsSearch("FN_HOADRES")
98            Text(42) = rsSearch("FN_HOKOTA")
99            Text(43) = rsSearch("FN_HOTELP")
100           Text(44) = rsSearch("FN_HONPWP")
101           Text(45) = rsSearch("FN_HOKDPLG")
102           Text(48) = rsSearch("NOSHOW")
103           TglPkp = Format(rsSearch("TGLPKP"))
104           Text(49) = Trim(rsSearch("FOLDERBACKUP"))
              
105           Text(50) = rsSearch("NOFONT")
106           Text(51) = rsSearch("NOHEIGHT")
107           Text(52) = rsSearch("LIMITDEALER")
108           Text(63) = rsSearch("FN_NO_BOOK")
109           Text(74) = Trim(rsSearch!discPartMax)
110           Text(89) = Trim(rsSearch!maxHariBooking)
111           If Log.discKacab = 1 Then
112               frameDiscKacab.Visible = True
113               Text(70) = Trim(rsSearch!discmax)
114               Text(72) = Trim(rsSearch!discPartMax)
115               Text(71) = Trim(rsSearch!DISCMAXKBG)
116               Text(73) = Trim(rsSearch!discpartkacabmax)
                  
117               Text(46).Enabled = False
118               Text(74).Enabled = False
119               Text(62).Enabled = False
                  
120           Else
121               frameDiscKacab.Visible = False
122               Text(46).Enabled = True
123               Text(74).Enabled = True
124               Text(62).Enabled = True
                  
125           End If
               
126           Text(50).Visible = False
127           Text(51).Visible = False
              
128           Text(52).Visible = True
              
129           Label62.Visible = False
130           Label63.Visible = True
              
              'nama server
131           Text(37).Enabled = False
132           Text(38).Enabled = False
133           Text(39).Enabled = False
134           txtPassHO.Enabled = False
135           Text(34).Enabled = False
136           Text(35).Enabled = False
137           Text(36).Enabled = False
138           txtpassSPK.Enabled = False
139           Text(31).Enabled = False
140           Text(32).Enabled = False
141           Text(33).Enabled = False
142           txtPassSPKHO.Enabled = False
              
143           If gUstat = 1 Then
144               Frame2.Visible = False
145               Frame3.Visible = False
146               Frame4.Visible = False
147               Frame5.Visible = False
148               Label10.Caption = "Kepala Divisi 1"
                  'Label11.Caption = "Kepala Divisi 2"
149               Label16.Caption = "Manager"
150               Label31.Caption = "User"
151           Else
152               Label10.Caption = "Kepala Bengkel"
                  'Label11.Caption = "Kepala Cabang"
153               Label16.Caption = "Section Manager"
154               Label31.Caption = "Administration Head"
155           End If
156       End If
157       cmdServer.Visible = False
158       If gTipeUser = "ADMIN" Then
159           Text(50).Visible = True
160           Text(51).Visible = True
161           Text(52).Visible = True
162           Label62.Visible = True
163           Label63.Visible = True
164           ChkIERC.Visible = True
165           chkSpart.Visible = True
166           ChkAAS.Visible = True
167           chkCash.Visible = True
168           cmdDisconnect.Visible = True
169           cmdServer.Visible = True
              'nama server
170           Text(37).Enabled = True
171           Text(38).Enabled = True
172           Text(39).Enabled = True
173           txtPassHO.Enabled = True
174           Text(34).Enabled = True
175           Text(35).Enabled = True
176           Text(36).Enabled = True
177           txtpassSPK.Enabled = True
178           Text(31).Enabled = True
179           Text(32).Enabled = True
180           Text(33).Enabled = True
181           txtPassSPKHO.Enabled = True
            
182       End If
183       Text(26).Enabled = False
184       Text(27).Enabled = False
185       Text(48).Enabled = False
186       Text(46).Enabled = False
187       Text(53).Enabled = False
188       Text(62).Enabled = False
189       Text(52).Enabled = False
          
190       Text(56).Enabled = False
191       Text(57).Enabled = False
192       Text(58).Enabled = False
193       Text(59).Enabled = False
194       frameDiscKacab.Visible = False
              
195       Text(89).Enabled = False
196       If gTipeUser = "HO" Or UCase(gTipeUser) = "ADMIN" Then
197           Text(89).Enabled = True
198           Text(26).Enabled = True
199           Text(27).Enabled = True
200           Text(48).Enabled = True
201           Text(46).Enabled = True
202           Text(53).Enabled = True
203           Text(62).Enabled = True
              'cmdDiscNRM.Visible = True
204           Text(52).Enabled = True
              
205           Text(56).Enabled = True
206           Text(57).Enabled = True
207           Text(58).Enabled = True
208           Text(59).Enabled = True
209           frameDMS.Visible = True
              
210           If Log.discKacab = 1 Then
211               frameDiscKacab.Visible = True
                  
212           End If
              
213           If Log.standarisasiSublet = 1 Then
214               Label108.Visible = True
215               Text(88).Visible = True
216               Text(88) = CInt(rsSearch!discMaxSublet)
217               Label107.Visible = True
218           Else
219               Label108.Visible = False
220               Text(88).Visible = False
221               Text(88) = 0
222               Label107.Visible = False
223           End If
224       Else
225           If Log.standarisasiSublet = 1 Then
226               Label108.Visible = False
227               Text(88).Visible = False
228               Text(88) = CInt(rsSearch!discMaxSublet)
229               Label107.Visible = False
230           Else
231               Label108.Visible = False
232               Text(88).Visible = False
233               Text(88) = 0
234               Label107.Visible = False
235           End If
236       End If
          
          Dim rsSt As ADODB.Recordset
237       Set rsSt = objSt.FetchQtyStall
238       If rsSt.RecordCount = 0 Then
239           Text(56) = 0
240           Text(57) = 0
241           Text(58) = 0
242           Text(59) = 0
243           Text(67) = 0
244           cEM = 0
245           cNP = 0
246           cPro = 0
247           cNPro = 0
248           cTHS = 0
249           cSP = 0
250       Else
251           Text(56) = rsSt.Fields(1).value
252           Text(57) = rsSt.Fields(2).value
253           Text(58) = rsSt.Fields(3).value
254           Text(59) = rsSt.Fields(4).value
255           Text(67) = rsSt.Fields(7).value
256           cEM = rsSt.Fields(1).value
257           cNP = rsSt.Fields(2).value
258           cPro = rsSt.Fields(3).value
259           cNPro = rsSt.Fields(4).value
260           cTHS = rsSt.Fields(6).value
261           cSP = rsSt.Fields(7).value
262       End If
          
          Dim rsDisc As ADODB.Recordset
263       Set rsDisc = objSt.FetchDiscNRM
264       Text(60) = rsDisc.Fields(0).value
265       Text(61) = rsDisc.Fields(1).value
          
266       frameHansaRenkei.Visible = False
267       If Log.hansaRenkei = 1 Then
268           cmdHansaRenkei.Visible = True
269       Else
270           cmdHansaRenkei.Visible = False
271       End If
End Sub


Private Sub Image1_Click()
    MsgBox Text(26).Text
End Sub

Private Sub tblSimpan_Click(Index As Integer)
    Dim rsPel As ADODB.Recordset
    rsKodeLgn.filter = ""
    rsKodeLgn.filter = "pilih=" & True
    If Not rsKodeLgn.EOF Then
        txt(0) = NulSID(rsKodeLgn!kodeLgn)
        Set rsPel = objO.FetchPelangganACC(" Where a.KodeLgn ='" & txt(0).Text & "'")
        If rsPel.RecordCount > 0 Then
            frmPilihKodeLgn.Visible = False
            Call BindToControlWO(rsPel)
        Else
            txt(100).Text = ""
            nkode = ""
        End If
        rsKodeLgn.filter = ""
    Else
        MsgBox "Kodelgn harus dipilih", vbInformation
        rsKodeLgn.filter = ""
        Exit Sub
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
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 29
                gMainKeyCari = Text(Index)
                gBantuan = "INTERNAL"
                gSearchOption = "DEALER"
                frmCari.show (vbModal)
            Case 45
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLGsid"
                frmCariAcess.show (vbModal)
            Case 54
                gMainKeyCari = Text(Index)
                gSearchOption = "TIPEESTMASI"
                frmCari.show (vbModal)
            Case 68
                gMainKeyCari = Text(Index)
                gSearchOption = "FORMATSMS"
                frmCari.show (vbModal)
        End Select
    End If
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
End Sub

Private Sub hapusForm()
    Dim i As Integer
    
    For i = 0 To 52
        Select Case i
            Case 48, 26, 46, 27, 47
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    FrameAR.Left = 480
    Text(67) = 0
    If Log.CR7 = 1 Then
        cmdMasterSettingCR.Visible = True
    Else
        cmdMasterSettingCR.Visible = False
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
      Dim rsSearch As ADODB.Recordset

1         Select Case Index
              Case 9
2                 Text(14) = Trim(Text(9))
3             Case 29
4                 If Text(Index) <> "" Then
                      
5                     If Log.RPLHistori = 0 Then
6                         Set rsSearch = objO.FetchData("SELECT * FROM T_DEALER WITH(NOLOCK) WHERE INTERNAL = 1 AND FS_KD_DEALER = '" & Trim(Text(Index)) & "'")
7                     Else
8                         Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_DEALER WITH(NOLOCK) WHERE INTERNAL = 1 AND FS_KD_DEALER = '" & Trim(Text(Index)) & "'")
9                     End If
10                    If rsSearch.EOF Then
11                        gMainKeyCari = Text(Index)
12                        gBantuan = "INTERNAL"
13                        gSearchOption = "DEALER"
14                        frmCari.show (vbModal)
15                    Else
16                        Text(Index + 1) = rsSearch("FS_NM_DEALER")
17                    End If
18                End If
19            Case 45
20                If Text(Index) <> "" Then
21                        Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp from Armst WITH(NOLOCK) where KodeLgn = '" & Trim(Text(Index)) & "'")
22                        If rsSearch.BOF Then
23                            Cancel = True
24                            gMainKeyCari = Text(Index)
25                            gSearchOption = "KODEPLG"
26                            frmCariAcess.show (vbModal)
27                        End If
28                End If
29            Case 50
30                If Text(Index) <> "" Then
31                    Text(51) = CDbl(Text(50)) / 8 * 240
32                End If
33            Case 54
34                If Text(Index) <> "" Then
35                    Set rsSearch = objO.FetchData("SELECT * FROM T_HITESTIMASI WITH(NOLOCK) WHERE KETERANGAN = '" & Trim(Text(Index)) & "'")
36                    If rsSearch.EOF Then
37                        gMainKeyCari = Text(Index)
38                        gSearchOption = "TIPEESTMASI"
39                        frmCari.show (vbModal)
40                    End If
41                End If
42        End Select
End Sub

'================================================== CONTROL AR ==========

Private Sub ViewOtoritas()
    Dim rsAr As ADODB.Recordset
    Set rsAr = objO.FetchData_Valid_AR
    If rsAr.RecordCount > 0 Then
        Grid.DataSource = rsAr
    Else
        Grid.DataSource = Nothing
    End If
End Sub


Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim rsPel As ADODB.Recordset
        Dim d As Double: d = 1
        sSQL = " select case when SUBSTRING(a.UserEntry,1,4) = 'GRP#' then 1 else 0 end as pilih,b.KodeLgn,b.NamaLgn,b.Alamat1 + ' ' + b.alamat2 + ' ' + b.Alamat3 as alamat,b.npwp"
        sSQL = sSQL & " from [SingleIDOri]..Bridging a"
        sSQL = sSQL & " inner join ARMST b on a.KodeLgn = b.kodelgn"
        sSQL = sSQL & " where a.KodeDealer = '" & Right(gDb, 2) & "' and a.idcust = '" & NulSID(txt(100)) & "'"
        Set rsTemp = objO.FetchDataACC(sSQL)
        If Not rsTemp.EOF Then
            Call createTableKodelgn
            Dim isGrp As Boolean
            If rsTemp.RecordCount > 1 Then
                rsTemp.MoveFirst
                isGrp = False
                Do While Not rsTemp.EOF
                    With rsKodeLgn
                        .AddNew
                        
                        !Pilih = IIf(rsTemp!Pilih = "1" And isGrp = False, True, False)
                        !kodeLgn = NulSID(rsTemp!kodeLgn)
                        !Namalgn = NulSID(rsTemp!Namalgn)
                        !Alamat = NulSID(rsTemp!Alamat)
                        !npwp = NulSID(rsTemp!npwp)
                        !nom = d
                        If rsTemp!Pilih = "1" Then
                            isGrp = True
                        End If
                        .Update
                    End With
                    d = d + 1
                    rsTemp.MoveNext
                Loop
                gridKodelgn.DataSource = rsKodeLgn
                frmPilihKodeLgn.Visible = True
            Else
                txt(0) = NulSID(rsTemp!kodeLgn)
            End If
        Else
            MsgBox "Kodelgn tidak ditemukan", vbInformation
            frmPilihKodeLgn.Visible = False
            Exit Sub
        End If
        Set rsPel = objO.FetchPelangganACC(" Where a.KodeLgn ='" & txt(0).Text & "'")
        If rsPel.RecordCount > 0 Then
            Call BindToControlWO(rsPel)
        Else
            txt(100).Text = ""
            nkode = ""
        End If
    End If
End Sub

Private Sub createTableKodelgn()

   On Error GoTo createTableKodelgn_Error
    If rsKodeLgn.State Then rsKodeLgn.Close
    Set rsKodeLgn = Nothing
    With rsKodeLgn
        .Fields.Append "Pilih", adBoolean
        .Fields.Append "KodeLgn", adVarChar, 20
        .Fields.Append "NamaLgn", adVarChar, 200
        .Fields.Append "alamat", adVarChar, 300
        .Fields.Append "Npwp", adVarChar, 50
        .Fields.Append "nom", adDouble
        .Open
    End With

   On Error GoTo 0
   Exit Sub

createTableKodelgn_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createTableKodelgn of Form frmtprofil"
End Sub

Private Sub BindToControlWO(rsG As ADODB.Recordset)
    With Me
        txt(100).Text = Trim(rsG.Fields(7).value)
        txt(0).Text = Trim(rsG.Fields(0).value)
        nkode = Trim(rsG.Fields(0).value)
        txt(107).Text = rsG.Fields(1).value
        txt(108).Text = rsG.Fields(2).value
        txt(109).Text = Nul(rsG.Fields(3).value)
        txt(110).Text = rsG.Fields(4).value
        Dim cAR As Double
        cAR = objO.Fetch_Current_Ar(nkode)
        If cAR < 0 Then
            MsgBox "Ada Titipan senilai " & Format(Abs(cAR), "###,###,###.00") & ", belum dilakukan DUM !"
            cAR = 0
        End If
        txtRp(1) = cAR
        txtRp(2) = objO.Fetch_BatasKredit_Ar(nkode)
        isOver = False
        Call isOverdueAR
        If txtRp(1) <> 0 And txtRp(2) <> 0 Then
            nAR.Caption = Round((txtRp(1) / IIf(txtRp(2) = 0, 1, txtRp(2))) * 100, 0) & "%"
        Else
            nAR.Caption = ""
        End If
    End With
    If isOver = True Then nAR.Caption = nAR.Caption & " OVERDUE"
    Call CexValidasi
End Sub

Private Sub isOverdueAR()
    Dim nOAR As ADODB.Recordset
    'Set nOAR = objO.FetchOverdueAR(Trim(Replace(txt(100).Text, ">", "")))
    Set nOAR = objO.FetchOverdueAR(Trim(Replace(txt(0).Text, ">", "")))
    If nOAR.RecordCount > 0 Then
        ''Set nOAR = objO.FetchOverdueAR_Data(Trim(Replace(txt(100).Text, ">", "")))
        Set nOAR = objO.FetchOverdueAR_Data(Trim(Replace(txt(0).Text, ">", "")))
        If nOAR.RecordCount > 0 Then
            nAR.Caption = nAR.Caption & "  ( AR OVERDUE )"
            isOver = True
        Else
            isOver = False
        End If
    Else
        isOver = False
    End If
End Sub

Private Sub CexValidasi()
    Dim rsV As ADODB.Recordset
    Set rsV = objO.Fetch_Valid_AR(nkode)
    If rsV.RecordCount > 0 Then
        MsgBox "Otoritas AR sudah diproses pada " & Chr(13) & rsV.Fields(5).value & "  ( " & rsV.Fields(4).value & " )"
        cmdUp2.Enabled = False
    Else
        cmdUp2.Enabled = True
    End If
    
    If isOver = False Then
        If (Round((txtRp(1) / IIf(txtRp(2) = 0, 1, txtRp(2))) * 100, 0) < 80) Then
            MsgBox "Otoritas hanya dapat diproses pada :" & Chr(13) & "1. Current AR > 100% Atas Nilai Plafond " & Chr(13) & "2. AR Overdue     "
            cmdUp2.Enabled = False
            Exit Sub
        End If
    End If
End Sub


Private Sub txt_GotFocus(Index As Integer)
    If txt(Index).BackColor = fColor Then Sendkeys ("{TAB}")
End Sub

Private Sub txtRp_GotFocus(Index As Integer)
    If txtRp(Index).BackColor = fColor Then Sendkeys ("{TAB}")
End Sub

Private Sub XPButton1_Click()
    FrameAR.Visible = False
End Sub

Private Sub XPButton2_Click()

End Sub

Private Sub XPButton3_Click()
    frmSimplify.Left = 444000
    
End Sub
