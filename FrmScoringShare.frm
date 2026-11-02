VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form FrmScoringShare 
   BorderStyle     =   0  'None
   Caption         =   "Customer Baru"
   ClientHeight    =   9045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   18870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Banner 
      BorderStyle     =   0  'None
      Height          =   8700
      Left            =   120
      TabIndex        =   11
      Top             =   300
      Width           =   18660
      Begin VB.VScrollBar VScroll1 
         Height          =   2175
         Left            =   18425
         TabIndex        =   18
         Top             =   0
         Width           =   225
      End
      Begin VB.Frame FrameTombol 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   18420
         Begin ucXPButton.XPButton CmdAdd 
            Height          =   450
            Left            =   360
            TabIndex        =   16
            Top             =   120
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   794
            BTYPE           =   3
            TX              =   "Search"
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmScoringShare.frx":0000
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ucXPButton.XPButton CmdQuit 
            Height          =   450
            Left            =   7560
            TabIndex        =   17
            Top             =   120
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   794
            BTYPE           =   3
            TX              =   "Exit"
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   14215660
            BCOLO           =   14215660
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmScoringShare.frx":001C
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   7965
         Left            =   0
         TabIndex        =   12
         Top             =   600
         Width           =   18420
         Begin VB.TextBox txtFindNama2 
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
            Height          =   320
            IMEMode         =   3  'DISABLE
            Left            =   5295
            MaxLength       =   30
            TabIndex        =   4
            Top             =   780
            Width           =   2655
         End
         Begin VB.TextBox txtFindNama1 
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
            Height          =   320
            IMEMode         =   3  'DISABLE
            Left            =   2535
            MaxLength       =   30
            TabIndex        =   3
            Top             =   780
            Width           =   2655
         End
         Begin VB.TextBox txtFindKTP 
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
            Height          =   320
            IMEMode         =   3  'DISABLE
            Left            =   2535
            MaxLength       =   50
            TabIndex        =   1
            Top             =   120
            Width           =   2655
         End
         Begin VB.TextBox TxtFindHP 
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
            Height          =   320
            IMEMode         =   3  'DISABLE
            Left            =   2535
            MaxLength       =   20
            TabIndex        =   2
            Top             =   450
            Width           =   2655
         End
         Begin VB.TextBox txtFindNama3 
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
            Height          =   320
            IMEMode         =   3  'DISABLE
            Left            =   8055
            MaxLength       =   150
            TabIndex        =   5
            Top             =   780
            Width           =   2655
         End
         Begin VB.Frame Frame3 
            Height          =   5010
            Left            =   360
            TabIndex        =   13
            Top             =   2880
            Width           =   18015
            Begin TrueOleDBGrid70.TDBGrid GridFind 
               Height          =   4485
               Left            =   120
               TabIndex        =   14
               Top             =   390
               Width           =   17655
               _ExtentX        =   31141
               _ExtentY        =   7911
               _LayoutType     =   4
               _RowHeight      =   -2147483647
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "Total"
               Columns(0).DataField=   "TotalScore"
               Columns(0).NumberFormat=   "#,##0.#0;(#,##0.#0)"
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "Add"
               Columns(1).DataField=   "AddScore"
               Columns(1).NumberFormat=   "#,##0.#0;(#,##0.#0)"
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "Nama"
               Columns(2).DataField=   "Score"
               Columns(2).NumberFormat=   "#,##0.#0;(#,##0.#0)"
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   0
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "Kode"
               Columns(3).DataField=   "KodeLgn"
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(4)._VlistStyle=   0
               Columns(4)._MaxComboItems=   5
               Columns(4).Caption=   "Tgl Entry"
               Columns(4).DataField=   "TglEntry"
               Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(5)._VlistStyle=   0
               Columns(5)._MaxComboItems=   5
               Columns(5).Caption=   "User Entry"
               Columns(5).DataField=   "UserEntry"
               Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(6)._VlistStyle=   0
               Columns(6)._MaxComboItems=   5
               Columns(6).Caption=   "No KTP"
               Columns(6).DataField=   "NoKTP"
               Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(7)._VlistStyle=   0
               Columns(7)._MaxComboItems=   5
               Columns(7).Caption=   "Nama 1"
               Columns(7).DataField=   "NamaCust1"
               Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(8)._VlistStyle=   0
               Columns(8)._MaxComboItems=   5
               Columns(8).Caption=   "Nama 2"
               Columns(8).DataField=   "NamaCust2"
               Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(9)._VlistStyle=   0
               Columns(9)._MaxComboItems=   5
               Columns(9).Caption=   "Nama 3"
               Columns(9).DataField=   "NamaCust3"
               Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(10)._VlistStyle=   0
               Columns(10)._MaxComboItems=   5
               Columns(10).Caption=   "Type"
               Columns(10).DataField=   "TypeCustomer"
               Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(11)._VlistStyle=   0
               Columns(11)._MaxComboItems=   5
               Columns(11).Caption=   "No HP 1"
               Columns(11).DataField=   "HP1"
               Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(12)._VlistStyle=   0
               Columns(12)._MaxComboItems=   5
               Columns(12).Caption=   "No HP 2"
               Columns(12).DataField=   "HP2"
               Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(13)._VlistStyle=   0
               Columns(13)._MaxComboItems=   5
               Columns(13).Caption=   "Tempat Lahir"
               Columns(13).DataField=   "TempatLahir"
               Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(14)._VlistStyle=   0
               Columns(14)._MaxComboItems=   5
               Columns(14).Caption=   "Tgl Lahir"
               Columns(14).DataField=   "TglLahir"
               Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(15)._VlistStyle=   0
               Columns(15)._MaxComboItems=   5
               Columns(15).Caption=   "Alamat"
               Columns(15).DataField=   "Alamat"
               Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(16)._VlistStyle=   0
               Columns(16)._MaxComboItems=   5
               Columns(16).Caption=   "Propinsi"
               Columns(16).DataField=   "Propinsi"
               Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(17)._VlistStyle=   0
               Columns(17)._MaxComboItems=   5
               Columns(17).Caption=   "Kota"
               Columns(17).DataField=   "Kota"
               Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(18)._VlistStyle=   0
               Columns(18)._MaxComboItems=   5
               Columns(18).Caption=   "Kecamatan"
               Columns(18).DataField=   "Kecamatan"
               Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(19)._VlistStyle=   0
               Columns(19)._MaxComboItems=   5
               Columns(19).Caption=   "Kelurahan"
               Columns(19).DataField=   "Kelurahan"
               Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(20)._VlistStyle=   0
               Columns(20)._MaxComboItems=   5
               Columns(20).Caption=   "Kode Pos"
               Columns(20).DataField=   "KodePos"
               Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   21
               Splits(0)._UserFlags=   0
               Splits(0).RecordSelectorWidth=   688
               Splits(0)._SavedRecordSelectors=   0   'False
               Splits(0).AlternatingRowStyle=   -1  'True
               Splits(0).DividerColor=   14215660
               Splits(0).FilterBar=   -1  'True
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=21"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=794"
               Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(5)=   "Column(1).Width=873"
               Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=794"
               Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(9)=   "Column(2).Width=873"
               Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=794"
               Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(13)=   "Column(3).Width=2117"
               Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=2037"
               Splits(0)._ColumnProps(16)=   "Column(3)._ColStyle=260"
               Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
               Splits(0)._ColumnProps(18)=   "Column(3)._FootDivider=0"
               Splits(0)._ColumnProps(19)=   "Column(4).Width=2725"
               Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
               Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
               Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
               Splits(0)._ColumnProps(23)=   "Column(5).Width=2725"
               Splits(0)._ColumnProps(24)=   "Column(5).DividerColor=0"
               Splits(0)._ColumnProps(25)=   "Column(5)._WidthInPix=2646"
               Splits(0)._ColumnProps(26)=   "Column(5).Order=6"
               Splits(0)._ColumnProps(27)=   "Column(6).Width=2725"
               Splits(0)._ColumnProps(28)=   "Column(6).DividerColor=0"
               Splits(0)._ColumnProps(29)=   "Column(6)._WidthInPix=2646"
               Splits(0)._ColumnProps(30)=   "Column(6).Order=7"
               Splits(0)._ColumnProps(31)=   "Column(7).Width=2725"
               Splits(0)._ColumnProps(32)=   "Column(7).DividerColor=0"
               Splits(0)._ColumnProps(33)=   "Column(7)._WidthInPix=2646"
               Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
               Splits(0)._ColumnProps(35)=   "Column(8).Width=2725"
               Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
               Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=2646"
               Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
               Splits(0)._ColumnProps(39)=   "Column(9).Width=2725"
               Splits(0)._ColumnProps(40)=   "Column(9).DividerColor=0"
               Splits(0)._ColumnProps(41)=   "Column(9)._WidthInPix=2646"
               Splits(0)._ColumnProps(42)=   "Column(9).Order=10"
               Splits(0)._ColumnProps(43)=   "Column(10).Width=2725"
               Splits(0)._ColumnProps(44)=   "Column(10).DividerColor=0"
               Splits(0)._ColumnProps(45)=   "Column(10)._WidthInPix=2646"
               Splits(0)._ColumnProps(46)=   "Column(10).Order=11"
               Splits(0)._ColumnProps(47)=   "Column(11).Width=2725"
               Splits(0)._ColumnProps(48)=   "Column(11).DividerColor=0"
               Splits(0)._ColumnProps(49)=   "Column(11)._WidthInPix=2646"
               Splits(0)._ColumnProps(50)=   "Column(11).Order=12"
               Splits(0)._ColumnProps(51)=   "Column(12).Width=2725"
               Splits(0)._ColumnProps(52)=   "Column(12).DividerColor=0"
               Splits(0)._ColumnProps(53)=   "Column(12)._WidthInPix=2646"
               Splits(0)._ColumnProps(54)=   "Column(12).Order=13"
               Splits(0)._ColumnProps(55)=   "Column(13).Width=2725"
               Splits(0)._ColumnProps(56)=   "Column(13).DividerColor=0"
               Splits(0)._ColumnProps(57)=   "Column(13)._WidthInPix=2646"
               Splits(0)._ColumnProps(58)=   "Column(13).Order=14"
               Splits(0)._ColumnProps(59)=   "Column(14).Width=2725"
               Splits(0)._ColumnProps(60)=   "Column(14).DividerColor=0"
               Splits(0)._ColumnProps(61)=   "Column(14)._WidthInPix=2646"
               Splits(0)._ColumnProps(62)=   "Column(14).Order=15"
               Splits(0)._ColumnProps(63)=   "Column(15).Width=2725"
               Splits(0)._ColumnProps(64)=   "Column(15).DividerColor=0"
               Splits(0)._ColumnProps(65)=   "Column(15)._WidthInPix=2646"
               Splits(0)._ColumnProps(66)=   "Column(15).Order=16"
               Splits(0)._ColumnProps(67)=   "Column(16).Width=2725"
               Splits(0)._ColumnProps(68)=   "Column(16).DividerColor=0"
               Splits(0)._ColumnProps(69)=   "Column(16)._WidthInPix=2646"
               Splits(0)._ColumnProps(70)=   "Column(16).Order=17"
               Splits(0)._ColumnProps(71)=   "Column(17).Width=2725"
               Splits(0)._ColumnProps(72)=   "Column(17).DividerColor=0"
               Splits(0)._ColumnProps(73)=   "Column(17)._WidthInPix=2646"
               Splits(0)._ColumnProps(74)=   "Column(17).Order=18"
               Splits(0)._ColumnProps(75)=   "Column(18).Width=2725"
               Splits(0)._ColumnProps(76)=   "Column(18).DividerColor=0"
               Splits(0)._ColumnProps(77)=   "Column(18)._WidthInPix=2646"
               Splits(0)._ColumnProps(78)=   "Column(18).Order=19"
               Splits(0)._ColumnProps(79)=   "Column(19).Width=2725"
               Splits(0)._ColumnProps(80)=   "Column(19).DividerColor=0"
               Splits(0)._ColumnProps(81)=   "Column(19)._WidthInPix=2646"
               Splits(0)._ColumnProps(82)=   "Column(19).Order=20"
               Splits(0)._ColumnProps(83)=   "Column(20).Width=2725"
               Splits(0)._ColumnProps(84)=   "Column(20).DividerColor=0"
               Splits(0)._ColumnProps(85)=   "Column(20)._WidthInPix=2646"
               Splits(0)._ColumnProps(86)=   "Column(20).Order=21"
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
               HeadLines       =   1,5
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
               _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFF&"
               _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H80FFFF&"
               _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
               _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=110,.parent=13"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=107,.parent=14"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=108,.parent=15"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=109,.parent=17"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=106,.parent=13"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=103,.parent=14"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=104,.parent=15"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=105,.parent=17"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=102,.parent=13"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=99,.parent=14"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=100,.parent=15"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=101,.parent=17"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14,.alignment=0"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
               _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=150,.parent=13"
               _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=147,.parent=14"
               _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=148,.parent=15"
               _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=149,.parent=17"
               _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
               _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
               _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
               _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
               _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=154,.parent=13"
               _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=151,.parent=14"
               _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=152,.parent=15"
               _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=153,.parent=17"
               _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
               _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
               _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
               _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
               _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=54,.parent=13"
               _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=51,.parent=14"
               _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=52,.parent=15"
               _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=53,.parent=17"
               _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=58,.parent=13"
               _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
               _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
               _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
               _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=62,.parent=13"
               _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=59,.parent=14"
               _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=60,.parent=15"
               _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=61,.parent=17"
               _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=13"
               _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=29,.parent=14"
               _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=30,.parent=15"
               _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=31,.parent=17"
               _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=46,.parent=13"
               _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=43,.parent=14"
               _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=44,.parent=15"
               _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=45,.parent=17"
               _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=66,.parent=13"
               _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=63,.parent=14"
               _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=64,.parent=15"
               _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=65,.parent=17"
               _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=74,.parent=13"
               _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=71,.parent=14"
               _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=72,.parent=15"
               _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=73,.parent=17"
               _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=78,.parent=13"
               _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=75,.parent=14"
               _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=76,.parent=15"
               _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=77,.parent=17"
               _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=82,.parent=13"
               _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=79,.parent=14"
               _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=80,.parent=15"
               _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=81,.parent=17"
               _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=86,.parent=13"
               _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=83,.parent=14"
               _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=84,.parent=15"
               _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=85,.parent=17"
               _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=90,.parent=13"
               _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=87,.parent=14"
               _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=88,.parent=15"
               _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=89,.parent=17"
               _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=94,.parent=13"
               _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=91,.parent=14"
               _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=92,.parent=15"
               _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=93,.parent=17"
               _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=98,.parent=13"
               _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=95,.parent=14"
               _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=96,.parent=15"
               _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=97,.parent=17"
               _StyleDefs(120) =   "Named:id=33:Normal"
               _StyleDefs(121) =   ":id=33,.parent=0"
               _StyleDefs(122) =   "Named:id=34:Heading"
               _StyleDefs(123) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(124) =   ":id=34,.wraptext=-1"
               _StyleDefs(125) =   "Named:id=35:Footing"
               _StyleDefs(126) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(127) =   "Named:id=36:Selected"
               _StyleDefs(128) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(129) =   "Named:id=37:Caption"
               _StyleDefs(130) =   ":id=37,.parent=34,.alignment=2"
               _StyleDefs(131) =   "Named:id=38:HighlightRow"
               _StyleDefs(132) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
               _StyleDefs(133) =   "Named:id=39:EvenRow"
               _StyleDefs(134) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
               _StyleDefs(135) =   "Named:id=40:OddRow"
               _StyleDefs(136) =   ":id=40,.parent=33"
               _StyleDefs(137) =   "Named:id=41:RecordSelector"
               _StyleDefs(138) =   ":id=41,.parent=34"
               _StyleDefs(139) =   "Named:id=42:FilterBar"
               _StyleDefs(140) =   ":id=42,.parent=33"
            End
            Begin VB.Label LblJumlah 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jumlah Data :"
               Height          =   195
               Left            =   120
               TabIndex        =   19
               Top             =   150
               Width           =   1575
            End
         End
         Begin AIFCmp1.asxPowerBanner asxPowerBanner2 
            Height          =   4845
            Left            =   -120
            Top             =   3000
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   8546
            StartColor      =   -2147483633
            EndColor        =   -2147483633
            FormatString    =   "Data Berdasarkan Searching / Scoring"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TextColor       =   0
         End
         Begin TDBDate6Ctl.TDBDate dtTglLahir 
            Height          =   315
            Left            =   2520
            TabIndex        =   6
            Top             =   1130
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   556
            Calendar        =   "FrmScoringShare.frx":0038
            Caption         =   "FrmScoringShare.frx":0164
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            DropDown        =   "FrmScoringShare.frx":01C0
            Keys            =   "FrmScoringShare.frx":01DE
            Spin            =   "FrmScoringShare.frx":023C
            AlignHorizontal =   0
            AlignVertical   =   0
            Appearance      =   1
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
            Text            =   "11/04/2007"
            ValidateMode    =   0
            ValueVT         =   6815751
            Value           =   39183
            CenturyMode     =   0
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboKecamatan 
            Height          =   285
            Left            =   2520
            TabIndex        =   9
            Tag             =   "Kode Wilayah"
            Top             =   2175
            Width           =   2655
            _Version        =   196616
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelColorHighlight=   -2147483634
            CheckBox3D      =   0   'False
            ForeColorEven   =   0
            BackColorEven   =   8454143
            BackColorOdd    =   65535
            RowHeight       =   423
            Columns(0).Width=   3200
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   4683
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboKelurahan 
            Height          =   285
            Left            =   2520
            TabIndex        =   10
            Tag             =   "Kode Wilayah"
            Top             =   2490
            Width           =   2655
            _Version        =   196616
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelColorHighlight=   -2147483634
            CheckBox3D      =   0   'False
            ForeColorEven   =   0
            BackColorEven   =   8454143
            BackColorOdd    =   65535
            RowHeight       =   423
            Columns(0).Width=   3200
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   4683
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoKecamatan 
            Height          =   345
            Left            =   2520
            Top             =   2175
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
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
            RecordSource    =   "select kodegl,namagl from glmst"
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
         Begin MSAdodcLib.Adodc AdoKelurahan 
            Height          =   345
            Left            =   2520
            Top             =   2490
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
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
            RecordSource    =   "select kodegl,namagl from glmst"
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
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboPropinsi 
            Height          =   285
            Left            =   2520
            TabIndex        =   7
            Tag             =   "Kode Wilayah"
            Top             =   1545
            Width           =   2655
            _Version        =   196616
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelColorHighlight=   -2147483634
            CheckBox3D      =   0   'False
            ForeColorEven   =   0
            BackColorEven   =   8454143
            BackColorOdd    =   65535
            RowHeight       =   423
            Columns(0).Width=   3200
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   4683
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cboKota 
            Height          =   285
            Left            =   2520
            TabIndex        =   8
            Tag             =   "Kode Wilayah"
            Top             =   1860
            Width           =   2655
            _Version        =   196616
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelColorHighlight=   -2147483634
            CheckBox3D      =   0   'False
            ForeColorEven   =   0
            BackColorEven   =   8454143
            BackColorOdd    =   65535
            RowHeight       =   423
            Columns(0).Width=   3200
            Columns(0).DataType=   8
            Columns(0).FieldLen=   4096
            _ExtentX        =   4683
            _ExtentY        =   503
            _StockProps     =   93
            ForeColor       =   -2147483640
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSAdodcLib.Adodc AdoPropinsi 
            Height          =   345
            Left            =   2520
            Top             =   1545
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
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
            RecordSource    =   "select kodegl,namagl from glmst"
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
         Begin MSAdodcLib.Adodc AdoKota 
            Height          =   345
            Left            =   2520
            Top             =   1860
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
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
            RecordSource    =   "select kodegl,namagl from glmst"
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
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Propinsi"
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
            Index           =   17
            Left            =   360
            TabIndex        =   32
            Top             =   1545
            Width           =   795
         End
         Begin VB.Label Label3 
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
            Index           =   16
            Left            =   360
            TabIndex        =   31
            Top             =   1860
            Width           =   435
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Filter :"
            Height          =   195
            Index           =   15
            Left            =   14760
            TabIndex        =   30
            Top             =   0
            Width           =   420
         End
         Begin VB.Label LblFilter 
            Caption         =   "Label1"
            Height          =   3375
            Left            =   15360
            TabIndex        =   29
            Top             =   0
            Width           =   6015
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kecamatan"
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
            Left            =   360
            TabIndex        =   28
            Top             =   2175
            Width           =   1080
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Kelurahan"
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
            Left            =   360
            TabIndex        =   27
            Top             =   2490
            Width           =   990
         End
         Begin VB.Label LblTglLahir 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tanggal Lahir"
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
            Left            =   360
            TabIndex        =   26
            Top             =   1125
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nama Customer"
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
            Left            =   360
            TabIndex        =   25
            Top             =   780
            Width           =   1545
         End
         Begin VB.Label Label6 
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
            Left            =   2415
            TabIndex        =   24
            Top             =   780
            Width           =   60
         End
         Begin VB.Label LblKtp 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. KTP"
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
            Left            =   360
            TabIndex        =   23
            Top             =   120
            Width           =   735
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
            Left            =   2415
            TabIndex        =   22
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
            Index           =   3
            Left            =   2415
            TabIndex        =   21
            Top             =   450
            Width           =   60
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. HP"
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
            Left            =   360
            TabIndex        =   20
            Top             =   450
            Width           =   630
         End
      End
   End
   Begin AIFCmp1.asxLabel LblHeader 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      Caption         =   "Scoring Pencarian Pelanggan"
      BackColor       =   -2147483631
      Alignment       =   2
      BackStyle       =   0
      HotTrackColor   =   255
      MouseIcon       =   "FrmScoringShare.frx":0264
   End
   Begin MSComDlg.CommonDialog ComDialog 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmScoringShare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsNewLgn        As ADODB.Recordset
Dim rsTemp          As ADODB.Recordset
Dim cols            As TrueOleDBGrid70.Columns
Dim sKodeLgn        As String
Dim namaCust        As String

Private Sub SetAdodc()
AdoPropinsi.ConnectionString = adoCon
AdoKota.ConnectionString = adoCon
AdoKecamatan.ConnectionString = adoCon
AdoKelurahan.ConnectionString = adoCon

AdoPropinsi.RecordSource = "SELECT DISTINCT Propinsi FROM TopologyDaerah ORDER BY Propinsi"
    With cboPropinsi
        .DataSourceList = AdoPropinsi
        .DataFieldList = "Propinsi"
    End With
cboPropinsi.Columns(0).Width = cboPropinsi.Width

End Sub

Private Sub cboKota_DropDown()
AdoKota.RecordSource = "SELECT DISTINCT Kota FROM TopologyDaerah where Propinsi = '" & cboPropinsi & "' ORDER BY Kota"
AdoKota.Refresh
    With cboKota
        .DataSourceList = AdoKota
        .DataFieldList = "Kota"
    End With
cboKota.Columns(0).Width = cboKota.Width

End Sub

Private Sub cboKecamatan_DropDown()
AdoKecamatan.RecordSource = "SELECT DISTINCT Kecamatan FROM TopologyDaerah where Propinsi = '" & cboPropinsi & "' AND Kota = '" & cboKota & "' ORDER BY Kecamatan"
AdoKecamatan.Refresh
    With cboKecamatan
        .DataSourceList = AdoKecamatan
        .DataFieldList = "Kecamatan"
    End With
cboKecamatan.Columns(0).Width = cboKecamatan.Width

End Sub

Private Sub cboKelurahan_DropDown()
AdoKelurahan.RecordSource = "SELECT DISTINCT Kelurahan FROM TopologyDaerah where Propinsi = '" & cboPropinsi & "' AND Kota = '" & cboKota & "' AND Kecamatan = '" & cboKecamatan & "' ORDER BY Kelurahan"
AdoKelurahan.Refresh
    With cboKelurahan
        .DataSourceList = AdoKelurahan
        .DataFieldList = "Kelurahan"
    End With
cboKelurahan.Columns(0).Width = cboKelurahan.Width

End Sub

Private Sub CreateTblFind(Optional Edit As Boolean = False)
1     On Error GoTo vb_error
2      Set rsTemp = New ADODB.Recordset
3      With rsTemp
4       .Fields.Append "Score", adDouble
5       .Fields.Append "AddScore", adDouble
6       .Fields.Append "TotalScore", adDouble
7       .Fields.Append "KODELGN", adVarChar, 13
8       .Fields.Append "TglEntry", adDate
9       .Fields.Append "UserEntry", adVarChar, 50
10      .Fields.Append "NoKTP", adVarChar, 30
11      .Fields.Append "NamaCust1", adVarChar, 300
12      .Fields.Append "NamaCust2", adVarChar, 300
13      .Fields.Append "NamaCust3", adVarChar, 300
14      .Fields.Append "TypeCustomer", adVarChar, 50
15      .Fields.Append "HP1", adVarChar, 30
16      .Fields.Append "HP2", adVarChar, 30
17      .Fields.Append "TempatLahir", adVarChar, 100
18      .Fields.Append "TglLahir", adDate
19      .Fields.Append "Alamat", adVarChar, 300
20      .Fields.Append "Propinsi", adVarChar, 100
21      .Fields.Append "Kota", adVarChar, 100
22      .Fields.Append "Kecamatan", adVarChar, 100
23      .Fields.Append "Kelurahan", adVarChar, 100
24      .Fields.Append "KodePos", adVarChar, 30
25      .Open
26      Set GridFind.DataSource = rsTemp
27      End With
28    Exit Sub
vb_error:
29        MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CreateTblFind"), vbExclamation, "Err Number : " & Erl
End Sub

Private Sub CmdAdd_Click()
         Dim vScore     As Double
         Dim vPembagi   As Double
         Dim vNilai     As Double
         Dim sFilter4   As String

1        Set rsFind = New ADODB.Recordset
2        CreateTblFind
         
3        If ProfilScoring.SearchingKTP = True Then
4            sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
5            sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
6            sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER "
7            sSQL = sSQL + "WHERE NoKTP = '" & txtFindKTP & "' "
8            sSQL = sSQL + "AND NoKTP not in (Select Kata from Exeption where Kolom = 'NoKTP') "
9            sSQL = sSQL + "ORDER BY KODELGN"
10           rsFind.Open sSQL, DbCon, adOpenKeyset
11           If rsFind.RecordCount <> 0 Then
12              LblJumlah = "Jumlah Data Dengan Keyword No. KTP: " & FormatNumber(rsFind.RecordCount, 0)
13              If MsgBox("Pencarian dengan No KTP menemukan hasil, pergunakan data ini?", vbYesNo + vbExclamation) = vbNo Then
14                 GoTo CariNoHP
15              Else
16                 rsFind.MoveFirst
17                 Do While Not rsFind.EOF
18                    rsTemp.AddNew
19                    rsTemp!kodeLgn = rsFind!kodeLgn
20                    rsTemp!tglentry = rsFind!tglentry
21                    rsTemp!userentry = Nul(rsFind!userentry)
22                    rsTemp!noKTP = Nul(rsFind!noKTP)
23                    rsTemp!namacust1 = Nul(rsFind!namacust1)
24                    rsTemp!namacust2 = Nul(rsFind!namacust2)
25                    rsTemp!namacust3 = Nul(rsFind!namacust3)
26                    rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
27                    rsTemp!hp1 = Nul(rsFind!hp1)
28                    rsTemp!hp2 = Nul(rsFind!hp2)
29                    rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
30                    rsTemp!tglLahir = rsFind!tglLahir
31                    rsTemp!Alamat = Nul(rsFind!Alamat)
32                    rsTemp!Propinsi = Nul(rsFind!Propinsi)
33                    rsTemp!Kota = Nul(rsFind!Kota)
34                    rsTemp!kecamatan = Nul(rsFind!kecamatan)
35                    rsTemp!kelurahan = Nul(rsFind!kelurahan)
36                    rsTemp!KodePos = Nul(rsFind!KodePos)
37                    rsTemp!Score = 0
38                    rsTemp!addScore = 0
39                    rsTemp!TotalScore = 0

40                    rsFind.MoveNext
41                 Loop
42                 Exit Sub
43              End If
44           End If
45       End If
         
CariNoHP:
46       If ProfilScoring.SearchingHP = True Then
47           Set rsFind = New ADODB.Recordset
48           sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
49           sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
50           sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER "
51           sSQL = sSQL + "WHERE HP1 = '" & TxtFindHP & "' "
52           sSQL = sSQL + "AND HP1 not in (Select Kata from Exeption where Kolom = 'NoHP') "
53           sSQL = sSQL + "ORDER BY KODELGN"
54           rsFind.Open sSQL, DbCon, adOpenKeyset
55           If rsFind.RecordCount <> 0 Then
56              LblJumlah = "Jumlah Data Dengan Keyword No. HP: " & FormatNumber(rsFind.RecordCount, 0)
57              If MsgBox("Pencarian dengan No HP menemukan hasil, pergunakan data ini?", vbYesNo + vbExclamation) = vbNo Then
58                 GoTo CariNama
59              Else
60                 rsFind.MoveFirst
61                 Do While Not rsFind.EOF
62                    rsTemp.AddNew
63                    rsTemp!kodeLgn = rsFind!kodeLgn
64                    rsTemp!tglentry = rsFind!tglentry
65                    rsTemp!userentry = Nul(rsFind!userentry)
66                    rsTemp!noKTP = Nul(rsFind!noKTP)
67                    rsTemp!namacust1 = Nul(rsFind!namacust1)
68                    rsTemp!namacust2 = Nul(rsFind!namacust2)
69                    rsTemp!namacust3 = Nul(rsFind!namacust3)
70                    rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
71                    rsTemp!hp1 = Nul(rsFind!hp1)
72                    rsTemp!hp2 = Nul(rsFind!hp2)
73                    rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
74                    rsTemp!tglLahir = rsFind!tglLahir
75                    rsTemp!Alamat = Nul(rsFind!Alamat)
76                    rsTemp!Propinsi = Nul(rsFind!Propinsi)
77                    rsTemp!Kota = Nul(rsFind!Kota)
78                    rsTemp!kecamatan = Nul(rsFind!kecamatan)
79                    rsTemp!kelurahan = Nul(rsFind!kelurahan)
80                    rsTemp!KodePos = Nul(rsFind!KodePos)
81                    rsTemp!Score = 0
82                    rsTemp!addScore = 0
83                    rsTemp!TotalScore = 0

84                    rsFind.MoveNext
85                 Loop
86                 Exit Sub
87              End If
88           End If
89       End If
         
CariNama:
90       If ProfilScoring.SearchingNama = True Then
91           Set rsFind = New ADODB.Recordset
92           sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
93           sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
94           sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER "
95           sSQL = sSQL + "WHERE NamaCust1 = '" & txtFindNama1 & "' AND NamaCust2 = '" & txtFindNama2 & "' "
96           sSQL = sSQL + "ORDER BY KODELGN"
97           rsFind.Open sSQL, DbCon, adOpenKeyset
98           If rsFind.RecordCount <> 0 Then
99              LblJumlah = "Jumlah Data Dengan Keyword Nama Identik: " & FormatNumber(rsFind.RecordCount, 0)
100             If MsgBox("Pencarian dengan Nama Identik menemukan hasil, pergunakan data ini?", vbYesNo + vbExclamation) = vbNo Then
101                GoTo ScoringNama1Khusus
102             Else
103                rsFind.MoveFirst
104                Do While Not rsFind.EOF
105                   rsTemp.AddNew
106                   rsTemp!kodeLgn = rsFind!kodeLgn
107                   rsTemp!tglentry = rsFind!tglentry
108                   rsTemp!userentry = Nul(rsFind!userentry)
109                   rsTemp!noKTP = Nul(rsFind!noKTP)
110                   rsTemp!namacust1 = Nul(rsFind!namacust1)
111                   rsTemp!namacust2 = Nul(rsFind!namacust2)
112                   rsTemp!namacust3 = Nul(rsFind!namacust3)
113                   rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
114                   rsTemp!hp1 = Nul(rsFind!hp1)
115                   rsTemp!hp2 = Nul(rsFind!hp2)
116                   rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
117                   rsTemp!tglLahir = rsFind!tglLahir
118                   rsTemp!Alamat = Nul(rsFind!Alamat)
119                   rsTemp!Propinsi = Nul(rsFind!Propinsi)
120                   rsTemp!Kota = Nul(rsFind!Kota)
121                   rsTemp!kecamatan = Nul(rsFind!kecamatan)
122                   rsTemp!kelurahan = Nul(rsFind!kelurahan)
123                   rsTemp!KodePos = Nul(rsFind!KodePos)
124                   rsTemp!Score = 0
125                   rsTemp!addScore = 0
126                   rsTemp!TotalScore = 0

127                   rsFind.MoveNext
128                Loop
129                Exit Sub
130             End If
131          End If
132      End If
         
ScoringNama1Khusus:
133      If ProfilScoring.ScoringNamaKhusus = True Then
134         If txtFindNama1 <> "" And txtFindNama2 = "" And txtFindNama3 = "" Then
135             Set rsFind = New ADODB.Recordset
136             sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
137             sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
138             sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER "
139             sSQL = sSQL + "WHERE Left(NamaCust1," & ProfilScoring.LockNama1Khusus & ") = '" & Left(txtFindNama1, ProfilScoring.LockNama1Khusus) & "' "
140             sSQL = sSQL + "ORDER BY KODELGN"
141             rsFind.Open sSQL, DbCon, adOpenKeyset
                
142             DoEvents
143             If rsFind.RecordCount <> 0 Then
144                rsFind.MoveFirst
145                Do While Not rsFind.EOF
146                   vScore = 0
147                   vPembagi = 0
148                   vNilai = 0
149                   rsTemp.AddNew
150                   rsTemp!kodeLgn = rsFind!kodeLgn
151                   rsTemp!tglentry = rsFind!tglentry
152                   rsTemp!userentry = Nul(rsFind!userentry)
153                   rsTemp!noKTP = Nul(rsFind!noKTP)
154                   rsTemp!namacust1 = Nul(rsFind!namacust1)
155                   rsTemp!namacust2 = Nul(rsFind!namacust2)
156                   rsTemp!namacust3 = Nul(rsFind!namacust3)
157                   rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
158                   rsTemp!hp1 = Nul(rsFind!hp1)
159                   rsTemp!hp2 = Nul(rsFind!hp2)
160                   rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
161                   rsTemp!tglLahir = rsFind!tglLahir
162                   rsTemp!Alamat = Nul(rsFind!Alamat)
163                   rsTemp!Propinsi = Nul(rsFind!Propinsi)
164                   rsTemp!Kota = Nul(rsFind!Kota)
165                   rsTemp!kecamatan = Nul(rsFind!kecamatan)
166                   rsTemp!kelurahan = Nul(rsFind!kelurahan)
167                   rsTemp!KodePos = Nul(rsFind!KodePos)

168                   If txtFindNama1 <> "" And Nul(rsFind!namacust1) <> "" Then
169                       vPembagi = (Len(rsFind!namacust1) + Len(txtFindNama1)) / 2
170                       For X = 1 To Len(txtFindNama1)
171                           vNilai = vNilai + IIf(Mid$(LCase(txtFindNama1), X, 1) = Mid$(LCase(rsFind!namacust1), X, 1), 1, 0)
172                       Next X
173                       vScore = vNilai / vPembagi
174                   End If
175                   rsTemp!Score = vScore
176                   rsFind.MoveNext
177                Loop
178             End If
179             rsTemp.filter = "Score > " & k2t(ProfilScoring.nScoringNamaKhusus)
180             rsTemp.Sort = "Score desc"
181             If rsTemp.RecordCount <> 0 Then
182                 LblJumlah = "Jumlah Data Dengan Scoring Nama 1 Khusus: " & FormatNumber(rsTemp.RecordCount, 0)
183                 If MsgBox("Pencarian dengan Nama 1 Khusus menemukan hasil, pergunakan data ini?", vbYesNo + vbExclamation) = vbNo Then
184                    GoTo ScoringNama2
185                 Else
186                    Exit Sub
187                 End If
188             End If
189         End If
190      End If

ScoringNama2:
191     If ProfilScoring.ScoringNamaSimultan = True Then
192             Set rsFind = New ADODB.Recordset
193             sSQL = "SELECT IDCUST AS KODELGN, TglEntry, UserEntry, NoKTP, NamaCust1, NamaCust2, NamaCust3, TypeCustomer, HP1, HP2, TempatLahir, TglLahir, "
194             sSQL = sSQL + "Alamat,Propinsi,Kota,Kecamatan,Kelurahan,KodePos "
195             sSQL = sSQL + "FROM SingleIdOri..NEWCUSTOMER "
196             sSQL = sSQL + "WHERE Left(NamaCust1," & ProfilScoring.LockNama1 & ") = '" & Left(txtFindNama1, ProfilScoring.LockNama1) & "' "
197             sSQL = sSQL + "AND Left(NamaCust2," & ProfilScoring.LockNama2 & ") = '" & Left(txtFindNama2, ProfilScoring.LockNama2) & "' "
198             sSQL = sSQL + "ORDER BY KODELGN"
199             rsFind.Open sSQL, DbCon, adOpenKeyset
                
200             DoEvents
201             If rsFind.RecordCount <> 0 Then
202                rsFind.MoveFirst
203                Do While Not rsFind.EOF
204                   vScore = 0
205                   vPembagi = 0
206                   vNilai = 0
207                   rsTemp.AddNew
208                   rsTemp!kodeLgn = rsFind!kodeLgn
209                   rsTemp!tglentry = rsFind!tglentry
210                   rsTemp!userentry = Nul(rsFind!userentry)
211                   rsTemp!noKTP = Nul(rsFind!noKTP)
212                   rsTemp!namacust1 = Nul(rsFind!namacust1)
213                   rsTemp!namacust2 = Nul(rsFind!namacust2)
214                   rsTemp!namacust3 = Nul(rsFind!namacust3)
215                   rsTemp!TypeCustomer = Nul(rsFind!TypeCustomer)
216                   rsTemp!hp1 = Nul(rsFind!hp1)
217                   rsTemp!hp2 = Nul(rsFind!hp2)
218                   rsTemp!TempatLahir = Nul(rsFind!TempatLahir)
219                   rsTemp!tglLahir = rsFind!tglLahir
220                   rsTemp!Alamat = Nul(rsFind!Alamat)
221                   rsTemp!Propinsi = Nul(rsFind!Propinsi)
222                   rsTemp!Kota = Nul(rsFind!Kota)
223                   rsTemp!kecamatan = Nul(rsFind!kecamatan)
224                   rsTemp!kelurahan = Nul(rsFind!kelurahan)
225                   rsTemp!KodePos = Nul(rsFind!KodePos)

226                   If txtFindNama2 <> "" And Nul(rsFind!namacust2) <> "" Then
227                       vPembagi = (Len(rsFind!namacust2) + Len(txtFindNama2)) / 2
228                       For X = 1 To Len(txtFindNama2)
229                           vNilai = vNilai + IIf(Mid$(LCase(txtFindNama2), X, 1) = Mid$(LCase(rsFind!namacust2), X, 1), 1, 0)
230                       Next X
231                       vScore = vNilai / vPembagi
232                   End If
233                   rsTemp!Score = vScore
234                   rsTemp!addScore = 0
235                   rsTemp!TotalScore = 0
236                   If (ProfilScoring.ScoringNoHP = True) And (Nul(Nul(rsFind!noKTP)) = TxtFindHP) Then rsTemp!addScore = rsTemp!addScore + 0.25
237                   If (ProfilScoring.ScoringTglLahir = True) And (rsFind!tglLahir = dtTglLahir) Then rsTemp!addScore = rsTemp!addScore + 0.25
238                   If (ProfilScoring.ScoringKelurahan = True) And (rsFind!kelurahan = cboKelurahan) Then rsTemp!addScore = rsTemp!addScore + 0.25
239                   If (ProfilScoring.ScoringKecamatan = True) And (rsFind!kecamatan = cboKecamatan) Then rsTemp!addScore = rsTemp!addScore + 0.25
240                   rsTemp!TotalScore = rsTemp!Score + rsTemp!addScore
241                   rsFind.MoveNext
242                Loop
243             End If
244             rsTemp.filter = "Score > " & k2t(ProfilScoring.nScoringNama2)
245             rsTemp.Sort = "TotalScore desc, Score desc, AddScore desc"
246             LblJumlah = "Jumlah Data Dengan Scoring Nama Simultan: " & FormatNumber(rsTemp.RecordCount, 0)
247             If rsTemp.RecordCount <> 0 Then
248                 MsgBox "Filter dengan batas Nama 2 " & rsTemp.filter & " menghasilkan " & rsTemp.RecordCount & " data", vbInformation, "Scoring Nama 2"
249             Else
250                 MsgBox "Filter dengan batas Nama 2 " & rsTemp.filter & " tidak menghasilkan data!", vbExclamation, "Scoring Nama 2"
251             End If
252     End If

End Sub

Private Sub Form_Load()
   Set frmPanggil = Me
   SetColorProfile Me
   CmdQuit.Left = (Frame3.Left + Frame3.Width) - CmdQuit.Width
SetAdodc

With ProfilScoring
'Searching Identik:
    .SearchingKTP = True
    .SearchingHP = True
    .SearchingNama = True
'Scoring Nama 1 Khusus:
    .ScoringNamaKhusus = True
    .nScoringNamaKhusus = 0.3
    .LockNama1Khusus = 2
'Scoring Nama Simultan:
    .ScoringNamaSimultan = True
    .LockNama1 = 2
    .LockNama2 = 2
    .nScoringNama2 = 0.5
'Kumulatif:
    .ScoringNoHP = True
    .ScoringTglLahir = True
    .ScoringKelurahan = True
    .ScoringKecamatan = True

    LblFilter = "Searching KTP = " & .SearchingKTP
    LblFilter = LblFilter & vbCrLf & "Searching HP = " & .SearchingHP
    LblFilter = LblFilter & vbCrLf & "Searching Nama = " & .SearchingNama
    LblFilter = LblFilter & vbCrLf & "Scoring Nama Khusus = " & .ScoringNamaKhusus
    LblFilter = LblFilter & vbCrLf & "N Scoring Nama Khusus = " & .nScoringNamaKhusus
    LblFilter = LblFilter & vbCrLf & "Lock Nama 1 Khusus = " & .LockNama1Khusus
    LblFilter = LblFilter & vbCrLf & "Scoring Nama Simultan = " & .ScoringNamaSimultan
    LblFilter = LblFilter & vbCrLf & "Lock Nama 1 = " & .LockNama1
    LblFilter = LblFilter & vbCrLf & "Lock Nama 2 = " & .LockNama2
    LblFilter = LblFilter & vbCrLf & "N Scoring Nama 2 = " & .nScoringNama2
    LblFilter = LblFilter & vbCrLf & "Scoring No HP = " & .ScoringNoHP
    LblFilter = LblFilter & vbCrLf & "Scoring Tgl Lahir = " & .ScoringTglLahir
    LblFilter = LblFilter & vbCrLf & "Scoring Kelurahan = " & .ScoringKelurahan
    LblFilter = LblFilter & vbCrLf & "Scoring Kecamatan = " & .ScoringKecamatan
End With

End Sub

Private Sub Form_Resize()
        Frame1.Top = FrameTombol.Top + FrameTombol.Height
'        Frame1.Height = Banner.Height - 455
        Frame1.Left = (Me.Width - Frame1.Width) / 2
        FrameTombol.Left = Frame1.Left
        SScroll VScroll1, Me
        VScroll1.Visible = True
        
End Sub

Private Sub CmdQuit_Click()
   Unload Me
End Sub

Private Sub CmdQuit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Objek_MouseMove CmdQuit, X, Y
End Sub

Private Sub CmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Objek_MouseMove cmdAdd, X, Y
End Sub

Private Sub Form_Unload(Cancel As Integer)
   HapusTab Me
End Sub

Private Sub GridFind_FilterChange()
      Dim C As Integer
1     On Error GoTo GridFind_Filter
2     Set cols = GridFind.Columns
3     C = GridFind.col
4     GridFind.HoldFields
5     rsTemp.filter = getFilter()
6       If rsTemp.EOF Then
7          MsgBox "Data yang dimaksud tidak dapat ditemukan !", vbInformation
8          For Each col In GridFind.Columns
9             col.FilterText = ""
10         Next col
11         rsTemp.filter = ""
12         Set GridFind.DataSource = rsTemp
13         GridFind.ReOpen
14         GridFind.col = C
15         LblJumlah = "Jumlah Data: " & FormatNumber(rsTemp.RecordCount, 0)
16         Exit Sub
17      End If

18    GridFind.col = C
19    GridFind.EditActive = True
20    LblJumlah = "Jumlah Data: " & FormatNumber(rsTemp.RecordCount, 0)
21    Exit Sub

GridFind_Filter:
22    If err.Number = -2147217825 Then
23      GridFind.Columns(C).FilterText = ""
24      Exit Sub
25    End If
26    MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".GridFind_FilterChange "), vbCritical

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
            tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
        End If
    Next col
    getFilter = tmp

End Function

Private Sub GridFind_HeadClick(ByVal ColIndex As Integer)
If rsTemp.Sort = GridFind.Columns(ColIndex).DataField Then
   rsTemp.Sort = GridFind.Columns(ColIndex).DataField & " Desc"

Else
   rsTemp.Sort = GridFind.Columns(ColIndex).DataField
End If
End Sub

Private Sub VScroll1_Change()
Frame1.Top = -VScroll1.value
If VScroll1.value <= 5 Then Frame1.Top = FrameTombol.Height
If VScroll1.value = VScroll1.Max Then Frame1.Top = -(VScroll1.value + (FrameTombol.Height / 4))
End Sub

Private Sub VScroll1_Scroll()
Frame1.Top = -VScroll1.value
If VScroll1.value <= 5 Then Frame1.Top = FrameTombol.Height
If VScroll1.value = VScroll1.Max Then Frame1.Top = -(VScroll1.value + (FrameTombol.Height / 4))
End Sub

