VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmOrderCuci 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Menu Cuci"
   ClientHeight    =   7620
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12765
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7620
   ScaleWidth      =   12765
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Finish Pengering"
      Height          =   1695
      Left            =   7320
      TabIndex        =   28
      Top             =   -240
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdFinishKePengering 
         BackColor       =   &H0080FF80&
         Caption         =   "PROSES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   960
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   30
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   1800
         TabIndex        =   31
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Posisi Kendaraan Awal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Wo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   32
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   1695
      Left            =   3360
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   5415
      Begin VB.CommandButton cmdFinishAkhir 
         BackColor       =   &H0080FF80&
         Caption         =   "PROSES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   960
         Width           =   2055
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   24
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Index           =   4
         Left            =   1800
         TabIndex        =   25
         Top             =   600
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "No Wo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   27
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Posisi Kendaraan Akhir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame FramePilihPengering 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pilih Pengering"
      Height          =   3015
      Left            =   2040
      TabIndex        =   17
      Top             =   2280
      Visible         =   0   'False
      Width           =   4215
      Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
         Height          =   2055
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3625
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "No Stall"
         Columns(0).DataField=   "NoStall"
         Columns(0).DropDown=   "TDBDropDown1"
         Columns(0).DropDown.vt=   8
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama Stall"
         Columns(1).DataField=   "NamaStall"
         Columns(1).DropDown=   "TDBDropDown1"
         Columns(1).DropDown.vt=   8
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Lead Time"
         Columns(2).DataField=   "LTCuci"
         Columns(2).DropDown=   "TDBDropDown1"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1773"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1693"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=3519"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3440"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=873"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=794"
         Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=27,.parent=17"
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
      Begin AIFCmp1.asxPowerButton btnPengeringPilih 
         Height          =   495
         Left            =   2760
         TabIndex        =   19
         Top             =   2400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         BackColor       =   33023
         Caption         =   "PILIH"
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
   End
   Begin VB.Timer Timer2 
      Interval        =   5000
      Left            =   9120
      Top             =   3360
   End
   Begin AIFCmp1.asxPowerButton cmdStart 
      Height          =   495
      Left            =   9720
      TabIndex        =   9
      Top             =   2760
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      BackColor       =   33023
      Caption         =   "START"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   7920
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
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
   Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbStallCuci 
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Top             =   600
      Width           =   2535
      _Version        =   196616
      Columns(0).Width=   3200
      _ExtentX        =   4471
      _ExtentY        =   661
      _StockProps     =   93
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8640
      Top             =   3360
   End
   Begin TrueOleDBGrid70.TDBGrid ListMech 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   2355
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
   Begin TrueOleDBGrid70.TDBGrid listPengering 
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   3413
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
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9).DropDown=   "TDBDropDown1"
      Columns(9).DropDown.vt=   8
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
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
      Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
      Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
      Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=70,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=74,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=71,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=72,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=73,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=32,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=29,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=30,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=31,.parent=17"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=39:EvenRow"
      _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=40:OddRow"
      _StyleDefs(92)  =   ":id=40,.parent=33"
      _StyleDefs(93)  =   "Named:id=41:RecordSelector"
      _StyleDefs(94)  =   ":id=41,.parent=34"
      _StyleDefs(95)  =   "Named:id=42:FilterBar"
      _StyleDefs(96)  =   ":id=42,.parent=33"
   End
   Begin ucXPButton.XPButton cmdKeluar 
      Height          =   495
      Left            =   11400
      TabIndex        =   7
      Top             =   6120
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
      MICON           =   "frmOrderCuci.frx":0000
      PICN            =   "frmOrderCuci.frx":001C
      PICH            =   "frmOrderCuci.frx":04F0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin AIFCmp1.asxPowerButton cmdFinish 
      Height          =   495
      Left            =   11280
      TabIndex        =   10
      Top             =   2760
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      BackColor       =   33023
      Caption         =   "FINISH"
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
   Begin VB.TextBox txtScanBarcode 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "TAMPILKAN SEMUA ANTRIAN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      TabIndex        =   20
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lblScanBarcode 
      Alignment       =   2  'Center
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
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   6360
      Width           =   4455
   End
   Begin VB.Label lblMenitCount 
      BackColor       =   &H00FFFFFF&
      Caption         =   "-99:99"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   600
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label lblLate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   840
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblOT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblNoPolisi 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AD-9999-YES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   3600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SEDANG CUCI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   1000
      Width           =   1815
   End
   Begin VB.Label lblJam 
      BackColor       =   &H00FFFFFF&
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
      Left            =   10800
      TabIndex        =   8
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblHari 
      BackColor       =   &H00FFFFFF&
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
      Left            =   9480
      TabIndex        =   6
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stall Cuci"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00FFFFFF&
      X1              =   12720
      X2              =   12120
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00FFFFFF&
      X1              =   12720
      X2              =   11880
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00FFFFFF&
      X1              =   12720
      X2              =   11760
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
      X1              =   12720
      X2              =   120
      Y1              =   6720
      Y2              =   6720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000080FF&
      X1              =   12720
      X2              =   120
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label hariNow 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Cuci"
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
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   10575
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
      Width           =   12675
   End
End
Attribute VB_Name = "frmOrderCuci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsTemp As ADODB.Recordset
Dim rsTempSimpan As ADODB.Recordset

Dim rsStallCuci As ADODB.Recordset
Dim rsTemStallCuci As New ADODB.Recordset

Dim rsAntriCuci As ADODB.Recordset
Dim rsTemAntriCuci As New ADODB.Recordset

Dim sMulai As String
Dim sSelesai As String
Dim batasMulai As String
Dim batasSelesai As String
Dim nHari As String
Dim jmlStall As Integer

Private jumlahStallCuci As Integer
Private jumlahStallPengering As Integer

Private noOrder As String

Private selectStallCuci As String
Private namaStallCuci As String

Private nopolAntrianAwal As String
Private nowoAntrianAwal As String
Private rsStallPengering As ADODB.Recordset
Dim stallPengering As String

Private Function convertMenit(menit As Double) As String

          Dim jam As Double
          Dim menitM As Double
1        On Error GoTo convertMenit_Error

2         If menit > 60 Then
3             jam = CInt(Val(menit / 60))
4             menitM = CInt(menit Mod 60)
              
5             convertMenit = Format(jam, "0#") & ":" & Format(menitM, "0#")
6         Else
7             convertMenit = "00:" & Format(menit, "0#")
8         End If

9        On Error GoTo 0
10       Exit Function

convertMenit_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure convertMenit of Form frmOrderCuci")
End Function

Private Sub refreshGridAntriCuci(isSemua As Boolean)
          ' LIHAT LAGI
          
1        On Error GoTo refreshGridAntriCuci_Error
          Dim rumus As String
2         If Log.prioritas <> "" Then
3             rumus = objO.FetchData_String("SELECT RUMUSAN FROM prioritasCuci WHERE PRIORITAS = '" & Log.prioritas & "'")
4         Else
              'rumus = "dateserah + timeserah DESC"
              ' revisi 30 jan 2017, harusnya paling atas yg janji serah terima paling cepat
5             rumus = "dateserah + timeserah"
6         End If
          'If rumus = "" Then rumus = "dateserah + timeserah DESC"
          ' revisi 30 jan 2017, harusnya paling atas yg janji serah terima paling cepat
7         If rumus = "" Then rumus = "dateserah + timeserah"
8         If isSemua Then
9             Set rsAntriCuci = objO.FetchData("SELECT pds,j.[code],j.kd_k,j.sbe,KLPKERJA1,KLPKERJA2,KLPKERJA3,KLPKERJA4,KLPKERJA5,KLPKERJA6,KLPKERJA7,j.pds,J.NO_POLISI AS NOPOLISI,A.*,J.FRONT,L.dateserah + l.timeserah as TIMESERAH,DATEDIFF(MINUTE, getdate(),dateserah + timeserah) as sisa,CONVERT(varchar(20),a.tanggal,108) as waktuAntri  FROM AntriCuci A WITH (NOLOCK) INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = A.NOORDER INNER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_WO = A.NOORDER WHERE tipestall = 'Cuci' and A.FLAG = '' ORDER BY " & rumus)
10        Else
11            Set rsAntriCuci = objO.FetchData("SELECT pds,j.[code],j.kd_k,j.sbe,KLPKERJA1,KLPKERJA2,KLPKERJA3,KLPKERJA4,KLPKERJA5,KLPKERJA6,KLPKERJA7,j.pds,J.NO_POLISI AS NOPOLISI,A.*,J.FRONT,L.dateserah + l.timeserah as TIMESERAH,DATEDIFF(MINUTE, getdate(),dateserah + timeserah) as sisa,CONVERT(varchar(20),a.tanggal,108) as waktuAntri  FROM AntriCuci A WITH (NOLOCK) INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = A.NOORDER INNER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_WO = A.NOORDER WHERE noStall = '" & selectStallCuci & "' AND A.FLAG = '' ORDER BY " & rumus)
12        End If
          Dim sisaWaktu As String
          Dim tipe As String
13        If createAntriCuci Then
14            If rsAntriCuci.RecordCount > 0 Then
15                rsAntriCuci.MoveFirst
16                nopolAntrianAwal = Trim(rsAntriCuci("NoPolisi"))
17                nowoAntrianAwal = Trim(rsAntriCuci("NoOrder"))
                  
18                Do While Not rsAntriCuci.EOF
19                    If Trim(rsAntriCuci("SISA")) >= 0 Then
20                        sisaWaktu = convertMenit(CDbl(Trim(rsAntriCuci("SISA"))))
21                    Else
22                        sisaWaktu = "-" & convertMenit(CDbl(Mid(Trim(rsAntriCuci("SISA")), 2, Len(Trim(rsAntriCuci("SISA")) - 1))))
23                    End If
24                    If CInt(rsAntriCuci("CODE")) = 1 Then
                          'em
25                        tipe = "EM"
26                    Else
                          'np
27                        If CInt(rsAntriCuci("KD_K")) = 1 Then
28                            tipe = "NP"
29                        Else
                              'sbi
30                            If rsAntriCuci("KLPKERJA1") = "SBI" Or rsAntriCuci("KLPKERJA2") = "SBI" Or rsAntriCuci("KLPKERJA3") = "SBI" Or rsAntriCuci("KLPKERJA4") = "SBI" Or rsAntriCuci("KLPKERJA5") = "SBI" Or rsAntriCuci("KLPKERJA6") = "SBI" Or rsAntriCuci("KLPKERJA7") = "SBI" Then
31                                tipe = "SBI"
32                            Else
                                  'pds
33                                If rsAntriCuci("PDS") = 1 Then
34                                    tipe = "PDS"
35                                Else
                                      'sbe
36                                    If rsAntriCuci("KLPKERJA1") = "SBE" Or rsAntriCuci("KLPKERJA2") = "SBE" Or rsAntriCuci("KLPKERJA3") = "SBE" Or rsAntriCuci("KLPKERJA4") = "SBE" Or rsAntriCuci("KLPKERJA5") = "SBE" Or rsAntriCuci("KLPKERJA6") = "SBE" Or rsAntriCuci("KLPKERJA7") = "SBE" Then
37                                        tipe = "SBE"
38                                    Else
                                          'gr
39                                        tipe = "GRP"
40                                    End If
41                                End If
42                            End If
43                        End If
44                    End If
45                    With rsTemAntriCuci
46                        .AddNew
47                        !NoPolisi = Trim(rsAntriCuci("NoPolisi"))
48                        !sa = Trim(rsAntriCuci("FRONT"))
49                        !noWO = Trim(rsAntriCuci("noorder"))
50                        !OTD = Trim(rsAntriCuci("TIMESERAH"))
51                        !tipe = tipe
52                        !sisaWaktu = sisaWaktu
53                        !waktu = Trim(rsAntriCuci!waktuAntri)
54                        !Mech = Trim(rsAntriCuci("MECH"))
55                        !posisiAwal = Trim(rsAntriCuci!posisiAwal)
56                        !posisiAkhir = Trim(rsAntriCuci!posisiAkhir)
57                        .Update
58                        .MoveNext
59                    End With
60                    rsAntriCuci.MoveNext
61                    j = j + 1
62                Loop
63            End If
64        End If
65        listPengering.DataSource = rsTemAntriCuci
66        Call headColumnAntriCuci
67        If rsTemAntriCuci.RecordCount > 0 Then
68            cmdStart.Enabled = True
69        Else
70            cmdStart.Enabled = False
              
71        End If

72       On Error GoTo 0
73       Exit Sub

refreshGridAntriCuci_Error:

74        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGridAntriCuci of Form frmOrderCuci")
End Sub

Private Function createAntriCuci() As Boolean
1         On Error GoTo vb_error
          
2         If rsTemAntriCuci.State Then rsTemAntriCuci.Close
3         Set rsTemAntriCuci = Nothing
4         With rsTemAntriCuci
5             .Fields.Append "NoPolisi", adVarChar, 25
6             .Fields.Append "SA", adVarChar, 5
7             .Fields.Append "NoWo", adVarChar, 15
8             .Fields.Append "OTD", adVarChar, 25
9             .Fields.Append "Waktu", adVarChar, 25
10            .Fields.Append "Tipe", adVarChar, 25
11            .Fields.Append "SisaWaktu", adVarChar, 25
12            .Fields.Append "MECH", adVarChar, 5
13            .Fields.Append "posisiAwal", adVarChar, 40
14            .Fields.Append "posisiAkhir", adVarChar, 40
15            .Open
16        End With
17        If rsTemAntriCuci.State Then createAntriCuci = True
18        Exit Function
          
vb_error:
19        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Function createStallCuci() As Boolean
           
1        On Error GoTo createStallCuci_Error

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
13            .Open
14        End With
15        If rsTemStallCuci.State Then createStallCuci = True
16        Exit Function

17       On Error GoTo 0
18       Exit Function

createStallCuci_Error:

19        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure createStallCuci of Form frmOrderCuci")
End Function

Private Function hitungEstWaktu(mulai As String, selesai As String) As String
          Dim jMulai As Integer
          Dim mMulai As Integer
          Dim jSelesai As Integer
          Dim mSelesai As Integer
          Dim total As Integer
          
1        On Error GoTo hitungEstWaktu_Error
2         jMulai = CInt(Left(mulai, 2))
3         mMulai = CInt(Mid(mulai, 4, 2))
4         jSelesai = CInt(Left(selesai, 2))
5         mSelesai = CInt(Mid(selesai, 4, 2))
6         total = (jMulai * 60) + (mMulai) + (jSelesai * 60) + (mSelesai)
7         If total > 60 Then
8             jSelesai = CInt(Val(total / 60))
9             mSelesai = CInt(total Mod 60)
10            hitungEstWaktu = Format(jSelesai, "0#") & ":" & Format(mSelesai, "0#")
11        Else
12            hitungEstWaktu = "00:" & Format(Abs(CInt(total Mod 60)), "0#")
13        End If
14       On Error GoTo 0
15       Exit Function
hitungEstWaktu_Error:
16        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hitungEstWaktu of Form frmOrderCuci")
End Function

Private Function hitungWaktu(waktuCuci As String, menitProses As Integer) As String
          Dim jamW As Integer
          Dim menitW As Integer
          Dim totalWaktu As Double
1        On Error GoTo hitungWaktu_Error

2         jamW = Left(waktuCuci, 2)
3         menitW = Mid(waktuCuci, 4, 2)
          
4         totalWaktu = (jamW * 60) + menitW
5         totalWaktu = totalWaktu - menitProses
6         If totalWaktu > 60 Then
7             jamW = CInt(Val(totalWaktu / 60))
8             menitW = CInt(totalWaktu Mod 60)
              
9             hitungWaktu = Format(jamW, "0#") & ":" & Format(menitW, "0#")
10        Else
              
11            hitungWaktu = "00:" & Format(Abs(CInt(totalWaktu Mod 60)), "0#")
12        End If

13       On Error GoTo 0
14       Exit Function

hitungWaktu_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hitungWaktu of Form frmOrderCuci")
End Function

Private Sub refreshGrid()
          ' LIHAT LAGI
          Dim waktuCuci As String
          Dim waktu As String
          Dim waktuSekarang As String
          Dim estWaktu As String
          Dim selisihWaktu As Double
          
          'waktuCuci = "00:30"
1        On Error GoTo refreshGrid_Error

2         cmdStart.Enabled = True
3         cmdFinish.Enabled = False
          
4         noOrder = ""
5         waktu = ""
6         Set rsStallCuci = objO.FetchData("SELECT datediff(minute,CONVERT(VARCHAR(5),getdate(),108),l.timeserah) as selisih,CONVERT(VARCHAR(5),getdate(),108) as waktuSekarang,datediff(minute,start,CONVERT(VARCHAR,getdate(),108)) as proses,T.*,J.NO_POLISI,J.FRONT,L.TIMESERAH,Q.LTCUCI FROM TempProgressCuci t WITH (NOLOCK) INNER JOIN QTY_STALLWASHING Q WITH (NOLOCK) ON Q.NOSTALL = T.NOSTALL AND ISDEL = '0' AND AKTIF = '1' INNER JOIN JOBDTA J WITH (NOLOCK) ON T.NOORDER = J.NO_WO INNER JOIN JOBDTL L WITH (NOLOCK) ON L.NO_WO = T.NOORDER WHERE T.noStall = '" & selectStallCuci & "' and FLAG = ''")
          Dim nopol As String
7         If createStallCuci Then
8             If rsStallCuci.RecordCount > 0 Then
9                 cmdStart.Enabled = False
10                cmdFinish.Enabled = True
11                rsStallCuci.MoveFirst
12                Do While Not rsStallCuci.EOF
13                    waktuCuci = Trim(rsStallCuci("LTCUCI"))
14                    nopol = Trim(rsStallCuci("NO_POLISI"))
15                    waktu = hitungWaktu(waktuCuci, rsStallCuci("proses"))
16                    waktuSekarang = rsStallCuci("waktuSekarang")
17                    estWaktu = hitungEstWaktu(rsStallCuci("START"), waktuCuci)
18                    selisihWaktu = rsStallCuci("selisih")
19                    noOrder = Trim(rsStallCuci("NOORDER"))
20                    With rsTemStallCuci
21                        .AddNew
22                        !NoPolisi = Trim(rsStallCuci("NO_POLISI"))
23                        !sa = Trim(rsStallCuci("FRONT"))
24                        !noWO = Trim(rsStallCuci("NOORDER"))
25                        !OTD = Trim(rsStallCuci("TIMESERAH"))
26                        !MulaiCuci = Trim(rsStallCuci("START"))
27                        !waktu = waktu
28                        !EstFinish = estWaktu
29                        !Mech = Trim(rsStallCuci("MECH"))
30                        .Update
31                        .MoveNext
32                    End With
33                    rsStallCuci.MoveNext
34                    j = j + 1
35                Loop
36            End If
37        End If
38        ListMech.DataSource = rsTemStallCuci
39        Call headColumns
40        If nopol <> "" Then
41            lblNoPolisi.Visible = True
42            lblNoPolisi.Caption = nopol
43        Else
44            lblNoPolisi.Visible = False
45        End If
          
46        If waktu <> "" Then
47            lblMenitCount.Visible = True
48            lblMenitCount.Caption = "-" & waktu
49        Else
50            lblMenitCount.Visible = False
51        End If
          
52        If waktu <> "" Then
53            If selisihWaktu >= 0 Then
54                lblLate.Visible = False '
55                lblOT.Visible = True
56            Else
57                lblOT.Visible = False
58                lblLate.Visible = True
59            End If
60        Else
61            lblLate.Visible = False
62            lblOT.Visible = False
63        End If

64       On Error GoTo 0
65       Exit Sub

refreshGrid_Error:

66        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form frmOrderCuci")
End Sub

Private Sub headColumnAntriCuci()
1        On Error GoTo headColumnAntriCuci_Error

2         With listPengering
3             .Columns(0).Caption = "No Polisi"
4             .Columns(0).DataField = "NoPolisi"
5             .Columns(0).Width = 1000
6             .Columns(1).Caption = "SA"
7             .Columns(1).DataField = "SA"
8             .Columns(1).Width = 700
9             .Columns(2).Caption = "No Wo"
10            .Columns(2).DataField = "NoWo"
11            .Columns(2).Width = 2000
12            .Columns(3).Caption = "OTD"
13            .Columns(3).DataField = "OTD"
14            .Columns(3).Width = 1000
15            .Columns(4).Caption = "Waktu Antri"
16            .Columns(4).DataField = "waktu"
17            .Columns(4).Width = 1500
18            .Columns(5).Caption = "R / Type"
19            .Columns(5).DataField = "tipe"
20            .Columns(5).Width = 1500
21            .Columns(6).Caption = "Sisa Waktu"
22            .Columns(6).DataField = "SisaWaktu"
23            .Columns(6).Width = 1000
24            .Columns(7).Caption = "MECH"
25            .Columns(7).DataField = "MECH"
26            .Columns(7).Width = 1000
27            .Columns(7).Visible = False
28            .Columns(8).Caption = "Posisi Awal"
29            .Columns(8).DataField = "posisiAwal"
30            .Columns(8).Width = 2000
31            .Columns(9).Caption = "Posisi Akhir"
32            .Columns(9).DataField = "posisiAkhir"
33            .Columns(9).Width = 2000
34        End With

35       On Error GoTo 0
36       Exit Sub

headColumnAntriCuci_Error:

37        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumnAntriCuci of Form frmOrderCuci")
End Sub

Private Sub headColumns()
1        On Error GoTo headColumns_Error

2         With ListMech
3             .Columns(0).Caption = "No Polisi"
4             .Columns(0).DataField = "NoPolisi"
5             .Columns(0).Width = 1000
6             .Columns(1).Caption = "SA"
7             .Columns(1).DataField = "SA"
8             .Columns(1).Width = 700
9             .Columns(2).Caption = "No Wo"
10            .Columns(2).DataField = "NoWo"
11            .Columns(2).Width = 1000
12            .Columns(3).Caption = "OTD"
13            .Columns(3).DataField = "OTD"
14            .Columns(3).Width = 1000
15            .Columns(4).Caption = "Mulai Cuci"
16            .Columns(4).DataField = "MulaiCuci"
17            .Columns(4).Width = 1000
18            .Columns(5).Caption = "Waktu"
19            .Columns(5).DataField = "Waktu"
20            .Columns(5).Width = 1000
21            .Columns(6).Caption = "Est Selesai Cuci"
22            .Columns(6).DataField = "EstFinish"
23            .Columns(6).Width = 1000
24        End With
          

25       On Error GoTo 0
26       Exit Sub

headColumns_Error:

27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumns of Form frmOrderCuci")
          
End Sub

Private Sub btnPengeringPilih_Click()
1        If Log.FIBeforeSublet = 1 Then
2           stallPengering = TDBGrid1.Columns(0).value
3           Text(1) = noOrder
4           Frame2.Visible = True
5           Exit Sub
6           FramePilihPengering.Visible = False
7        Else
8         Call objHT.OpenTransactionBengkel
9         On Error GoTo errorTrans
10        erTranAll = False
          
11        Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'D' WHERE KETERANGAN LIKE '%" & noOrder & "%'")
12        Call objHT.InsertData_Bengkel("UPDATE TempProgressCuci SET FLAG = 'D',FINISH = CONVERT(VARCHAR,getdate(),108) WHERE NOSTALL = '" & selectStallCuci & "' AND  NOORDER = '" & noOrder & "'")
          
13        Call objHT.InsertData_Bengkel("INSERT INTO AntriCuci (NoOrder, Tanggal, NoStall, TipeStall, Waktu, Mech, flag) VALUES ('" & ListMech.Columns(2).value & "',GETDATE(),'" & TDBGrid1.Columns(0).value & "','Pengering',CONVERT(VARCHAR,getdate(),108),'" & rsTemStallCuci("mech") & "','')")
14        If erTranAll = False Then
15            Call objHT.CloseTransactionBengkel
16            MsgBox "Proses Finish Cuci Stall Berhasil !", vbInformation + vbOKOnly
17            FramePilihPengering.Visible = False
18            txtScanBarcode.Text = ""
19            txtScanBarcode.SetFocus
20            lblScanBarcode.Caption = ""
                                      
21            Call Timer2_Timer
22            Exit Sub
23        Else
24            Call objHT.CloseTransactionBengkel_err
25            MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
26            Exit Sub
27        End If
28        Call Timer2_Timer
29        Exit Sub
30       End If
errorTrans:
31        Call objHT.CloseTransactionBengkel_err
32        MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
33        Exit Sub
End Sub

Private Sub cbStallCuci_Change()
    If txtScanBarcode.Text = "" Then
    selectStallCuci = cbStallCuci.Columns(0).value
    refreshGrid
    Call refreshGridAntriCuci(False)
    End If
End Sub


Private Sub cbStallCuci_Click()
    If cbStallCuci <> "" Then
        Call cbStallCuci_Change
    End If
End Sub

Private Sub Check1_Click()
    If selectStallCuci <> "" Then
        If Check1.value = 1 Then
            'menampilkan semua antrian
            Call refreshGridAntriCuci(True)
        Else
            Call refreshGridAntriCuci(False)
        End If
    End If
End Sub

Private Sub cmdFinish_Click()
1         If noOrder = "" Then
2             MsgBox "Tidak ada No Order yang diproses", vbCritical
3             Exit Sub
4         End If
5         stallPengering = ""
6         Set rsTemp = objO.FetchData("SELECT * FROM QTY_StallWashing WITH (NOLOCK) WHERE TIPESTALL = 'PENGERING' AND ISDEL = '0'  and aktif = '1'")
7         If rsTemp.RecordCount > 0 Then
              'ada proses pengeringan
8             If rsTemp.RecordCount > 1 Then
9                 FramePilihPengering.Visible = True
10                Set rsStallPengering = objO.FetchData("SELECT * FROM QTY_StallWashing WHERE TipeStall <> 'Cuci'  AND ISDEL = '0'  and aktif = '1'")
11                TDBGrid1.DataSource = rsStallPengering
12            Else
13                rsTemp.MoveFirst
14                stallPengering = Trim(rsTemp!noStall)
15                If Log.FIBeforeSublet = 1 Then
16                   Text(1) = noOrder
17                   Frame2.Visible = True
18                   Exit Sub
19                Else
                     ''rsTemp.MoveFirst
20                   Call objHT.OpenTransactionBengkel
21                   On Error GoTo errorTrans
22                   erTranAll = False
                     
23                   Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'D' WHERE KETERANGAN LIKE '%" & noOrder & "%'")
24                   Call objHT.InsertData_Bengkel("UPDATE TempProgressCuci SET FLAG = 'D',FINISH = CONVERT(VARCHAR,getdate(),108) WHERE NOSTALL = '" & selectStallCuci & "' AND  NOORDER = '" & noOrder & "'")
25                   Call objHT.InsertData_Bengkel("INSERT INTO AntriCuci (NoOrder, Tanggal, NoStall, TipeStall, Waktu, Mech, flag ) VALUES ('" & noOrder & "',GETDATE(),'" & rsTemp("NoStall") & "','Pengering',CONVERT(VARCHAR,getdate(),108),'" & rsTemStallCuci("mech") & "','')")

         
26                   If erTranAll = False Then
27                       Call objHT.CloseTransactionBengkel
28                       MsgBox "Proses Finish Cuci Stall Berhasil !", vbInformation + vbOKOnly
29                       FramePilihPengering.Visible = False
30                       txtScanBarcode.Visible = True
31                       txtScanBarcode.Text = ""
32                       txtScanBarcode.SetFocus
33                       lblScanBarcode.Caption = ""
                                                 
34                       Call Timer2_Timer
35                       Exit Sub
36                   Else
37                       Call objHT.CloseTransactionBengkel_err
38                       MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
39                       Exit Sub
40                   End If
41                   Call Timer2_Timer
42                   Exit Sub
43                End If
                  
44            End If
45        Else
46            If Log.FIBeforeSublet = 1 Then
47                Text(0) = noOrder
48                Frame1.Visible = True
49                Exit Sub
50            Else
51                If Trim(noOrder) = "" Then
52                   MsgBox "No Wo kosong tidak dapat melanjutkan proses Finish", vbInformation
53                   Exit Sub
54                End If
55                Call objHT.OpenTransactionBengkel
56                On Error GoTo errorTrans
                  
57                erTranAll = False
58                Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'F' WHERE KETERANGAN LIKE '%" & noOrder & "%'")
59                Call objHT.InsertData_Bengkel("UPDATE TempProgressCuci SET FLAG = 'F',FINISH = CONVERT(VARCHAR,getdate(),108) WHERE NOSTALL = '" & selectStallCuci & "' AND  NOORDER = '" & noOrder & "'")
                  ''2016-05-16 ""
60                Call objHT.InsertData_Bengkel("UPDATE ANTRICUCI SET FLAG = 'FINISH' WHERE NOSTALL = '" & selectStallCuci & "' AND  NOORDER = '" & noOrder & "'")
                  
61                If erTranAll = False Then
62                    Call objHT.CloseTransactionBengkel
63                    Frame1.Visible = False
64                    MsgBox "Proses Finish Cuci Stall Berhasil !", vbInformation + vbOKOnly
65                    Call Timer2_Timer
66                    Exit Sub
67                Else
68                    Call objHT.CloseTransactionBengkel_err
69                    MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
70                    Exit Sub
71                End If
72                Call Timer2_Timer
                  
73                Exit Sub
74            End If
              
75        End If
errorTrans:
76        Call objHT.CloseTransactionBengkel_err
77        MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
78        Exit Sub

End Sub

Private Sub cmdFinishAkhir_Click()
1        If Trim(noOrder) = "" Then
2           MsgBox "No Wo kosong tidak dapat melanjutkan proses Finish", vbInformation
3           Exit Sub
4        End If
         
5         Call objHT.OpenTransactionBengkel
6         On Error GoTo errorTrans
          
7         erTranAll = False
          
8         Call objHT.InsertData_Bengkel("UPDATE ANTRICUCI SET POSISIAKHIR = '" & Trim(Text(4)) & "' WHERE NOORDER = '" & noOrder & "'")
9         Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'F' WHERE KETERANGAN LIKE '%" & noOrder & "%'")
10        Call objHT.InsertData_Bengkel("UPDATE TempProgressCuci SET FLAG = 'F',FINISH = CONVERT(VARCHAR,getdate(),108) WHERE NOSTALL = '" & selectStallCuci & "' AND  NOORDER = '" & noOrder & "'")
          '' 2016-05-16 ""
11        Call objHT.InsertData_Bengkel("UPDATE ANTRICUCI SET FLAG = 'FINISH' WHERE NOSTALL = '" & selectStallCuci & "' AND  NOORDER = '" & noOrder & "'")
          
12        If erTranAll = False Then
13            Call objHT.CloseTransactionBengkel
14            Frame1.Visible = False
15            MsgBox "Proses Finish Cuci Stall Berhasil !", vbInformation + vbOKOnly
16            Call Timer2_Timer
17            Exit Sub
18        Else
19            Call objHT.CloseTransactionBengkel_err
20            MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
21            Exit Sub
22        End If
23        Call Timer2_Timer
          
24        Exit Sub
errorTrans:
25        Call objHT.CloseTransactionBengkel_err
26        MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
27        Exit Sub
End Sub

Private Sub cmdFinishKePengering_Click()

1        On Error GoTo cmdFinishKePengering_Click_Error
2                    Call objHT.OpenTransactionBengkel
                     ''On Error GoTo errorTrans
3                    erTranAll = False
                     
4                    Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'D' WHERE KETERANGAN LIKE '%" & Trim(Text(1)) & "%'")
5                    Call objHT.InsertData_Bengkel("UPDATE TempProgressCuci SET FLAG = 'D',FINISH = CONVERT(VARCHAR,getdate(),108) WHERE NOSTALL = '" & selectStallCuci & "' AND  NOORDER = '" & Trim(Text(1)) & "'")
6                    Call objHT.InsertData_Bengkel("INSERT INTO AntriCuci ([NoOrder],[Tanggal],[NoStall],[TipeStall],[Waktu],[Mech],[flag],[posisiAwal],[posisiAkhir]) VALUES ('" & Trim(Text(1)) & "',GETDATE(),'" & stallPengering & "','Pengering',CONVERT(VARCHAR,getdate(),108),'" & rsTemStallCuci("mech") & "','','" & Trim(Text(2)) & "','')")

         
7                    If erTranAll = False Then
8                        Call objHT.CloseTransactionBengkel
9                        Frame2.Visible = False
10                       MsgBox "Proses Finish Cuci Stall Berhasil !", vbInformation + vbOKOnly
11                       FramePilihPengering.Visible = False
12                       txtScanBarcode.Text = ""
13                       txtScanBarcode.Visible = True
14                       lblScanBarcode.Caption = ""
15                       txtScanBarcode.SetFocus
16                       Call Timer2_Timer
17                       Exit Sub
18                   Else
19                       Call objHT.CloseTransactionBengkel_err
20                        Frame2.Visible = False
21                       MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
22                       Exit Sub
23                   End If
24                   Call Timer2_Timer
25                   Exit Sub
         
26       On Error GoTo 0
27       Exit Sub

cmdFinishKePengering_Click_Error:
         'Call objHT.CloseTransactionBengkel_err
28       MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdFinishKePengering_Click of Form frmOrderCuci"
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdStart_Click()
1        On Error GoTo cmdStart_Click_Error
2         If rsStallCuci.RecordCount > 0 Then
3             MsgBox "Stall sedang proses Cuci!", vbCritical
4         Else
5             Call objO.InsertData("INSERT INTO TempProgressCuci VALUES ('" & listPengering.Columns(2).value & "',GETDATE(),'" & selectStallCuci & "',CONVERT(VARCHAR,getdate(),108),'','" & listPengering.Columns(7).value & "','')")
6             Call objO.DeleteData("UPDATE AntriCuci SET FLAG = 'PROSES' WHERE NOORDER = '" & listPengering.Columns(2).value & "'")
7             Call Timer2_Timer
8         End If

9        On Error GoTo 0
10       Exit Sub

cmdStart_Click_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdStart_Click of Form frmOrderCuci")
End Sub

Private Sub Form_Load()
    Call FetchCombo
    selectStallCuci = ""
    Set rsTemp = objO.FetchData("SELECT NoStall,NamaStall FROM qty_stallwashing WHERE tipestall = 'Cuci' AND ISDEL = '0' and aktif = '1'")
    If rsTemp.RecordCount = 1 Then
        cbStallCuci.value = rsTemp("NoStall")
    End If
    If Log.barcode = 1 Then
        txtScanBarcode.Visible = True
    End If
End Sub

Private Sub FetchCombo()
1        On Error GoTo FetchCombo_Error

2         Adodc1.ConnectionString = SQLCONBengkel
3         Adodc1.RecordSource = "SELECT NoStall,NamaStall,petugas FROM qty_stallwashing WHERE tipestall = 'Cuci' AND ISDEL = '0'  and aktif = '1'"
4         cbStallCuci.DataSourceList = Adodc1
5         cbStallCuci.DataFieldList = "NoStall"
6         cbStallCuci.DataFieldToDisplay = "NamaStall"
7         cbStallCuci.Columns(0).Width = 1200
8         cbStallCuci.Columns(1).Width = 4000
9         cbStallCuci.Columns(0).Caption = "No Stall"
10        cbStallCuci.Columns(1).Caption = "Nama Stall"

11       On Error GoTo 0
12       Exit Sub

FetchCombo_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure FetchCombo of Form frmOrderCuci")
          
End Sub

Private Sub TDBGrid1_Click()
    'btnPengeringPilih.Enabled = True
    Call btnPengeringPilih_Click
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    Dim n As String
    
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            n = Text(Index).Mask
            Text(Index).Mask = ""
            Text(Index) = gMainKeyCari
            gMainKeyCari = ""
            Text(Index).Mask = n
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
            Case 4
                gMainKeyCari = Text(Index)
                gSearchOption = "POSISI_MOBIL"
                frmCari.show (vbModal)
        End Select
    End If
End Sub

Private Sub Timer1_Timer()
    'MsgBox Format(Time, "HH:MM:SS")
    lblHari.Caption = WeekdayName(IIf(Weekday(isnow) = 1, 7, Weekday(isnow) - 1)) & ", " & Format(isnow, "DD - MMMM - YYYY")
    lblJam.Caption = Format(Time, "HH:MM:SS")
End Sub

Private Sub Timer2_Timer()
    If selectStallCuci <> "" Then
        refreshGrid
        Call refreshGridAntriCuci(IIf(Check1.value = 1, True, False))
    End If
End Sub

Private Sub txtScanBarcode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'MsgBox txtScanBarcode.text
         If Len(txtScanBarcode.Text) >= 5 Then

            If UCase(Trim(txtScanBarcode.Text)) = "START" Then
                If selectStallCuci <> "" Then
                    txtScanBarcode.Text = ""
                    txtScanBarcode.SetFocus
                    If rsStallCuci.RecordCount > 0 Then
                        MsgBox "Stall sedang proses Cuci!", vbCritical
                    Else
                        Call objO.InsertData("INSERT INTO TempProgressCuci VALUES ('" & listPengering.Columns(2).value & "',GETDATE(),'" & selectStallCuci & "',CONVERT(VARCHAR,getdate(),108),'','" & listPengering.Columns(6).value & "','')")
                        Call objO.DeleteData("UPDATE AntriCuci SET FLAG = 'PROSES' WHERE NOORDER = '" & listPengering.Columns(2).value & "'")
                        Call Timer2_Timer
                    End If
                Else
                    MsgBox "Silahkan pilih stall cuci terlebih dahulu", vbCritical
                End If
            ElseIf UCase(Trim(txtScanBarcode.Text)) = "FINISH" Then
                Call cmdFinish_Click
            ElseIf FramePilihPengering.Visible = True Then
                If Not rsStallPengering.EOF Then
                    rsStallPengering.MoveFirst
                    Do While Not rsStallPengering.EOF
                        If rsStallPengering("NoStall") = Trim(txtScanBarcode.Text) Then
                                Call objHT.OpenTransactionBengkel
                                On Error GoTo errorTrans
                                erTranAll = False
                                
                                Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'D' WHERE KETERANGAN LIKE '%" & noOrder & "%'")
                                Call objHT.InsertData_Bengkel("UPDATE TempProgressCuci SET FLAG = 'D',FINISH = CONVERT(VARCHAR,getdate(),108) WHERE NOSTALL = '" & selectStallCuci & "' AND  NOORDER = '" & noOrder & "'")
                                
                                Call objHT.InsertData_Bengkel("INSERT INTO AntriCuci (NoOrder, Tanggal, NoStall, TipeStall, Waktu, Mech, flag ) VALUES ('" & ListMech.Columns(2).value & "',GETDATE(),'" & rsStallPengering("NoStall") & "','Pengering',CONVERT(VARCHAR,getdate(),108),'" & rsTemStallCuci("mech") & "','')")
                                If erTranAll = False Then
                                    Call objHT.CloseTransactionBengkel
                                    MsgBox "Proses Finish Cuci Stall Berhasil !", vbInformation + vbOKOnly
                                    FramePilihPengering.Visible = False
                                    txtScanBarcode.Text = ""
                                    txtScanBarcode.SetFocus
                                    lblScanBarcode.Caption = ""
                                                            
                                    Call Timer2_Timer
                                    Exit Sub
                                Else
                                    Call objHT.CloseTransactionBengkel_err
                                    MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
                                    Exit Sub
                                End If
                                Call Timer2_Timer
                                Exit Sub
errorTrans:
                                Call objHT.CloseTransactionBengkel_err
                                MsgBox "Tidak berhasil proses Finish Cuci stall !", vbInformation + vbOKOnly
                                Exit Sub
                        End If
                        rsStallPengering.MoveNext
                    Loop
                End If
                
            Else
                Set rsTemp = objO.FetchData("SELECT NoStall,NamaStall FROM qty_stallwashing WHERE tipestall = 'Cuci' AND NOSTALL = '" & Trim(txtScanBarcode.Text) & "'  and aktif = '1'")
    
                If rsTemp.RecordCount > 0 Then
                        selectStallCuci = Trim(rsTemp("NoStall"))
                        namaStallCuci = Trim(rsTemp("NamaStall"))
                        cbStallCuci.value = selectStallCuci
                        Call refreshGrid
                        Call refreshGridAntriCuci(False)
                Else
                    'pengecekan no antrian
                    If Len(Trim(txtScanBarcode.Text)) >= 16 And Len(Trim(txtScanBarcode.Text)) <= 18 Then
                        Dim scanAntrian As String
                        Dim tanggalSA As Date
                        Dim kodeSA As String
                        Dim nomorSA As Integer
                        Dim noOrderScan As String
                        
                        scanAntrian = Trim(txtScanBarcode.Text)
                        
                        tanggalSA = CDate(Mid(scanAntrian, InStr(scanAntrian, "-") + 2, Len(scanAntrian)))
                        kodeSA = Mid(scanAntrian, 1, InStr(scanAntrian, "-") - 2)
                        
                        Select Case Len(kodeSA)
                        Case 3
                            nomorSA = Mid(kodeSA, 2, Len(kodeSA))
                             kodeSA = Mid(kodeSA, 1, 1)
                        Case 4
                            If Mid(kodeSA, 2, 1) = "1" Then
                                nomorSA = Mid(kodeSA, 2, Len(kodeSA))
                                kodeSA = Mid(kodeSA, 1, 1)
                                
                            Else
                                nomorSA = Mid(kodeSA, 3, Len(kodeSA))
                                kodeSA = Mid(kodeSA, 1, 2)
                            End If
                        End Select
                    
                        Set rsTemp = objO.FetchData("select top 10 l.*,m.initial from listantrian l inner join " & Log.LogQueue & "..menu m on l.tipe = m.tipe WHERE m.INITIAL = '" & kodeSA & "' and l.tanggal = '" & Format(tanggalSA, "MM/DD/YYYY") & "' and l.queue = '" & nomorSA & "' order by tanggal desc")
                        If rsTemp.RecordCount > 0 Then
                            noOrderScan = Trim(rsTemp("WO"))
                        Else
                            MsgBox "Kode Barcode tidak dikenali", vbCritical
                            txtScanBarcode.Text = ""
                            txtScanBarcode.SetFocus
                            lblScanBarcode.Caption = "Kode barcode tidak dikenali"
                            Exit Sub
                        End If
                        
                        If selectStallCuci = "" Then
                            MsgBox "Tidak ada antrian pada stall cuci", vbCritical
                            txtScanBarcode.Text = ""
                            txtScanBarcode.SetFocus
                            lblScanBarcode.Caption = ""
                            Exit Sub
                        ElseIf rsTemAntriCuci.RecordCount = 0 And noOrder = "" Then
                            MsgBox "Tidak ada antrian pada stall cuci", vbCritical
                            txtScanBarcode.Text = ""
                            txtScanBarcode.SetFocus
                            lblScanBarcode.Caption = ""
                            Exit Sub
                        End If
                        If noOrder <> "" And noOrder = noOrderScan Then
                            'lakukan proses finish
                            If selectStallCuci <> "" Then
                                Call cmdFinish_Click
                                txtScanBarcode.Text = ""
                                txtScanBarcode.SetFocus
                                lblScanBarcode.Caption = ""
                                Exit Sub
                            End If
                            
                        Else
                            'lakukan start
                            If rsStallCuci.RecordCount > 0 Then
                                MsgBox "Stall sedang proses Cuci!", vbCritical
                                txtScanBarcode.Text = ""
                                txtScanBarcode.SetFocus
                                lblScanBarcode.Caption = ""
                                Exit Sub
                            Else
                                Set rsTemp = objO.FetchData("SELECT MECH FROM JOBDTA WHERE NO_WO = '" & noOrderScan & "'")
                                If rsTemp.RecordCount > 0 Then
                                    rsTemp.MoveFirst
                                    If nowoAntrianAwal <> noOrderScan Then
                                        If MsgBox("Apakah ingin mendahulukan cuci untuk no order " & noOrderScan & " ?", vbYesNo + vbDefaultButton2, "Konfirmasi Proses Cuci") = vbYes Then
                                            Call objO.InsertData("INSERT INTO TempProgressCuci VALUES ('" & noOrderScan & "',GETDATE(),'" & selectStallCuci & "',CONVERT(VARCHAR,getdate(),108),'','" & Trim(rsTemp("MECH")) & "','')")
                                            Call objO.DeleteData("UPDATE AntriCuci SET FLAG = 'PROSES' WHERE NOORDER = '" & noOrderScan & "'")
                                            Call Timer2_Timer
                                        End If
                                    Else
                                        Call objO.InsertData("INSERT INTO TempProgressCuci VALUES ('" & noOrderScan & "',GETDATE(),'" & selectStallCuci & "',CONVERT(VARCHAR,getdate(),108),'','" & Trim(rsTemp("MECH")) & "','')")
                                        Call objO.DeleteData("UPDATE AntriCuci SET FLAG = 'PROSES' WHERE NOORDER = '" & noOrderScan & "'")
                                        Call Timer2_Timer
                                    End If
                                    txtScanBarcode.Text = ""
                                    txtScanBarcode.SetFocus
                                    lblScanBarcode.Caption = ""
                                    Exit Sub

                                Else
                                    MsgBox "Terjadi kesalahan dalam mengambil data, proses start cuci dibatalkan.", vbCritical
                                    txtScanBarcode.Text = ""
                                    txtScanBarcode.SetFocus
                                    lblScanBarcode.Caption = ""
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                    MsgBox "Kode Barcode tidak dikenali", vbCritical
                    txtScanBarcode.Text = ""
                    txtScanBarcode.SetFocus
                    lblScanBarcode.Caption = "Kode barcode tidak dikenali"
                    Exit Sub
                End If
            End If
            txtScanBarcode.Text = ""
            txtScanBarcode.SetFocus
            lblScanBarcode.Caption = "Scan Barcode, Stall Cuci " & namaStallCuci
        Else
            txtScanBarcode.Text = ""
            txtScanBarcode.SetFocus
            lblScanBarcode.Caption = ""
        End If
    End If
End Sub
