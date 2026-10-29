VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmAturProgress 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "PENGATURAN PEKERJAAN TUNDA"
   ClientHeight    =   6330
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   12945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton tblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke JPCB"
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
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5760
      Width           =   2415
   End
   Begin VB.CommandButton tblSimpan 
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
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5760
      Width           =   1695
   End
   Begin VB.Frame framPindah 
      BackColor       =   &H00C0FFFF&
      Height          =   3615
      Left            =   6000
      TabIndex        =   6
      Top             =   960
      Width           =   3255
      Begin VB.CommandButton cmdBatal 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&Batal"
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   3120
         Width           =   1455
      End
      Begin VB.CommandButton cmdPindah 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&OK"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3120
         Width           =   1455
      End
      Begin VB.ComboBox cboStall 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   2040
         Width           =   1395
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
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
         Index           =   1
         Left            =   1320
         TabIndex        =   11
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   12
         Top             =   720
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
         Index           =   3
         Left            =   1320
         TabIndex        =   1
         Top             =   2400
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   1320
         TabIndex        =   2
         Top             =   2760
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   270
         Left            =   2160
         TabIndex        =   17
         Top             =   2760
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
         Index           =   5
         Left            =   1320
         TabIndex        =   18
         Top             =   960
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">CCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mekanik"
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
         Left            =   450
         TabIndex        =   19
         Top             =   960
         Width           =   690
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mekanik"
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
         Left            =   375
         TabIndex        =   16
         Top             =   2760
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jam"
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
         Left            =   720
         TabIndex        =   15
         Top             =   2400
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stall Baru"
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
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "PINDAH KE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   960
         TabIndex        =   13
         Top             =   1560
         Width           =   1440
      End
      Begin VB.Line Line1 
         BorderWidth     =   25
         X1              =   0
         X2              =   3240
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stall Lama"
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
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. WO"
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
         TabIndex        =   8
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Polisi"
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
         Left            =   420
         TabIndex        =   7
         Top             =   480
         Width           =   720
      End
   End
   Begin TrueOleDBGrid70.TDBGrid grid 
      Height          =   5655
      Left            =   480
      TabIndex        =   5
      Top             =   0
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   9975
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NOORDER"
      Columns(0).DataField=   "NOORDER"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "NOPOL"
      Columns(1).DataField=   "NOPOL"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "CUSTOMER"
      Columns(2).DataField=   "CUSTOMER"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "STALL"
      Columns(3).DataField=   "STALL"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "TANGGAL"
      Columns(4).DataField=   "TANGGAL"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "JAMMULAI"
      Columns(5).DataField=   "JAMMULAI"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "JAMSELESAI"
      Columns(6).DataField=   "JAMSELESAI"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "RATE"
      Columns(7).DataField=   "RATE"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "GROUP"
      Columns(8).DataField=   "GROUP"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "MECH"
      Columns(9).DataField=   "MECH"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   4
      Columns(10)._MaxComboItems=   5
      Columns(10).Caption=   "LANJUT"
      Columns(10).DataField=   "LANJUT"
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   4
      Columns(11)._MaxComboItems=   5
      Columns(11).Caption=   "PINDAH"
      Columns(11).DataField=   "PINDAH"
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   12
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=12"
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
      Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=2725"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
      PrintInfos(0).PageFooterFont=   "Size=9,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Trebuchet MS"
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Trebuchet MS"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Trebuchet MS"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=900,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Trebuchet MS"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HF1E4BE&"
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
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=54,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=51,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=52,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=53,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=50,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=47,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=48,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=49,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=46,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=43,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=44,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=45,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=28,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=25,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=26,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=27,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=32,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=29,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=30,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=31,.parent=17"
      _StyleDefs(84)  =   "Named:id=33:Normal"
      _StyleDefs(85)  =   ":id=33,.parent=0"
      _StyleDefs(86)  =   "Named:id=34:Heading"
      _StyleDefs(87)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(88)  =   ":id=34,.wraptext=-1"
      _StyleDefs(89)  =   "Named:id=35:Footing"
      _StyleDefs(90)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(91)  =   "Named:id=36:Selected"
      _StyleDefs(92)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(93)  =   "Named:id=37:Caption"
      _StyleDefs(94)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(95)  =   "Named:id=38:HighlightRow"
      _StyleDefs(96)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(97)  =   "Named:id=39:EvenRow"
      _StyleDefs(98)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(99)  =   "Named:id=40:OddRow"
      _StyleDefs(100) =   ":id=40,.parent=33"
      _StyleDefs(101) =   "Named:id=41:RecordSelector"
      _StyleDefs(102) =   ":id=41,.parent=34"
      _StyleDefs(103) =   "Named:id=42:FilterBar"
      _StyleDefs(104) =   ":id=42,.parent=33"
   End
   Begin VB.Image Image1 
      Height          =   6255
      Left            =   0
      Picture         =   "frmAturProgress.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmAturProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHlp As ADODB.Recordset
Dim rsGrid As New ADODB.Recordset
Dim nTgl As String
Dim objProgress As New BClsProgress
Dim cols As TrueOleDBGrid70.Columns

Private Function createRecord() As Boolean
          
1         On Error GoTo vb_error
2         createRecord = False
          
3         If rsGrid.State Then rsGrid.Close
4         Set rsGrid = Nothing
5         With rsGrid
6             .Fields.Append "NOORDER", adVarChar, 10
7             .Fields.Append "NOPOL", adVarChar, 12
8             .Fields.Append "CUSTOMER", adVarChar, 50
9             .Fields.Append "STALL", adVarChar, 10
10            .Fields.Append "TANGGAL", adVarChar, 10
11            .Fields.Append "JAMMULAI", adVarChar, 5
12            .Fields.Append "JAMSELESAI", adVarChar, 5
13            .Fields.Append "RATE", adDouble
14            .Fields.Append "GROUP", adVarChar, 10
15            .Fields.Append "MECH", adVarChar, 10
16            .Fields.Append "LANJUT", adBoolean
17            .Fields.Append "PINDAH", adBoolean
18            .Fields.Append "PINDAHAN", adVarChar, 10
19            .Fields.Append "JAMPINDAHAN", adVarChar, 5
20            .Open
21        End With
22        If rsGrid.State Then createRecord = True
23        Set grid.DataSource = rsGrid
24        Exit Function
          
vb_error:
25        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub refreshGrid()
    
    Set rsHlp = objO.FetchData("SELECT SUBSTRING(KETERANGAN, CHARINDEX(' |',KETERANGAN) + 2, CHARINDEX('| ',KETERANGAN) - (CHARINDEX(' |',KETERANGAN)+2)) AS NOORDER, " _
                            & "LEFT(KETERANGAN, CHARINDEX(' |', KETERANGAN)-1) AS NOPOL, SUBSTRING(KETERANGAN, CHARINDEX('| ', KETERANGAN) +2, LEN(KETERANGAN) - CHARINDEX('| ', KETERANGAN)) AS CUSTOMER, " _
                            & "NOSTALL AS STALL, TANGGAL, JAMMULAI AS MULAI, JAMSELESAI AS SELESAI, RATE, [GROUP], MECH FROM STALL WITH(NOLOCK) WHERE TANGGAL < '" & Format(isnow, "MM/DD/YYYY") & "' AND FLAG IN ('P','N') ORDER BY TANGGAL DESC, NOORDER") 'TANGGAL = '" & Format(DateAdd("d", -1, isnow), "mm/dd/yyyy") & "' AND
    If Not rsHlp.EOF Then
        Set grid.DataSource = Nothing
        For Each col In grid.Columns
            col.FilterText = ""
        Next col
        rsGrid.filter = ""
        If createRecord Then
            Do While Not rsHlp.EOF
                With rsGrid
                    .AddNew
                    !noOrder = Trim(rsHlp("NOORDER"))
                    !noPol = Trim(rsHlp("NOPOL"))
                    !CUSTOMER = Trim(rsHlp("CUSTOMER"))
                    !stall = Trim(rsHlp("STALL"))
                    !Tanggal = Trim(TglDDMMYYYY(rsHlp("TANGGAL")))
                    !jamMulai = Trim(rsHlp("MULAI"))
                    !jamSelesai = Trim(rsHlp("SELESAI"))
                    !Rate = objProgress.newRate(Trim(rsHlp("NOORDER")))  'CDbl(rsHlp("RATE"))
                    !Group = Trim(rsHlp("GROUP"))
                    !Mech = Trim(rsHlp("MECH"))
                    !lanjut = False
                    !pindah = False
                    .Update
                End With
                rsHlp.MoveNext
            Loop
        End If
    End If
    If rsGrid.State <> adStateClosed Then rsGrid.MoveFirst
    Set grid.DataSource = rsGrid
    Call HeadColumn
End Sub

Private Sub isiCombo()
Dim hlp As String
    
    cboStall.Clear
    Set rsHlp = objO.FetchData("SELECT FN_JML_STALL FROM T_PROFIL WITH(NOLOCK)")
    If Not rsHlp.EOF Then
        For i = 1 To rsHlp("FN_JML_STALL")
            hlp = "Stall " & i
            cboStall.AddItem hlp
        Next
    End If
    
    Set rsHlp = objO.FetchData("SELECT FN_JML_THS FROM T_PROFIL WITH(NOLOCK)")
    If Not rsHlp.EOF Then
        For i = 1 To rsHlp("FN_JML_THS")
            hlp = "THS " & i
            cboStall.AddItem hlp
        Next
    End If
End Sub

Private Sub HeadColumn()
    
    With grid
        .Columns(0).Width = 1000
        .Columns(0).Caption = "NO WO"
        .Columns(0).DataField = "NOORDER"
        .Splits(0).Columns(0).Locked = True
        
        .Columns(1).Width = 1000
        .Columns(1).Caption = "NO POLISI"
        .Columns(1).DataField = "NOPOL"
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Width = 3500
        .Columns(2).Caption = "CUSTOMER"
        .Columns(2).DataField = "CUSTOMER"
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Width = 1000
        .Columns(3).Caption = "STALL"
        .Columns(3).DataField = "STALL"
        .Splits(0).Columns(3).Locked = True
        
        .Columns(4).Width = 1200
        .Columns(4).Caption = "TANGGAL"
        .Columns(4).DataField = "TANGGAL"
        .Splits(0).Columns(4).Locked = True
        
        .Columns(5).Width = 800
        .Columns(5).Caption = "MULAI"
        .Columns(5).DataField = "JAMMULAI"
        .Splits(0).Columns(5).Locked = True
        
        .Columns(6).Width = 0
        .Columns(6).Caption = "SELESAI"
        .Columns(6).DataField = "JAMSELESAI"
        .Columns(6).Visible = False
        .Splits(0).Columns(6).Locked = True
        
        .Columns(7).Width = 800
        .Columns(7).Caption = "RATE"
        .Columns(7).DataField = "RATE"
        .Splits(0).Columns(7).Locked = True
        
        .Columns(8).Width = 800
        .Columns(8).Caption = "GROUP"
        .Columns(8).DataField = "GROUP"
        .Columns(8).Visible = False
        .Splits(0).Columns(8).Locked = True
        
        .Columns(9).Width = 800
        .Columns(9).Caption = "MEKANIK"
        .Columns(9).DataField = "MECH"
        .Splits(0).Columns(9).Locked = True
        
        .Columns(10).Width = 800
        .Columns(10).Caption = "LANJUT PROSES"
        .Columns(10).DataField = "LANJUT"
        .Columns(10).Alignment = dbgCenter
        
        .Columns(11).Width = 800
        .Columns(11).Caption = "PINDAH STALL"
        .Columns(11).DataField = "PINDAH"
        .Columns(11).Alignment = dbgCenter
    End With
End Sub

Private Sub cmdBatal_Click()
    framPindah.Visible = False
End Sub

Private Sub cmdPindah_Click()
    If text(3) <> ">>:>>" And cboStall <> "" Then
        If Trim(text(4)) = "" Then
            If MsgBox("Mekanik tidak diganti? ", vbYesNo, "Ganti Mekanik") = vbNo Then
                text(4).SelLength = Len(text(4))
                text(4).SetFocus
                Exit Sub
            End If
        End If
        rsGrid.Find "NOORDER = '" & Trim(text(0)) & "'", , adSearchForward, 1
        If Not rsGrid.EOF Then
            rsGrid!PINDAHAN = cboStall
            rsGrid!JAMPINDAHAN = Trim(text(3))
            If Trim(text(4)) <> "" Then
                rsGrid!Mech = Trim(text(4))
            End If
        End If
        MsgBox "Kendaraan dengan No. Polisi " & rsGrid!noPol & " Akan dikerjakan di " & cboStall & " , jam " & Trim(text(3))
        framPindah.Visible = False
    Else
        MsgBox "Lengkapi data ! "
        text(3).SetFocus
    End If
End Sub

Private Sub Form_Load()
    nTgl = Format(isnow, "MM/DD/YYYY")
    Call refreshGrid
    framPindah.Visible = False
End Sub

Private Sub grid_AfterColUpdate(ByVal ColIndex As Integer)
Dim pesan As String
Dim harusPindah As Boolean
Dim nDay As String
'
    nDay = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
    With grid
        pesan = ""
        harusPindah = False
        Select Case ColIndex
            Case 10     'TOMBOL LANJUT
                If .Columns(ColIndex).value Then
                    pesan = "Kendaraan dengan No. Polisi " & .Columns(1) & " Akan dilanjutkan pada " & .Columns(3) & ", Jam " & Format(isnow, "HH:NN")
                    If Not objProgress.searchMechanic(.Columns(9), nDay, Format(isnow, "DD/MM/YYYY"), Format(isnow, "HH:NN")) Then
                        pesan = pesan + " " & vbCr & "Mekanik " & grid.Columns(9) & " tidak masuk, atau jam kerjanya penuh !"
                        harusPindah = True
                    End If
                    MsgBox pesan
                    If harusPindah Then
                        .Columns(11).value = True
                        Call isiCombo
                        text(0) = grid.Columns(0)
                        text(1) = grid.Columns(1)
                        text(2) = grid.Columns(3)
                        text(5) = grid.Columns(9)
                        framPindah.Visible = True
                    End If
                ElseIf Not .Columns(ColIndex).value Then
                    MsgBox "Kendaraan dengan No. Polisi " & .Columns(1) & " Batal dilanjutkan."
                End If
            Case 11     'TOMBOL PINDAH
                If .Columns(10).value Then
                    If .Columns(ColIndex).value = -1 Then
                        ' PROSES PINDAH STALL
                        Call isiCombo
                        text(0) = grid.Columns(0)
                        text(1) = grid.Columns(1)
                        text(2) = grid.Columns(3)
                        text(5) = grid.Columns(9)
                        framPindah.Visible = True
                    ElseIf .Columns(ColIndex).value <> -1 Then
                        ' BATAL PINDAH
                        rsGrid.Find "NOORDER = '" & Trim(grid.Columns(0)) & "'", , adSearchForward, 1
                        If Not rsGrid.EOF Then
                            rsGrid!PINDAHAN = ""
                            rsGrid!JAMPINDAHAN = ""
                        End If
                    End If
                Else
                    If .Columns(11).value = -1 Then
                        MsgBox "Centang Pilihan Lanjut !!"
                        .Columns(11).value = False
                        Exit Sub
                    End If
                End If
        End Select
    End With
End Sub

Private Sub grid_FilterChange()
On Error GoTo errHandler
    Set cols = grid.Columns
    Dim C As Integer
    
    C = grid.col
    grid.HoldFields
    rsGrid.filter = getFilter()
    If rsGrid.EOF Then
        For Each col In grid.Columns
            col.FilterText = ""
        Next col
        rsGrid.filter = ""
    End If
    grid.col = C
    grid.EditActive = True
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

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSimpan_Click()
Dim rsHlp2 As ADODB.Recordset
Dim nRangka As String
Dim nDay As String
Dim nFlag As String
Dim nFO As String
Dim unitEM As Boolean
Dim sProses As Boolean

    If rsGrid.State <> adStateClosed Then
        rsGrid.MoveFirst
        
        nDay = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
        Do While Not rsGrid.EOF
            Dim rsHlp3 As ADODB.Recordset
            Set rsHlp3 = objO.FetchData("SELECT * FROM MEKANIK WITH(NOLOCK) WHERE MECH = '" & Trim(rsGrid("MECH")) & "'")
            If Not rsHlp3.EOF Then
                nFO = Trim(rsHlp3("GROUP"))
            End If
            
            unitEM = False
            Set rsHlp3 = objO.FetchData("SELECT * FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & Trim(rsGrid("NOORDER")) & "'")
            If Not rsHlp3.EOF Then
                unitEM = IIf(rsHlp3("CODE") = 1, True, False)
            End If
            
            If rsGrid("LANJUT") Then
                If Not rsGrid("PINDAH") Then
                    ' PROSES LANJUT
                    Set rsHlp = objO.FetchData("SELECT SUBSTRING(KETERANGAN, CHARINDEX(' |',KETERANGAN) + 2, CHARINDEX('| ',KETERANGAN) - (CHARINDEX(' |',KETERANGAN)+2)) AS NOORDER, * FROM STALL WITH(NOLOCK) WHERE NOSTALL = '" & rsGrid("STALL") & "' AND TANGGAL = '" & nTgl & "' " _
                                            & "AND KETERANGAN NOT LIKE '%" & rsGrid("NOORDER") & "%' AND FLAG <> 'X' and ((DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Left(hitungWaktu(rsGrid("RATE"), Format(isnow, "HH:NN"), nTgl), 5) & "') > 0) " _
                                            & "or (DATEDIFF(n,JAMMULAI, '" & Format(isnow, "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Left(hitungWaktu(rsGrid("RATE"), Format(isnow, "HH:NN"), nTgl), 5) & "') > 0))")
                                            
                    If Not rsHlp.EOF Then
                        MsgBox "Digunakan untuk unit kendaraan lain. "
                        Exit Sub
                    Else
                        ' PROSES LANJUT DIJALANKAN
                        Set rsHlp2 = objO.FetchData("SELECT NO_CHASSIS FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & rsGrid("NOORDER") & "'")
                        If Not rsHlp2.EOF Then
                            nRangka = Trim(rsHlp2("NO_CHASSIS"))
                        End If
                        
                        nFlag = "P"
                        If Not unitEM Then
                            Call objO.InsertData("INSERT INTO TEMPPROGRESS VALUES ('" & rsGrid("NOORDER") & "', '" & nTgl & "', '" & rsGrid("STALL") & "', '" & Format(isnow, "HH:NN") & "', '','','" & rsGrid("MECH") & "','" & Trim(nFO) & "')")
                            Call objO.InsertData("INSERT INTO PROGRESS VALUES ('" & rsGrid("NOORDER") & "', '" & nRangka & "', 'S', '" & rsGrid("STALL") & "', '" & nTgl & "', '" & Format(isnow, "HH:NN") & "', '" & gSA & "', '', '" & rsGrid("MECH") & "')")
                            Call objO.UpdateData("UPDATE STALL SET TANGGAL = '" & nTgl & "', JAMMULAI = '" & pembulatanJam(Format(isnow, "HH:NN")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(rsGrid("RATE")) < 0.3, 0.3, CDbl(rsGrid("RATE"))), Format(isnow, "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "', FLAG = 'S', RATE = '" & Replace(IIf(CDbl(rsGrid("RATE")) < 0.3, 0.3, CDbl(rsGrid("RATE"))), ",", ".") & "', MECH = '" & Trim(rsGrid("MECH")) & "' WHERE KETERANGAN LIKE '%|" & rsGrid("NOORDER") & "|%' AND TANGGAL = '" & Format(rsGrid("TANGGAL"), "mm/dd/yyyy") & "' AND NOSTALL = '" & rsGrid("STALL") & "'")
                            Call objO.UpdateData("UPDATE JOBDTA SET MECH = '" & Trim(rsGrid("MECH")) & "', [GROUP] = '" & Trim(nFO) & "', RPR_COND = 'P' WHERE NO_WO = '" & rsGrid("NOORDER") & "'")
                            Call objO.UpdateData("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI + " & Replace((CDbl(rsGrid("RATE")) * 60), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & rsGrid("MECH") & "'")
                            nFlag = "S"
                        End If
                        Call objO.UpdateData("UPDATE STALL SET TANGGAL = '" & nTgl & "', JAMMULAI = '" & pembulatanJam(Format(isnow, "HH:NN")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(CDbl(rsGrid("RATE")) < 0.3, 0.3, CDbl(rsGrid("RATE"))), Format(isnow, "HH:NN"), Format(isnow, "MM/DD/YYYY"), "progress"), 5) & "', FLAG = '" & nFlag & "', RATE = '" & Replace(IIf(CDbl(rsGrid("RATE")) < 0.3, 0.3, CDbl(rsGrid("RATE"))), ",", ".") & "', MECH = '" & Trim(rsGrid("MECH")) & "' WHERE KETERANGAN LIKE '%|" & rsGrid("NOORDER") & "|%' AND TANGGAL = '" & Format(rsGrid("TANGGAL"), "mm/dd/yyyy") & "' AND NOSTALL = '" & rsGrid("STALL") & "'")
                        
                    End If
                ElseIf rsGrid("PINDAH") Then
                    Set rsHlp = objO.FetchData("SELECT SUBSTRING(KETERANGAN, CHARINDEX(' |',KETERANGAN) + 2, CHARINDEX('| ',KETERANGAN) - (CHARINDEX(' |',KETERANGAN)+2)) AS NOORDER, * FROM STALL WITH(NOLOCK) WHERE NOSTALL = '" & rsGrid("PINDAHAN") & "' AND TANGGAL = '" & nTgl & "' " _
                                            & "AND KETERANGAN NOT LIKE '%" & rsGrid("NOORDER") & "%' AND FLAG <> 'X' and ((DATEDIFF(n,JAMMULAI, '" & Format(rsGrid("JAMPINDAHAN"), "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Left(hitungWaktu(rsGrid("RATE"), rsGrid("JAMPINDAHAN"), nTgl), 5) & "') > 0) " _
                                            & "or (DATEDIFF(n,JAMMULAI, '" & Format(rsGrid("JAMPINDAHAN"), "HH:NN") & "') < 0 AND DATEDIFF(n,JAMSELESAI, '" & Left(hitungWaktu(rsGrid("RATE"), rsGrid("JAMPINDAHAN"), nTgl), 5) & "') > 0))")
                                            
                    If Not rsHlp.EOF Then
                        MsgBox "Digunakan untuk unit kendaraan lain. "
                        grid.Columns(11).value = False
                        Exit Sub
                    Else
                        ' PROSES LANJUT DIJALANKAN --> DENGAN PINDAH STALL / MEKANIK / KEDUANYA
                        Set rsHlp2 = objO.FetchData("SELECT NO_CHASSIS FROM JOBDTA WITH(NOLOCK) WHERE NO_WO = '" & rsGrid("NOORDER") & "'")
                        If Not rsHlp2.EOF Then
                            nRangka = Trim(rsHlp2("NO_CHASSIS"))
                        End If
                        nFlag = "P"
                        If Not unitEM Then
                            If Trim(rsGrid("JAMPINDAHAN")) >= Trim(Format(isnow, "HH:NN")) And Trim(rsGrid("JAMPINDAHAN")) <= Trim(Format(DateAdd("n", 5, Format(isnow, "HH:NN")), "HH:NN")) Then
                                Call objO.InsertData("INSERT INTO TEMPPROGRESS VALUES ('" & rsGrid("NOORDER") & "', '" & nTgl & "', '" & rsGrid("PINDAHAN") & "', '" & Format(rsGrid("JAMPINDAHAN"), "HH:NN") & "', '','', '" & rsGrid("MECH") & "','" & Trim(nFO) & "')")
                                Call objO.InsertData("INSERT INTO PROGRESS VALUES ('" & rsGrid("NOORDER") & "', '" & nRangka & "', 'S', '" & rsGrid("PINDAHAN") & "', '" & nTgl & "', '" & Format(rsGrid("JAMPINDAHAN"), "HH:NN") & "', '" & gSA & "', '', '" & rsGrid("MECH") & "')")
                                Call objO.UpdateData("UPDATE JOBDTA SET RPR_COND = 'P' WHERE NO_WO = '" & rsGrid("NOORDER") & "'")
                                Call objO.UpdateData("UPDATE T_MEKANIK SET TERPAKAI = TERPAKAI + " & Replace((CDbl(rsGrid("RATE")) * 60), ",", ".") & " WHERE HARI = '" & nDay & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND MECH = '" & rsGrid("MECH") & "'")
                                Call objO.UpdateData("UPDATE JOBDTA SET MECH = '" & Trim(rsGrid("MECH")) & "', [GROUP] = '" & Trim(nFO) & "' WHERE NO_WO = '" & rsGrid("NOORDER") & "'")
                                nFlag = "S"
                            End If
                        End If
                        Call objO.UpdateData("UPDATE STALL SET NOSTALL = '" & rsGrid("PINDAHAN") & "', TANGGAL = '" & nTgl & "', JAMMULAI = '" & pembulatanJam(Format(rsGrid("JAMPINDAHAN"), "HH:NN")) & "', JAMSELESAI = '" & Left(hitungWaktu(IIf(rsGrid("RATE") < 0.3, 0.3, rsGrid("RATE")), rsGrid("JAMPINDAHAN"), nTgl, "progress"), 5) & "', FLAG = '" & nFlag & "', RATE = '" & Replace(IIf(CDbl(rsGrid("RATE")) < 0.3, 0.3, CDbl(rsGrid("RATE"))), ",", ".") & "', MECH = '" & Trim(rsGrid("MECH")) & "' WHERE KETERANGAN LIKE '%|" & rsGrid("NOORDER") & "|%' AND TANGGAL = '" & Format(rsGrid("TANGGAL"), "mm/dd/yyyy") & "' AND NOSTALL = '" & rsGrid("STALL") & "'")
                    End If
                End If
            End If
            rsGrid.MoveNext
        Loop
        TblKeluar_Click
    End If
End Sub

Private Sub text_GotFocus(Index As Integer)
Dim DiTab As Boolean
    
    DiTab = False
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
            text(Index) = gMainKeyCari
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
       
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 4
                gMainKeyCari = text(Index)
                gBantuan = "ORDER"
                gBantuan2 = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
                gBantuan3 = Format(isnow, "DD/MM/YYYY")
                gSearchOption = "MEKANIK"
                gBantuan4 = ""
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    Select Case Index
        Case 3
            KeyAscii = CekNumerik(KeyAscii)
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset
    
    Select Case Index
        Case 3
            If Not CekJam(text(Index)) Then
                MsgBox "Format jam salah"
                text(Index) = ">>:>>"
                text(Index).SetFocus
                Exit Sub
            Else
                If DateDiff("n", Format(isnow, "HH:NN"), Trim(text(Index))) > 0 Then
                    ' CARI DI STALL APAKAH BELUM TERPAKAI
                    rsGrid.Find "NOORDER = '" & Trim(text(0)) & "'", , adSearchForward, 1
                    If Not rsGrid.EOF Then
                        Set rsHlp = objO.FetchData("SELECT SUBSTRING(KETERANGAN, CHARINDEX(' |',KETERANGAN) + 2, CHARINDEX('| ',KETERANGAN) - (CHARINDEX(' |',KETERANGAN)+2)) AS NOORDER, * FROM STALL WITH(NOLOCK) WHERE NOSTALL = '" & Trim(cboStall) & "' AND TANGGAL = '" & nTgl & "' " _
                                                & "AND KETERANGAN NOT LIKE '%" & Trim(text(0)) & "%' AND (JAMMULAI BETWEEN '" & Trim(text(3)) & "' AND '" & Left(hitungWaktu(rsGrid("RATE"), Trim(text(3)), nTgl), 5) & "' " _
                                                & "OR JAMSELESAI BETWEEN '" & Trim(text(3)) & "' AND '" & Left(hitungWaktu(rsGrid("RATE"), Trim(text(3)), nTgl), 5) & "')")
                        If Not rsHlp.EOF Then
                            MsgBox "Digunakan untuk unit kendaraan lain. "
                            text(3) = ">>:>>"
                            text(3).SetFocus
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Case 4
            If Trim(text(Index)) <> "" Then
                Set rsSearch = objO.FetchData("select A.MECH,[NAME],[GROUP], LEVMECH, MENITKERJA, TERPAKAI from mekanik A WITH(NOLOCK) LEFT JOIN T_MEKANIK B WITH(NOLOCK) ON A.MECH = B.MECH where hari = '" & IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1) & "' AND TANGGAL = '" & Format(isnow, "DD/MM/YYYY") & "' AND (MENITKERJA-TERPAKAI) > 0 AND MASUK = 1 AND JAMMULAI <= '" & Format(isnow, "HH:NN") & "' AND JAMSELESAI >= '" & Format(isnow, "HH:NN") & "' AND A.MECH = '" & Trim(text(Index)) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gBantuan = "ORDER"
                    gBantuan2 = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
                    gBantuan3 = Format(isnow, "MM/DD/YYYY")
                    gSearchOption = "MEKANIK"
                    gBantuan4 = ""
                    frmCari.Show (vbModal)
                Else
                    If (rsSearch("MENITKERJA") - rsSearch("TERPAKAI")) <= 0 Then
                        MsgBox "Jam Kerja Mekanik " & rsSearch("NAME") & " tidak memenuhi." & vbCr & "Ganti dengan mekanik lain !", vbCritical, "Jam Kerja Kurang"
                        text(4) = ""
                        text(4).SelLength = Len(text(4))
                        text(4).SetFocus
                        Exit Sub
                    End If
                End If
            End If
    End Select
End Sub
