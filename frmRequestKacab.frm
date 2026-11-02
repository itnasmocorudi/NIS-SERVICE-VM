VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmRequestKacab 
   Caption         =   "Validasi Request Kacab"
   ClientHeight    =   6930
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11520
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   11520
   WindowState     =   2  'Maximized
   Begin VB.Frame frameReq 
      Height          =   6015
      Left            =   60
      TabIndex        =   4
      Top             =   120
      Width           =   14115
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   0
         Left            =   1250
         MaxLength       =   12
         TabIndex        =   43
         Top             =   960
         Width           =   3735
      End
      Begin VB.TextBox txt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   320
         Index           =   1
         Left            =   1250
         MaxLength       =   12
         TabIndex        =   42
         Top             =   1270
         Width           =   3735
      End
      Begin VB.Frame Frame1 
         Caption         =   "PERSETUJUAN DISKON"
         Height          =   1905
         Left            =   120
         TabIndex        =   31
         Top             =   2760
         Width           =   4950
         Begin VB.TextBox txtReqJasa 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1830
            MaxLength       =   2
            TabIndex        =   36
            Text            =   "99"
            Top             =   225
            Width           =   495
         End
         Begin VB.TextBox txtReqPart 
            Alignment       =   2  'Center
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4095
            MaxLength       =   2
            TabIndex        =   35
            Text            =   "99"
            Top             =   180
            Width           =   495
         End
         Begin VB.TextBox txtInfo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   990
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   975
            Width           =   3885
         End
         Begin VB.TextBox txtDisc 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4110
            MaxLength       =   2
            TabIndex        =   33
            Text            =   "99"
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtDiscJasa 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1830
            MaxLength       =   2
            TabIndex        =   32
            Text            =   "99"
            Top             =   570
            Width           =   495
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "REQ DISC JASA : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   180
            TabIndex        =   41
            Top             =   225
            Width           =   1635
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Caption         =   "REQ DISC PART : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2385
            TabIndex        =   40
            Top             =   210
            Width           =   1680
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "Note :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   225
            TabIndex        =   39
            Top             =   1065
            Width           =   720
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "DISC PART : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2520
            TabIndex        =   38
            Top             =   615
            Width           =   1530
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "DISC JASA : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   345
            TabIndex        =   37
            Top             =   660
            Width           =   1515
         End
      End
      Begin VB.CommandButton cmdBatal 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Batal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   4680
         Width           =   1890
      End
      Begin VB.CommandButton cmdValidasi 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Validasi"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4680
         Width           =   1890
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   0
         Left            =   1245
         TabIndex        =   8
         Top             =   675
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   1
         Left            =   1250
         TabIndex        =   9
         Top             =   990
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   1250
         TabIndex        =   10
         Top             =   1290
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   1245
         TabIndex        =   11
         Top             =   1590
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   1250
         TabIndex        =   12
         Top             =   1860
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   1250
         TabIndex        =   13
         Top             =   2160
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   1250
         TabIndex        =   14
         Top             =   2445
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin TrueOleDBGrid70.TDBGrid JobGrid 
         Height          =   1890
         Left            =   5040
         TabIndex        =   25
         Top             =   855
         Width           =   8865
         _ExtentX        =   15637
         _ExtentY        =   3334
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
         Columns(6).NumberFormat=   "Standard"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).DataField=   ""
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).DataField=   ""
         Columns(8).NumberFormat=   "Standard"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).DataField=   ""
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   4
         Columns(10)._MaxComboItems=   5
         Columns(10).DataField=   ""
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   11
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=11"
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
         Caption         =   "JOB ORDER"
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
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=32,.parent=13,.locked=0"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=29,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=30,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=31,.parent=17"
         _StyleDefs(80)  =   "Named:id=33:Normal"
         _StyleDefs(81)  =   ":id=33,.parent=0"
         _StyleDefs(82)  =   "Named:id=34:Heading"
         _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(84)  =   ":id=34,.wraptext=-1"
         _StyleDefs(85)  =   "Named:id=35:Footing"
         _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(87)  =   "Named:id=36:Selected"
         _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(89)  =   "Named:id=37:Caption"
         _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(91)  =   "Named:id=38:HighlightRow"
         _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(93)  =   "Named:id=39:EvenRow"
         _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(95)  =   "Named:id=40:OddRow"
         _StyleDefs(96)  =   ":id=40,.parent=33"
         _StyleDefs(97)  =   "Named:id=41:RecordSelector"
         _StyleDefs(98)  =   ":id=41,.parent=34"
         _StyleDefs(99)  =   "Named:id=42:FilterBar"
         _StyleDefs(100) =   ":id=42,.parent=33"
      End
      Begin TrueOleDBGrid70.TDBGrid partGrid 
         Height          =   1590
         Left            =   5070
         TabIndex        =   27
         Top             =   3105
         Width           =   8880
         _ExtentX        =   15663
         _ExtentY        =   2805
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
         Columns(3).NumberFormat=   "Standard"
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
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         Caption         =   "PART ORDER"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=900,.italic=0"
         _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
         _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
         _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
         _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
         _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
         _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
         _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
         _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
         _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
         _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
         _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
         _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
         _StyleDefs(70)  =   "Named:id=33:Normal"
         _StyleDefs(71)  =   ":id=33,.parent=0"
         _StyleDefs(72)  =   "Named:id=34:Heading"
         _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(74)  =   ":id=34,.wraptext=-1"
         _StyleDefs(75)  =   "Named:id=35:Footing"
         _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   "Named:id=36:Selected"
         _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(79)  =   "Named:id=37:Caption"
         _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(81)  =   "Named:id=38:HighlightRow"
         _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(83)  =   "Named:id=39:EvenRow"
         _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(85)  =   "Named:id=40:OddRow"
         _StyleDefs(86)  =   ":id=40,.parent=33"
         _StyleDefs(87)  =   "Named:id=41:RecordSelector"
         _StyleDefs(88)  =   ":id=41,.parent=34"
         _StyleDefs(89)  =   "Named:id=42:FilterBar"
         _StyleDefs(90)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lblPDS 
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3660
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label13 
         Caption         =   "* ITEM JOB ORDER DAN PART BISA JADI MASIH ADA PERUBAHAN"
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
         Height          =   510
         Left            =   210
         TabIndex        =   24
         Top             =   5175
         Width           =   8730
      End
      Begin VB.Label Label12 
         Caption         =   "DAFTAR PART"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5070
         TabIndex        =   23
         Top             =   2820
         Width           =   1605
      End
      Begin VB.Label Label11 
         Caption         =   "DAFTAR JASA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5070
         TabIndex        =   22
         Top             =   525
         Width           =   1605
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MoU"
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
         Left            =   855
         TabIndex        =   19
         Top             =   2505
         Width           =   315
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Unit"
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
         Left            =   420
         TabIndex        =   18
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label Label6 
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
         Left            =   555
         TabIndex        =   17
         Top             =   1875
         Width           =   585
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telp"
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
         Left            =   855
         TabIndex        =   16
         Top             =   1620
         Width           =   300
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
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
         Left            =   675
         TabIndex        =   15
         Top             =   1335
         Width           =   495
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Wo"
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
         Left            =   720
         TabIndex        =   7
         Top             =   750
         Width           =   480
      End
      Begin VB.Label Label24 
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
         Left            =   510
         TabIndex        =   6
         Top             =   1050
         Width           =   690
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: DAFTAR REQUEST DISC KACAB ::."
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
         Left            =   120
         TabIndex        =   5
         Top             =   210
         Width           =   13755
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   45
         Shape           =   4  'Rounded Rectangle
         Top             =   135
         Width           =   13935
      End
   End
   Begin VB.CommandButton tblRefresh 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   10335
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5640
      Width           =   1890
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
      Height          =   390
      Left            =   12240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5640
      Width           =   1890
   End
   Begin TrueOleDBGrid70.TDBGrid gridData 
      Height          =   4980
      Left            =   105
      TabIndex        =   0
      Top             =   600
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   8784
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "No WO"
      Columns(0).DataField=   "NO_WO"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Tanggal Request"
      Columns(1).DataField=   "tglReq"
      Columns(1).NumberFormat=   "Short Date"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nama Customer"
      Columns(2).DataField=   "Customer"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "SA"
      Columns(3).DataField=   "front"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Req Jasa"
      Columns(4).DataField=   "jasa"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Req Part"
      Columns(5).DataField=   "part"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).Locked=   -1  'True
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
      Splits(0)._ColumnProps(9)=   "Column(2).Width=5292"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=5186"
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
      Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
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
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
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
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=32,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=29,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=30,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=31,.parent=17"
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
      _StyleDefs(85)  =   "Named:id=87:SudahDitagih"
      _StyleDefs(86)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(87)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(88)  =   ":id=87,.fontname=Tahoma"
   End
   Begin VB.Label Label15 
      Caption         =   "* KLIK 2X UNTUK VALIDASI"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   345
      Left            =   7860
      TabIndex        =   30
      Top             =   5685
      Width           =   2400
   End
   Begin VB.Label Label1 
      BackColor       =   &H000080FF&
      Caption         =   "TOTAL DATA = "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11685
      TabIndex        =   29
      Top             =   165
      Width           =   1410
   End
   Begin VB.Label lblJumlah 
      BackColor       =   &H000080FF&
      Caption         =   "Label2"
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
      Height          =   300
      Left            =   13125
      TabIndex        =   28
      Top             =   165
      Width           =   990
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: DAFTAR REQUEST DISC KACAB ::."
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
      Left            =   5235
      TabIndex        =   1
      Top             =   195
      Width           =   3840
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   105
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   14055
   End
End
Attribute VB_Name = "frmRequestKacab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private rsData As New ADODB.Recordset
Private rsJob As New ADODB.Recordset
Private rsPart As New ADODB.Recordset
Private cols As TrueOleDBGrid70.Columns
Private rsTemp As ADODB.Recordset
Private nilaiDiscJasa As Integer
Private nilaiDiscPart As Integer
Private objBilling As New BClsBilling


Private Sub doCreatePart()
1        On Error GoTo doCreatePart_Error
2         If rsPart.State Then rsPart.Close
          'Set rsPart = Nothing
3         With rsPart
             ' If .State Then .Close
4             .Fields.Append "KODEITEM", adVarChar, 50
5             .Fields.Append "NAMAITEM", adVarChar, 200
6             .Fields.Append "QTYOrder", adInteger
7             .Fields.Append "QTY", adInteger
8             .Fields.Append "QTYBO", adInteger
9             .Fields.Append "QTYCANCEL", adInteger
10            .Fields.Append "HARGA", adDouble
11            .Fields.Append "OLI", adBoolean
12            .Fields.Append "EXTDISC", adDouble
13            .Open
14        End With
15        Call doHeadPart
          
16       On Error GoTo 0
17       Exit Sub

doCreatePart_Error:

18        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreatePart of Form frmOrderValidasi")
End Sub

Private Sub doHeadPart()
1        On Error GoTo doHeadPart_Error

2         With partGrid
3             .Columns(0).Width = 2000
4             .Columns(0).Caption = "KODE PART"
5             .Columns(0).DataField = "KODEITEM"
6             .Splits(0).Columns(0).Locked = True
              
7             .Columns(1).Width = 2500
8             .Columns(1).Caption = "NAMA PART"
9             .Columns(1).DataField = "NAMAITEM"
10            .Splits(0).Columns(1).Locked = True
              
11            .Columns(2).Width = 1000
12            .Columns(2).Caption = "QTY Order"
13            .Columns(2).DataField = "QTYOrder"
14            .Splits(0).Columns(2).Locked = True
              
15            .Columns(3).Width = 1000
16            .Columns(3).Caption = "QTY Supply"
17            .Columns(3).DataField = "QTY"
18            .Splits(0).Columns(3).Locked = True
              
19            .Columns(4).Width = 1000
20            .Columns(4).Caption = "QTY BO"
21            .Columns(4).DataField = "QTYBO"
22            .Splits(0).Columns(3).Locked = True
23            On Error GoTo Column_Ini_Untuk_Grid_FillRate
              
              
24            .Columns(5).Visible = True
25            .Columns(5).Width = 1000
26            .Columns(5).Caption = "HARGA"
27            .Columns(5).DataField = "Harga"
28            .Splits(0).Columns(5).Locked = True
              
29            .Columns(6).Visible = False
30            .Columns(6).Caption = "OLI"
31            .Columns(6).DataField = "OLI"
32            .Splits(0).Columns(6).Locked = True
              
33            .Columns(7).Width = 1000
34            .Columns(7).Caption = "QTY CANCEL"
35            .Columns(7).DataField = "QTYCANCEL"
36            .Splits(0).Columns(7).Locked = True
Column_Ini_Untuk_Grid_FillRate:
37        End With

38       On Error GoTo 0
39       Exit Sub

doHeadPart_Error:

40        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadPart of Form frmOrderValidasi")
End Sub


Private Sub doCreateJob()
1        On Error GoTo doCreateJob_Error
2         If rsJob.State Then rsJob.Close
          Set rsJob = Nothing
3         With rsJob
              'If .State Then .Close
4             .Fields.Append "JOB", adVarChar, 50
5             .Fields.Append "RATE", adDouble
6             .Fields.Append "ACTRT", adDouble
7             .Fields.Append "GRP", adVarChar, 6
8             .Fields.Append "MECH", adVarChar, 10
9             .Fields.Append "ADDJOB", adBoolean
10            .Fields.Append "HRG", adDouble
11            .Fields.Append "DISC", adDouble
12            .Fields.Append "DISCRP", adDouble
13            .Fields.Append "STLDISC", adDouble
14            .Fields.Append "VALID", adVarChar, 10
15            .Fields.Append "FREEC", adBoolean           ' TANDA JIKA DIBEBANKAN KE CUSTOMER
16            .Open
17        End With
          
18        Call doHeadJob
19
20       On Error GoTo 0
21       Exit Sub

doCreateJob_Error:

22        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreateJob of Form frmOrderValidasi")
End Sub


Private Sub doHeadJob()
1        On Error GoTo doHeadJob_Error

2         With JobGrid
3             .Columns(0).Width = 3000
4             .Columns(0).Caption = "JOB ORDER"
5             .Columns(0).DataField = "JOB"
6             .Splits(0).Columns(0).Locked = True
              
7             .Columns(1).Width = 800
8             .Columns(1).Caption = "RATE"
9             .Columns(1).DataField = "RATE"
10            .Splits(0).Columns(1).Locked = True
              
11            .Columns(2).Width = 0
12            .Columns(2).Caption = "ACTRATE"
13            .Columns(2).DataField = "ACTRT"
14            .Columns(2).Visible = False
15            .Splits(0).Columns(2).Locked = True
              
16            .Columns(3).Width = 0
17            .Columns(3).Caption = "GROUP"
18            .Columns(3).DataField = "GRP"
19            .Columns(3).Visible = False
20            .Splits(0).Columns(3).Locked = True
              
21            .Columns(4).Width = 0
22            .Columns(4).Caption = "MEKANIK"
23            .Columns(4).DataField = "MECH"
24            .Columns(4).Visible = False
25            .Splits(0).Columns(4).Locked = True
              
26            .Columns(5).Width = 0
27            .Columns(5).Caption = "Additional Job"
28            .Columns(5).DataField = "ADDJOB"
29            .Columns(5).Visible = False
30            .Splits(0).Columns(5).Locked = True
              
31            .Columns(6).Width = 1200
32            .Columns(6).Caption = "HARGA"
33            .Columns(6).DataField = "HRG"
34            .Splits(0).Columns(6).Locked = True
                     
35            .Columns(7).Width = 600
36            .Columns(7).Caption = "DISC"
37            .Columns(7).DataField = "DISC"
              
38            .Columns(8).Width = 1200
39            .Columns(8).Caption = "DISC RUPIAH"
40            .Columns(8).DataField = "DISCRP"
              
41            .Columns(9).Width = 1200
42            .Columns(9).Caption = "SUB TOTAL"
43            .Columns(9).DataField = "STLDISC"
44            .Splits(0).Columns(9).Locked = True
              
45            .Columns(10).Width = 800
46            .Columns(10).Caption = "Customer"
47            .Columns(10).DataField = "FREEC"
48            .Columns(10).Alignment = dbgCenter
49            .Splits(0).Columns(10).Locked = False
50            .Columns(10).Visible = False
51            'If nDuoInv Then
52            '    .Columns(10).Visible = True
53            'End If
54        End With

55       On Error GoTo 0
56       Exit Sub

doHeadJob_Error:

57        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadJob of Form frmOrderValidasi")
End Sub


Private Sub refreshGrid()
    sSQL = "SELECT a.No_Wo,a.tglReq,b.CUSTOMER,b.FRONT,convert(varchar(10),cast(a.DiscJasaReq as int)) + ' %' as jasa,convert(varchar(10),cast(a.DiscPartReq as int)) + ' %' as part FROM RequestDisc a WITH (NOLOCK) INNER JOIN JOBDTA b WITH (NOLOCK) ON a.No_Wo = b.NO_WO AND a.No_Chassis = b.NO_CHASSIS WHERE ISNULL(TGLVALID,'') = '' "
    Set rsData = objO.FetchData(sSQL)
    gridData.DataSource = rsData
    LblJumlah.Caption = rsData.RecordCount
    
End Sub

Private Sub cmdBatal_Click()
    frameReq.Visible = False
    Call refreshGrid
End Sub

Private Sub cmdValidasi_Click()
    If MsgBox("Apakah yakin validasi disc untuk no wo " & Trim(Text(0)), vbYesNo) = vbYes Then
        sSQL = "UPDATE RequestDisc SET discjasavalid = " & Trim(txtDiscJasa.Text) & " , discpartvalid = " & Trim(txtDisc.Text) & ",infokcb = '" & Trim(txtInfo.Text) & "',tglValid = getdate(),uservalid = '" & gSA & "' WHERE no_wo = '" & Trim(Text(0)) & "'"
        Call objO.InsertData(sSQL)
        MsgBox "Data berhasil divalidasi", vbInformation
        frameReq.Visible = False
        Call refreshGrid
    End If
End Sub

Private Sub Form_Load()
    frameReq.Visible = False
    Call refreshGrid
End Sub

Private Function typeJob(statusJob As String) As String
1        On Error GoTo typeJob_Error

2         If statusJob = "W" Then
3             typeJob = "WARRANTY"
4         ElseIf statusJob = "P" Then
5             typeJob = "PROGRAM"
6         Else
7             typeJob = "REGULAR"
8         End If

9        On Error GoTo 0
10       Exit Function

typeJob_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure typeJob of Form frmOrderValidasi")
End Function

Private Sub gridData_DblClick()
          Dim rsHlp As ADODB.Recordset
          Dim rsHlp2 As ADODB.Recordset
          Dim rsHlp3 As ADODB.Recordset
          Dim fOth As Boolean
          Dim nHarga As Double
1         lblPDS.Visible = False
2         If rsData.State <> adStateClosed Then
3             If Not rsData.EOF Then
                  
4                 If Log.RPLHistori = 0 Then
5                     sSQL = "SELECT b.fs_kd_plg, a.No_Wo,b.CUSTOMER,b.CUST_ADRES,b.TELPON,b.NO_POLISI,b.FMDL,ISNULL(c.PROGRAM,'') MOU,a.DiscJasaReq as jasa, a.DiscPartReq discPart,B.* FROM RequestDisc a WITH (NOLOCK) INNER JOIN JOBDTA b WITH (NOLOCK) ON a.No_Wo = b.NO_WO AND a.No_Chassis = b.NO_CHASSIS LEFT JOIN HO_PROGRAM c WITH (NOLOCK) ON c.id_Program = b.KODETIPE WHERE a.No_Wo = '" & rsData!no_wo & "'"
6                 Else
7                     sSQL = "SELECT b.fs_kd_plg, a.No_Wo,b.CUSTOMER,b.CUST_ADRES,b.TELPON,b.NO_POLISI,b.FMDL,ISNULL(c.PROGRAM,'') MOU,a.DiscJasaReq as jasa, a.DiscPartReq discPart,B.* FROM RequestDisc a WITH (NOLOCK) INNER JOIN JOBDTA b WITH (NOLOCK) ON a.No_Wo = b.NO_WO AND a.No_Chassis = b.NO_CHASSIS LEFT JOIN [" & dbReplikasiHO & "]..HO_PROGRAM c WITH (NOLOCK) ON c.FS_LOG_DEL = '' AND c.id_Program = b.KODETIPE WHERE a.No_Wo = '" & rsData!no_wo & "'"
8                 End If
9                 Set rsTemp = objO.FetchData(sSQL)
10                If Not rsTemp.EOF Then
11                    frameReq.Visible = True
                      
                      ''2017-08-01 Revisi untuk membaca nilai part
                ''--------------------------------------------------------------
                Dim noInv As Boolean
12              noInv = True
13              If rsTemp("NOINV") <> "" And rsTemp("RPR_COND") <> "K" Then
14                noInv = False
15              ElseIf rsTemp("NOINV") <> "" And rsTemp("RPR_COND") = "K" Then
16                noInv = False
17              ElseIf rsTemp("RPR_COND") = "K" Then
18                noInv = False
19              End If
                ''--------------------------------------------------------------
                
                      
20                    Call doCreateJob
21                    Call doCreatePart
                      
22                    Text(0) = Nul(rsTemp!no_wo)
23                    txt(0) = Nul(rsTemp!customer) 'text(1) = Nul(rsTemp!CUSTOMER)
24                    txt(1) = Nul(rsTemp!cust_adres)
25                    Text(2) = Nul(rsTemp!FS_KD_PLG)
26                    Text(3) = Nul(rsTemp!Telpon)
27                    Text(4) = Nul(rsTemp!NO_POLISI)
28                    Text(5) = Nul(rsTemp!Fmdl)
29                    Text(6) = Nul(rsTemp!MOU)
                      
                      ''2017-08-01 Penambahan nilai request sa
                ''----------------------------------------------
30              txtReqJasa.Text = CInt(Nul(rsTemp!jasa))
31              txtReqPart.Text = CInt(Nul(rsTemp!discPart))
                ''----------------------------------------------
                      
32                    txtDiscJasa.Text = CInt(Nul(rsTemp!jasa))
33                    txtDisc.Text = CInt(Nul(rsTemp!discPart))
                      
34                    nilaiDiscJasa = CInt(Nul(rsTemp!jasa))
35                    nilaiDiscPart = CInt(Nul(rsTemp!discPart))
                      
36                    lblPDS.Visible = IIf(rsTemp("PDS") = "1", True, False)
                      
                      Dim nilaiRate As Double
37                    For i = 1 To 7
38                        hlp = "REP" & i
39                        hlp2 = "RTE" & i
40                        hlp3 = "KLPKERJA" & i
41                        hlp4 = "ACTRATE" & i
42                        hlp5 = "MEKANIK" & i
43                        If rsTemp(hlp) <> "" Then
44                            fOth = False
45                            Set rsHlp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "' AND MDL = '" & rsTemp("MDL") & "'")
46                            If rsHlp.EOF Then
47                                fOth = True
48                            End If
49                            If rsHlp.State <> adStateClosed Then rsHlp.Close
                               
50                            If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
51                                nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
52                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
53                            ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
54                                If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
55                                    If rsTemp("FS_RC") = "R001" Then
56                                        If Trim(rsTemp(hlp3)) = "SBI" Then
                                               
57                                            Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
58                                            If Not rsHlp2.EOF Then
59                                                nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
60                                                If nStatusDealer = 1 Then
61                                                    nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
62                                                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
63                                                Else
64                                                    If nStatusDealer = 2 Then
65                                                        nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
66                                                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
67                                                    ElseIf nStatusDealer = 3 Then
68                                                        nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
69                                                        If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
70                                                    End If
71                                                End If
72                                            End If
73                                            If rsHlp2.State <> adStateClosed Then rsHlp2.Close
74                                        Else
75                                            nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
76                                            If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
77                                        End If
78                                    Else
                                           
79                                        Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
80                                        If Not rsHlp2.EOF Then
                                              'Dim nStatusDealer As Integer
81                                            nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
82                                            If nStatusDealer = 1 Then
83                                                nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0  '
84                                                If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
85                                            Else
86                                                If nStatusDealer = 2 Then
87                                                    nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
88                                                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
89                                                ElseIf nStatusDealer = 3 Then
90                                                    nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
91                                                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
92                                                End If
93                                            End If
                                              'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
94                                        End If
95                                        If rsHlp2.State <> adStateClosed Then rsHlp2.Close
96                                    End If
97                                Else
98                                    nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
99                                    If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
100                               End If
101                           ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
102                               If lblPDS.Visible Then
103                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
104                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
105                               ElseIf rsTemp("FS_RC") = "R001" Then
                                       
106                                   Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
107                                   If Not rsHlp2.EOF Then
108                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
109                                       If nStatusDealer = 1 Then
110                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
111                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
112                                       Else
113                                           If nStatusDealer = 2 Then
114                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
115                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
116                                           ElseIf nStatusDealer = 3 Then
117                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
118                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
119                                           End If
120                                       End If
                                          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
121                                   End If
122                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
123                               ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
124                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
125                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
126                               Else
127                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
128                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
129                               End If
130                           End If
                              
131                           nilaiRate = CDbl(rsTemp(hlp2))
                          
132                           With rsJob
133                               If .State = adStateClosed Then doCreateJob
134                              .AddNew
135                              !job = Trim(rsTemp(hlp))
136                              !Rate = nilaiRate 'CDbl(rsTemp(hlp2))
137                              !ACTRT = CDbl(rsTemp(hlp4))
138                              !grp = Trim(rsTemp(hlp3))
139                              !Mech = Trim(rsTemp(hlp5))
140                              !hrg = CDbl(nHarga) * nilaiRate
141                              !addJob = False
142                              !disc = 0
143                              !DISCRP = CDbl(!hrg) * (CDbl(0) / 100)
144                              !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
145                              !valid = gVerifikator
146                              !FREEC = False
147                              .Update
148                           End With
149                       End If
150                   Next
151                   Set JobGrid.DataSource = rsJob
152                   JobGrid.Refresh
                      ''XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                      ''PART
                      ''xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
153                   gSpart = True
                      Dim HargaDisc As Double
154                   If rsPart.State Then rsPart.Close
                      
                      ' /// AMBIL PART YG BAWA SENDIRI //
                      
155                   Set rsHlp = objO.FetchData("SELECT KODE_ITEM, NAMA_ITEM, PARTBAWA, QTY, isnull(fs_kd_oli,'') as OLI, TIPE, GENUINE, BAWA FROM jobpart a WITH (NOLOCK) LEFT JOIN T_oli b WITH (NOLOCK) ON LEFT(a.kode_item,1) = b.fs_kd_Oli WHERE no_WO = '" & noWO & "' " _
                                              & "AND no_Chassis = '" & rsTemp("NO_CHASSIS") & "' AND GANTI = 'Y' ") 'AND BAWA <> 'Y'
156                   If Not rsHlp.EOF Then
157                       rsHlp.MoveFirst
158                       Do While Not rsHlp.EOF
159                           With rsPart
160                           If .State = adStateClosed Then doCreatePart
161                               gSpart = True
                                  
162                               If rsHlp("BAWA") = "Y" Then
163                                   .AddNew
164                                   !KodeItem = "BAWA SENDIRI"
165                                   !NAMAITEM = rsHlp("PARTBAWA")
166                                   !qty = rsHlp("QTY")
167                                   !QtyOrder = rsHlp("QTY")
168                                   !QtyBO = 0
169                                   !QtyCancel = 0
170                                   !Harga = 0
171                                   !oli = IIf(rsHlp("TIPE") <> "P", True, False)
172                                   .Update
173                               Else
174                                   If Not gFPart Then
                                          '///// TIDAK CONNECT ACCOUNTING SPAREPART ////
175                                       .AddNew
176                                       !KodeItem = IIf(rsHlp("BAWA") = "Y", "BAWA SENDIRI", rsHlp("KODE_ITEM"))
177                                       !NAMAITEM = IIf(rsHlp("BAWA") = "Y", rsHlp("PARTBAWA"), rsHlp("NAMA_ITEM"))
178                                       !oli = IIf(rsHlp("TIPE") <> "P", True, False)
179                                       !qty = rsHlp("QTY")
                                              'Set rsHlp2 = objO.FetchDataACC("SELECT hargajualA FROM inmst WHERE kodeItem = '" & rsHlp("KODE_ITEM") & "'")
180                                           Set rsHlp2 = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & rsHlp("KODE_ITEM") & "'")
181                                           If Not rsHlp2.EOF Then
182                                               HargaDisc = CDbl(rsHlp2("HARGAJUALA"))
183                                               !Harga = CDbl(HargaDisc)
184                                               If rsHlp("OLI") <> "" Then
185                                                   If rsHlp("TIPE") = "P" Then
186                                                       If rsHlp2("TGenuine") = False Then
                                                              ' PART GENUINE
187                                                           subPartNGenuine = subPartNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
188                                                       ElseIf rsHlp2("TGenuine") = True Then
                                                              ' PART NON GENUINE
189                                                           subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
190                                                       End If
                                                         ' text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
191                                                   ElseIf rsHlp("TIPE") = "O" Then
192                                                       If rsHlp2("TGenuine") = False Then
                                                              ' OLI GENUINE
193                                                           subOliNGenuine = subOliNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
194                                                       ElseIf rsHlp2("TGenuine") = True Then
                                                              ' OLI NON GENUINE
195                                                           subOliGenuine = subOliGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
196                                                       End If
                                                         ' text(25) = text(25) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
197                                                   End If
198                                               Else
                                                      ' PART GENUINE
199                                                   subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
                                                      'text(24) = text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
200                                               End If
201                                           End If
202                                           If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                                          
203                                       .Update
204                                   End If
205                               End If
206                           End With
207                           rsHlp.MoveNext
208                       Loop
209                   End If
210                   If rsHlp.State <> adStateClosed Then rsHlp.Close
                      
211                   If gFPart Then
                          '/// KONEKSI KE SPAREPART
212                       If Get_ConnectionNisSP(NisSPDb, NisSPUSER, NisSPPass, NisSPKomp) Then
213                           With rsPart
214                               If noInv Then
                                      Dim strSql As String
215                                   If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
216                                       strSql = "select (case When KodeItem = '' Then namaBYSA Else KodeItem end) as KodeItem, " & _
                                               "Keterangan as Namabarang, SatuanNs as Satuan, Qty as QtyOrder, (isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0)) AS Jml, " & _
                                               "(Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) as QtyBO, HargaJual as HargaJualA, NOSO as NoRack, " & _
                                               "0 as DisItem, 0 as ExtDis, Case BO When 'CANCEL' THEN (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) ELSE 0 END AS QTYCANCEL from SOORD WITH (NOLOCK) Where NOSO = '" & Trim(Text(0)) & "'"
217                                   Else
218                                       strSql = "Select Z.*, 0 as QtyOrder, 0 as QtyBO FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                              "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                              "(INMST.HARGAJUALA ) as HARGAJUALA, " & _
                                              "INMST.NORACK,0 as DisItem, 0 as ExtDis, 0 as QtyCANCEL FROM INMST " & _
                                              "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                              "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                              "Where Z.Norack = '" & Trim(Text(0)) & "'"
219                                   End If
220                                   Set rsHlp3 = objO.FetchDataACC(strSql)
              
221                                   If Not rsHlp3.EOF Then
222                                       Do While Not rsHlp3.EOF
223                                           If .State = adStateClosed Then doCreatePart
224                                           .AddNew
225                                           !KodeItem = Trim(rsHlp3("KODEITEM"))
226                                           !NAMAITEM = Trim(rsHlp3("NAMABARANG"))
227                                           !QtyOrder = CDbl(rsHlp3("QtyOrder"))
228                                           !qty = CDbl(rsHlp3("JML"))
229                                            If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
230                                               !QtyBO = IIf(UCase(isBO) = "CANCEL", 0, objO.FetchData_DoubleACC("select Qty FROM BOFILE WITH (NOLOCK) Where TutupOrderFlag = 0 and (NoSO = '" & Trim(Text(0)) & "' Or NOSO =(Select Distinct NoSPK From SOORD WITH (NOLOCK) Where NOSO ='" & Trim(Text(0)) & "' and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "')) and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "'"))
231                                            Else
232                                               !QtyBO = 0
233                                            End If
                                              
234                                           !QtyCancel = CDbl(rsHlp3("QtyCancel"))
235                                           !Harga = CDbl(rsHlp3("HARGAJUALA"))
           
236                                               Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp3("KODEITEM") & "'")
237                                               If Not rsHlp.EOF Then
238                                                   If Left(rsHlp("SCC"), 3) = "OLI" Then
239                                                       !oli = True
                                                         ' text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
240                                                   Else
241                                                       !oli = False
                                                         ' text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
242                                                   End If
243                                               Else
244                                                   !oli = False
                                                      'text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
245                                               End If
246                                           .Update
247                                           rsHlp3.MoveNext
248                                       Loop
249                                   Else 'If rsHlp3.EOF And newValid Then
                                          Dim rsInvPart As New ADODB.Recordset
                                          Dim vHitung As Double
250                                       Set rsInvPart = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                                                           & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(Text(0)) & "' and KodeLgn = '" & Text(2) & "' ")
251                                       If Not rsInvPart.EOF Then
252                                           Do While Not rsInvPart.EOF
253                                               If .State = adStateClosed Then doCreatePart
254                                               .AddNew
255                                               !KodeItem = Trim(rsInvPart("KODEITEM"))
256                                               !NAMAITEM = Trim(rsInvPart("KETERANGAN"))
257                                               vHitung = (Koma2Titik(rsInvPart("Qty")) * (Koma2Titik(rsInvPart("HargaJual")) - (Koma2Titik(rsInvPart("ItemDisc")) / Abs(Koma2Titik(rsInvPart("Qty"))))))
258                                               vHitung = vHitung - (Koma2Titik(rsInvPart("ExtDisPersen")) * vHitung)
259                                               If Trim(UCase(rsInvPart("TypePpn"))) = "E" Then
260                                                   vHitung = Round(vHitung + vHitung * (profilACC.nilaiPPn / 100), 2)
261                                               End If
262                                               If gJurnalPart Then
263                                                   vHitung = cekIncludePPn(vHitung)
264                                               End If
265                                               !qty = Val(rsInvPart("QTY"))
266                                               If !qty = 0 Then
267                                                   !Harga = 0
268                                               Else
269                                                   !Harga = CDbl(vHitung / !qty)
270                                               End If
271                                               !QtyOrder = Val(rsInvPart("QTY"))
272                                               !QtyBO = 0
273                                               !QtyCancel = 0
                                                   
274                                               Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsInvPart("KODEITEM") & "'")
275                                               If Not rsHlp.EOF Then
276                                                   If Left(rsHlp("SCC"), 3) = "OLI" Then
277                                                       !oli = True
                                                          'text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
278                                                   Else
279                                                       !oli = False
                                                         ' text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
280                                                   End If
281                                               Else
282                                                   !oli = False
                                                    '  text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
283                                               End If
284                                               .Update
285                                               rsInvPart.MoveNext
286                                           Loop
287                                       Else
288                                           gSpart = False
289                                       End If
290                                       If rsInvPart.State <> adStateClosed Then rsInvPart.Close
291                                   End If
292                                   If rsHlp3.State <> adStateClosed Then rsHlp3.Close
293                               Else 'Not newValid Then
                                       
294                                   Set rsHlp2 = objO.FetchData("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsTemp!no_wo) & "' AND NOINV LIKE '" & Trim(rsTemp!noInv) & "%'")
295                                   If Not rsHlp2.EOF Then
296                                       Do While Not rsHlp2.EOF
297                                           If .State = adStateClosed Then doCreatePart
298                                           .AddNew
299                                           !KodeItem = Trim(rsHlp2("KODEITEM"))
300                                           !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
301                                           !qty = CDbl(rsHlp2("QTY"))
302                                           !Harga = CDbl(rsHlp2("HARGA")) '* ((100 - CDbl(rsHlp2("FN_EXT_DISKON"))) / 100)
303                                           !oli = Trim(rsHlp2("OLI")) 'IIf(Trim(rsHlp2("OLI")) = 1, True, False)
304                                           If !oli Then
                                                 ' text(25) = text(25) + CDbl(!Harga)
305                                           Else
                                                  'text(24) = text(24) + CDbl(!Harga)
306                                           End If
307                                           !EXTDISC = CDbl(rsHlp2("FN_EXT_DISKON"))
308                                           !QtyOrder = CDbl(rsHlp2("QTY"))
309                                           !QtyBO = 0
310                                           !QtyCancel = 0
311                                           .Update
312                                           rsHlp2.MoveNext
313                                       Loop
314                                   Else
315                                       gSpart = False
316                                   End If
317                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
318                               End If
319                           End With
320                       End If
                      
321                       If Not gSpart Then
322                               With rsPart
323                                   Set rsHlp3 = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
                                                                     & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(Text(0)) & "' and KodeLgn = '" & Text(2) & "' ")
                                      
324                                   If Not rsHlp3.EOF Then
325                                       Do While Not rsHlp3.EOF
326                                           If .State = adStateClosed Then doCreatePart
327                                           .AddNew
328                                           !KodeItem = Trim(rsHlp3("KODEITEM"))
329                                           !NAMAITEM = Trim(rsHlp3("KETERANGAN"))
330                                           vHitung = (Koma2Titik(rsHlp3("Qty")) * (Koma2Titik(rsHlp3("HargaJual")) - (Koma2Titik(rsHlp3("ItemDisc")) / Abs(Koma2Titik(rsHlp3("Qty"))))))
331                                           vHitung = vHitung - (Koma2Titik(rsHlp3("ExtDisPersen")) * vHitung)
332                                           If Trim(UCase(rsHlp3("TypePpn"))) = "E" Then
333                                               vHitung = Round(vHitung + vHitung * (profilACC.nilaiPPn / 100), 2)
334                                           End If
335                                           If gJurnalPart Then
336                                               vHitung = cekIncludePPn(vHitung)
337                                           End If
338                                           !qty = CDbl(rsHlp3("Qty"))
339                                           !Harga = CDbl(vHitung / !qty)
340                                           Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp3("KODEITEM"), 1) & "'")
341                                           If Not rsHlp.EOF Then
342                                               If rsHlp("TIPE") = "O" Then
343                                                   !oli = True
                                                     ' text(25) = text(25) + (CDbl(!Harga) * CDbl(!qty))
344                                               ElseIf rsHlp("TIPE") = "P" Then
345                                                   !oli = False
                                                     ' text(24) = text(24) + (CDbl(!Harga) * CDbl(!qty))
346                                               End If
347                                           End If
348                                           !QtyOrder = CDbl(rsHlp3("Qty"))
349                                           !QtyBO = 0
350                                           !QtyCancel = 0
351                                           .Update
352                                           rsHlp3.MoveNext
353                                       Loop
354                                       gSpart = True
                                      
355                                   End If
356                                   If rsHlp3.State <> adStateClosed Then rsHlp3.Close
357                               End With
358                           Else
359                               gSpart = True
360                           End If
361                       End If
362                   End If
                      ''__________________________________
363                   Set partGrid.DataSource = rsPart
364                   partGrid.Refresh
                      
365               End If
366           End If
            
          
End Sub

Private Sub gridData_FilterChange()
    On Error GoTo errHandler
    Set cols = gridData.Columns
    Dim C As Integer
    
    C = gridData.col
    gridData.HoldFields
    rsData.filter = getFilter()
    If rsData.EOF Then
        For Each col In gridData.Columns
            col.FilterText = ""
        Next col
        rsRecord.filter = ""
    End If
    gridData.col = C
    gridData.EditActive = True
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

Private Sub gridData_HeadClick(ByVal ColIndex As Integer)
On Error GoTo err
   Urut = IIf(Urut = "DESC", "ASC", "DESC")
   If ColIndex <= 12 Then rsRecord.Sort = gridData.Columns(ColIndex).DataField & " " & Urut
   gridData.SetFocus
   Sendkeys "{Right " & Val(ColIndex) & "}"
Exit Sub
err:
   MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblRefresh_Click()
    Call refreshGrid
End Sub

Private Sub txtDisc_KeyPress(KeyAscii As Integer)
    KeyAscii = CekNumerik(KeyAscii)
End Sub

Private Sub txtDisc_LostFocus()
    If CInt(txtDisc.Text) > Log.discPartMax Then
        MsgBox "Nilai disc kacab melebihi max disc sebesar " & Log.discPartMax & " %", vbInformation
        txtDisc.Text = nilaiDiscPart
    End If
End Sub

Private Sub txtDiscJasa_KeyPress(KeyAscii As Integer)
    KeyAscii = CekNumerik(KeyAscii)
End Sub

Private Sub txtDiscJasa_LostFocus()
    If CInt(txtDiscJasa.Text) > Log.discMaxJasa Then
        MsgBox "Nilai disc kacab melebihi max disc sebesar " & Log.discMaxJasa & " %", vbInformation
        txtDiscJasa.Text = nilaiDiscJasa
    End If
End Sub
