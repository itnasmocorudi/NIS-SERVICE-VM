VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmtActRate 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Master Actual Rate"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7710
   ScaleWidth      =   11850
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   480
      TabIndex        =   9
      Top             =   480
      Width           =   11295
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5640
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   900
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   2760
         TabIndex        =   32
         Top             =   915
         Visible         =   0   'False
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Frame frmeCopyJob 
         BackColor       =   &H00FF0000&
         Caption         =   "             Copy Job In Group                "
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   3480
         TabIndex        =   26
         Top             =   3120
         Visible         =   0   'False
         Width           =   4695
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   10
            Left            =   105
            TabIndex        =   27
            Top             =   720
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   503
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
            Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Visible         =   0   'False
            Width           =   4425
            _ExtentX        =   7805
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            Enabled         =   0   'False
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
         Begin AIFCmp1.asxPowerButton cmdProses 
            Height          =   360
            Left            =   720
            TabIndex        =   30
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            BackColor       =   14482683
            BorderStyle     =   3
            Caption         =   "Pr&oses"
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
         Begin AIFCmp1.asxPowerButton cmdTutup 
            Height          =   360
            Left            =   2520
            TabIndex        =   31
            Top             =   1560
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   635
            BackColor       =   14482683
            BorderStyle     =   3
            Caption         =   "Tutu&p"
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
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   1920
            Top             =   720
            Width           =   60
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Group Model"
            ForeColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   405
            Width           =   915
         End
      End
      Begin ucXPButton.XPButton cmdDisplay 
         Height          =   855
         Left            =   9720
         TabIndex        =   24
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "Display All"
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
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmActRate.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.CheckBox chOther 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Other"
         Height          =   255
         Left            =   6840
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   2760
         TabIndex        =   1
         Top             =   600
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   2760
         TabIndex        =   3
         Top             =   915
         Width           =   3860
         _ExtentX        =   6800
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   2760
         TabIndex        =   7
         Top             =   1860
         Width           =   3945
         _ExtentX        =   6959
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
      Begin TrueOleDBGrid70.TDBGrid GridActRate 
         Height          =   4335
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   1
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=1"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(55)  =   "Named:id=40:OddRow"
         _StyleDefs(56)  =   ":id=40,.parent=33"
         _StyleDefs(57)  =   "Named:id=41:RecordSelector"
         _StyleDefs(58)  =   ":id=41,.parent=34"
         _StyleDefs(59)  =   "Named:id=42:FilterBar"
         _StyleDefs(60)  =   ":id=42,.parent=33"
         _StyleDefs(61)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(62)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(63)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(64)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   3
         Left            =   3480
         TabIndex        =   11
         Top             =   600
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   0
         Left            =   2760
         TabIndex        =   0
         Top             =   240
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   4695
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Height          =   285
         Index           =   7
         Left            =   2760
         TabIndex        =   5
         Top             =   1560
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">99,99"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   2760
         TabIndex        =   4
         Top             =   1230
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   3480
         TabIndex        =   22
         Top             =   1230
         Visible         =   0   'False
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   9
         Left            =   5760
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   ">99,99"
         PromptChar      =   ">"
      End
      Begin ucXPButton.XPButton cmdCopyJob 
         Height          =   855
         Left            =   9720
         TabIndex        =   25
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   1508
         BTYPE           =   3
         TX              =   "Copy All Job in Group"
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
         BCOL            =   14933984
         BCOLO           =   14933984
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmActRate.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate Terjual"
         Height          =   195
         Left            =   4800
         TabIndex        =   23
         Top             =   1605
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3360
         Top             =   1230
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level Mekanik"
         Height          =   195
         Left            =   1440
         TabIndex        =   21
         Top             =   1275
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   6640
         Top             =   930
         Width           =   60
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3360
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   195
         Left            =   1665
         TabIndex        =   17
         Top             =   1905
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Order"
         Height          =   195
         Left            =   1800
         TabIndex        =   16
         Top             =   965
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe Job Order"
         Height          =   195
         Left            =   1440
         TabIndex        =   15
         Top             =   645
         Width           =   1050
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group Model"
         Height          =   195
         Left            =   1575
         TabIndex        =   14
         Top             =   285
         Width           =   915
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4575
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   195
         Left            =   2145
         TabIndex        =   13
         Top             =   1605
         Width           =   345
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   7200
      TabIndex        =   8
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Tambah"
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
   Begin AIFCmp1.asxPowerButton cmdUbah 
      Height          =   360
      Left            =   8160
      TabIndex        =   18
      Top             =   7320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Ubah"
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
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   9120
      TabIndex        =   19
      Top             =   7320
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "&Hapus"
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
   Begin AIFCmp1.asxPowerButton cmdKeluar 
      Height          =   360
      Left            =   10080
      TabIndex        =   20
      Top             =   7320
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   14482683
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
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
   Begin VB.Image asxPowerBanner1 
      Height          =   7455
      Left            =   0
      Picture         =   "frmActRate.frx":0038
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmtActRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsTemp As ADODB.Recordset
Dim rsGrid As ADODB.Recordset
Dim fSBE As Boolean
Dim rateHO As Double
Dim cols As TrueOleDBGrid70.Columns

Private Sub ber_sih()
    For i = 0 To 9
        Select Case i
            Case 7, 9
                Text(i) = "00,00"
                Text(i).Mask = "99,99"
            Case Else
                Text(i) = ""
        End Select
    Next
    rateHO = 0
    Text(4).Enabled = True
    Combo1.Visible = False
    Text(12).Visible = False
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    For i = 0 To 8
        Select Case i
            Case 0, 2, 4, 5
                Text(i).Enabled = Flg
            Case 6, 7
                Text(i).Enabled = flg1
            Case 1, 3, 8
                Text(i).Enabled = False
        End Select
    Next
End Sub

Private Sub refreshGrid(idx As Integer)
    
    GridActRate.ClearFields
    If idx = 1 Then
        
        'If Log.RPLHistori = 0 Then
            Set rsGrid = objO.FetchData("SELECT TOP 20 A.MDL, A.MODEL, A.JOB, A.FS_ITEM, A.RATE, B.RATE as flate, A.KETERANGAN, A.OTHER, A.SBE FROM ACTRATE a WITH (NOLOCK) INNER JOIN HO_FLATE B WITH (NOLOCK) ON A.MDL = B.MDL AND A.FS_ITEM = B.FS_ITEM WHERE A.FS_LOG_DEL = '' ORDER BY A.MDL, A.JOB, A.FS_ITEM")
        'Else
        '    Set rsGrid = objO.FetchData("SELECT TOP 20 A.MDL, A.MODEL, A.JOB, A.FS_ITEM, A.RATE, B.RATE as flate, A.KETERANGAN, A.OTHER, A.SBE FROM [" & dbReplikasiHO & "]..ACTRATE a WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..HO_FLATE B WITH (NOLOCK) ON A.MDL = B.MDL AND A.FS_ITEM = B.FS_ITEM WHERE A.FS_LOG_DEL = '' ORDER BY A.MDL, A.JOB, A.FS_ITEM")
        'End If
    ElseIf idx = 2 Then
        
       ' If Log.RPLHistori = 0 Then
            Set rsGrid = objO.FetchData("SELECT TOP 20 A.MDL, A.MODEL, A.JOB, A.FS_ITEM, A.RATE, B.RATE as flate, A.KETERANGAN, A.OTHER, A.SBE FROM ACTRATE a WITH (NOLOCK) INNER JOIN HO_FLATE B WITH (NOLOCK) ON A.MDL = B.MDL AND A.FS_ITEM = B.FS_ITEM WHERE A.FS_LOG_DEL = '' AND A.MDL LIKE '" & Trim(Text(0)) & "%' AND A.FS_ITEM LIKE '%" & Trim(Text(4)) & "%' ORDER BY A.MDL, A.JOB, A.FS_ITEM")
        'Else
       '     Set rsGrid = objO.FetchData("SELECT TOP 20 A.MDL, A.MODEL, A.JOB, A.FS_ITEM, A.RATE, B.RATE as flate, A.KETERANGAN, A.OTHER, A.SBE FROM [" & dbReplikasiHO & "]..ACTRATE a WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..HO_FLATE B WITH (NOLOCK) ON A.MDL = B.MDL AND A.FS_ITEM = B.FS_ITEM WHERE A.FS_LOG_DEL = '' AND A.MDL LIKE '" & Trim(text(0)) & "%' AND A.FS_ITEM LIKE '%" & Trim(text(4)) & "%' ORDER BY A.MDL, A.JOB, A.FS_ITEM")
       ' End If
    Else
        
      '  If Log.RPLHistori = 0 Then
            Set rsGrid = objO.FetchData("SELECT A.MDL, A.MODEL, A.JOB, A.FS_ITEM, A.RATE, B.RATE as flate, A.KETERANGAN, A.OTHER, A.SBE FROM ACTRATE a WITH (NOLOCK) INNER JOIN HO_FLATE B WITH (NOLOCK) ON A.MDL = B.MDL AND A.FS_ITEM = B.FS_ITEM WHERE A.FS_LOG_DEL = '' ORDER BY A.MDL, A.JOB, A.FS_ITEM")
      '  Else
      '      Set rsGrid = objO.FetchData("SELECT A.MDL, A.MODEL, A.JOB, A.FS_ITEM, A.RATE, B.RATE as flate, A.KETERANGAN, A.OTHER, A.SBE FROM [" & dbReplikasiHO & "]..ACTRATE a WITH (NOLOCK) INNER JOIN [" & dbReplikasiHO & "]..HO_FLATE B WITH (NOLOCK) ON A.MDL = B.MDL AND A.FS_ITEM = B.FS_ITEM WHERE A.FS_LOG_DEL = '' ORDER BY A.MDL, A.JOB, A.FS_ITEM")
      '  End If
    End If
    If Not rsGrid.EOF Then
        GridActRate.DataSource = rsGrid
        Call headColumns
    End If
End Sub

Private Sub headColumns()
    With GridActRate
        .Columns(0).Caption = "KODE MODEL"
        .Columns(0).Width = 2500
        .Columns(0).DataField = "MDL"
        .Columns(1).Caption = "Model"
        .Columns(1).Width = 3000
        .Columns(1).DataField = "MODEL"
        .Columns(2).Caption = "GROUP"
        .Columns(2).Width = 1000
        .Columns(2).DataField = "JOB"
        .Columns(3).Caption = "Job Order"
        .Columns(3).Width = 4000
        .Columns(3).DataField = "FS_ITEM"
        .Columns(4).Caption = "Rate"
        .Columns(4).Width = 1000
        .Columns(4).DataField = "RATE"
        .Columns(5).Caption = "Flate"
        .Columns(5).Width = 1000
        .Columns(5).DataField = "FLATE"
        
        .Columns(6).Caption = "Keterangan"
        .Columns(6).Width = 2500
        .Columns(6).DataField = "KETERANGAN"
        .Columns(7).Caption = "Other"
        .Columns(7).Visible = False
        .Columns(7).DataField = "OTHER"
        .Columns(8).Caption = "Sbe"
        .Columns(8).Visible = False
        .Columns(8).DataField = "SBE"
    End With
End Sub

Private Sub chOther_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
    chOther_Validate False
End Sub

Private Sub chOther_Validate(Cancel As Boolean)
    
    If Not Cancel Then
        If chOther = Checked Then
            Label6.Visible = True
            Text(9).Visible = True
        Else
            Label6.Visible = False
            Text(9).Visible = False
        End If
        Text(9) = "00,00"
    End If
End Sub

Private Sub cmdCopyJob_Click()
    frmeCopyJob.Visible = True
    Text(10).SelLength = Len(Trim(Text(10)))
    Text(10).SetFocus
End Sub

Private Sub cmdDisplay_Click()
    Call refreshGrid(3)
End Sub

Private Sub cmdHapus_Click()
    
   ' If Log.RPLHistori = 0 Then
        Set rsTemp = objO.FetchData("SELECT * FROM ACTRATE WITH (NOLOCK) WHERE MDL = '" & Text(0) & "' AND JOB = '" & Text(2) & "' AND FS_ITEM = '" & Text(4) & "' AND LEVELMECH = '" & Text(5) & "' AND FS_LOG_DEL = ''")
   ' Else
   '     Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) WHERE MDL = '" & text(0) & "' AND JOB = '" & text(2) & "' AND FS_ITEM = '" & text(4) & "' AND LEVELMECH = '" & text(5) & "' AND FS_LOG_DEL = ''")
   ' End If
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus ?", vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
           ' If Log.RPLHistori = 0 Then
                Call objO.UpdateData("UPDATE ACTRATE SET FS_LOG_DEL = '" & gSA & "', FD_TGL_DEL = CONVERT(varchar(10),GETDATE(),1) WHERE MDL = '" & Text(0) & "' AND JOB = '" & Text(2) & "' AND FS_ITEM = '" & Text(4) & "' AND LEVELMECH = '" & Text(5) & "' and FS_LOG_DEL = ''")
           ' Else
                
            '    Call objO.UpdateData("UPDATE [" & dbReplikasiHO & "]..ACTRATE SET FS_LOG_DEL = '" & gSA & "', FD_TGL_DEL = CONVERT(varchar(10),GETDATE(),1) WHERE MDL = '" & text(0) & "' AND JOB = '" & text(2) & "' AND FS_ITEM = '" & text(4) & "' AND LEVELMECH = '" & text(5) & "' and FS_LOG_DEL = ''")
           ' End If
        End If
        Call ber_sih
        Call buka_Tutup(False, False)
        Call refreshGrid(2)
    Else
        MsgBox "Pilih data yang akan dihapus"
        Exit Sub
    End If
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdProses_Click()
Dim rsMdl As New ADODB.Recordset
Dim ulang As Integer
    
    If Trim(Text(10)) <> "" Then
            
            If Log.RPLHistori = 0 Then
                Set rsMdl = objO.FetchData("SELECT * FROM MODEL WITH (NOLOCK) WHERE MDL LIKE '" & Trim(Text(10)) & "%' ORDER BY MDL")
            Else
                Set rsMdl = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MODEL WITH (NOLOCK) WHERE MDL LIKE '" & Trim(Text(10)) & "%' ORDER BY MDL")
            End If
            If Not rsMdl.EOF Then
                Do While Not rsMdl.EOF
                    'If Log.RPLHistori = 0 Then
                        Call objO.InsertData("INSERT INTO HO_FLATE (FS_GRP_MDL,MDL,MODEL,JOB,REMARK,FS_ITEM,Rate,KETERANGAN,TIPEBP,HARGA,FS_LOG_DEL,FD_TGL_DEL,SBE) SELECT DISTINCT '" & Trim(rsMdl("FS_GRP_MDL")) & "','" & Trim(rsMdl("MDL")) & "','" & Trim(rsMdl("DESC")) & "', JOB, REMARK, FS_ITEM, RATE,'CBG','',0,'','',SBE " & _
                                   "FROM HO_FLATE  WHERE MDL LIKE '" & Trim(Text(10)) & "%' AND FS_ITEM NOT IN (SELECT DISTINCT FS_ITEM FROM HO_FLATE WHERE MDL = '" & Trim(rsMdl("MDL")) & "')")
                   ' Else
                   '     Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..HO_FLATE (FS_GRP_MDL,MDL,MODEL,JOB,REMARK,FS_ITEM,Rate,KETERANGAN,TIPEBP,HARGA,FS_LOG_DEL,FD_TGL_DEL,SBE) SELECT DISTINCT '" & Trim(rsMdl("FS_GRP_MDL")) & "','" & Trim(rsMdl("MDL")) & "','" & Trim(rsMdl("DESC")) & "', JOB, REMARK, FS_ITEM, RATE,'CBG','',0,'','',SBE " & _
                                   "FROM [" & dbReplikasiHO & "]..HO_FLATE  WHERE MDL LIKE '" & Trim(text(10)) & "%' AND FS_ITEM NOT IN (SELECT DISTINCT FS_ITEM FROM HO_FLATE WHERE MDL = '" & Trim(rsMdl("MDL")) & "')")
                    'End If
                    'If Log.RPLHistori = 0 Then
                        Call objO.InsertData("INSERT INTO ACTRATE (MDL,MODEL,JOB,REMARK,FS_ITEM,LEVELMECH,KETLEV,Rate,KETERANGAN,OTHER,SBE,FS_LOG_DEL,FD_TGL_DEL) SELECT DISTINCT '" & Trim(rsMdl("MDL")) & "','" & Trim(rsMdl("DESC")) & "',JOB,REMARK,FS_ITEM,'','',Rate,'CBG',0,SBE,'','' " & _
                                    "FROM ACTRATE WHERE MDL LIKE '" & Trim(Text(10)) & "%' AND FS_ITEM NOT IN (SELECT DISTINCT FS_ITEM FROM ACTRATE WHERE MDL = '" & Trim(rsMdl("MDL")) & "')")
                   ' Else
                   '     Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..ACTRATE (MDL,MODEL,JOB,REMARK,FS_ITEM,LEVELMECH,KETLEV,Rate,KETERANGAN,OTHER,SBE,FS_LOG_DEL,FD_TGL_DEL) SELECT DISTINCT '" & Trim(rsMdl("MDL")) & "','" & Trim(rsMdl("DESC")) & "',JOB,REMARK,FS_ITEM,'','',Rate,'CBG',0,SBE,'','' " & _
                                   "FROM [" & dbReplikasiHO & "]..ACTRATE WHERE MDL LIKE '" & Trim(text(10)) & "%' AND FS_ITEM NOT IN (SELECT DISTINCT FS_ITEM FROM [" & dbReplikasiHO & "]..ACTRATE WHERE MDL = '" & Trim(rsMdl("MDL")) & "')")
                   ' End If
                    rsMdl.MoveNext
                Loop
                MsgBox "Proses copy berhasil !"
            End If
    End If
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
          
1         If cmdSimpan.Caption = "&Tambah" Then
2             Call ber_sih
3             Call buka_Tutup(True, True)
4             Text(0).SetFocus
5             cmdSimpan.Caption = "&Simpan"
6             cmdUbah.Caption = "&Batal"
7         ElseIf cmdSimpan.Caption = "&Simpan" Then
              
8             boleh = True
9             pesan = "Lengkapi data "
10            If Text(0) = "" Then
11                pesan = pesan + "Model. "
12                boleh = False
13            End If
14            If Text(2) = "" Then
15                pesan = pesan + "Group. "
16                boleh = False
17            End If
18            If Text(4) = "" Then
19                If Trim(Text(2)) = "SBE" Or Trim(Text(2)) = "SBI" Then
20                   Text(4) = Trim(Text(12)) & " " & Trim(Combo1.Text) & " KM"
21                Else
22                    pesan = pesan + "Job Order. "
23                    boleh = False
24                End If
25            End If
26            If Text(7) = ">>,>>" Then
27               pesan = pesan + "Rate. "
28               boleh = False
29            End If
              
30            If boleh Then
              
                  Dim rsTemp As ADODB.Recordset
                  Dim rsHlp As ADODB.Recordset
                  
31                If Log.RPLHistori = 0 Then
32                    Set rsTemp = objO.FetchData("SELECT * FROM MODEL WITH (NOLOCK) WHERE LEFT(MDL, CHARINDEX('-',MDL)-1) = '" & Trim(Text(0)) & "' ORDER BY MDL")
33                Else
34                    Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MODEL WITH (NOLOCK) WHERE LEFT(MDL, CHARINDEX('-',MDL)-1) = '" & Trim(Text(0)) & "' ORDER BY MDL")
                      ''Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..MODEL WITH (NOLOCK) WHERE MDL = 'KUN40R-GKPNHD P0' ORDER BY MDL")
35                End If
36                If Not rsTemp.EOF Then
37                    rsTemp.MoveFirst
38                    Do While Not rsTemp.EOF
                          ' SIMPAN DI ACTRATE
39                     '   If Log.RPLHistori = 0 Then
40                            Set rsHlp = _
                                  objO.FetchData("SELECT * FROM ACTRATE WITH (NOLOCK) WHERE MDL = '" _
                                  & Trim(rsTemp("MDL")) & "' AND FS_ITEM = '" & Trim(Text(4)) _
                                  & "'")
41                     '   Else
42                       '     Set rsHlp = _
                                  objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..ACTRATE WITH (NOLOCK) WHERE MDL = '" _
                                  & Trim(rsTemp("MDL")) & "' AND FS_ITEM = '" & Trim(text(4)) _
                                  & "'")
                          
43                        'End If
44                        If Not rsHlp.EOF Then
45                            If Trim(Text(2)) = "SBE" Or Trim(Text(2)) = "SBI" Then
46                                'If Log.RPLHistori = 0 Then
47                                    Call objO.UpdateData("UPDATE ACTRATE SET RATE = '" _
                                          & Replace(CDbl(Replace(Text(7), ">", "")), ",", _
                                          ".") & "', OTHER = '" & IIf(chOther.value, 1, 0) & _
                                          "' WHERE MDL = '" & Trim(rsTemp("MDL")) & _
                                          "' AND FS_ITEM = '" & Trim(Text(12)) & " " & _
                                          Trim(Combo1.Text) & " KM'")
48                                'Else
49                                 '   Call objO.UpdateData("UPDATE [" & dbReplikasiHO & "]..ACTRATE SET RATE = '" _
                                          & Replace(CDbl(Replace(text(7), ">", "")), ",", _
                                          ".") & "', OTHER = '" & IIf(chOther.value, 1, 0) & _
                                          "' WHERE MDL = '" & Trim(rsTemp("MDL")) & _
                                          "' AND FS_ITEM = '" & Trim(text(12)) & " " & _
                                          Trim(Combo1.text) & " KM'")
50                                'End If
51                            Else
52                              '  If Log.RPLHistori = 0 Then
53                                    Call objO.UpdateData("UPDATE ACTRATE SET RATE = '" _
                                      & Replace(CDbl(Replace(Text(7), ">", "")), ",", _
                                      ".") & "', OTHER = '" & IIf(chOther.value, 1, 0) & _
                                      "' WHERE MDL = '" & Trim(rsTemp("MDL")) & _
                                      "' AND FS_ITEM = '" & Trim(Text(4)) & "'")
54                               ' Else
55                                '    Call objO.UpdateData("UPDATE [" & dbReplikasiHO & "]..ACTRATE SET RATE = '" _
                                      & Replace(CDbl(Replace(text(7), ">", "")), ",", _
                                      ".") & "', OTHER = '" & IIf(chOther.value, 1, 0) & _
                                      "' WHERE MDL = '" & Trim(rsTemp("MDL")) & _
                                      "' AND FS_ITEM = '" & Trim(text(4)) & "'")
56                                'End If
                                  
57                            End If
58                        Else
59                            If Trim(Text(2)) = "SBE" Or Trim(Text(2)) = "SBI" Then
                                  
60                                'If Log.RPLHistori = 0 Then
61                                    Call objO.InsertData("INSERT INTO ACTRATE (MDL,MODEL,JOB,REMARK,FS_ITEM,LEVELMECH,KETLEV,Rate,KETERANGAN,OTHER,SBE,FS_LOG_DEL,FD_TGL_DEL) VALUES ('" _
                                          & Trim(rsTemp("MDL")) & "','" & _
                                          Trim(rsTemp("DESC")) & "','" & Trim(Text(2)) & _
                                          "','" & Trim(Text(3)) & "','" & Trim(Text(12)) & " " _
                                          & Combo1.Text & " KM','','','" & _
                                          Replace(CDbl(Replace(Text(7), ">", "")), ",", ".") _
                                          & "','" & Trim(Text(6)) & "','" & IIf(chOther.value, _
                                          1, 0) & "','" & IIf(Trim(fSBE), 1, 0) & "','','')")
62                               ' Else
63                                 '   Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..ACTRATE (MDL,MODEL,JOB,REMARK,FS_ITEM,LEVELMECH,KETLEV,Rate,KETERANGAN,OTHER,SBE,FS_LOG_DEL,FD_TGL_DEL) VALUES ('" _
                                          & Trim(rsTemp("MDL")) & "','" & _
                                          Trim(rsTemp("DESC")) & "','" & Trim(text(2)) & _
                                          "','" & Trim(text(3)) & "','" & Trim(text(12)) & " " _
                                          & Combo1.text & " KM','','','" & _
                                          Replace(CDbl(Replace(text(7), ">", "")), ",", ".") _
                                          & "','" & Trim(text(6)) & "','" & IIf(chOther.value, _
                                          1, 0) & "','" & IIf(Trim(fSBE), 1, 0) & "','','')")
64                               ' End If
65                            Else
                                  
66                               ' If Log.RPLHistori = 0 Then
67                                    Call objO.InsertData("INSERT INTO ACTRATE (MDL,MODEL,JOB,REMARK,FS_ITEM,LEVELMECH,KETLEV,Rate,KETERANGAN,OTHER,SBE,FS_LOG_DEL,FD_TGL_DEL) VALUES ('" _
                                        & Trim(rsTemp("MDL")) & "','" & _
                                        Trim(rsTemp("DESC")) & "','" & Trim(Text(2)) & _
                                        "','" & Trim(Text(3)) & "','" & Trim(Text(4)) & _
                                        "','','','" & Replace(CDbl(Replace(Text(7), ">", _
                                        "")), ",", ".") & "','" & Trim(Text(6)) & "','" & _
                                        IIf(chOther.value, 1, 0) & "','" & IIf(Trim(fSBE), _
                                        1, 0) & "','','')")
68                               ' Else
69                                 '   Call objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..ACTRATE (MDL,MODEL,JOB,REMARK,FS_ITEM,LEVELMECH,KETLEV,Rate,KETERANGAN,OTHER,SBE,FS_LOG_DEL,FD_TGL_DEL) VALUES ('" _
                                        & Trim(rsTemp("MDL")) & "','" & _
                                        Trim(rsTemp("DESC")) & "','" & Trim(text(2)) & _
                                        "','" & Trim(text(3)) & "','" & Trim(text(4)) & _
                                        "','','','" & Replace(CDbl(Replace(text(7), ">", _
                                        "")), ",", ".") & "','" & Trim(text(6)) & "','" & _
                                        IIf(chOther.value, 1, 0) & "','" & IIf(Trim(fSBE), _
                                        1, 0) & "','','')")
70                               ' End If
                                  
71                            End If
                              
72                        End If
                          
                          ' JIKA OTHER DI CENTANG, SIMPAN DI HO_FLATE
73                        If chOther.value Then
                             '
74                            'If Log.RPLHistori = 0 Then
75                                Set rsHlp = _
                                      objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE MDL = '" _
                                      & Trim(rsTemp("MDL")) & "' AND FS_ITEM = '" & _
                                      Trim(Text(4)) & "'")
76                           ' Else
77                            '    Set rsHlp = _
                                      objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HO_FLATE WITH (NOLOCK) WHERE MDL = '" _
                                      & Trim(rsTemp("MDL")) & "' AND FS_ITEM = '" & _
                                      Trim(text(4)) & "'")
78                            'End If
                              
79                            If Not rsHlp.EOF Then
                                  
80                               ' If Log.RPLHistori = 0 Then
81                                    Call objO.UpdateData("UPDATE HO_FLATE SET RATE = '" _
                                          & Replace(CDbl(Replace(Text(9), ">", "")), ",", _
                                          ".") & "', KETERANGAN = 'CBG' WHERE MDL = '" & _
                                          Trim(rsTemp("MDL")) & "' AND FS_ITEM = '" & _
                                          Trim(Text(4)) & "'")
82                              '  Else
83                               '     Call objO.UpdateData("UPDATE [" & dbReplikasiHO & "]..HO_FLATE SET RATE = '" _
                                          & Replace(CDbl(Replace(text(9), ">", "")), ",", _
                                          ".") & "', KETERANGAN = 'CBG' WHERE MDL = '" & _
                                          Trim(rsTemp("MDL")) & "' AND FS_ITEM = '" & _
                                          Trim(text(4)) & "'")
84                               ' End If
                                  
85                            Else
86                                If Trim(Text(2)) = "SBE" Or Trim(Text(2)) = "SBI" _
                                      Then
                                      
87                                   'If Log.RPLHistori = 0 Then
88                                        Call _
                                          objO.InsertData("INSERT INTO HO_FLATE (FS_GRP_MDL,MDL,MODEL,JOB,REMARK,FS_ITEM,Rate,KETERANGAN,TIPEBP,HARGA,FS_LOG_DEL,FD_TGL_DEL,SBE) VALUES ('" _
                                              & Trim(rsTemp("FS_GRP_MDL")) & "','" & _
                                              Trim(rsTemp("MDL")) & "','" & _
                                              Trim(rsTemp("DESC")) & "','" & Trim(Text(2)) & _
                                              "','" & Trim(Text(3)) & "','" & Trim(Text(12)) _
                                              & " " & Combo1.Text & " KM','" & _
                                              Replace(CDbl(Replace(Text(9), ">", "")), ",", _
                                              ".") & "','CBG','','0','','','" & _
                                              IIf(Trim(fSBE), 1, 0) & "')")
89                                   ' Else
90                                    '    Call _
                                          objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..HO_FLATE (FS_GRP_MDL,MDL,MODEL,JOB,REMARK,FS_ITEM,Rate,KETERANGAN,TIPEBP,HARGA,FS_LOG_DEL,FD_TGL_DEL,SBE) VALUES ('" _
                                              & Trim(rsTemp("FS_GRP_MDL")) & "','" & _
                                              Trim(rsTemp("MDL")) & "','" & _
                                              Trim(rsTemp("DESC")) & "','" & Trim(text(2)) & _
                                              "','" & Trim(text(3)) & "','" & Trim(text(12)) _
                                              & " " & Combo1.text & " KM','" & _
                                              Replace(CDbl(Replace(text(9), ">", "")), ",", _
                                              ".") & "','CBG','','0','','','" & _
                                              IIf(Trim(fSBE), 1, 0) & "')")
91                                   ' End If
                                      
92                                Else
93                                  '  If Log.RPLHistori = 0 Then
94                                        Call _
                                              objO.InsertData("INSERT INTO HO_FLATE (FS_GRP_MDL,MDL,MODEL,JOB,REMARK,FS_ITEM,Rate,KETERANGAN,TIPEBP,HARGA,FS_LOG_DEL,FD_TGL_DEL,SBE) VALUES ('" _
                                              & Trim(rsTemp("FS_GRP_MDL")) & "','" & _
                                              Trim(rsTemp("MDL")) & "','" & _
                                              Trim(rsTemp("DESC")) & "','" & Trim(Text(2)) & _
                                              "','" & Trim(Text(3)) & "','" & Trim(Text(4)) & _
                                              "','" & Replace(CDbl(Replace(Text(9), ">", "")), _
                                              ",", ".") & "','CBG','','0','','','" & _
                                              IIf(Trim(fSBE), 1, 0) & "')")
95                                 '   Else
96                                   '     Call _
                                              objO.InsertData("INSERT INTO [" & dbReplikasiHO & "]..HO_FLATE (FS_GRP_MDL,MDL,MODEL,JOB,REMARK,FS_ITEM,Rate,KETERANGAN,TIPEBP,HARGA,FS_LOG_DEL,FD_TGL_DEL,SBE) VALUES ('" _
                                              & Trim(rsTemp("FS_GRP_MDL")) & "','" & _
                                              Trim(rsTemp("MDL")) & "','" & _
                                              Trim(rsTemp("DESC")) & "','" & Trim(text(2)) & _
                                              "','" & Trim(text(3)) & "','" & Trim(text(4)) & _
                                              "','" & Replace(CDbl(Replace(text(9), ">", "")), _
                                              ",", ".") & "','CBG','','0','','','" & _
                                              IIf(Trim(fSBE), 1, 0) & "')")
97                                   ' End If
                                      
98                                End If
                                  
99                            End If
                               
100                       End If
                          
101                       rsTemp.MoveNext
102                   Loop
103               End If
                  
104               cmdSimpan.Caption = "&Tambah"
105               cmdUbah.Caption = "&Ubah"
106               Call ber_sih
107               Call buka_Tutup(False, False)
108               Call refreshGrid(1)
109           Else
110               MsgBox pesan, vbCritical, "Validasi Data"
111               Exit Sub
112           End If
113       End If
End Sub

Private Sub cmdTutup_Click()
    frmeCopyJob.Visible = False
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call buka_Tutup(False, True)
        Text(7).SetFocus
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        Call ber_sih
        Call refreshGrid(2)
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
    End If
End Sub

Private Sub Combo1_Click()
    Text(9).Mask = "99,99"
    Text(9) = "00,00"
    Text(7).Mask = "99,99"
    Text(7) = "00,00"
    Set rsSearch = objO.FetchData("SELECT DISTINCT FS_ITEM, RATE, SBE, MDL FROM HO_FLATE WITH (NOLOCK) WHERE FS_Item = '" & Trim(Text(12)) & " " & Combo1.Text & " KM' AND FS_LOG_DEL = '' AND JOB = '" & Trim(Text(2)) & "' AND MDL LIKE '" & Trim(Text(0)) & "%'")
    If rsSearch.BOF Then
        If chOther = Unchecked Then
            Cancel = True
'            gMainKeyCari = Text(Index)
'            gSearchOption = "JOB_ORDER"
'            gBantuan = Text(2)
'            frmCari.Show (vbModal)
        End If
    Else
        If Left(Trim(rsSearch("MDL")), InStr(Trim(rsSearch("MDL")), "-") - 1) = Trim(Text(0)) Then
            rateHO = CDbl(rsSearch("RATE"))
            fSBE = IIf(rsSearch("SBE") = "Y", True, False)
            Text(9).Mask = ""
            Text(9) = Format(CDbl(rateHO), "00.00")
            Text(7).Mask = ""
            Text(7) = Format((CDbl(rateHO) * 0.8), "00.00")
        End If
    End If
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid(1)
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub GridActRate_FilterChange()
    On Error GoTo errHandler
    Set cols = GridActRate.Columns
    Dim C As Integer
    
    C = GridActRate.col
    GridActRate.HoldFields
    rsGrid.filter = getFilter()
    GridActRate.col = C
    GridActRate.EditActive = True
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

Private Sub GridActRate_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If LastRow > 1 Then
        If Not rsGrid.EOF Then
            If rsGrid.RecordCount > 0 Then
                Text(0) = Left(rsGrid("MDL"), InStr(rsGrid("MDL"), "-") - 1)
                Text(1) = rsGrid("MODEL")
                Text(2) = rsGrid("JOB")
                Text(4) = rsGrid("FS_ITEM")
                Text(6) = rsGrid("KETERANGAN")
                If rsGrid("RATE") <> "" Then
                    Text(7) = Format(Replace(CDbl(rsGrid("RATE")), ".", ","), "00.00")
                End If
                chOther = IIf(rsGrid("OTHER") = 1, Checked, Unchecked)
                fSBE = IIf(rsGrid("SBE") = "Y", True, False)
            End If
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
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "GRPMDL"
                frmCari.show (vbModal)
            Case 2
                If Text(0) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.show (vbModal)
                Else
                    MsgBox "Isi Model Kendaraan", vbCritical, "KELENGKAPAN"
                    Text(0).SetFocus
                    Exit Sub
                End If
            Case 4
                If Text(2) <> "" Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = Text(2)
                    frmCari.show (vbModal)
                Else
                    MsgBox "Isi Tipe Job Order", vbCritical, "KELENGKAPAN"
                    Text(2).SetFocus
                    Exit Sub
                End If
            Case 5
                gMainKeyCari = Text(Index)
                gSearchOption = "LEVMEKANIK"
                frmCari.show (vbModal)
            Case 10
                gMainKeyCari = Text(Index)
                gSearchOption = "GRPMDL"
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

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                If Log.RPLHistori = 1 Then
                    Set rsSearch = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..model WITH (NOLOCK) WHERE LEFT(mdl, CHARINDEX('-',MDL)-1) = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                Else
                    Set rsSearch = objO.FetchData("SELECT * FROM model WITH (NOLOCK) WHERE LEFT(mdl, CHARINDEX('-',MDL)-1) = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                End If
                If rsSearch.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "GRPMDL"
                    frmCari.show (vbModal)
                Else
                    Text(1) = ""
                End If
            End If
        Case 2
            If Text(Index) <> "" Then
                If Text(0) <> "" Then
                    Set rsSearch = objO.FetchData("SELECT DISTINCT JOB, REMARK FROM HO_FLATE WITH (NOLOCK) WHERE job = '" & Text(Index) & "' AND FS_LOG_DEL = ''")
                    If rsSearch.BOF Then
                        Cancel = True
                        gMainKeyCari = Text(Index)
                        gSearchOption = "JENIS_KERJA"
                        frmCari.show (vbModal)
                    Else
                        Text(3) = rsSearch("REMARK")
                        If Text(2) = "SBE" Or Text(2) = "SBI" Then
                            Text(4).Enabled = False
                            Combo1.Visible = True
                            Text(12).Visible = True
                            Text(12) = "Servis Berkala "
                            Combo1.Clear
                            If Text(2) = "SBE" Then
                                
                                If Log.RPLHistori = 0 Then
                                    Set rsTemp = objO.FetchData("SELECT * FROM T_REGULaR WHERE fs_keterangan not in ('1.000','5.000') ")
                                Else
                                    Set rsTemp = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..T_REGULaR WHERE fs_keterangan not in ('1.000','5.000') ")
                                End If
                                If Not rsTemp.EOF Then
                                    rsTemp.MoveFirst
                                    Do While Not rsTemp.EOF
                                        Combo1.AddItem (Trim(rsTemp!fs_keterangan))
                                        rsTemp.MoveNext
                                    Loop
                                End If
                            Else
                                Combo1.AddItem ("1.000")
                                Combo1.AddItem ("5.000")
                            End If
                        Else
                            Text(4).Enabled = True
                            Combo1.Visible = False
                            Text(12).Visible = False
                        End If
                    End If
                Else
                    MsgBox "Isi Model Kendaraan", vbCritical, "KELENGKAPAN"
                    Text(0).SetFocus
                    Sendkeys ("+{TAB}")
                    Exit Sub
                End If
            End If
        Case 4
            If Text(2) = "" Then
                MsgBox "Isi Group Order terlebih dahulu !!!"
                Text(2).SelLength = Len(Text(2))
                Text(2).SetFocus
                Sendkeys ("+{TAB}")
                Exit Sub
            Else
                If Text(Index) <> "" Then
                    Dim is03 As Boolean
                    is03 = False
                    Text(Index) = Text(Index) & " "
                    Dim nLen As Double
                    nLen = Len(Text(Index))
                    Dim is0 As Integer
                    is0 = 0
                    For i = nLen To 2 Step -1
                        If Mid(Text(Index).Text, i - 1, 1) = "0" Then
                            is0 = is0 + 1
                            If is0 = 3 Then
                                If Mid(Text(Index), i - 2, 1) <> "." Then
                                    Text(Index) = Left(Text(Index), i - 2) & "." & Right(Text(Index), nLen - (i - 2))
                                End If
                                is03 = True
                                Exit For
                            End If
                        End If
                    Next i
                    
                    If Right(Text(Index).Text, 1) = " " Then
                        Text(Index).Text = Left(Text(Index).Text, Len(Text(Index)) - 1)
                    End If
                    
                    If UCase(Left(Trim(Text(2).Text), 2)) = "SB" Then
                        If is03 = False Then
                            MsgBox "Pekerjaan Service Berkala, harus memiliki Nilai KM (Kelipatan 1.000)!!!", vbCritical + vbOKOnly, "Service Rate"
                            Text(Index).Text = ""
                            Exit Sub
                        End If
                    End If
                    
                    Text(9).Mask = "99,99"
                    Text(9) = "00,00"
                    Text(7).Mask = "99,99"
                    Text(7) = "00,00"
                    Set rsSearch = objO.FetchData("SELECT DISTINCT FS_ITEM, RATE, SBE, MDL FROM HO_FLATE WITH (NOLOCK) WHERE FS_Item = '" & Text(Index) & "' AND FS_LOG_DEL = '' AND JOB = '" & Trim(Text(2)) & "' AND MDL LIKE '" & Trim(Text(0)) & "%'")
                    If rsSearch.BOF Then
                        If chOther = Unchecked Then
                            Cancel = True
                            gMainKeyCari = Text(Index)
                            gSearchOption = "JOB_ORDER"
                            gBantuan = Text(2)
                            frmCari.show (vbModal)
                        End If
                    Else
                        If Left(Trim(rsSearch("MDL")), InStr(Trim(rsSearch("MDL")), "-") - 1) = Trim(Text(0)) Then
                            rateHO = CDbl(rsSearch("RATE"))
                            fSBE = IIf(rsSearch("SBE") = "Y", True, False)
                            Text(9).Mask = ""
                            Text(9) = Format(CDbl(rateHO), "00.00")
                            Text(7).Mask = ""
                            Text(7) = Format((CDbl(rateHO) * 0.8), "00.00")
                        End If
                    End If
                End If
            End If
        Case 5
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM T_MECHLEV WITH (NOLOCK) WHERE levelmech = '" & Text(Index) & "'")
                If rsSearch.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "LEVMEKANIK"
                    frmCari.show (vbModal)
                Else
                    Text(8) = rsSearch("KETERANGAN")
                End If
            End If
        Case 10
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM model WITH (NOLOCK) WHERE LEFT(mdl, CHARINDEX('-',MDL)-1) = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If rsSearch.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "GRPMDL"
                    frmCari.show (vbModal)
                Else
                    Text(1) = ""
                End If
            End If
    End Select
End Sub

