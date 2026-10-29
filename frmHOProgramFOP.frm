VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmHOProgramFOP 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MASTER FOP PROGRAM"
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7500
   ScaleWidth      =   12585
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12840
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   6495
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   14055
      Begin TDBDate6Ctl.TDBDate tglRetailAwal 
         Height          =   285
         Left            =   1440
         TabIndex        =   46
         Top             =   1420
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         Calendar        =   "frmHOProgramFOP.frx":0000
         Caption         =   "frmHOProgramFOP.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOProgramFOP.frx":0198
         Keys            =   "frmHOProgramFOP.frx":01B6
         Spin            =   "frmHOProgramFOP.frx":0214
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
         Text            =   "22/05/2020"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   43973
         CenturyMode     =   0
      End
      Begin VB.ComboBox cbTagihan 
         Height          =   315
         ItemData        =   "frmHOProgramFOP.frx":023C
         Left            =   1440
         List            =   "frmHOProgramFOP.frx":0246
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   780
         Width           =   2415
      End
      Begin VB.CheckBox chPart 
         BackColor       =   &H000080FF&
         Caption         =   "Free Part"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   39
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frameSubstitusi 
         BackColor       =   &H000080FF&
         Caption         =   "Substitusi Part"
         Height          =   2655
         Left            =   7440
         TabIndex        =   30
         Top             =   120
         Width           =   6495
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   3075
            TabIndex        =   31
            Top             =   240
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Index           =   12
            Left            =   555
            TabIndex        =   32
            Top             =   240
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Height          =   285
            Index           =   13
            Left            =   555
            TabIndex        =   33
            Top             =   540
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
            Mask            =   ">999"
            PromptChar      =   ">"
         End
         Begin TrueOleDBGrid70.TDBGrid gridSub 
            Height          =   1335
            Left            =   120
            TabIndex        =   36
            Top             =   1155
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   2355
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
            _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H83F3EE&"
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
         Begin AIFCmp1.asxPowerButton cmdHapusSub 
            Height          =   360
            Left            =   2160
            TabIndex        =   37
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "&Hapus"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin AIFCmp1.asxPowerButton cmdSimpanSub 
            Height          =   360
            Left            =   1200
            TabIndex        =   38
            Top             =   540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Simpan"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
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
            Left            =   2955
            Top             =   240
            Width           =   60
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part"
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
            Left            =   120
            TabIndex        =   35
            Top             =   285
            Width           =   300
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty"
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
            Left            =   150
            TabIndex        =   34
            Top             =   585
            Width           =   270
         End
      End
      Begin VB.Frame frmeGanti 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   960
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   10335
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   21
            Top             =   765
            Width           =   3465
            _ExtentX        =   6112
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
            Index           =   8
            Left            =   915
            TabIndex        =   22
            Top             =   405
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
            Height          =   285
            Index           =   9
            Left            =   6675
            TabIndex        =   24
            Top             =   720
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
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
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   10
            Left            =   6675
            TabIndex        =   23
            Top             =   360
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
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
         Begin AIFCmp1.asxPowerButton cmdProses 
            Height          =   360
            Left            =   7965
            TabIndex        =   25
            Top             =   1320
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   635
            BackColor       =   16761024
            BorderStyle     =   3
            Caption         =   "&Proses"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin VB.Image Image1 
            Height          =   855
            Left            =   4680
            Picture         =   "frmHOProgramFOP.frx":0254
            Stretch         =   -1  'True
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Baru"
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
            Left            =   5880
            TabIndex        =   27
            Top             =   405
            Width           =   675
         End
         Begin VB.Shape Shape4 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   3315
            Top             =   405
            Width           =   60
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part Lama"
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
            Left            =   120
            TabIndex        =   26
            Top             =   450
            Width           =   720
         End
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   180
         Width           =   5865
         _ExtentX        =   10345
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
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   2040
         Width           =   3465
         _ExtentX        =   6112
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
      Begin TrueOleDBGrid70.TDBGrid GridPart 
         Height          =   3615
         Left            =   120
         TabIndex        =   8
         Top             =   2760
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   6376
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HFFFFFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H83F3EE&"
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
         Index           =   2
         Left            =   3960
         TabIndex        =   3
         Top             =   1740
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   1740
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Height          =   285
         Index           =   5
         Left            =   1440
         TabIndex        =   1
         Top             =   480
         Width           =   2385
         _ExtentX        =   4207
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
         Index           =   6
         Left            =   3960
         TabIndex        =   15
         Top             =   480
         Width           =   3465
         _ExtentX        =   6112
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
         Left            =   1440
         TabIndex        =   5
         Top             =   2340
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
         Mask            =   ">999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   1440
         TabIndex        =   43
         Top             =   1125
         Width           =   825
         _ExtentX        =   1455
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
      Begin TDBDate6Ctl.TDBDate tglRetailAkhir 
         Height          =   285
         Left            =   3000
         TabIndex        =   48
         Top             =   1420
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         Calendar        =   "frmHOProgramFOP.frx":3527
         Caption         =   "frmHOProgramFOP.frx":3653
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOProgramFOP.frx":36BF
         Keys            =   "frmHOProgramFOP.frx":36DD
         Spin            =   "frmHOProgramFOP.frx":373B
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
         Text            =   "22/05/2020"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   43973
         CenturyMode     =   0
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   15
         Left            =   3255
         TabIndex        =   49
         Top             =   2340
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   12
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">9999999999999"
         PromptChar      =   ">"
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deposit"
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
         Left            =   2610
         TabIndex        =   50
         Top             =   2385
         Width           =   540
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s / d"
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
         Left            =   2640
         TabIndex        =   47
         Top             =   1470
         Width           =   315
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode Retail"
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
         Left            =   285
         TabIndex        =   45
         Top             =   1465
         Width           =   990
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Tahun NIK"
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
         Left            =   120
         TabIndex        =   44
         Top             =   1170
         Width           =   1155
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tagihan"
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
         Left            =   705
         TabIndex        =   41
         Top             =   840
         Width           =   570
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Discount Part"
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
         Left            =   5145
         TabIndex        =   40
         Top             =   1710
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Left            =   1035
         TabIndex        =   17
         Top             =   2385
         Width           =   270
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3840
         Top             =   487
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model Kendaran"
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
         Left            =   120
         TabIndex        =   16
         Top             =   525
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Part"
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
         Left            =   1005
         TabIndex        =   14
         Top             =   1785
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3840
         Top             =   1740
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Job Order"
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
         Left            =   585
         TabIndex        =   10
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
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
         Left            =   465
         TabIndex        =   9
         Top             =   2085
         Width           =   840
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   7320
         Top             =   187
         Width           =   60
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7275
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   12832
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "  Tabel Master FOP Program"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   4194304
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12060
      _ExtentX        =   21273
      _ExtentY        =   503
      BackColor       =   16777215
      HighlightColor  =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   7920
      TabIndex        =   7
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Tambah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      Left            =   8880
      TabIndex        =   11
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Ubah"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      Left            =   9840
      TabIndex        =   12
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Hapus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
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
      Left            =   10800
      TabIndex        =   13
      Top             =   7080
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdCetak 
      Height          =   360
      Left            =   6600
      TabIndex        =   18
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Cetak"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdGanti 
      Height          =   360
      Left            =   480
      TabIndex        =   19
      Top             =   7080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "&Ganti Part"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      Appearance      =   1
   End
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   5040
      TabIndex        =   28
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   12648447
      Caption         =   "Import Excell"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ARProgBarCtrl.ARProgressBar ProgBar 
      Height          =   255
      Left            =   1200
      TabIndex        =   29
      Top             =   7080
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   450
      Value           =   0
      ForeColor       =   52479
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionForeColor=   128
      IniColor        =   52479
   End
End
Attribute VB_Name = "frmHOProgramFOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOProgramFOP
' DateTime  : 26/05/2020 10:48
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

 

Dim rsGrid      As New ADODB.Recordset
Dim rsGridSub   As New ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim rsModel     As ADODB.Recordset
Dim newRecord   As Boolean
Dim KDPART_LAMA As String
Dim boleh       As Boolean
Dim pesan       As String
Dim cols        As TrueOleDBGrid70.Columns

Private Sub ber_sih()
    For i = 0 To 7
        Select Case i
            Case 4
            Case 7
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    Text(11) = ""
    Text(12) = ""
    Text(13) = 0
    
    Text(14) = ""
    tglRetailAwal = isnow
    tglRetailAkhir = isnow
    ''cbTagihan.Text = ""
    cbTagihan.Clear
    cbTagihan.AddItem "NRM"
    cbTagihan.AddItem "TAM"
    
    Label15.Visible = True
    Text(15).Visible = True
    
    chPart.value = vbUnchecked
End Sub

Private Sub buka_tutupSubtitusi(Flg As Boolean)
    Text(11).Enabled = Flg
    Text(12).Enabled = Flg
    Text(13).Enabled = Flg
    cmdSimpanSub.Enabled = Flg
    cmdHapusSub.Enabled = Flg
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    Text(0).Enabled = Flg
    Text(1).Enabled = flg1
    Text(2).Enabled = flg1
    Text(3).Enabled = flg1
    Text(5).Enabled = flg1
    Text(7).Enabled = flg1
    chPart.Enabled = flg1
    
    Text(14).Enabled = flg1
    tglRetailAwal.Enabled = flg1
    tglRetailAkhir.Enabled = flg1
    cbTagihan.Enabled = flg1
    
    Label15.Enabled = flg1
    Text(15).Enabled = flg1
    
End Sub

Private Sub refreshGrid()
    GridPart.ClearFields
    Set rsGrid = objO.FetchDataHO("SELECT FOP,NIK,RETAILAWAL,RETAILAKHIR,mdl as [KODE MODEL], MODEL, fs_item as [JOB ORDER], item_part as [KODE PART], nama_part as [NAMA PART], QTY, KETERANGAN FROM HO_FOP_PROGRAM WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY FS_item, ITEM_part")
    gridSub.ClearFields
    Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_PARTPACKSUBSTITUSI WITH (NOLOCK) WHERE FS_LOG_DEL = '' and 1=0")
    
    If Not rsGrid.EOF Then
        GridPart.DataSource = rsGrid
        Call HeaderColom
    End If
End Sub

Private Sub HeaderColom()
    GridPart.Columns(0).Caption = "TAGIHAN"
    GridPart.Columns(0).Width = 800
    GridPart.Columns(1).Caption = "KODE TAHUN NIK"
    GridPart.Columns(1).Width = 2000
    GridPart.Columns(2).Caption = "RETAIL AWAL"
    GridPart.Columns(2).Width = 2000
    GridPart.Columns(3).Caption = "RETAIL AKHIR"
    GridPart.Columns(3).Width = 2000
    GridPart.Columns(4).Caption = "KODE MODEL"
    GridPart.Columns(4).Width = 2000
    GridPart.Columns(5).Caption = "MODEL"
    GridPart.Columns(5).Width = 3500
    GridPart.Columns(6).Caption = "JOB ORDER"
    GridPart.Columns(6).Width = 6000
    GridPart.Columns(7).Caption = "KODE PART"
    GridPart.Columns(7).Width = 2000
    GridPart.Columns(8).Caption = "NAMA PART"
    GridPart.Columns(8).Width = 3000
    GridPart.Columns(9).Caption = "QTY"
    GridPart.Columns(9).Width = 600
    GridPart.Columns(9).NumberFormat = "#,#0"
    GridPart.Columns(10).Caption = "KETERANGAN"
    GridPart.Columns(10).Width = 3000
End Sub


Private Sub HeaderColomSubstitusi()
    
    gridSub.Columns(0).Caption = "KODE PART"
    gridSub.Columns(0).Width = 2000
    gridSub.Columns(1).Caption = "NAMA PART"
    gridSub.Columns(1).Width = 3000
    gridSub.Columns(2).Caption = "QTY"
    gridSub.Columns(2).Width = 600
    gridSub.Columns(2).NumberFormat = "#,#0"
End Sub

Private Sub cbTagihan_Change()
    If cbTagihan.Text = "NRM" Then
        Label15.Visible = True
        Text(15).Visible = True
    Else
        Label15.Visible = False
        Text(15).Visible = False
    End If
End Sub

Private Sub cbTagihan_Click()
    If cbTagihan.Text = "NRM" Then
        Label15.Visible = True
        Text(15).Visible = True
    Else
        Label15.Visible = False
        Text(15).Visible = False
    End If
End Sub

Private Sub cmdGanti_Click()
    Text(8) = ""
    Text(4) = ""
    Text(10) = ""
    Text(9) = ""
    frmeGanti.Visible = True
End Sub

Private Sub cmdHapus_Click()
    Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_fop_program WITH (NOLOCK) WHERE FOP = '" & cbTagihan.Text & "' and NIK = '" & NulSID(Text(14)) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & Text(5) & "' AND item_part = '" & Text(1) & "' and FS_LOG_DEL = ''")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus data part : " & Text(1) & " ?" & vbCr & "Untuk Job Order : " & Text(0), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE ho_fop_program SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE FOP = '" & cbTagihan.Text & "' and NIK = '" & NulSID(Text(14)) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & Text(5) & "' AND item_part = '" & Text(1) & "' and FS_LOG_DEL = ''")
            Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('PARTPACK', '" & Text(0) & "', 'D')")
        End If
        Call ber_sih
        Call refreshGrid
    Else
        MsgBox "Pilih data yang ingin dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdHapusSub_Click()
    Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and mdl = '" & NulSID(rsGrid.Fields(4)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "'  and ITEM_PARTSUB = '" & NulSID(rsGridSub.Fields(0)) & "' and NAMA_PARTSUB = '" & NulSID(rsGridSub.Fields(1)) & "' and QTYSUB = '" & NulSID(rsGridSub.Fields(2)) & "' and FS_LOG_DEL = ''")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus data sub part : " & NulSID(rsGridSub.Fields(0)) & " ?" & vbCr & "Untuk Job Order : " & Text(0), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE HO_FOP_PROGRAM_SUB SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and mdl = '" & NulSID(rsGrid.Fields(4)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "'  and ITEM_PARTSUB = '" & NulSID(rsGridSub.Fields(0)) & "' and NAMA_PARTSUB = '" & NulSID(rsGridSub.Fields(1)) & "' and QTYSUB = '" & NulSID(rsGridSub.Fields(2)) & "' and FS_LOG_DEL = ''")
        End If
        Text(11) = ""
        Text(12) = ""
        Text(13) = 0
        gridSub.ClearFields
        Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(8)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and MDL = '" & NulSID(rsGrid.Fields(4)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
        gridSub.DataSource = rsGridSub
        Call HeaderColomSubstitusi
        gridSub.ReBind
    Else
        MsgBox "Pilih data yang ingin dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdImport_Click()
      Dim ErrBegin As Boolean, konter As Long
1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim rsTempExcel As ADODB.Recordset
         Dim Jml As Integer
         Dim itemPart As String
         
        ' Dim konter As Integer
         
2       konter = 0
3       ProgBar.Min = 0
4       CommonDialog1.FileName = ""
5       CommonDialog1.InitDir = App.Path
6       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
7       CommonDialog1.ShowOpen
8       If CommonDialog1.FileName <> "" Then
9          If CnExel.State Then CnExel.Close
10         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
11         GoSub Proses
12     Else
13        Exit Sub
14     End If
15     Exit Sub
       
Proses:
16    ErrBegin = True
         Dim VarNmFile As String
         Dim model As String
         
17       Call objHT.OpenTransactionHO
18        gBantuan = "SBE"
19       If RcExcel.State Then RcExcel.Close
20       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
21       If RcExcel.RecordCount > 0 Then
22          ProgBar.Visible = True
23          ProgBar.Max = RcExcel.RecordCount
24          RcExcel.MoveFirst
25          Do While Not RcExcel.EOF
26             DoEvents
27             konter = konter + 1
28             ProgBar.value = konter
29             model = ""
              
30             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
31             If objO.FetchDataHO_Count("select * from JOBORDER WITH(NOLOCK) WHERE FS_LOG_DEL = '' and FS_ITEM = '" & Trim(RcExcel.Fields(0).value) & "'") = 0 Then
32                Call objHT.CloseTransactionHO_err
33                ProgBar.Visible = False
34                MsgBox "JOB ORDER = " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " salah!", vbCritical, "Error"
35                Exit Sub
36             End If
               
37             If Nul(RcExcel.Fields(1).value) = "" Then
38                Call objHT.CloseTransactionHO_err
39                ProgBar.Visible = False
40                MsgBox "JOB ORDER baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
41                Exit Sub
42             Else
43                Set rsTempExcel = objO.FetchDataHO("SELECT * FROM MODEL WHERE FS_LOG_DEL = '' AND MDL = '" & Trim(RcExcel.Fields(1).value) & "'")
44                If rsTempExcel.RecordCount > 0 Then
45                    model = Trim(rsTempExcel("DESC"))
46                Else
                      'jika tidak ada data
47                    Call objHT.CloseTransactionHO_err
48                    ProgBar.Visible = False
49                    MsgBox "JOB ORDER baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
50                    Exit Sub
51                End If
52             End If
               
53             If Nul(RcExcel.Fields(2).value) = "" Then
54                If Nul(RcExcel.Fields(6).value) = "" Then
55                    Call objHT.CloseTransactionHO_err
56                    ProgBar.Visible = False
57                    MsgBox "ITEM PART baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
58                    Exit Sub
59                Else
60                    itemPart = Trim(Replace(RcExcel.Fields(6).value, "'", ""))
61                End If
62             Else
63                itemPart = Nul(RcExcel.Fields(2).value)
64             End If
               
65             If Nul(RcExcel.Fields(3).value) = "" Then
66                Call objHT.CloseTransactionHO_err
67                ProgBar.Visible = False
68                MsgBox "NAMA PART baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
69                Exit Sub
70             End If
               
71             If Nul(RcExcel.Fields(4).value) = "" Then
72                Call objHT.CloseTransactionHO_err
73                ProgBar.Visible = False
74                MsgBox "QTY PART baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
75                Exit Sub
76             Else
77                If CInt(RcExcel.Fields(4).value) <= 0 Then
78                    Call objHT.CloseTransactionHO_err
79                    ProgBar.Visible = False
80                    MsgBox "QTY PART baris ke-" & konter & " harus lebih dari 0 !", vbCritical, "Error"
81                    Exit Sub
82                End If
83             End If
84             Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_partpack WITH (NOLOCK) WHERE fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(RcExcel.Fields(1).value) & "' AND item_part = '" & Trim(itemPart) & "' and FS_LOG_DEL = ''")
85            If Not rsTemp.EOF Then
86                Call objHT.CloseTransactionHO_err
87                ProgBar.Visible = False
88                MsgBox "Job Order " & Trim(RcExcel.Fields(0).value) & " untuk model " & Trim(RcExcel.Fields(1).value) & ", item part " & itemPart & " sudah ada !", vbCritical, "Error"
89                Exit Sub
90            Else
91                Call objHT.InsertData_Ho("INSERT INTO ho_partpack (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL) VALUES ('" & gBantuan & "', '" & Trim(RcExcel.Fields(0).value) & "', '" & Trim(RcExcel.Fields(1).value) & "','" & model & "', '" & Trim(itemPart) & "','" & Trim(RcExcel.Fields(3).value) & "','" & CDbl(Trim(RcExcel.Fields(4).value)) & "','" & Trim(RcExcel.Fields(5).value) & "','','')")
92            End If
93            Call objHT.InsertData_Ho("INSERT INTO TEMPKIRIM VALUES ('PARTPACK', '" & Trim(RcExcel.Fields(0).value) & "', 'S')")
              
94             RcExcel.MoveNext
95          Loop
96       End If
97       RcExcel.Close
98       CnExel.Close
      '   ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
99        Call objHT.CloseTransactionHO
100      MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
101      ProgBar.Visible = False
102      Call ber_sih
103      Call buka_Tutup(False, False)
104      Call refreshGrid
105      cmdUbah.Caption = "&Ubah"
106      cmdSimpan.Caption = "&Tambah"
      '   GridAFI.Refresh
107      Return
108      Exit Sub
vb_error:
109      If ErrBegin = True Then
           ' DbCon.RollbackTrans
110      End If
111      Call objHT.CloseTransactionHO_err
112      ProgBar.Visible = False
113      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdProses_Click()
    Call objO.InsertDataHO("INSERT INTO HO_PARTPACK  (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL) SELECT DISTINCT JOB, FS_ITEM, MDL, MODEL, '" & Trim(Text(10)) & "', '" & Trim(Text(9)) & "', QTY, KETERANGAN, FS_LOG_DEL, FD_TGL_DEL FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & Trim(Text(8)) & "'")
    Call objO.UpdateDataHO("UPDATE HO_PARTPACK SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE ITEM_PART = '" & Trim(Text(8)) & "'")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(Text(10)) & "', 'S')")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(Text(8)) & "', 'D')")
    frmeGanti.Visible = False
End Sub

Private Sub cmdSimpan_Click()
          Dim isBegin As Boolean
          Dim SBE As String
1        On Error GoTo cmdSimpan_Click_Error

2         isBegin = False
3         If cmdSimpan.Caption = "&Tambah" Then
4             newRecord = True
5             Call ber_sih
6             Call buka_Tutup
7             Call buka_tutupSubtitusi(False)
8             Text(0).SetFocus
9             cmdUbah.Caption = "&Batal"
10            cmdSimpan.Caption = "&Simpan"
11        ElseIf cmdSimpan.Caption = "&Simpan" Then
12            boleh = True
13            pesan = "Lengkapi data berikut "
              
14            If cbTagihan.Text = "" Then
15                pesan = pesan + "Jenis Tagihan. "
16                cbTagihan.SetFocus
17                boleh = False
18            End If
              
19            If Text(14) = "" Then
20                boleh = False
21                pesan = pesan + "Kode NIK Tahun. "
22            End If
              
23            If CDbl(Text(7)) <= 0 Then
24                pesan = pesan + "Jumlah item (QTY) "
25                Text(7) = 0
26                Text(7).SetFocus
27                boleh = False
28            End If
29            If Text(0) = "" Then
30                boleh = False
31                pesan = pesan + "Job Order. "
32            End If
33            If Text(5) = "" Then
34                boleh = False
35                pesan = pesan + "Model. "
36            ElseIf Len(NulSID(Text(5))) < 7 Then
37                boleh = False
38                pesan = pesan + "Model min 7 karakter. "
39            End If
40            If Text(1) = "" Then
41                boleh = False
42                pesan = pesan + "Part. "
43            End If
              
44            If boleh Then
45                If newRecord Then
46                    Set rsModel = objO.FetchDataHO("SELECT * FROM model WITH (NOLOCK) WHERE MDL LIKE '" & NulSID(Text(5)) & "%'")
47                    If rsModel.RecordCount > 0 Then
48                        If MsgBox("Apakah yakin akan menambahkan data ke " & rsModel.RecordCount & " model ?", vbYesNo, "Konfirmasi tambah data") = vbYes Then
49                            GoTo loopModel
50                        Else
51                            Exit Sub
52                        End If
53                    Else
54                        MsgBox "Model kendaraan tidak ditemukan." & vbCr & "Silahkan ubah model kendaraan"
55                        Exit Sub
56                    End If
57                End If
58                Call objHT.OpenTransactionHO
59                isBegin = True
60                Set rsTemp = objHT.FetchDataHO("SELECT * FROM ho_fop_program WITH (NOLOCK) WHERE FOP = '" & cbTagihan.Text & "' and NIK = '" & NulSID(Text(14)) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & NulSID(Text(5)) & "' AND item_part = '" & Text(1) & "' and FS_LOG_DEL = ''")
61                If Not rsTemp.EOF Then
62                    If newRecord Then
63                        MsgBox "Sudah ada data yang sama." & vbCr & "Jika ingin mengubah pilih ubah"
64                        Exit Sub
65                    Else
66                        Call objHT.InsertData_Ho("UPDATE ho_fop_program SET item_part = '" & Text(1) & "', nama_part = '" & Text(2) & "', keterangan = '" & Text(3) & "', QTY = '" & CDbl(Text(7)) & "',FOP = '" & cbTagihan.Text & "',NIK = '" & NulSID(Text(14)) & "',RETAILAWAL = '" & Format(tglRetailAwal, "yyyy-mm-dd") & "',RETAILAKHIR = '" & Format(tglRetailAkhir, "yyyy-mm-dd") & "' WHERE FOP = '" & NulSID(rsGrid.Fields(0)) & "' and NIK = '" & NulSID(rsGrid.Fields(1)) & "' and fs_item = '" & NulSID(rsGrid.Fields(6)) & "' AND mdl = '" & NulSID(rsGrid.Fields(4)) & "' AND item_part = '" & NulSID(rsGrid.Fields(7)) & "' and FS_LOG_DEL = ''")
67                    End If
68                Else
69                    Call objHT.InsertData_Ho("INSERT INTO ho_fop_program (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL,FOP,NIK,RETAILAWAL,RETAILAKHIR,USERENTRY,TGLENTRY) VALUES ('" & gBantuan & "', '" & Text(0) & "','" & Left(NulSID(Text(5)), 20) & "','" & Left(NulSID(Text(6)), 50) & "', '" & Text(1) & "','" & Text(2) & "','" & CDbl(Text(7)) & "','" & Text(3) & "','','','" & cbTagihan.Text & "','" & NulSID(Text(14)) & "','" & Format(tglRetailAwal, "yyyy-mm-dd") & "','" & Format(tglRetailAkhir, "yyyy-mm-dd") & "','" & gSA & "',getdate())")
70                End If
                  ''cek mvp m_deposit
                  ''--------------------------------------------------------------
71                If cbTagihan.Text = "NRM" Then
                      
                      
72                    Set rsTemp = objHT.FetchDataHO("SELECT * FROM mvp..m_deposit WITH (NOLOCK) WHERE KodeKategory = '" & Left(NulSID(Text(5)), 20) & "' ")
73                    SBE = objHT.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(Text(0)) & "'")
74                    If SBE = "R010" Then
75                        SBE = "sb10k"
76                    ElseIf SBE = "R020" Then
77                        SBE = "sb20k"
78                    ElseIf SBE = "R030" Then
79                        SBE = "sb30k"
80                    ElseIf SBE = "R040" Then
81                        SBE = "sb40k"
82                    ElseIf SBE = "R050" Then
83                        SBE = "sb50k"
84                    Else
85                        SBE = "DEPOSIT"
86                    End If
87                    If Not rsTemp.EOF Then
88                        sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(Text(15)) & " WHERE KodeKategory = '" & Left(NulSID(Text(5)), 20) & "' "
89                    Else
90                        sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & ",tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(Text(5)), 20) & "'," & k2t(Text(15)) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
91                    End If
92                    Call objHT.InsertData_Ho(sSQL)
93                End If
94                Call objHT.CloseTransactionHO
95                GoTo Lanjutkan
loopModel:
96                rsModel.MoveFirst
97                Call objHT.OpenTransactionHO
98                isBegin = True
99                pesan = ""
100               Do While Not rsModel.EOF
101                   Set rsTemp = objHT.FetchDataHO("SELECT * FROM ho_fop_program WITH (NOLOCK) WHERE FOP = '" & cbTagihan.Text & "' and NIK = '" & NulSID(Text(14)) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & Left(NulSID(rsModel!MDL), 20) & "' AND item_part = '" & Text(1) & "' and FS_LOG_DEL = ''")
102                   If Not rsTemp.EOF Then
103                       If newRecord Then
104                           pesan = pesan & vbCr & "Kode model " & NulSID(rsModel!MDL) & " sudah ada item part " & NulSID(Text(0)) & vbCrLf & "Silahkan menggunakan tombol ubah untuk mengedit data"
105                       Else
106                           Call objHT.InsertData_Ho("UPDATE ho_fop_program SET item_part = '" & Text(1) & "', nama_part = '" & Text(2) & "', keterangan = '" & Text(3) & "', QTY = '" & CDbl(Text(7)) & "',FOP = '" & cbTagihan.Text & "',NIK = '" & NulSID(Text(14)) & "',RETAILAWAL = '" & Format(tglRetailAwal, "yyyy-mm-dd") & "',RETAILAKHIR = '" & Format(tglRetailAkhir, "yyyy-mm-dd") & "' WHERE FOP = '" & NulSID(rsGrid.Fields(0)) & "' and NIK = '" & NulSID(rsGrid.Fields(1)) & "' and fs_item = '" & NulSID(rsGrid.Fields(6)) & "' AND mdl = '" & Left(NulSID(rsModel!MDL), 20) & "' AND item_part = '" & NulSID(rsGrid.Fields(7)) & "' and FS_LOG_DEL = ''")
107                       End If
108                   Else
109                       Call objHT.InsertData_Ho("INSERT INTO ho_fop_program (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL,FOP,NIK,RETAILAWAL,RETAILAKHIR,USERENTRY,TGLENTRY) VALUES ('" & gBantuan & "', '" & Text(0) & "','" & Left(NulSID(rsModel!MDL), 20) & "','" & Left(NulSID(rsModel!desc), 50) & "', '" & Text(1) & "','" & Text(2) & "','" & CDbl(Text(7)) & "','" & Text(3) & "','','','" & cbTagihan.Text & "','" & NulSID(Text(14)) & "','" & Format(tglRetailAwal, "yyyy-mm-dd") & "','" & Format(tglRetailAkhir, "yyyy-mm-dd") & "','" & gSA & "',getdate())")
110                   End If
                      ''cek mvp m_deposit
                      ''--------------------------------------------------------------
111                   If cbTagihan.Text = "NRM" Then
                          ''Dim sbe As String
                          
112                       Set rsTemp = objHT.FetchDataHO("SELECT * FROM mvp..m_deposit WITH (NOLOCK) WHERE KodeKategory = '" & Left(NulSID(rsModel!MDL), 20) & "' ")
113                       SBE = objHT.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(Text(0)) & "'")
114                       If SBE = "R010" Then
115                           SBE = "sb10k"
116                       ElseIf SBE = "R020" Then
117                           SBE = "sb20k"
118                       ElseIf SBE = "R030" Then
119                           SBE = "sb30k"
120                       ElseIf SBE = "R040" Then
121                           SBE = "sb40k"
122                       ElseIf SBE = "R050" Then
123                           SBE = "sb50k"
124                       Else
125                           SBE = "DEPOSIT"
126                       End If
127                       If Not rsTemp.EOF Then
128                           sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(Text(15)) & " WHERE KodeKategory = '" & Left(NulSID(rsModel!MDL), 20) & "' "
129                       Else
130                           sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & ",tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(rsModel!MDL), 20) & "'," & k2t(Text(15)) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
131                       End If
132                       Call objHT.InsertData_Ho(sSQL)
133                   End If
134                   rsModel.MoveNext
135               Loop
136               Call objHT.CloseTransactionHO
137               If pesan <> "" Then
138                 MsgBox pesan, vbInformation
                    
139               End If
Lanjutkan:
140               Call ber_sih
141               Call buka_Tutup(False, False)
142               Call refreshGrid
143               cmdUbah.Caption = "&Ubah"
144               cmdSimpan.Caption = "&Tambah"
145           Else
146               MsgBox pesan, vbCritical, "VALID DATA"
147               Exit Sub
148           End If
149       End If

150      On Error GoTo 0
151      Exit Sub

cmdSimpan_Click_Error:
152       If isBegin Then
153           Call objHT.CloseTransactionHO_err
154       End If
155       MsgBox "Error " & Erl & " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmHOProgramFOP"
End Sub

Private Sub cmdSimpanSub_Click()
    Dim isBegin     As Boolean
   On Error GoTo cmdSimpanSub_Click_Error

    isBegin = False
    boleh = True
    
    If CDbl(Text(13)) <= 0 Then
        pesan = pesan + "Jumlah item (QTY) substitusi "
        Text(13) = 0
        Text(13).SetFocus
        boleh = False
    End If
    
    If Text(12) = "" Then
        boleh = False
        pesan = pesan + "Part substitusi. "
    End If
    
    If boleh Then
        Call objHT.OpenTransactionHO
        isBegin = True
        sSQL = "SELECT * FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' AND  MDL = '" & NulSID(rsGrid.Fields(4)) & "' AND ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' AND QTY = '" & NulSID(rsGrid.Fields(9)) & "' AND ITEM_PARTSUB = '" & NulSID(Text(12)) & "'"
        Set rsTemp = objHT.FetchDataHO(sSQL)
        If Not rsTemp.EOF Then
            sSQL = "UPDATE HO_FOP_PROGRAM_SUB SET ITEM_PARTSUB = '" & NulSID(Text(12)) & "',NAMA_PARTSUB = '" & NulSID(Text(11)) & "',QTYSUB = '" & NulSID(Text(13)) & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' AND  MDL = '" & NulSID(rsGrid.Fields(4)) & "' AND ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' AND QTY = '" & NulSID(rsGrid.Fields(9)) & "' AND ITEM_PARTSUB = '" & NulSID(Text(12)) & "'"
        Else
            sSQL = "INSERT INTO HO_FOP_PROGRAM_SUB (FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,ITEM_PARTSUB,NAMA_PARTSUB,QTYSUB) VALUES ('" & NulSID(rsGrid.Fields(6)) & "','" & NulSID(rsGrid.Fields(4)) & "','" & NulSID(rsGrid.Fields(5)) & "','" & NulSID(rsGrid.Fields(7)) & "','" & NulSID(rsGrid.Fields(8)) & "','" & NulSID(rsGrid.Fields(9)) & "','" & NulSID(Text(12)) & "','" & NulSID(Text(11)) & "','" & NulSID(Text(13)) & "')"
        End If
        Call objHT.InsertData_Ho(sSQL)
        Call objHT.CloseTransactionHO
        
        Text(11) = ""
        Text(12) = ""
        Text(13) = 0
        gridSub.ClearFields
        Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(8)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and MDL = '" & NulSID(rsGrid.Fields(4)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
        gridSub.DataSource = rsGridSub
        Call HeaderColomSubstitusi
        gridSub.ReBind
    Else
        MsgBox pesan, vbCritical, "VALID DATA"
        Exit Sub
    End If

   On Error GoTo 0
   Exit Sub

cmdSimpanSub_Click_Error:
    If isBegin Then
        Call objHT.CloseTransactionHO_err
    End If
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSimpanSub_Click of Form frmHOProgramFOP"
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        newRecord = False
        Call buka_Tutup(False, True)
        Call buka_tutupSubtitusi(True)
        KDPART_LAMA = Text(1)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        Call buka_tutupSubtitusi(False)
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
    End If
End Sub

Private Sub Form_Load()
'    Call ber_sih
    Call buka_Tutup(False, False)
    Call buka_tutupSubtitusi(False)
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridPart.Width = Frame1.Width - 200
    GridPart.Height = Frame1.Height - Text(7).Top - Text(7).Height - 200
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdGanti.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    ProgBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    cmdImport.Left = cmdCetak.Left - cmdImport.Width - 100
    ProgBar.Left = cmdImport.Left - ProgBar.Width - 100
End Sub

Private Sub GridPart_FilterChange()
    On Error GoTo errHandler
    Set cols = GridPart.Columns
    Dim C As Integer
    
    C = GridPart.col
    GridPart.HoldFields
    rsGrid.filter = getFilter()
    GridPart.col = C
    GridPart.EditActive = True
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
 
Private Sub GridPart_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
1         If Not rsGrid.EOF Then
2             Text(5) = rsGrid.Fields(4)      ''model kendaraan   'GridPart.Columns(0)
3             Text(6) = rsGrid.Fields(5)      ''nama model        'GridPart.Columns(1)
4             Text(0) = rsGrid.Fields(6)      ''job order         'GridPart.Columns(2)
5             Text(1) = rsGrid.Fields(7)      ''kode part         'GridPart.Columns(3)
6             Text(2) = rsGrid.Fields(8)      ''nama part         'GridPart.Columns(4)
7             Text(3) = rsGrid.Fields(10)     ''keterangan        'GridPart.Columns(6)
8             Text(7) = rsGrid.Fields(9)      ''qty               'GridPart.Columns(5)
              ''chPart.Enabled = IIf(rsGrid.Fields(7).value, vbChecked, vbUnchecked)
              
9             Text(14) = rsGrid.Fields(1)     '' NIK TAHUN
10            tglRetailAwal.value = Format(rsGrid.Fields(2), "")
11            tglRetailAkhir.value = Format(rsGrid.Fields(3), "")
12            cbTagihan.Text = rsGrid.Fields(0)
              Dim SBE As String
              
13            If cbTagihan.Text = "NRM" Then
14                Label15.Visible = True
15                Text(15).Visible = True
16                SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "'")
17                If SBE = "R010" Then
18                    SBE = "sb10k"
19                ElseIf SBE = "R020" Then
20                    SBE = "sb20k"
21                ElseIf SBE = "R030" Then
22                    SBE = "sb30k"
23                ElseIf SBE = "R040" Then
24                    SBE = "sb40k"
25                ElseIf SBE = "R050" Then
26                    SBE = "sb50k"
27                Else
28                    SBE = "DEPOSIT"
29                End If
30                Text(15) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(rsGrid.Fields(4)) & "'")
31            Else
32                Label15.Visible = False
33                Text(15).Visible = False
                          
34            End If
              
35            gridSub.ClearFields
36            Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(8)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and MDL = '" & NulSID(rsGrid.Fields(4)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
              ''If Not rsGridSub.EOF Then
37                Text(12) = ""
38                Text(11) = ""
39                Text(13) = 0
                                  
40                gridSub.DataSource = rsGridSub
41                Call HeaderColomSubstitusi
42                gridSub.ReBind
              ''End If
43        End If
End Sub

 
Private Sub gridSub_FilterChange()
On Error GoTo errHandler
    Set cols = gridSub.Columns
    Dim C As Integer
    
    C = gridSub.col
    gridSub.HoldFields
    rsGridSub.filter = getFilter()
    gridSub.col = C
    gridSub.EditActive = True
    Exit Sub
    
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub gridSub_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGridSub.EOF Then
        Text(12) = rsGridSub.Fields(0)  ' kodeitem
        Text(11) = rsGridSub.Fields(1)  ' nama part
        Text(13) = rsGridSub.Fields(2)  ' qty
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
                gSearchOption = "JOB_ORDER_FOP"
                gBantuan = "SBE"    '/ "ERC" '/
                frmCariHO.show (vbModal)
            Case 1, 8, 12
                ' PART DARI HO
                gMainKeyCari = Text(Index)
                gSearchOption = "PARTHO"
                frmCariHO.show (vbModal)
            Case 5
                Cancel = True
                gMainKeyCari = Text(Index)
                gSearchOption = "MDL"
                frmCariHO.show (vbModal)
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
    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("SELECT * FROM ho_fop_job WITH (NOLOCK) WHERE FS_Item = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JOB_ORDER_FOP"
                    gBantuan = "SBE"
                    frmCariHO.show (vbModal)
                End If
                oRs.Close
            End If
        Case 1, 8
            If Text(Index) <> "" Then
            ' PART MANUAL INPUTAN HO
                Set oRs = objO.FetchDataHO("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & Text(Index) & "'")
                If Not oRs.BOF Then
                    Text(4) = Trim(oRs("NAMA_PART"))
                End If
                oRs.Close
            End If
        Case 12
            If Text(Index) <> "" Then
            ' PART MANUAL INPUTAN HO
                Set oRs = objO.FetchDataHO("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & Text(Index) & "'")
                If Not oRs.BOF Then
                    Text(11) = Trim(oRs("NAMA_PART"))
                End If
                oRs.Close
            End If
        Case 5
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL like '" & Text(5) & "%' and FS_LOG_DEL = ''")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "MDL"
                    frmCariHO.show (vbModal)
                Else
                    Text(6) = oRs("Desc")
                End If
                oRs.Close
            End If
    End Select
End Sub
