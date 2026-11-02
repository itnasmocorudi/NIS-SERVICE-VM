VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOProgramFOPNew 
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
      TabIndex        =   9
      Top             =   360
      Width           =   14055
      Begin VB.Frame frameNikKendaraan 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kode NIK Kendaraan"
         Height          =   3615
         Left            =   14400
         TabIndex        =   38
         Top             =   1080
         Visible         =   0   'False
         Width           =   5415
         Begin TrueOleDBGrid70.TDBGrid GridNik 
            Height          =   3015
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   5318
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE"
            Columns(0).DataField=   "kode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "TAHUN"
            Columns(1).DataField=   "tahun"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   4
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "OK"
            Columns(2).DataField=   "IKUT"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   3
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=3"
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
            MultipleLines   =   0
            CellTipsWidth   =   0
            MultiSelect     =   2
            DeadAreaBackColor=   14737632
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=112,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
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
         Begin AIFCmp1.asxPowerButton cmdTutup 
            Height          =   360
            Left            =   120
            TabIndex        =   40
            Top             =   3240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "&Tutup"
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
      End
      Begin AIFCmp1.asxPowerButton cmdTampilkanKodeNik 
         Height          =   285
         Left            =   1440
         TabIndex        =   37
         Top             =   1125
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Pilih NIK"
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbKodeNik 
         Height          =   285
         Left            =   1440
         TabIndex        =   36
         Top             =   1125
         Visible         =   0   'False
         Width           =   1455
         AllowInput      =   0   'False
         _Version        =   196616
         BackColorOdd    =   8438015
         Columns(0).Width=   3200
         _ExtentX        =   2566
         _ExtentY        =   503
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin TDBDate6Ctl.TDBDate tglRetailAwal 
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1420
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         Calendar        =   "frmHOProgramFOPNew.frx":0000
         Caption         =   "frmHOProgramFOPNew.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOProgramFOPNew.frx":0198
         Keys            =   "frmHOProgramFOPNew.frx":01B6
         Spin            =   "frmHOProgramFOPNew.frx":0214
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
         ItemData        =   "frmHOProgramFOPNew.frx":023C
         Left            =   1440
         List            =   "frmHOProgramFOPNew.frx":0246
         Style           =   2  'Dropdown List
         TabIndex        =   3
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
         Left            =   11400
         TabIndex        =   28
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame frmeGanti 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   960
         TabIndex        =   18
         Top             =   3000
         Visible         =   0   'False
         Width           =   10335
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   19
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
            TabIndex        =   20
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   23
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
            Picture         =   "frmHOProgramFOPNew.frx":0254
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
            TabIndex        =   25
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
            TabIndex        =   24
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
         Index           =   1
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
         Index           =   2
         Left            =   3960
         TabIndex        =   2
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
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   1125
         Visible         =   0   'False
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
         TabIndex        =   6
         Top             =   1420
         Width           =   1095
         _Version        =   65536
         _ExtentX        =   1931
         _ExtentY        =   503
         Calendar        =   "frmHOProgramFOPNew.frx":3527
         Caption         =   "frmHOProgramFOPNew.frx":3653
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOProgramFOPNew.frx":36BF
         Keys            =   "frmHOProgramFOPNew.frx":36DD
         Spin            =   "frmHOProgramFOPNew.frx":373B
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
         Index           =   5
         Left            =   4485
         TabIndex        =   34
         Top             =   1860
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
      Begin AIFCmp1.asxPowerButton cmdRefresh 
         Height          =   360
         Left            =   4200
         TabIndex        =   7
         Top             =   1345
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Tampilkan Paket Item"
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
      Begin TrueOleDBGrid70.TDBGrid gridPart 
         Height          =   4095
         Left            =   120
         TabIndex        =   8
         Top             =   2280
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   7223
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Terlibat"
         Columns(0).DataField=   "terlibat"
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
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=20"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=20"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3175"
         Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=20"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=3281"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=3175"
         Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=20"
         Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(21)=   "Column(4).Width=3281"
         Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=3175"
         Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=20"
         Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=91,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=88,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=89,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=90,.parent=17"
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
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=28,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=25,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=26,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=27,.parent=17"
         _StyleDefs(56)  =   "Named:id=33:Normal"
         _StyleDefs(57)  =   ":id=33,.parent=0,.valignment=2"
         _StyleDefs(58)  =   "Named:id=34:Heading"
         _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   ":id=34,.wraptext=-1"
         _StyleDefs(61)  =   "Named:id=35:Footing"
         _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=36:Selected"
         _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=37:Caption"
         _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(67)  =   "Named:id=38:HighlightRow"
         _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(69)  =   "Named:id=39:EvenRow"
         _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(71)  =   "Named:id=40:OddRow"
         _StyleDefs(72)  =   ":id=40,.parent=33"
         _StyleDefs(73)  =   "Named:id=41:RecordSelector"
         _StyleDefs(74)  =   ":id=41,.parent=34"
         _StyleDefs(75)  =   "Named:id=42:FilterBar"
         _StyleDefs(76)  =   ":id=42,.parent=33"
         _StyleDefs(77)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(78)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(79)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(80)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         _StyleDefs(81)  =   "Named:id=87:SudahDitagih"
         _StyleDefs(82)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(83)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(84)  =   ":id=87,.fontname=Tahoma"
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   375
         Left            =   11400
         Top             =   1680
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   661
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
         Left            =   3840
         TabIndex        =   35
         Top             =   1905
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         Left            =   10305
         TabIndex        =   29
         Top             =   270
         Visible         =   0   'False
         Width           =   960
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
         TabIndex        =   15
         Top             =   525
         Width           =   1155
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
         TabIndex        =   11
         Top             =   225
         Width           =   720
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
      TabIndex        =   10
      Top             =   7080
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
   Begin AIFCmp1.asxPowerButton cmdUbah 
      Height          =   360
      Left            =   8880
      TabIndex        =   12
      Top             =   7080
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Batal"
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
      TabIndex        =   13
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
      TabIndex        =   14
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
      Left            =   6720
      TabIndex        =   16
      Top             =   7080
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Template"
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
      TabIndex        =   17
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
      TabIndex        =   26
      Top             =   7080
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BackColor       =   12648447
      Caption         =   "Import Excel"
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
      TabIndex        =   27
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
Attribute VB_Name = "frmHOProgramFOPNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOProgramFOPNew
' DateTime  : 26/05/2020 10:48
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

Dim rsDownload  As New ADODB.Recordset
Dim rsData      As New ADODB.Recordset
Dim rsGrid      As New ADODB.Recordset
Dim rsGridSub   As New ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim rsModel     As ADODB.Recordset
Dim newRecord   As Boolean
Dim KDPART_LAMA As String
Dim boleh       As Boolean
Dim pesan       As String
Dim cols        As TrueOleDBGrid70.Columns
Dim rsNik       As New ADODB.Recordset
Dim daftarNIK   As String

Private Sub ber_sih()
    For i = 0 To 5
        Select Case i
            Case 5
                text(i) = 0
            Case Else
                text(i) = ""
        End Select
    Next
    tglRetailAwal = isnow
    tglRetailAkhir = isnow
    ''cbTagihan.Text = ""
    cbTagihan.Clear
    cbTagihan.AddItem "NRM"
    cbTagihan.AddItem "TAM"
    
    Call refreshComboNIK
    chPart.value = vbUnchecked
    
    Label15.Visible = False
    text(5).Visible = False
    
    cmdSimpan.Enabled = False
    cmdUbah.Enabled = False
    cmdHapus.Enabled = False
    
    gridPart.DataSource = Nothing
End Sub

Private Sub buka_tutupSubtitusi(Flg As Boolean)
     
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    text(0).Enabled = Not Flg
    text(1).Enabled = Not flg1
    text(2).Enabled = Not flg1
    text(3).Enabled = Not flg1
    text(5).Enabled = Not flg1
    chPart.Enabled = Not flg1
    cbKodeNik.Enabled = Not flg1
    
    tglRetailAwal.Enabled = Not flg1
    tglRetailAkhir.Enabled = Not flg1
    cbTagihan.Enabled = Not flg1
     
End Sub

Private Sub refreshGrid()
1         gridPart.ClearFields
2         If Right(daftarNIK, 1) = "," Then
3             daftarNIK = Left(daftarNIK, Len(daftarNIK) - 1)
4         End If
          
          ''sSQL = "select distinct a.* from (select ITEM_PART kode,NAMA_PART nama ,QTY,ISNULL((select '1' from HO_FOP_Program with (nolock) where FS_ITEM = a.FS_ITEM and MDL = a.MDL and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & cbTagihan.Text & "' and ITEM_PART = a.ITEM_PART ),0) as terlibat from HO_PARTPACK a with (nolock) where rekomendasi = 0 and FS_ITEM = '" & NulSID(Text(0)) & "' and MDL = '" & NulSID(Text(1)) & "'"
          ''sSQL = sSQL & " Union"
          ''sSQL = sSQL & " select KDSUBLET,NAMASUBLET,QTY,ISNULL((select '1' from HO_FOP_Program with (nolock) where FS_ITEM = b.FS_ITEM and MDL = b.MDL and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & cbTagihan.Text & "' and ITEM_PART = b.KDSUBLET),0) as terlibat from HO_SUBLETPACK b with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL = '" & NulSID(Text(1)) & "' ) a "
5         sSQL = "select distinct a.* from (select ITEM_PART kode,NAMA_PART nama ,QTY,ISNULL((select '1' from HO_FOP_Program with (nolock) where FS_ITEM = a.FS_ITEM and MDL = a.MDL and NIK = '" & NulSID(cbKodeNik.text) & "' and FOP = '" & cbTagihan.text & "' and ITEM_PART = a.ITEM_PART ),0) as terlibat from HO_PARTPACK a with (nolock) where rekomendasi = 0 and FS_ITEM = '" & NulSID(text(0)) & "' and MDL like '" & NulSID(text(1)) & "%'"
6         sSQL = sSQL & " Union"
7         sSQL = sSQL & " select KDSUBLET,NAMASUBLET,QTY,ISNULL((select '1' from HO_FOP_Program with (nolock) where FS_ITEM = b.FS_ITEM and MDL = b.MDL and NIK = '" & NulSID(cbKodeNik.text) & "' and FOP = '" & cbTagihan.text & "' and ITEM_PART = b.KDSUBLET),0) as terlibat from HO_SUBLETPACK b with (nolock) where FS_ITEM = '" & NulSID(text(0)) & "' and MDL like '" & NulSID(text(1)) & "%' ) a "
          
8         sSQL = "select distinct a.* from (select ITEM_PART kode,NAMA_PART nama ,QTY,"
          ''sSQL = sSQL & " ISNULL((select top 1 '1' from HO_FOP_Program with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL like '" & NulSID(Text(1)) & "%' and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & NulSID(cbTagihan.Text) & "' and ITEM_PART = a.ITEM_PART  and fs_log_del = ''),0) as terlibat"
9         sSQL = sSQL & " ISNULL((select top 1 '1' from HO_FOP_Program with (nolock) where FS_ITEM = '" & NulSID(text(0)) & "' and MDL like '" & NulSID(text(1)) & "%' and NIK IN (" & daftarNIK & ") and FOP = '" & NulSID(cbTagihan.text) & "' and ITEM_PART = a.ITEM_PART  and fs_log_del = ''),0) as terlibat"
10        sSQL = sSQL & " from HO_PARTPACK a with (nolock)"
11        sSQL = sSQL & " where rekomendasi = 0 and FS_ITEM = '" & NulSID(text(0)) & "'  and FS_LOG_DEL = ''"
12        sSQL = sSQL & " and MDL like '" & NulSID(text(1)) & "%'"
13        sSQL = sSQL & " Union"
          ''sSQL = sSQL & " select sub.KDSUBLET,sub.NAMASUBLET,sub.QTY,ISNULL((select top 1 '1' from HO_FOP_Program with (nolock) where FS_ITEM = '" & NulSID(Text(0)) & "' and MDL like '" & NulSID(Text(1)) & "%'  and NIK = '" & NulSID(cbKodeNik.Text) & "' and FOP = '" & NulSID(cbTagihan.Text) & "' and ITEM_PART = sub.KDSUBLET and fs_log_del = ''),0) as terlibat  from ("
14        sSQL = sSQL & " select sub.KDSUBLET,sub.NAMASUBLET,sub.QTY,ISNULL((select top 1 '1' from HO_FOP_Program with (nolock) where FS_ITEM = '" & NulSID(text(0)) & "' and MDL like '" & NulSID(text(1)) & "%'  and NIK in (" & daftarNIK & ") and FOP = '" & NulSID(cbTagihan.text) & "' and ITEM_PART = sub.KDSUBLET and fs_log_del = ''),0) as terlibat  from ("
15        sSQL = sSQL & " select b.KDSUBLET,c.NAMASUBLET,max(b.QTY) QTY"
16        sSQL = sSQL & " from HO_SUBLETPACK b with (nolock) inner join MSTSUBLET c on b.KDSUBLET = c.KDSUBLET  and c.FS_LOG_DEL = '' where b.FS_ITEM = '" & NulSID(text(0)) & "' and b.MDL like '" & NulSID(text(1)) & "%' and b.FS_LOG_DEL = ''"
17        sSQL = sSQL & " group by b.KDSUBLET,c.NAMASUBLET ) sub"
18        sSQL = sSQL & " ) a"
          
19        Set rsGrid = objO.FetchDataHO(sSQL)
20        Call createRSData
          ''Call HeaderColom
21        If Not rsGrid.EOF Then
22            rsGrid.MoveFirst
23            Do While Not rsGrid.EOF
                  
24                With rsData
25                    .AddNew
26                    !kode = NulSID(rsGrid!kode)
27                    !Nama = NulSID(rsGrid!Nama)
28                    !qty = NulSID(rsGrid!qty)
29                    !qtyasli = NulSID(rsGrid!qty)
30                    !terlibat = IIf(NulSID(rsGrid!terlibat) = 1, True, False)
31                    .Update
32                End With
                  'Debug.Print gridPart.Splits(0).Columns(0).value
                  'Call createRSData
33                rsGrid.MoveNext
34            Loop
35        End If
         ' gridPart.Splits(0).Columns(0).ValueItems.Presentation = dbgCheckBox
         '' Set gridPart.DataSource = rsData
36        Call HeaderColom
          ''gridPart.ReBind
End Sub

Private Sub createRSData()
    If rsData.State = adStateOpen Then rsData.Close
    With rsData
        .Fields.Append "terlibat", adBoolean
        .Fields.Append "kode", adVarChar, 50
        .Fields.Append "nama", adVarChar, 200
        .Fields.Append "qty", adInteger
        .Fields.Append "qtyAsli", adInteger
        .Open
    End With
    
    gridPart.DataSource = rsData
    ''Call HeaderColom
End Sub


Private Sub HeaderColom()
    With gridPart
        .Columns(0).Caption = "Terlibat"
        .Columns(0).Width = 1000
        .Columns(0).Alignment = dbgCenter
        .Columns(0).DataField = "terlibat"
        .Splits(0).Columns(0).Locked = False
        .Splits(0).Columns(0).ValueItems.Presentation = dbgCheckBox
        
        .Columns(1).Caption = "Kode Item"
        .Columns(1).DataField = "kode"
        .Columns(1).Width = 2000
        .Columns(1).Alignment = dbgLeft
        .Splits(0).Columns(1).Locked = True
        
        .Columns(2).Caption = "Nama Item"
        .Columns(2).Width = 4000
        .Columns(2).DataField = "nama"
        .Columns(2).Alignment = dbgLeft
        .Splits(0).Columns(2).Locked = True
        
        .Columns(3).Caption = "Qty"
        .Columns(3).Width = 500
        .Columns(3).DataField = "QTY"
        .Columns(3).Alignment = dbgRight
        .Columns(3).NumberFormat = "#,#0"
        .Columns(3).Locked = False
        
        .Columns(4).Caption = "Qty Asli"
        .Columns(4).Width = 500
        .Columns(4).DataField = "QTYasli"
        .Columns(4).Visible = False
    End With
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
    If cbTagihan.text = "NRM" Then
        Label15.Visible = True
        text(5).Visible = True
        SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "'")
        If SBE = "R010" Then
            SBE = "sb10k"
        ElseIf SBE = "R020" Then
            SBE = "sb20k"
        ElseIf SBE = "R030" Then
            SBE = "sb30k"
        ElseIf SBE = "R040" Then
            SBE = "sb40k"
        ElseIf SBE = "R050" Then
            SBE = "sb50k"
        Else
            SBE = "DEPOSIT"
        End If
        text(15) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(rsGrid.Fields(4)) & "'")
    Else
        Label15.Visible = False
        text(5).Visible = False
    End If
End Sub

Private Sub cbTagihan_Click()
    If cbTagihan.text = "NRM" Then
        Label15.Visible = True
        text(5).Visible = True
        SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(text(0)) & "'")
        If SBE = "R010" Then
            SBE = "sb10k"
        ElseIf SBE = "R020" Then
            SBE = "sb20k"
        ElseIf SBE = "R030" Then
            SBE = "sb30k"
        ElseIf SBE = "R040" Then
            SBE = "sb40k"
        ElseIf SBE = "R050" Then
            SBE = "sb50k"
        Else
            SBE = "DEPOSIT"
        End If
        'Text(5) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(Text(1)) & "'")
        text(5) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory like '" & NulSID(text(1)) & "%'")

    Else
        Label15.Visible = False
        text(5).Visible = False
    End If
End Sub

Private Sub cmdCetak_Click()

   On Error GoTo cmdCetak_Click_Error
    Dim Er As New ErrInfo
    If rsDownload.State Then rsDownload.Close
    Set rsDownload = Nothing
    With rsDownload
        .Fields.Append "NAMAPEKERJAAN", adVarChar, 50
        .Fields.Append "KODEMODEL", adVarChar, 50
        .Fields.Append "TAGIHAN", adVarChar, 50
        .Fields.Append "KODENIK", adVarChar, 50
        .Fields.Append "TGLRETAILAWAL", adVarChar, 50
        .Fields.Append "TGLRETAILAKHIR", adVarChar, 50
        .Fields.Append "KODEITEM", adVarChar, 50
        .Fields.Append "NOMINALDEPOSITO", adVarChar, 50
        .Open
 
    End With
      
    Dim MyPath As String
    Dim onVal As String
    MyPath = ""
    Dim f As New FormExplorer
    f.show (vbModal)
    MyPath = f.PathSaving
    If MyPath = "" Then Exit Sub
    If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
    If f.FileName = "" Then Exit Sub
    onVal = MyPath & f.FileName & ".xls"
    Call ProsesExport(onVal)
            Er.HeadCap = "Export to excel !"
            Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                         "File : " & onVal
            
            Beep
            Er.show (vbModal)
    
    
    

   On Error GoTo 0
   Exit Sub

cmdCetak_Click_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdCetak_Click of Form frmHOProgramFOPNew"


End Sub

Private Sub cmdGanti_Click()
    text(8) = ""
    text(4) = ""
    text(10) = ""
    text(9) = ""
    frmeGanti.Visible = True
End Sub

Private Sub cmdHapus_Click()
    Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_fop_program WITH (NOLOCK) WHERE FOP = '" & cbTagihan.text & "' and NIK = '" & NulSID(cbKodeNik.text) & "' and fs_item = '" & text(0) & "' AND mdl like '" & text(1) & "%' and FS_LOG_DEL = ''")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus data " & rsTemp.RecordCount & " model kendaraan ?" & vbCr & "Untuk Job Order : " & text(0), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE ho_fop_program SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE FOP = '" & cbTagihan.text & "' and NIK = '" & NulSID(cbKodeNik.text) & "' and fs_item = '" & text(0) & "' AND mdl like '" & text(1) & "%' and FS_LOG_DEL = ''")
            ''Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('PARTPACK', '" & Text(0) & "', 'D')")
        End If
        Call ber_sih
        ''Call refreshGrid
    Else
        MsgBox "Pilih data yang ingin dihapus"
        Exit Sub
    End If
End Sub

Private Sub cmdHapusSub_Click()
    Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and mdl = '" & NulSID(rsGrid.Fields(4)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "'  and ITEM_PARTSUB = '" & NulSID(rsGridSub.Fields(0)) & "' and NAMA_PARTSUB = '" & NulSID(rsGridSub.Fields(1)) & "' and QTYSUB = '" & NulSID(rsGridSub.Fields(2)) & "' and FS_LOG_DEL = ''")
    If Not rsTemp.EOF Then
        temp = MsgBox("Yakin akan menghapus data sub part : " & NulSID(rsGridSub.Fields(0)) & " ?" & vbCr & "Untuk Job Order : " & text(0), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.UpdateDataHO("UPDATE HO_FOP_PROGRAM_SUB SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and mdl = '" & NulSID(rsGrid.Fields(4)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "'  and ITEM_PARTSUB = '" & NulSID(rsGridSub.Fields(0)) & "' and NAMA_PARTSUB = '" & NulSID(rsGridSub.Fields(1)) & "' and QTYSUB = '" & NulSID(rsGridSub.Fields(2)) & "' and FS_LOG_DEL = ''")
        End If
        text(11) = ""
        text(12) = ""
        text(13) = 0
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
          Dim tglAwal  As Date
          Dim tglAkhir As Date
          Dim kodePart As String
          Dim namaPart As String
          Dim qty      As Double
          Dim isBegin  As Boolean
          Dim SBE As String
          Dim namaPekerjaan As String
          Dim itemPekerjaan As String
2         isBegin = False
         
        ' Dim konter As Integer
         
3       konter = 0
4       ProgBar.Min = 0
5       CommonDialog1.FileName = ""
6       CommonDialog1.InitDir = App.Path
7       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
8       CommonDialog1.ShowOpen
9       If CommonDialog1.FileName <> "" Then
10         If CnExel.State Then CnExel.Close
11         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
12         GoSub Proses
13     Else
14        Exit Sub
15     End If
16     Exit Sub
       
Proses:
17    ErrBegin = True
         Dim VarNmFile As String
         Dim model As String
         
18        gBantuan = "SBE"
19       If RcExcel.State Then RcExcel.Close
20       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
21       If RcExcel.RecordCount > 0 Then
22            Call objHT.OpenTransactionHO
23            isBegin = True

24          ProgBar.Visible = True
25          ProgBar.Max = RcExcel.RecordCount
26          RcExcel.MoveFirst
27          Do While Not RcExcel.EOF
28             DoEvents
29             konter = konter + 1
30             ProgBar.value = konter
31             model = ""
              
32             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
33             If objO.FetchDataHO_Count("select * from HO_FOP_JOB WITH(NOLOCK) WHERE FS_LOG_DEL = '' and FS_ITEM = '" & Trim(RcExcel.Fields(0).value) & "'") = 0 Then
34                Call objHT.CloseTransactionHO_err
35                ProgBar.Visible = False
36                MsgBox "JOB ORDER = " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " salah!", vbCritical, "Error"
37                Exit Sub
38             End If
               
39             If Nul(RcExcel.Fields(1).value) = "" Then
40                Call objHT.CloseTransactionHO_err
41                ProgBar.Visible = False
42                MsgBox "KODE MODEL baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
43                Exit Sub
44             Else
45                Set rsTempExcel = objO.FetchDataHO("SELECT * FROM MODEL WHERE FS_LOG_DEL = '' AND MDL = '" & Trim(RcExcel.Fields(1).value) & "'")
46                If rsTempExcel.RecordCount > 0 Then
47                    model = Trim(rsTempExcel("DESC"))
48                Else
                      'jika tidak ada data
49                    Call objHT.CloseTransactionHO_err
50                    ProgBar.Visible = False
51                    MsgBox "Model baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
52                    Exit Sub
53                End If
54             End If
               
55             If Nul(RcExcel.Fields(2).value) = "" Then
56                Call objHT.CloseTransactionHO_err
57                ProgBar.Visible = False
58                MsgBox "TAGIHAN baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
59                Exit Sub
60             ElseIf Not (NulSID(RcExcel.Fields(2).value) = "TAM" Or NulSID(RcExcel.Fields(2).value) = "NRM") Then
61                Call objHT.CloseTransactionHO_err
62                ProgBar.Visible = False
63                MsgBox "TAGIHAN baris ke-" & konter & " harus NRM / TAM !", vbCritical, "Error"
64                Exit Sub
65             End If
               
               
66             If Nul(RcExcel.Fields(3).value) = "" Then
67                Call objHT.CloseTransactionHO_err
68                ProgBar.Visible = False
69                MsgBox "KODE TAHUN NIK Kendaraan baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
70                Exit Sub
71             Else
72                Set rsTempExcel = objO.FetchDataHO("SELECT  kode,tahun from listTahunNikFOP WHERE kode = '" & Trim(RcExcel.Fields(3).value) & "'")
73                If rsTempExcel.RecordCount = 0 Then
                      'jika tidak ada data
74                    Call objHT.CloseTransactionHO_err
75                    ProgBar.Visible = False
76                    MsgBox "KODE TAHUN NIK baris ke-" & konter & " tidak terdaftar !", vbCritical, "Error"
77                    Exit Sub
78                End If
79             End If
               
               
80            If NulSID(RcExcel.Fields(2).value) = "NRM" Then
81                If Nul(RcExcel.Fields(7).value) = "" Then
82                   Call objHT.CloseTransactionHO_err
83                   ProgBar.Visible = False
84                   MsgBox "Deposit baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
85                   Exit Sub
86                ElseIf IsNumeric(Nul(RcExcel.Fields(7).value)) = False Then
87                   Call objHT.CloseTransactionHO_err
88                   ProgBar.Visible = False
89                   MsgBox "Deposit baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
90                   Exit Sub
91                End If
92            End If
                  
93             tglAwal = "01/01/1900"
94             tglAkhir = "01/01/1900"
95             If NulSID(RcExcel.Fields(4).value) <> "" Then
96                If IsDate(NulSID(RcExcel.Fields(4).value)) Then
97                    tglAwal = Format(NulSID(RcExcel.Fields(4).value), "dd/mm/yyyy")
98                End If
99             End If
               
100            If NulSID(RcExcel.Fields(5).value) <> "" Then
101               If IsDate(NulSID(RcExcel.Fields(5).value)) Then
102                   tglAwal = Format(NulSID(RcExcel.Fields(5).value), "dd/mm/yyyy")
103               End If
104            End If
               
105            kodePart = ""
106            namaPart = ""
107            qty = 0
108            If NulSID(RcExcel.Fields(6).value) = "" Then
109               Call objHT.CloseTransactionHO_err
110               ProgBar.Visible = False
111               MsgBox "ITEM PART / SUBLET baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
112               Exit Sub
113            Else
                  ''cek part terdaftar / tidak
114               sSQL = "select ITEM_PART kode ,NAMA_PART nama,QTY from HO_PARTPACK where FS_ITEM = '" & NulSID(RcExcel.Fields(0).value) & "' and MDL = '" & NulSID(RcExcel.Fields(1).value) & "' and ITEM_PART = '" & NulSID(RcExcel.Fields(6).value) & "' and FS_LOG_DEL = ''"
115               Set rsTemp = objHT.FetchDataHO(sSQL)
116               If rsTemp.RecordCount = 0 Then
                      ''cek sublet terdaftar ./ tidak
117                   sSQL = "select KDSUBLET kode ,NAMASUBLET nama,QTY from HO_SUBLETPACK where FS_ITEM = '" & NulSID(RcExcel.Fields(0).value) & "' and MDL = '" & NulSID(RcExcel.Fields(1).value) & "' and KDSUBLET = '" & NulSID(RcExcel.Fields(6).value) & "' and FS_LOG_DEL = ''"
118                   Set rsTemp = objHT.FetchDataHO(sSQL)
119                   If rsTemp.RecordCount = 0 Then
120                       Call objHT.CloseTransactionHO_err
121                       ProgBar.Visible = False
122                       MsgBox "ITEM PART / SUBLET baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
123                       Exit Sub
124                   Else
125                       GoTo isiDataPartSublet
126                   End If
127               Else
isiDataPartSublet:
128                   kodePart = NulSID(rsTemp!kode)
129                   namaPart = NulSID(rsTemp!Nama)
130                   qty = NulSID(rsTemp!qty)
131               End If
132            End If
              ''Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_FOP_Program WITH (NOLOCK) WHERE FOP = '" & Trim(RcExcel.Fields(2).value) & "' and NIK = '" & Trim(RcExcel.Fields(3).value) & "' and fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(RcExcel.Fields(1).value) & "' AND item_part = '" & Trim(kodePart) & "' and FS_LOG_DEL = ''")
133           Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_FOP_Program WITH (NOLOCK) WHERE NIK = '" & Trim(RcExcel.Fields(3).value) & "' and fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(RcExcel.Fields(1).value) & "' AND item_part = '" & Trim(kodePart) & "' and FS_LOG_DEL = ''")
134           If Not rsTemp.EOF Then
135               sSQL = "UPDATE HO_FOP_Program SET MODEL = '" & NulSID(model) & "',RETAILAWAL = '" & Format(tglAwal, "mm/dd/yyyy") & "',RETAILAKHIR = '" & Format(tglAkhir, "mm/dd/yyyy") & "',NAMA_PART = '" & namaPart & "',QTY = " & k2t(qty) & " WHERE FOP = '" & Trim(RcExcel.Fields(2).value) & "' and NIK = '" & Trim(RcExcel.Fields(3).value) & "' and fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(RcExcel.Fields(1).value) & "' AND item_part = '" & Trim(kodePart) & "' and FS_LOG_DEL = ''"
136           Else
137               sSQL = "INSERT INTO HO_FOP_Program (JOB,FS_ITEM,MDL,MODEL,FOP,NIK,RETAILAWAL,RETAILAKHIR,ITEM_PART,NAMA_PART,QTY,KETERANGAN,USERENTRY,TGLENTRY,FS_LOG_DEL,FD_TGL_DEL) VALUES "
138               sSQL = sSQL & " ('" & gBantuan & "','" & Trim(RcExcel.Fields(0).value) & "','" & Trim(RcExcel.Fields(1).value) & "','" & model & "','" & Trim(RcExcel.Fields(2).value) & "','" & Trim(RcExcel.Fields(3).value) & "','" & Format(tglAwal, "mm/dd/yyyy") & "','" & Format(tglAkhir, "mm/dd/yyyy") & "','" & Trim(kodePart) & "','" & Trim(namaPart) & "'," & k2t(qty) & ",'Import Excel','" & gSA & "',getdate(),'','')"
139           End If
140           Call objHT.InsertData_Ho(sSQL)
              
141           If NulSID(RcExcel.Fields(2).value) = "NRM" Then
                  ''Dim sbe As String
                  
142               Set rsTemp = objHT.FetchDataHO("SELECT * FROM mvp..m_deposit WITH (NOLOCK) WHERE KodeKategory = '" & Left(NulSID(RcExcel.Fields(1).value), 20) & "' ")
143               SBE = objHT.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(RcExcel.Fields(0).value) & "'")
144               If SBE = "R010" Then
145                   SBE = "sb10k"
146               ElseIf SBE = "R020" Then
147                   SBE = "sb20k"
148               ElseIf SBE = "R030" Then
149                   SBE = "sb30k"
150               ElseIf SBE = "R040" Then
151                   SBE = "sb40k"
152               ElseIf SBE = "R050" Then
153                   SBE = "sb50k"
154               Else
155                   SBE = "DEPOSIT"
156               End If
157               If Not rsTemp.EOF Then
158                   If NulSID(SBE) = NulSID("DEPOSIT") Then
159                       sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(RcExcel.Fields(7).value) & " ,DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(RcExcel.Fields(7).value) & " WHERE KodeKategory = '" & Left(NulSID(RcExcel.Fields(1).value), 20) & "' "
160                   Else
161                       itemPekerjaan = ""
162                       Set rsTemp = objHT.FetchDataHO("SELECT * FROM mvp..m_deposit WITH (NOLOCK) WHERE KodeKategory = '" & Left(NulSID(model), 20) & "' and  ket_" & SBE & " like '%" & itemPekerjaan & "%'")
163                       If rsTemp.EOF Then
164                           itemPekerjaan = namaPart
165                           sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(RcExcel.Fields(7).value) & " ,ket_" & SBE & " = CONVERT(VARCHAR(MAX), ket_" & SBE & ") + ' " & itemPekerjaan & "',DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(RcExcel.Fields(7).value) & " WHERE KodeKategory = '" & Left(NulSID(RcExcel.Fields(1).value), 20) & "' "
166                       Else
167                           sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(RcExcel.Fields(7).value) & ",DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(RcExcel.Fields(7).value) & " WHERE KodeKategory = '" & Left(NulSID(RcExcel.Fields(1).value), 20) & "' "
168                       End If
                          
169                   End If
170               Else
171                   If NulSID(SBE) = NulSID("DEPOSIT") Then
172                       sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & ", ket_" & SBE & ",deposit,tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(RcExcel.Fields(1).value), 20) & "'," & k2t(RcExcel.Fields(7).value) & ",'" & itemPekerjaan & "'," & k2t(RcExcel.Fields(7).value) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
173                   Else
174                       sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & " ,deposit,tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(RcExcel.Fields(1).value), 20) & "'," & k2t(RcExcel.Fields(7).value) & " ," & k2t(RcExcel.Fields(7).value) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
175                   End If
176               End If
177               Call objHT.InsertData_Ho(sSQL)
178           End If
              
179            RcExcel.MoveNext
180         Loop
181      End If
182      RcExcel.Close
183      CnExel.Close
      '   ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
184       Call objHT.CloseTransactionHO
185      MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
186      ProgBar.Visible = False
187      Call ber_sih
188      Call buka_Tutup(False, False)
         'Call refreshGrid
         'cmdUbah.Caption = "&Ubah"
         'cmdSimpan.Caption = "&Tambah"
      '   GridAFI.Refresh
189      Return
190      Exit Sub
vb_error:
191      If isBegin = True Then
           ' DbCon.RollbackTrans
192       Call objHT.CloseTransactionHO_err
193      End If
194      ProgBar.Visible = False
195       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdImport_Click of Form frmHOProgramFOPNew"

196      On Error GoTo 0
197      Exit Sub
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdProses_Click()
    Call objO.InsertDataHO("INSERT INTO HO_PARTPACK  (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL) SELECT DISTINCT JOB, FS_ITEM, MDL, MODEL, '" & Trim(text(10)) & "', '" & Trim(text(9)) & "', QTY, KETERANGAN, FS_LOG_DEL, FD_TGL_DEL FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & Trim(text(8)) & "'")
    Call objO.UpdateDataHO("UPDATE HO_PARTPACK SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE ITEM_PART = '" & Trim(text(8)) & "'")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(text(10)) & "', 'S')")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(text(8)) & "', 'D')")
    frmeGanti.Visible = False
End Sub

Private Sub cmdRefresh_Click()
    Dim boleh As Boolean
    Dim pesan As String
    pesan = ""
    boleh = True
    daftarNIK = ""
    
    If cbTagihan.text = "" Then
         pesan = pesan + "Jenis Tagihan. "
         cbTagihan.SetFocus
         boleh = False
    End If
     
'    If NulSID(cbKodeNik.Text) = "" Then ''If Text(3) = "" Then
'        boleh = False
'        pesan = pesan + "Kode NIK Tahun. "
'    End If

    If rsNik.State = adStateClosed Then Call createTbl
    
    If Not rsNik.EOF Then
        rsNik.filter = "ikut=true"
        If rsNik.RecordCount = 0 Then
            boleh = False
            pesan = pesan + "Kode NIK Tahun. "
        Else
            rsNik.MoveFirst
            Do While Not rsNik.EOF
                daftarNIK = daftarNIK & "'" & rsNik!kode & "',"
                
                rsNik.MoveNext
            Loop
        End If
        rsNik.filter = ""
    Else
        boleh = False
        pesan = pesan + "Kode NIK Tahun. "
    End If
    
    If boleh Then
        Call refreshGrid
        
        cmdSimpan.Enabled = True
        cmdUbah.Enabled = True
        cmdHapus.Enabled = True
    Else
        MsgBox pesan, vbCritical, "VALID DATA"
    End If
End Sub

Private Sub cmdSimpan_Click()
          Dim isBegin As Boolean
          Dim SBE As String
          Dim namaPekerjaan As String
          Dim itemPekerjaan As String
          
1        On Error GoTo cmdSimpan_Click_Error

2         isBegin = False
           
3             boleh = True
4             pesan = "Lengkapi data berikut "
              
5             If cbTagihan.text = "" Then
6                 pesan = pesan + "Jenis Tagihan. "
7                 cbTagihan.SetFocus
8                 boleh = False
9             End If
              
      '        If NulSID(cbKodeNik.Text) = "" Then  ''If Text(3) = "" Then
      '            boleh = False
      '            pesan = pesan + "Kode NIK Tahun. "
      '        End If
10            If rsNik.State = adStateClosed Then Call createTbl
          
11            If Not rsNik.EOF Then
12                rsNik.filter = "ikut=true"
13                If rsNik.RecordCount = 0 Then
14                    boleh = False
15                    pesan = pesan + "Kode NIK Tahun. "
16                End If
17            Else
18                boleh = False
19                pesan = pesan + "Kode NIK Tahun. "
20            End If
               
21            If text(0) = "" Then
22                boleh = False
23                pesan = pesan + "Job Order. "
24            End If
25            If text(1) = "" Then
26                boleh = False
27                pesan = pesan + "Model. "
28            ElseIf Len(NulSID(text(1))) < 6 Then
29                boleh = False
30                pesan = pesan + "Model min 6 karakter. "
31            End If
              
32            If rsData.State = adStateClosed Then
33                boleh = False
34                pesan = pesan + "Item part / sublet harus diisi. "
35            Else
36                rsData.MoveFirst
37                rsData.filter = "terlibat=true"
38                If rsData.RecordCount = 0 Then
39                    boleh = False
40                    pesan = pesan + "Item part / sublet harus diisi. "
41                End If
42                rsData.filter = ""
43            End If
              
44            If boleh Then
                  
45                rsNik.filter = "ikut=true"
                  
                  ''Set rsModel = objO.FetchDataHO("SELECT * FROM model WITH (NOLOCK) WHERE MDL LIKE '" & NulSID(Text(1)) & "%'")
                  ''Set rsModel = objO.FetchDataHO("SELECT * FROM model WITH (NOLOCK) WHERE MDL = '" & NulSID(Text(1)) & "'")
                  ''20200625 dibuka kembali untuk model 6 digit
46                Set rsModel = objO.FetchDataHO("SELECT * FROM model WITH (NOLOCK) WHERE MDL LIKE '" & NulSID(text(1)) & "%'")
47                If rsModel.RecordCount > 0 Then
48                    If rsModel.RecordCount = 1 Then
49                        GoTo loopModel
50                    ElseIf MsgBox("Apakah yakin akan menambahkan data ke " & rsModel.RecordCount & " model ?", vbYesNo, "Konfirmasi tambah data") = vbYes Then
51                        GoTo loopModel
52                    Else
53                        Exit Sub
54                    End If
55                Else
56                    MsgBox "Model kendaraan tidak ditemukan." & vbCr & "Silahkan ubah model kendaraan"
57                    Exit Sub
58                End If
loopModel:
59            rsNik.MoveFirst
60            Do While Not rsNik.EOF
61                rsModel.MoveFirst
62                Call objHT.OpenTransactionHO
63                isBegin = True
64                pesan = ""
65                Do While Not rsModel.EOF
66                    rsData.MoveFirst
67                    itemPekerjaan = ""
68                    Do While Not rsData.EOF
69                        If rsData!terlibat Then
70                            namaPekerjaan = ""
71                            namaPekerjaan = objHT.FetchDataHO_String("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE mdl = '" & Left(NulSID(rsModel!MDL), 20) & "' and fs_item = '" & text(0) & "' AND item_part = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = '' ")
72                            If NulSID(namaPekerjaan, "") = "" Then
73                                namaPekerjaan = objHT.FetchDataHO_String("SELECT NAMASUBLET FROM HO_SUBLETPACK WITH (NOLOCK) WHERE mdl = '" & Left(NulSID(rsModel!MDL), 20) & "' and fs_item = '" & text(0) & "' AND KDSUBLET = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = '' ")
74                            End If
75                            If namaPekerjaan <> "" Then
                                  ''Set rsTemp = objHT.FetchDataHO("SELECT * FROM ho_fop_program WITH (NOLOCK) WHERE FOP = '" & cbTagihan.Text & "' and NIK = '" & NulSID(cbKodeNik.Text) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & Left(NulSID(rsModel!MDL), 20) & "' AND item_part = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = ''")
76                                Set rsTemp = objHT.FetchDataHO("SELECT * FROM ho_fop_program WITH (NOLOCK) WHERE NIK = '" & NulSID(cbKodeNik.text) & "' and fs_item = '" & text(0) & "' AND mdl = '" & Left(NulSID(rsModel!MDL), 20) & "' AND item_part = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = ''")
77                                If Not rsTemp.EOF Then
                                      ''Call objHT.InsertData_Ho("UPDATE ho_fop_program SET item_part = '" & NulSID(rsData!kode) & "', nama_part = '" & NulSID(namaPekerjaan) & "' , QTY = '" & CDbl(rsData!qty) & "',FOP = '" & cbTagihan.Text & "',NIK = '" & NulSID(cbKodeNik.Text) & "',RETAILAWAL = '" & Format(tglRetailAwal, "yyyy-mm-dd") & "',RETAILAKHIR = '" & Format(tglRetailAkhir, "yyyy-mm-dd") & "' WHERE FOP = '" & cbTagihan.Text & "' and NIK = '" & NulSID(cbKodeNik.Text) & "' and fs_item = '" & Text(0) & "' AND mdl = '" & Left(NulSID(rsModel!MDL), 20) & "' AND item_part = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = ''")
78                                    Call objHT.InsertData_Ho("UPDATE ho_fop_program SET item_part = '" & NulSID(rsData!kode) & "', nama_part = '" & NulSID(namaPekerjaan) & "' , QTY = '" & CDbl(rsData!qty) & "',FOP = '" & cbTagihan.text & "',NIK = '" & NulSID(rsNik!kode) & "',RETAILAWAL = '" & Format(tglRetailAwal, "yyyy-mm-dd") & "',RETAILAKHIR = '" & Format(tglRetailAkhir, "yyyy-mm-dd") & "' WHERE FOP = '" & cbTagihan.text & "' and NIK = '" & NulSID(rsNik!kode) & "' and fs_item = '" & text(0) & "' AND mdl = '" & Left(NulSID(rsModel!MDL), 20) & "' AND item_part = '" & NulSID(rsData!kode) & "' and FS_LOG_DEL = ''")
79                                Else
80                                    sSQL = "INSERT INTO ho_fop_program (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL,FOP,NIK,RETAILAWAL,RETAILAKHIR,USERENTRY,TGLENTRY) "
                                      ''sSQL = sSQL & " VALUES ('" & gBantuan & "', '" & Text(0) & "','" & Left(NulSID(rsModel!MDL), 20) & "','" & Left(NulSID(rsModel!desc), 50) & "', '" & NulSID(rsData!kode) & "','" & NulSID(namaPekerjaan) & "','" & CDbl(rsData!qty) & "','','','','" & cbTagihan.Text & "','" & NulSID(cbKodeNik.Text) & "','" & Format(tglRetailAwal, "yyyy-mm-dd") & "','" & Format(tglRetailAkhir, "yyyy-mm-dd") & "','" & gSA & "',getdate())"
81                                    sSQL = sSQL & " VALUES ('" & gBantuan & "', '" & text(0) & "','" & Left(NulSID(rsModel!MDL), 20) & "','" & Left(NulSID(rsModel!desc), 50) & "', '" & NulSID(rsData!kode) & "','" & NulSID(namaPekerjaan) & "','" & CDbl(rsData!qty) & "','','','','" & cbTagihan.text & "','" & NulSID(rsNik!kode) & "','" & Format(tglRetailAwal, "yyyy-mm-dd") & "','" & Format(tglRetailAkhir, "yyyy-mm-dd") & "','" & gSA & "',getdate())"
82                                    Call objHT.InsertData_Ho(sSQL)
83                                End If
84                                itemPekerjaan = itemPekerjaan & " " & NulSID(namaPekerjaan)
85                            Else
86                                pesan = pesan & "Item " & NulSID(rsData!kode) & " tidak terdaftar untuk model " & Left(NulSID(rsModel!MDL), 20) & vbCrLf
87                            End If
88                        End If
89                        rsData.MoveNext
90                    Loop
                      ''cek mvp m_deposit
                      ''--------------------------------------------------------------
91                    If cbTagihan.text = "NRM" Then
                          ''Dim sbe As String
                          
92                        Set rsTemp = objHT.FetchDataHO("SELECT * FROM mvp..m_deposit WITH (NOLOCK) WHERE KodeKategory = '" & Left(NulSID(rsModel!MDL), 20) & "' ")
93                        SBE = objHT.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(text(0)) & "'")
94                        If SBE = "R010" Then
95                            SBE = "sb10k"
96                        ElseIf SBE = "R020" Then
97                            SBE = "sb20k"
98                        ElseIf SBE = "R030" Then
99                            SBE = "sb30k"
100                       ElseIf SBE = "R040" Then
101                           SBE = "sb40k"
102                       ElseIf SBE = "R050" Then
103                           SBE = "sb50k"
104                       Else
105                           SBE = "DEPOSIT"
106                       End If
107                       If Not rsTemp.EOF Then
108                           If NulSID(SBE) = NulSID("DEPOSIT") Then
109                               sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(text(5)) & " ,DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(text(5)) & " WHERE KodeKategory = '" & Left(NulSID(rsModel!MDL), 20) & "' "
110                           Else
111                               sSQL = "UPDATE MVP..M_DEPOSIT SET " & SBE & " = " & k2t(text(5)) & " ,ket_" & SBE & " = '" & itemPekerjaan & "',DEPOSIT = (DEPOSIT - " & SBE & ")+ " & k2t(text(5)) & " WHERE KodeKategory = '" & Left(NulSID(rsModel!MDL), 20) & "' "
112                           End If
113                       Else
114                           If NulSID(SBE) = NulSID("DEPOSIT") Then
115                               sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & ", ket_" & SBE & ",deposit,tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(rsModel!MDL), 20) & "'," & k2t(text(5)) & ",'" & itemPekerjaan & "'," & k2t(text(5)) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
116                           Else
117                               sSQL = "INSERT INTO MVP..M_DEPOSIT (KodeKategory," & SBE & " ,deposit,tgl_entry,ip_entry,id_user_entry) VALUES ('" & Left(NulSID(rsModel!MDL), 20) & "'," & k2t(text(5)) & " ," & k2t(text(5)) & ",getdate(),'" & Left(ComptName, 50) & "',0)"
118                           End If
119                       End If
120                       Call objHT.InsertData_Ho(sSQL)
121                   End If
122                   rsModel.MoveNext
123               Loop
124               rsNik.MoveNext
125           Loop
126               Call objHT.CloseTransactionHO
127               If pesan <> "" Then
128                   MsgBox pesan, vbInformation
129               End If
      '            If pesan <> "" Then
      '              MsgBox pesan, vbInformation
      '
      '            End If
Lanjutkan:
130               Call ber_sih
131               Call buka_Tutup(False, False)
                  'Call refreshGrid
132           Else
133               MsgBox pesan, vbCritical, "VALID DATA"
134               Exit Sub
135           End If

136      On Error GoTo 0
137      Exit Sub

cmdSimpan_Click_Error:
138       If isBegin Then
139           Call objHT.CloseTransactionHO_err
140       End If
141       MsgBox "Error " & Erl & " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmHOProgramFOPNew"
End Sub

Private Sub cmdSimpanSub_Click()
    Dim isBegin     As Boolean
   On Error GoTo cmdSimpanSub_Click_Error

    isBegin = False
    boleh = True
    
    If CDbl(text(13)) <= 0 Then
        pesan = pesan + "Jumlah item (QTY) substitusi "
        text(13) = 0
        text(13).SetFocus
        boleh = False
    End If
    
    If text(12) = "" Then
        boleh = False
        pesan = pesan + "Part substitusi. "
    End If
    
    If boleh Then
        Call objHT.OpenTransactionHO
        isBegin = True
        sSQL = "SELECT * FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' AND  MDL = '" & NulSID(rsGrid.Fields(4)) & "' AND ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' AND QTY = '" & NulSID(rsGrid.Fields(9)) & "' AND ITEM_PARTSUB = '" & NulSID(text(12)) & "'"
        Set rsTemp = objHT.FetchDataHO(sSQL)
        If Not rsTemp.EOF Then
            sSQL = "UPDATE HO_FOP_PROGRAM_SUB SET ITEM_PARTSUB = '" & NulSID(text(12)) & "',NAMA_PARTSUB = '" & NulSID(text(11)) & "',QTYSUB = '" & NulSID(text(13)) & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' AND  MDL = '" & NulSID(rsGrid.Fields(4)) & "' AND ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' AND QTY = '" & NulSID(rsGrid.Fields(9)) & "' AND ITEM_PARTSUB = '" & NulSID(text(12)) & "'"
        Else
            sSQL = "INSERT INTO HO_FOP_PROGRAM_SUB (FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,ITEM_PARTSUB,NAMA_PARTSUB,QTYSUB) VALUES ('" & NulSID(rsGrid.Fields(6)) & "','" & NulSID(rsGrid.Fields(4)) & "','" & NulSID(rsGrid.Fields(5)) & "','" & NulSID(rsGrid.Fields(7)) & "','" & NulSID(rsGrid.Fields(8)) & "','" & NulSID(rsGrid.Fields(9)) & "','" & NulSID(text(12)) & "','" & NulSID(text(11)) & "','" & NulSID(text(13)) & "')"
        End If
        Call objHT.InsertData_Ho(sSQL)
        Call objHT.CloseTransactionHO
        
        text(11) = ""
        text(12) = ""
        text(13) = 0
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
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSimpanSub_Click of Form frmHOProgramFOPNew"
End Sub

Private Sub cmdTampilkanKodeNik_Click()
    Call refreshNikKendaraan
    frameNikKendaraan.Visible = True
    frameNikKendaraan.Left = 1440
End Sub


Private Sub createTbl()
    
    If rsNik.State Then rsNik.Close
    Set rsNik = Nothing
    With rsNik
        .Fields.Append "kode", adVarChar, 15
        .Fields.Append "tahun", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridNik.DataSource = rsNik
    GridNik.AlternatingRowStyle = True
End Sub

Private Sub refreshNikKendaraan()
    Set rsTemp = objO.FetchDataHO("select kode,tahun from listTahunNikFOP with (nolock)")
    Call createTbl
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            rsNik.AddNew
            rsNik!kode = Trim(rsTemp("kode"))
            rsNik!Tahun = Trim(rsTemp("tahun"))
            rsNik!ikut = False
            rsNik.Update
            rsTemp.MoveNext
        Loop

        rsNik.MoveFirst
        GridNik.Columns(0).Width = 1200
        GridNik.Columns(1).Width = 2000
        GridNik.Columns(2).Alignment = dbgCenter
        GridNik.Columns(2).Width = 800
    End If
End Sub

Private Sub cmdTutup_Click()
    frameNikKendaraan.Visible = False
    frameNikKendaraan.Left = 14400
End Sub

Private Sub cmdUbah_Click()
    Call cmdRefresh_Click
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False, False)
    Call buka_tutupSubtitusi(False)
    
    ''Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    gridPart.Width = Frame1.Width - 200
    gridPart.Height = Frame1.Height - text(5).Top - text(5).Height - 200
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

Private Sub gridPart_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 3 Then
        If gridPart.Columns(ColIndex).value > rsData!qtyasli Then
            gridPart.Columns(ColIndex).value = rsData!qtyasli
            rsData!qty = rsData!qtyasli
        ElseIf gridPart.Columns(ColIndex).value < 1 Then
            gridPart.Columns(ColIndex).value = rsData!qtyasli
            rsData!qty = rsData!qtyasli
        End If
    End If
   ' gridPart.ReBind
End Sub

Private Sub GridPart_FilterChange()
    On Error GoTo errHandler
    Set cols = gridPart.Columns
    Dim C As Integer
    
    C = gridPart.col
    gridPart.HoldFields
    rsGrid.filter = getFilter()
    gridPart.col = C
    gridPart.EditActive = True
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
 
Private Sub gridPart_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        If rsData.State Then
            rsData.MoveFirst
            Do While Not rsData.EOF
                rsData!terlibat = IIf(rsData!terlibat = True, False, True)
                rsData.MoveNext
            Loop
            rsData.MoveFirst
        End If
    End If
End Sub

Private Sub GridPart_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If Not rsGrid.EOF Then
'        Text(5) = rsGrid.Fields(4)      ''model kendaraan   'GridPart.Columns(0)
'        Text(6) = rsGrid.Fields(5)      ''nama model        'GridPart.Columns(1)
'        Text(0) = rsGrid.Fields(6)      ''job order         'GridPart.Columns(2)
'        Text(1) = rsGrid.Fields(7)      ''kode part         'GridPart.Columns(3)
'        Text(2) = rsGrid.Fields(8)      ''nama part         'GridPart.Columns(4)
'        Text(3) = rsGrid.Fields(10)     ''keterangan        'GridPart.Columns(6)
'        Text(7) = rsGrid.Fields(9)      ''qty               'GridPart.Columns(5)
'        ''chPart.Enabled = IIf(rsGrid.Fields(7).value, vbChecked, vbUnchecked)
'
'        Text(14) = rsGrid.Fields(1)     '' NIK TAHUN
'        tglRetailAwal.value = Format(rsGrid.Fields(2), "")
'        tglRetailAkhir.value = Format(rsGrid.Fields(3), "")
'        cbTagihan.Text = rsGrid.Fields(0)
'        Dim SBE As String
'
'        If cbTagihan.Text = "NRM" Then
'            Label15.Visible = True
'            Text(15).Visible = True
'            SBE = objO.FetchDataHO_String("SELECT FS_RC FROM HO_FOP_JOB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "'")
'            If SBE = "R010" Then
'                SBE = "sb10k"
'            ElseIf SBE = "R020" Then
'                SBE = "sb20k"
'            ElseIf SBE = "R030" Then
'                SBE = "sb30k"
'            ElseIf SBE = "R040" Then
'                SBE = "sb40k"
'            ElseIf SBE = "R050" Then
'                SBE = "sb50k"
'            Else
'                SBE = "DEPOSIT"
'            End If
         '  Text(15) = objO.FetchDataHO_Double("select " & SBE & " from MVP..m_deposit where kodekategory = '" & NulSID(rsGrid.Fields(4)) & "'")
'        Else
'            Label15.Visible = False
'            Text(15).Visible = False
'
'        End If
'
'        gridSub.ClearFields
'        Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(7)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(8)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(6)) & "' and MDL = '" & NulSID(rsGrid.Fields(4)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
'        ''If Not rsGridSub.EOF Then
'            Text(12) = ""
'            Text(11) = ""
'            Text(13) = 0
'
'            gridSub.DataSource = rsGridSub
'            Call HeaderColomSubstitusi
'            gridSub.ReBind
'        ''End If
'    End If
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
        text(12) = rsGridSub.Fields(0)  ' kodeitem
        text(11) = rsGridSub.Fields(1)  ' nama part
        text(13) = rsGridSub.Fields(2)  ' qty
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
        
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index) = gMainKeyCari
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
                gMainKeyCari = text(Index)
                gSearchOption = "JOB_ORDER_FOP"
                gBantuan = "SBE"    '/ "ERC" '/
                frmCariHO.show (vbModal)
            Case 8, 12
                ' PART DARI HO
                gMainKeyCari = text(Index)
                gSearchOption = "PARTHO"
                frmCariHO.show (vbModal)
            Case 1
                Cancel = True
                gMainKeyCari = text(Index)
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
            If text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("SELECT * FROM ho_fop_job WITH (NOLOCK) WHERE FS_Item = '" & text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = text(Index)
                    gSearchOption = "JOB_ORDER_FOP"
                    gBantuan = "SBE"
                    frmCariHO.show (vbModal)
                End If
                oRs.Close
            End If
        Case 8
            If text(Index) <> "" Then
            ' PART MANUAL INPUTAN HO
                Set oRs = objO.FetchDataHO("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & text(Index) & "'")
                If Not oRs.BOF Then
                    text(4) = Trim(oRs("NAMA_PART"))
                End If
                oRs.Close
            End If
        Case 12
            If text(Index) <> "" Then
            ' PART MANUAL INPUTAN HO
                Set oRs = objO.FetchDataHO("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & text(Index) & "'")
                If Not oRs.BOF Then
                    text(11) = Trim(oRs("NAMA_PART"))
                End If
                oRs.Close
            End If
        Case 1
            If text(Index) <> "" Then
                'Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL like '" & Text(1) & "%' and FS_LOG_DEL = ''")
                '' ttidak bisa menggunakan like karena bisa jadi data ho_part / ho_sublet tidak ada d model tertentu
                
                ''Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL = '" & Text(1) & "' and FS_LOG_DEL = ''")
                ''20200625 Dibuka kembali
                Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL like '" & text(1) & "%' and FS_LOG_DEL = ''")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = text(Index)
                    gSearchOption = "MDL"
                    frmCariHO.show (vbModal)
                Else
                    text(2) = oRs("Desc")
                End If
                oRs.Close
            End If
    End Select
End Sub


Private Sub ProsesExport(onV As String)
    Dim oApp    As Excel.Application
    Dim oWB     As Excel.Workbook
    Dim nAsc    As Integer
    Set oApp = New Excel.Application
    oApp.Visible = False
    Set oWB = oApp.Workbooks.Add
    With oWB
'        For C = 0 To rsDownload.Fields.Count - 1
'            .Sheets("Sheet1").Cells(1, C + 1).value = rsDownload.Fields(C).Name
'            nAsc = nAsc + 1
'        Next
        .Sheets("Sheet1").Cells(1, 1).value = "NAMAPEKERJAAN (max 50 Kar - sesuai master job)"
        .Sheets("Sheet1").Cells(1, 2).value = "KODEMODEL (max 20 kar - sesuai master model)"
        .Sheets("Sheet1").Cells(1, 3).value = "TAGIHAN ( pilihan NRM / TAM)"
        .Sheets("Sheet1").Cells(1, 4).value = "KODENIK (inisial Tahun NIK - sesuai master)"
        .Sheets("Sheet1").Cells(1, 5).value = "TGLRETAILAWAL (DD/MM/YYYY)"
        .Sheets("Sheet1").Cells(1, 6).value = "TGLRETAILAKHIR (DD/MM/YYYY)"
        .Sheets("Sheet1").Cells(1, 7).value = "KODEITEM (max 15 - sesuai master part)"
        .Sheets("Sheet1").Cells(1, 8).value = "NOMINAL DEPOSITO"

        .Sheets("Sheet1").Cells(2, 1).CopyFromRecordset rsDownload
        .Sheets("Sheet1").Range("1:1").Font.Bold = True
        .Sheets("Sheet1").Select
        .Close SaveChanges:=True, FileName:=onV
    End With
    Set oWB = Nothing
    oApp.Quit
    Set oApp = Nothing
End Sub

Private Sub refreshComboNIK()
    sSQL = "select kode,tahun from listTahunNikFOP with (nolock) "

    Adodc1.ConnectionString = SQLCONBengkel
    Adodc1.RecordSource = sSQL
    cbKodeNik.DataSourceList = Adodc1
    cbKodeNik.DataFieldList = "kode"
    
    cbKodeNik.Columns(0).Width = 2000
    cbKodeNik.Columns(1).Width = 3500
    cbKodeNik.Columns(0).Caption = "Kode"
    cbKodeNik.Columns(1).Caption = "Tahun"
    cbKodeNik.text = ""
End Sub
