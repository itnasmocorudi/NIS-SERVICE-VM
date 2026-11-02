VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmHOPartPacket 
   BackColor       =   &H00FFFFFF&
   Caption         =   "PACKAGE PART"
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
      Left            =   480
      TabIndex        =   6
      Top             =   360
      Width           =   12015
      Begin VB.Frame frameSubstitusi 
         BackColor       =   &H000080FF&
         Caption         =   "Substitusi Part"
         Height          =   4575
         Left            =   360
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   10215
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   3075
            TabIndex        =   33
            Top             =   240
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
            Index           =   12
            Left            =   555
            TabIndex        =   34
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
            TabIndex        =   35
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
            Height          =   3135
            Left            =   120
            TabIndex        =   36
            Top             =   1320
            Width           =   9855
            _ExtentX        =   17383
            _ExtentY        =   5530
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
            Left            =   1560
            TabIndex        =   37
            Top             =   900
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
            Left            =   555
            TabIndex        =   38
            Top             =   900
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
         Begin AIFCmp1.asxPowerButton cmdTutupSub 
            Height          =   360
            Left            =   2560
            TabIndex        =   41
            Top             =   900
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Tutup"
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
            TabIndex        =   40
            Top             =   585
            Width           =   270
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
            TabIndex        =   39
            Top             =   285
            Width           =   300
         End
         Begin VB.Shape Shape5 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   2955
            Top             =   240
            Width           =   60
         End
      End
      Begin VB.CheckBox chRekomendasi 
         BackColor       =   &H000080FF&
         Caption         =   "Rekomendasi"
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
         Left            =   1440
         TabIndex        =   30
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Frame frmeGanti 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   1080
         TabIndex        =   20
         Top             =   2400
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
            Picture         =   "frmHOPartPacket.frx":0000
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
         Top             =   120
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
         Top             =   1065
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
         Height          =   4335
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   11775
         _ExtentX        =   20770
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
         Top             =   750
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
         Top             =   750
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
         Top             =   435
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
         Top             =   435
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
         Top             =   1380
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
      Begin AIFCmp1.asxPowerButton cmdSubtitusi 
         Height          =   360
         Left            =   3960
         TabIndex        =   31
         Top             =   1440
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         BackColor       =   12648447
         BorderStyle     =   3
         Caption         =   "Substitusi Part"
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
         Top             =   1425
         Width           =   270
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3840
         Top             =   450
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
         Left            =   135
         TabIndex        =   16
         Top             =   480
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
         Top             =   795
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3840
         Top             =   750
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
         Top             =   165
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
         Top             =   1110
         Width           =   840
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   7320
         Top             =   120
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
      FormatString    =   "  Tabel Part Package"
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
      Left            =   6480
      TabIndex        =   18
      Top             =   7080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      BackColor       =   12648447
      BorderStyle     =   3
      Caption         =   "Download Excel"
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
Attribute VB_Name = "frmHOPartPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 10 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================
'Tgl Modifikasi : Desember 2014 - Rudi

Dim rsGrid As New ADODB.Recordset
Dim rsDownload  As New ADODB.Recordset
Dim rsGridSub   As New ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim rsModel As ADODB.Recordset
Dim newRecord As Boolean
Dim KDPART_LAMA As String
Dim boleh As Boolean
Dim pesan As String
Dim cols As TrueOleDBGrid70.Columns

Private Sub ber_sih()
    For i = 0 To 7
        Select Case i
            Case 4
            Case 7
                text(i) = 0
            Case Else
                text(i) = ""
        End Select
    Next
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    text(0).Enabled = Flg
    text(1).Enabled = flg1
    text(2).Enabled = flg1
    text(3).Enabled = flg1
    text(5).Enabled = flg1
    text(7).Enabled = flg1
End Sub

Private Sub refreshGrid()
    gridPart.ClearFields
    Set rsGrid = objO.FetchDataHO("SELECT mdl as [KODE MODEL], MODEL, fs_item as [JOB ORDER], item_part as [KODE PART], nama_part as [NAMA PART], QTY, KETERANGAN,[rekomendasi] FROM HO_PARTPACK WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY FS_item, ITEM_part")
    If Not rsGrid.EOF Then
        gridPart.DataSource = rsGrid
        Call HeaderColom
    End If
End Sub

Private Sub HeaderColom()
    gridPart.Columns(0).Caption = "KODE MODEL"
    gridPart.Columns(0).Width = 2000
    gridPart.Columns(1).Caption = "MODEL"
    gridPart.Columns(1).Width = 3500
    gridPart.Columns(2).Caption = "JOB ORDER"
    gridPart.Columns(2).Width = 6000
    gridPart.Columns(3).Caption = "KODE PART"
    gridPart.Columns(3).Width = 2000
    gridPart.Columns(4).Caption = "NAMA PART"
    gridPart.Columns(4).Width = 3000
    gridPart.Columns(5).Caption = "QTY"
    gridPart.Columns(5).Width = 600
    gridPart.Columns(5).NumberFormat = "#,#0"
    gridPart.Columns(6).Caption = "KETERANGAN"
    gridPart.Columns(6).Width = 3000
End Sub

Private Sub cmdCetak_Click()
              
1        On Error GoTo cmdCetak_Click_Error
          Dim Er As New ErrInfo
2         If rsDownload.State Then rsDownload.Close
3         Set rsDownload = Nothing
4         With rsDownload
5             .Fields.Append "NAMAPEKERJAAN", adVarChar, 50
6             .Fields.Append "KODEMODEL", adVarChar, 50
7             .Fields.Append "KODEPART", adVarChar, 50
8             .Fields.Append "NAMAPART", adVarChar, 50
9             .Fields.Append "QTY", adVarChar, 50
10            .Fields.Append "KETERANGAN", adVarChar, 50
11            .Fields.Append "REKOMENDASI", adVarChar, 50
12            .Open
       
13        End With
            
          Dim MyPath As String
          Dim onVal As String
14        MyPath = ""
          Dim f As New FormExplorer
15        f.show (vbModal)
16        MyPath = f.PathSaving
17        If MyPath = "" Then Exit Sub
18        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
19        If f.FileName = "" Then Exit Sub
20        onVal = MyPath & f.FileName & ".xls"
21        Call ProsesExport(onVal)
22                Er.HeadCap = "Export to excel !"
23                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
24                Beep
25                Er.show (vbModal)
          
          

26       On Error GoTo 0
27       Exit Sub

cmdCetak_Click_Error:

28        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdCetak_Click of Form frmHOPartPacket"


End Sub

Private Sub cmdGanti_Click()
    text(8) = ""
    text(4) = ""
    text(10) = ""
    text(9) = ""
    frmeGanti.Visible = True
End Sub

Private Sub cmdHapus_Click()
1         Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_partpack WITH (NOLOCK) WHERE fs_item = '" & text(0) & "' AND item_part = '" & text(1) & "' and FS_LOG_DEL = ''")
2         If Not rsTemp.EOF Then
3             temp = MsgBox("Yakin akan menghapus data part : " & text(1) & " ?" & vbCr & "Untuk Job Order : " & text(0), vbYesNo, "HAPUS DATA")
4             If temp = vbYes Then
5                 Call objO.UpdateDataHO("UPDATE ho_partpack SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE fs_item = '" & NulSID(rsGrid.Fields(2)) & "' AND item_part = '" & text(1) & "' and FS_LOG_DEL = ''")
6                 Call objO.UpdateDataHO("UPDATE HO_FOP_PROGRAM_SUB SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(2)) & "' and mdl = '" & NulSID(rsGrid.Fields(0)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(3)) & "' and FS_LOG_DEL = ''")
7                 Call objO.UpdateDataHO("INSERT INTO TEMPKIRIM VALUES ('PARTPACK', '" & text(0) & "', 'D')")
8             End If
9             Call ber_sih
10            Call refreshGrid
11        Else
12            MsgBox "Pilih data yang ingin dihapus"
13            Exit Sub
14        End If
End Sub

Private Sub cmdHapusSub_Click()
          
1        On Error GoTo cmdHapusSub_Click_Error
2             Set rsTemp = objO.FetchDataHO("SELECT * FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(2)) & "' and mdl = '" & NulSID(rsGrid.Fields(0)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(3)) & "'  and ITEM_PARTSUB = '" & NulSID(rsGridSub.Fields(0)) & "' and NAMA_PARTSUB = '" & NulSID(rsGridSub.Fields(1)) & "' and QTYSUB = '" & NulSID(rsGridSub.Fields(2)) & "' and FS_LOG_DEL = ''")
3             If Not rsTemp.EOF Then
4                 temp = MsgBox("Yakin akan menghapus data sub part : " & NulSID(rsGridSub.Fields(0)) & " ?" & vbCr & "Untuk Job Order : " & text(0), vbYesNo, "HAPUS DATA")
5                 If temp = vbYes Then
6                     Call objO.UpdateDataHO("UPDATE HO_FOP_PROGRAM_SUB SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(2)) & "' and mdl = '" & NulSID(rsGrid.Fields(0)) & "' and ITEM_PART = '" & NulSID(rsGrid.Fields(3)) & "'  and ITEM_PARTSUB = '" & NulSID(rsGridSub.Fields(0)) & "' and NAMA_PARTSUB = '" & NulSID(rsGridSub.Fields(1)) & "' and QTYSUB = '" & NulSID(rsGridSub.Fields(2)) & "' and FS_LOG_DEL = ''")
7                 End If
8                 text(11) = ""
9                 text(12) = ""
10                text(13) = 0
11                gridSub.ClearFields
12                Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(3)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(4)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(2)) & "' and MDL = '" & NulSID(rsGrid.Fields(0)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
13                gridSub.DataSource = rsGridSub
14                Call HeaderColomSubstitusi
15                gridSub.ReBind
16            Else
17                MsgBox "Pilih data yang ingin dihapus"
18                Exit Sub
19            End If
          

20       On Error GoTo 0
21       Exit Sub

cmdHapusSub_Click_Error:

22        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdHapusSub_Click of Form frmHOPartPacket"


End Sub

Private Sub cmdImport_Click()
      Dim ErrBegin As Boolean, konter As Long
      Dim isBegin As Boolean
1     isBegin = False
2       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim rsTempExcel As ADODB.Recordset
         Dim Jml As Integer
         Dim itemPart As String
         
        ' Dim konter As Integer
         
3       konter = 0
4       progBar.Min = 0
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
         Dim rekomendasi As Integer
         Dim rsModel  As ADODB.Recordset
         
18       Call objHT.OpenTransactionHO
19       isBegin = True
20        gBantuan = "SBE"
21       If RcExcel.State Then RcExcel.Close
22       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
23       If RcExcel.RecordCount > 0 Then
24          progBar.Visible = True
25          progBar.Max = RcExcel.RecordCount
26          RcExcel.MoveFirst
27          Do While Not RcExcel.EOF
28             DoEvents
29             konter = konter + 1
30             progBar.value = konter
31             model = ""
32             rekomendasi = 0
33             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
34             If objO.FetchDataHO_Count("select * from JOBORDER WITH(NOLOCK) WHERE FS_LOG_DEL = '' and FS_ITEM = '" & Trim(RcExcel.Fields(0).value) & "'") = 0 Then
35                Call objHT.CloseTransactionHO_err
36                progBar.Visible = False
37                MsgBox "JOB ORDER = " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " salah!", vbCritical, "Error"
38                Exit Sub
39             End If
               
40             If Nul(RcExcel.Fields(1).value) = "" Then
41                Call objHT.CloseTransactionHO_err
42                progBar.Visible = False
43                MsgBox "KODE MODEL baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
44                Exit Sub
45             ElseIf Len(Nul(RcExcel.Fields(1).value)) < 6 Then
46                Call objHT.CloseTransactionHO_err
47                progBar.Visible = False
48                MsgBox "KODE MODEL baris ke-" & konter & " min 6 karakter !", vbCritical, "Error"
49                Exit Sub
50             Else
51                Set rsModel = objO.FetchDataHO("SELECT * FROM MODEL WHERE FS_LOG_DEL = '' AND MDL like '" & Trim(RcExcel.Fields(1).value) & "%'")
52                If rsModel.RecordCount > 0 Then
53                    model = Trim(rsModel("DESC"))
54                Else
                      'jika tidak ada data
55                    Call objHT.CloseTransactionHO_err
56                    progBar.Visible = False
57                    MsgBox "KODE MODEL baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
58                    Exit Sub
59                End If
60             End If
               
61             If Nul(RcExcel.Fields(2).value) = "" Then
62                Call objHT.CloseTransactionHO_err
63                progBar.Visible = False
64                MsgBox "KODE PART baris ke-" & konter & " harus diisi !", vbCritical, "Error"
65                Exit Sub
66             Else
67                itemPart = Nul(RcExcel.Fields(2).value)
68             End If
               
      '         If Nul(RcExcel.Fields(3).value) = "" Then
      '            Call objHT.CloseTransactionHO_err
      '            ProgBar.Visible = False
      '            MsgBox "NAMA PART baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
      '            Exit Sub
      '         End If
               
69             If Nul(RcExcel.Fields(4).value) = "" Then
70                Call objHT.CloseTransactionHO_err
71                progBar.Visible = False
72                MsgBox "QTY PART baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
73                Exit Sub
74             Else
75                If CInt(RcExcel.Fields(4).value) <= 0 Then
76                    Call objHT.CloseTransactionHO_err
77                    progBar.Visible = False
78                    MsgBox "QTY PART baris ke-" & konter & " harus lebih dari 0 !", vbCritical, "Error"
79                    Exit Sub
80                End If
81             End If
               
               'If Nul(RcExcel.Fields(5).value) = "" Then
      'pesanRekomendasi:
      '            Call objHT.CloseTransactionHO_err
      '            ProgBar.Visible = False
      '            MsgBox "Rekomendasi baris ke-" & konter & " harus diisi 0 / 1 !", vbCritical, "Error"
      '            Exit Sub
82             If Nul(RcExcel.Fields(6).value) = "1" Or Nul(RcExcel.Fields(6).value) = "0" Then
83                rekomendasi = Nul(RcExcel.Fields(6).value)
84             Else
85                rekomendasi = 0 ''GoTo pesanRekomendasi
86             End If
87             If rsModel.RecordCount > 0 Then
88                rsModel.MoveFirst
89                Do While Not rsModel.EOF
90                    Set rsTemp = objO.FetchDataHO("SELECT * FROM ho_partpack WITH (NOLOCK) WHERE fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(rsModel!MDL) & "' AND item_part = '" & Trim(itemPart) & "' and FS_LOG_DEL = ''")
91                    If Not rsTemp.EOF Then
92                        Call objHT.UpdateData_Ho("UPDATE ho_partpack SET item_part = '" & NulSID(itemPart) & "', nama_part = '" & Trim(RcExcel.Fields(3).value) & "', keterangan = '" & Trim(RcExcel.Fields(5).value) & "', QTY = '" & CDbl(Trim(RcExcel.Fields(4).value)) & "',rekomendasi = '" & rekomendasi & "' WHERE fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(rsModel!MDL) & "' AND item_part = '" & Trim(itemPart) & "' and FS_LOG_DEL = ''")
                          ''update ke ho_fop_program_sub
                          '' -----------------------------------
93                        Call objHT.UpdateData_Ho("UPDATE HO_FOP_PROGRAM_SUB SET item_part = '" & Trim(itemPart) & "', nama_part = '" & Trim(RcExcel.Fields(3).value) & "' , QTY = '" & CDbl(Trim(RcExcel.Fields(4).value)) & "'  WHERE fs_item = '" & Trim(RcExcel.Fields(0).value) & "' AND mdl = '" & Trim(rsModel!MDL) & "' AND item_part = '" & Trim(itemPart) & "' and FS_LOG_DEL = ''")
94                    Else
95                       sSQL = "INSERT INTO ho_partpack (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL,rekomendasi) VALUES ('" & gBantuan & "', '" & Trim(RcExcel.Fields(0).value) & "', '" & Trim(rsModel!MDL) & "','" & NulSID(rsModel!desc) & "', '" & Trim(itemPart) & "','" & Trim(RcExcel.Fields(3).value) & "','" & CDbl(Trim(RcExcel.Fields(4).value)) & "','" & Trim(RcExcel.Fields(5).value) & "','','','" & rekomendasi & "')"
                         
96                        Call objHT.InsertData_Ho(sSQL)
97                    End If
98                    rsModel.MoveNext
99                Loop
100            End If
101            RcExcel.MoveNext
102         Loop
103      End If
104      RcExcel.Close
105      CnExel.Close
      '   ProgBar.Visible = False
      '  ' IsiGrid
      '   DbCon.CommitTrans
      '   ErrBegin = False
106       Call objHT.CloseTransactionHO
107      MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
108      progBar.Visible = False
109      Call ber_sih
110      Call buka_Tutup(False, False)
111      Call refreshGrid
112      cmdUbah.Caption = "&Ubah"
113      cmdSimpan.Caption = "&Tambah"
      '   GridAFI.Refresh
114      Return
115      Exit Sub
vb_error:
116      If ErrBegin = True Then
           ' DbCon.RollbackTrans
117      End If
118      If isBegin Then
119          Call objHT.CloseTransactionHO_err
120      End If
121      progBar.Visible = False
122      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdProses_Click()
    Call objO.InsertDataHO("INSERT INTO HO_PARTPACK  (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL) SELECT DISTINCT JOB, FS_ITEM, MDL, MODEL, '" & Trim(text(10)) & "', '" & Trim(text(9)) & "', QTY, KETERANGAN, FS_LOG_DEL, FD_TGL_DEL FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & Trim(text(8)) & "'")
    Call objO.UpdateDataHO("UPDATE HO_PARTPACK SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE ITEM_PART = '" & Trim(text(8)) & "'")
    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
    frmeGanti.Visible = False
End Sub

Private Sub cmdSimpan_Click()
1        On Error GoTo cmdSimpan_Click_Error
          Dim isBegin     As Boolean
2         isBegin = False
3         If cmdSimpan.Caption = "&Tambah" Then
4             newRecord = True
5             Call ber_sih
6             Call buka_Tutup
7             text(0).SetFocus
8             cmdUbah.Caption = "&Batal"
9             cmdSimpan.Caption = "&Simpan"
10        ElseIf cmdSimpan.Caption = "&Simpan" Then
11            boleh = True
12            pesan = "Lengkapi data berikut "
              
13            If CDbl(text(7)) <= 0 Then
14                pesan = pesan + "Jumlah item (QTY) "
15                text(7) = 0
16                text(7).SetFocus
17                boleh = False
18            End If
19            If text(0) = "" Then
20                boleh = False
21                pesan = pesan + "Job Order. "
22            End If
23            If text(5) = "" Then
24                boleh = False
25                pesan = pesan + "Model. "
26            End If
27            If text(1) = "" Then
28                boleh = False
29                pesan = pesan + "Part. "
30            End If
              
31            If boleh Then
32                If newRecord Then
33                    Set rsModel = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL like '" & text(5) & "%' and FS_LOG_DEL = ''")
34                    If Not rsModel.EOF Then
35                        If MsgBox("Apakah yakin akan menambahkan data ke " & rsModel.RecordCount & " model ?", vbYesNo) = vbYes Then
36                            GoTo lanjutLooping
37                        Else
38                            Exit Sub
39                        End If
40                    Else
41                        MsgBox "Model tidak ditemukan", vbInformation
42                        Exit Sub
43                    End If
44                End If
lanjutLooping:
45                Call objHT.OpenTransactionHO
46                isBegin = True
47                If newRecord Then

48                    rsModel.MoveFirst
49                    Do While Not rsModel.EOF
50                        Set rsTemp = objHT.FetchDataHO("SELECT * FROM ho_partpack WITH (NOLOCK) WHERE fs_item = '" & text(0) & "' AND mdl = '" & NulSID(rsModel!MDL) & "' AND item_part = '" & text(1) & "' and FS_LOG_DEL = ''")
51                        If Not rsTemp.EOF Then
52                            Call objHT.UpdateData_Ho("UPDATE ho_partpack SET item_part = '" & text(1) & "', nama_part = '" & text(2) & "', keterangan = '" & text(3) & "', QTY = '" & CDbl(text(7)) & "',rekomendasi = '" & IIf(chRekomendasi.value, "1", "0") & "' WHERE fs_item = '" & text(0) & "' AND mdl = '" & NulSID(rsModel!MDL) & "' AND item_part = '" & NulSID(text(1)) & "' and FS_LOG_DEL = ''")
                              ''update ke ho_fop_program_sub
                              '' -----------------------------------
53                            Call objHT.UpdateData_Ho("UPDATE HO_FOP_PROGRAM_SUB SET item_part = '" & text(1) & "', nama_part = '" & text(2) & "' , QTY = '" & CDbl(text(7)) & "'  WHERE fs_item = '" & text(0) & "' AND mdl = '" & NulSID(rsModel!MDL) & "' AND item_part = '" & NulSID(text(1)) & "' and FS_LOG_DEL = ''")
54                        Else
55                            Call objHT.InsertData_Ho("INSERT INTO ho_partpack (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL,rekomendasi) VALUES ('" & gBantuan & "', '" & text(0) & "','" & Left(NulSID(rsModel!MDL), 20) & "','" & Left(NulSID(rsModel!desc), 50) & "', '" & text(1) & "','" & text(2) & "','" & CDbl(text(7)) & "','" & text(3) & "','','','" & IIf(chRekomendasi.value, "1", "0") & "')")
56                        End If
57                        rsModel.MoveNext
58                    Loop
59                Else
60                    Set rsTemp = objHT.FetchDataHO("SELECT * FROM ho_partpack WITH (NOLOCK) WHERE fs_item = '" & text(0) & "' AND mdl = '" & text(5) & "' AND item_part = '" & text(1) & "' and FS_LOG_DEL = ''")
61                    If Not rsTemp.EOF Then
62                        If newRecord Then
63                            Call objHT.CloseTransactionHO_err
64                            MsgBox "Sudah ada data yang sama." & vbCr & "Jika ingin mengubah pilih ubah"
65                            Exit Sub
66                        Else
67                            Call objHT.InsertData_Ho("UPDATE ho_partpack SET item_part = '" & text(1) & "', nama_part = '" & text(2) & "', keterangan = '" & text(3) & "', QTY = '" & CDbl(text(7)) & "',rekomendasi = '" & IIf(chRekomendasi.value, "1", "0") & "' WHERE fs_item = '" & text(0) & "' AND mdl = '" & text(5) & "' AND item_part = '" & KDPART_LAMA & "' and FS_LOG_DEL = ''")
                              ''update ke ho_fop_program_sub
                              '' -----------------------------------
68                            Call objHT.UpdateData_Ho("UPDATE HO_FOP_PROGRAM_SUB SET item_part = '" & text(1) & "', nama_part = '" & text(2) & "' , QTY = '" & CDbl(text(7)) & "'  WHERE fs_item = '" & text(0) & "' AND mdl = '" & NulSID(text(5)) & "' AND item_part = '" & NulSID(KDPART_LAMA) & "' and FS_LOG_DEL = ''")
69                        End If
70                    Else
71                        Call objHT.InsertData_Ho("INSERT INTO ho_partpack (JOB,FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL,rekomendasi) VALUES ('" & gBantuan & "', '" & text(0) & "','" & text(5) & "','" & text(6) & "', '" & text(1) & "','" & text(2) & "','" & CDbl(text(7)) & "','" & text(3) & "','','','" & IIf(chRekomendasi.value, "1", "0") & "')")
72                    End If
73                End If
74                Call objHT.InsertData_Ho("INSERT INTO TEMPKIRIM VALUES ('PARTPACK', '" & text(0) & "', 'S')")
75                Call objHT.CloseTransactionHO
76                Call ber_sih
77                Call buka_Tutup(False, False)
78                Call refreshGrid
79                cmdUbah.Caption = "&Ubah"
80                cmdSimpan.Caption = "&Tambah"
81            Else
82                MsgBox pesan, vbCritical, "VALID DATA"
83                Exit Sub
84            End If
85        End If

86       On Error GoTo 0
87       Exit Sub

cmdSimpan_Click_Error:
88        If isBegin Then
89            Call objHT.CloseTransactionHO_err
90        End If
91        MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmHOPartPacket"
End Sub

Private Sub cmdSimpanSub_Click()
          Dim isBegin     As Boolean
1        On Error GoTo cmdSimpanSub_Click_Error

2         isBegin = False
3         boleh = True
          
4         If CDbl(text(13)) <= 0 Then
5             pesan = pesan + "Jumlah item (QTY) substitusi "
6             text(13) = 0
7             text(13).SetFocus
8             boleh = False
9         End If
          
10        If text(12) = "" Then
11            boleh = False
12            pesan = pesan + "Part substitusi. "
13        End If
          
14        If boleh Then
15            Call objHT.OpenTransactionHO
16            isBegin = True
17            sSQL = "SELECT * FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(2)) & "' AND  MDL = '" & NulSID(rsGrid.Fields(0)) & "' AND ITEM_PART = '" & NulSID(rsGrid.Fields(3)) & "' AND QTY = '" & NulSID(rsGrid.Fields(5)) & "' AND ITEM_PARTSUB = '" & NulSID(text(12)) & "'"
18            Set rsTemp = objHT.FetchDataHO(sSQL)
19            If Not rsTemp.EOF Then
20                sSQL = "UPDATE HO_FOP_PROGRAM_SUB SET ITEM_PARTSUB = '" & NulSID(text(12)) & "',NAMA_PARTSUB = '" & NulSID(text(11)) & "',QTYSUB = '" & NulSID(text(13)) & "' WHERE FS_ITEM = '" & NulSID(rsGrid.Fields(2)) & "' AND  MDL = '" & NulSID(rsGrid.Fields(0)) & "' AND ITEM_PART = '" & NulSID(rsGrid.Fields(3)) & "' AND QTY = '" & NulSID(rsGrid.Fields(5)) & "' AND ITEM_PARTSUB = '" & NulSID(text(12)) & "'"
21            Else
22                sSQL = "INSERT INTO HO_FOP_PROGRAM_SUB (FS_ITEM,MDL,MODEL,ITEM_PART,NAMA_PART,QTY,ITEM_PARTSUB,NAMA_PARTSUB,QTYSUB) "
23                sSQL = sSQL & " VALUES ('" & NulSID(rsGrid.Fields(2)) & "','" & NulSID(rsGrid.Fields(0)) & "','" & NulSID(rsGrid.Fields(1)) & "','" & NulSID(rsGrid.Fields(3)) & "','" & NulSID(rsGrid.Fields(4)) & "','" & NulSID(rsGrid.Fields(5)) & "','" & NulSID(text(12)) & "','" & NulSID(text(11)) & "','" & NulSID(text(13)) & "')"
24            End If
25            Call objHT.InsertData_Ho(sSQL)
26            Call objHT.CloseTransactionHO
              
27            text(11) = ""
28            text(12) = ""
29            text(13) = 0
30            gridSub.ClearFields
31            Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(3)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(4)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(2)) & "' and MDL = '" & NulSID(rsGrid.Fields(0)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
32            gridSub.DataSource = rsGridSub
33            Call HeaderColomSubstitusi
34            gridSub.ReBind
35        Else
36            MsgBox pesan, vbCritical, "VALID DATA"
37            Exit Sub
38        End If

39       On Error GoTo 0
40       Exit Sub
       
cmdSimpanSub_Click_Error:
41        If isBegin Then
42            Call objHT.CloseTransactionHO_err
43        End If
44        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpanSub_Click of Form frmHOPartPacket"


End Sub

Private Sub cmdSubtitusi_Click()
    Call buka_Tutup(False, False)
    Set rsGridSub = objO.FetchDataHO("SELECT ITEM_PARTSUB as [KODE PART], NAMA_PARTSUB as [NAMA PART], QTYSUB FROM HO_FOP_PROGRAM_SUB WITH (NOLOCK) WHERE FS_LOG_DEL = '' and ITEM_PART = '" & NulSID(rsGrid.Fields(3)) & "' and NAMA_PART = '" & NulSID(rsGrid.Fields(4)) & "' and FS_ITEM = '" & NulSID(rsGrid.Fields(2)) & "' and MDL = '" & NulSID(rsGrid.Fields(0)) & "' ORDER BY ITEM_PARTSUB, NAMA_PARTSUB")
      
    gridSub.ClearFields
      
    text(12) = ""
    text(11) = ""
    text(13) = 0
                    
    gridSub.DataSource = rsGridSub
    Call HeaderColomSubstitusi
    gridSub.ReBind
    frameSubstitusi.Visible = True
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


Private Sub cmdTutupSub_Click()
    Call buka_Tutup(False, True)
    frameSubstitusi.Visible = False
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        newRecord = False
        Call buka_Tutup(False, True)
        KDPART_LAMA = text(1)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
        cmdSubtitusi.Enabled = True
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
        cmdSubtitusi.Enabled = False
    End If
End Sub

Private Sub Form_Load()
'    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
    cmdSubtitusi.Enabled = False
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    gridPart.Width = Frame1.Width - 200
    gridPart.Height = Frame1.Height - text(7).Top - text(7).Height - 200
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdGanti.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    progBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    cmdImport.Left = cmdCetak.Left - cmdImport.Width - 100
    progBar.Left = cmdImport.Left - progBar.Width - 100
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

Private Sub GridPart_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not rsGrid.EOF Then
        text(5) = rsGrid.Fields(0)      'GridPart.Columns(0)
        text(6) = rsGrid.Fields(1)      'GridPart.Columns(1)
        text(0) = rsGrid.Fields(2)      'GridPart.Columns(2)
        text(1) = rsGrid.Fields(3)      'GridPart.Columns(3)
        text(2) = rsGrid.Fields(4)      'GridPart.Columns(4)
        text(3) = rsGrid.Fields(6)      'GridPart.Columns(6)
        text(7) = rsGrid.Fields(5)      'GridPart.Columns(5)
        chRekomendasi.value = IIf(rsGrid.Fields(7) = "1", vbChecked, vbUnchecked)
    End If
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
                gSearchOption = "JOB_ORDER"
                gBantuan = "SBE"    '/ "ERC" '/
                frmCariHO.show (vbModal)
            Case 1, 8, 12
                ' PART DARI HO
                gMainKeyCari = text(Index)
                gSearchOption = "PARTHO"
                frmCariHO.show (vbModal)
            Case 5
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
                Set oRs = objO.FetchDataHO("SELECT * FROM jobOrder WITH (NOLOCK) WHERE FS_Item = '" & text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = text(Index)
                    gSearchOption = "JOB_ORDER"
                    gBantuan = "SBE"
                    frmCariHO.show (vbModal)
                End If
                oRs.Close
            End If
        Case 1, 8, 12
            If text(Index) <> "" Then
            ' PART MANUAL INPUTAN HO
                Set oRs = objO.FetchDataHO("SELECT NAMA_PART FROM HO_PARTPACK WITH (NOLOCK) WHERE ITEM_PART = '" & text(Index) & "'")
                If Not oRs.BOF Then
                    text(4) = Trim(oRs("NAMA_PART"))
                End If
                oRs.Close
            End If
        Case 5
            If text(Index) <> "" Then
                ''Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL ='" & Text(5) & "' and FS_LOG_DEL = ''")
                Set oRs = objO.FetchDataHO("Select * from MODEL WITH (NOLOCK) where MDL like '" & text(5) & "%' and FS_LOG_DEL = ''")
                If oRs.BOF Then
                    Cancel = True
                    gMainKeyCari = text(Index)
                    gSearchOption = "MDL"
                    frmCariHO.show (vbModal)
                Else
                    text(6) = oRs("Desc")
                End If
                oRs.Close
            End If
    End Select
End Sub

Private Sub ProsesExport(onV As String)
          Dim oApp    As Excel.Application
          Dim oWB     As Excel.Workbook
          Dim nAsc    As Integer
1         Set oApp = New Excel.Application
2         oApp.Visible = False
3         Set oWB = oApp.Workbooks.Add
4         With oWB
      '        For C = 0 To rsDownload.Fields.Count - 1
      '            .Sheets("Sheet1").Cells(1, C + 1).value = rsDownload.Fields(C).Name
      '            nAsc = nAsc + 1
      '        Next
5             .Sheets("Sheet1").Cells(1, 1).value = "NAMAPEKERJAAN (max 50 kar- sesuai master job)"
6             .Sheets("Sheet1").Cells(1, 2).value = "KODEMODEL (max 20 Kar -sesuai master model)"
7             .Sheets("Sheet1").Cells(1, 3).value = "KODEPART (max 15 kar - sesuai master part)"
8             .Sheets("Sheet1").Cells(1, 4).value = "NAMAPART (max 40 Karakter)"
9             .Sheets("Sheet1").Cells(1, 5).value = "QTY (dalam angka)"
10            .Sheets("Sheet1").Cells(1, 6).value = "KETERANGAN (max 30 Kar)"
11            .Sheets("Sheet1").Cells(1, 7).value = "REKOMENDASI (isian 0/1)"
                          

12            .Sheets("Sheet1").Cells(2, 1).CopyFromRecordset rsDownload
13            .Sheets("Sheet1").Range("1:1").Font.Bold = True
14            .Sheets("Sheet1").Select
15            .Close SaveChanges:=True, FileName:=onV
16        End With
17        Set oWB = Nothing
18        oApp.Quit
19        Set oApp = Nothing
End Sub
