VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmHOInsentifSetResultKPICabang 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SET RESULT KPI CABANG"
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
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   1920
      TabIndex        =   2
      Top             =   2040
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
   Begin AIFCmp1.asxPowerButton cmdDownloadExcel 
      Height          =   360
      Left            =   8880
      TabIndex        =   10
      Top             =   1920
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
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   375
      Left            =   7440
      TabIndex        =   20
      Top             =   1920
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12840
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   6495
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   12015
      Begin TDBDate6Ctl.TDBDate Periode 
         Height          =   375
         Left            =   1200
         TabIndex        =   26
         Top             =   240
         Width           =   735
         _Version        =   65536
         _ExtentX        =   1296
         _ExtentY        =   661
         Calendar        =   "frmHOInsentifSetResultKPICabang.frx":0000
         Caption         =   "frmHOInsentifSetResultKPICabang.frx":012C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOInsentifSetResultKPICabang.frx":0198
         Keys            =   "frmHOInsentifSetResultKPICabang.frx":01B6
         Spin            =   "frmHOInsentifSetResultKPICabang.frx":0214
         AlignHorizontal =   0
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         CursorPosition  =   0
         DataProperty    =   0
         DisplayFormat   =   "mm-yyyy"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         FirstMonth      =   4
         ForeColor       =   -2147483640
         Format          =   "mm/yyyy"
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
         Text            =   "11/2024"
         ValidateMode    =   0
         ValueVT         =   7
         Value           =   45611
         CenturyMode     =   0
      End
      Begin VB.ComboBox cboAdjKabeng 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmHOInsentifSetResultKPICabang.frx":023C
         Left            =   8520
         List            =   "frmHOInsentifSetResultKPICabang.frx":0246
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame frmeGanti 
         BackColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   96000
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   10335
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   4
            Left            =   915
            TabIndex        =   13
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
            TabIndex        =   14
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   17
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
            Picture         =   "frmHOInsentifSetResultKPICabang.frx":025B
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   450
            Width           =   720
         End
      End
      Begin TrueOleDBGrid70.TDBGrid GridPart 
         Height          =   4335
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   7646
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Kode KPI"
         Columns(0).DataField=   "kdkpi"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nama KPI"
         Columns(1).DataField=   "NamaKPI"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Kode Cabang"
         Columns(2).DataField=   "kdCabang"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nama Cabang"
         Columns(3).DataField=   "nmCabang"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Target"
         Columns(4).DataField=   "TargetKPICbg"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Actual"
         Columns(5).DataField=   "ActKpiCbg"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Achievement"
         Columns(6).DataField=   "AchvKpiCbg"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "kategori"
         Columns(7).DataField=   "kategori"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "setach"
         Columns(8).DataField=   "SetAchv"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=9"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
         Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
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
         Splits(0)._ColumnProps(25)=   "Column(6).Width=3281"
         Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
         Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(29)=   "Column(7).Width=3281"
         Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
         Splits(0)._ColumnProps(32)=   "Column(7).Visible=0"
         Splits(0)._ColumnProps(33)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(34)=   "Column(8).Width=3281"
         Splits(0)._ColumnProps(35)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(36)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(37)=   "Column(8).Visible=0"
         Splits(0)._ColumnProps(38)=   "Column(8).Order=9"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=66,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=28,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=25,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=26,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=27,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=32,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
         _StyleDefs(72)  =   "Named:id=33:Normal"
         _StyleDefs(73)  =   ":id=33,.parent=0"
         _StyleDefs(74)  =   "Named:id=34:Heading"
         _StyleDefs(75)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(76)  =   ":id=34,.wraptext=-1"
         _StyleDefs(77)  =   "Named:id=35:Footing"
         _StyleDefs(78)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(79)  =   "Named:id=36:Selected"
         _StyleDefs(80)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=37:Caption"
         _StyleDefs(82)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(83)  =   "Named:id=38:HighlightRow"
         _StyleDefs(84)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(85)  =   "Named:id=39:EvenRow"
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
         _StyleDefs(93)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(94)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(95)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(96)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   7560
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   9360
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   4185
         _ExtentX        =   7382
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
         Index           =   1
         Left            =   1440
         TabIndex        =   23
         Top             =   840
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin AIFCmp1.asxPowerButton cmdCekData 
         Height          =   375
         Left            =   2040
         TabIndex        =   28
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BackColor       =   12648447
         Caption         =   "Cek Data"
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   1440
         TabIndex        =   29
         Top             =   1200
         Width           =   825
         _ExtentX        =   1455
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Porsi Insentif"
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
         Left            =   360
         TabIndex        =   30
         Top             =   1200
         Width           =   945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   120
         X2              =   11880
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
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
         Left            =   495
         TabIndex        =   27
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adj Kabeng"
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
         Left            =   7575
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hari Kerja"
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
         TabIndex        =   9
         Top             =   840
         Width           =   705
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   9240
         Top             =   240
         Visible         =   0   'False
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cabang"
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
         Left            =   6855
         TabIndex        =   8
         Top             =   285
         Visible         =   0   'False
         Width           =   555
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
      FormatString    =   "  Tabel Insentif Set Result KPI Cabang"
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
   Begin AIFCmp1.asxPowerButton cmdUbah 
      Height          =   360
      Left            =   8880
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
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
      TabIndex        =   5
      Top             =   7080
      Visible         =   0   'False
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
      TabIndex        =   6
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
   Begin AIFCmp1.asxPowerButton cmdGanti 
      Height          =   360
      Left            =   480
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
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
   Begin ARProgBarCtrl.ARProgressBar ProgBar 
      Height          =   255
      Left            =   1200
      TabIndex        =   21
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
   Begin TrueOleDBGrid70.TDBGrid gridDownload 
      Height          =   1215
      Left            =   4680
      TabIndex        =   22
      Top             =   7440
      Visible         =   0   'False
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   2143
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "NAMA PEKERJAAN"
      Columns(0).DataField=   "NAMAPEKERJAAN"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "KODE MODEL"
      Columns(1).DataField=   "KODE MODEL"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "KODE SUBLET"
      Columns(2).DataField=   "KODESUBLET"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "QTY"
      Columns(3).DataField=   "QTY"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "KETERANGAN"
      Columns(4).DataField=   "KETERANGAN"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(9)=   "Column(2).Width=3281"
      Splits(0)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
      Splits(0)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(13)=   "Column(3).Width=3281"
      Splits(0)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
      Splits(0)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(17)=   "Column(4).Width=3281"
      Splits(0)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
      Splits(0)._ColumnProps(20)=   "Column(4).Order=5"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
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
   End
End
Attribute VB_Name = "frmHOInsentifSetResultKPICabang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOInsentifSetResultKPICabang
' DateTime  : 13/05/2020 10:01
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

 
Dim rsGrid      As New ADODB.Recordset
Dim rsDownload As New ADODB.Recordset
Dim rsModel     As ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim newRecord   As Boolean
Dim KDPART_LAMA As String
Dim boleh       As Boolean
Dim pesan       As String
Dim cols        As TrueOleDBGrid70.Columns

Private Sub ber_sih(semua As Boolean)
    For i = 0 To 6
        If semua = False Then
            Text(1) = ">>"
            Text(1).Mask = "99"
            Text(2) = ">>,>>"
            Text(2).Mask = "99,99"
        Else
            Select Case i
                Case 5
                Case 1
                    Text(i) = ">>"
                    Text(i).Mask = "99"
                Case 2
                    Text(i) = ">>,>>"
                    Text(i).Mask = "99,99"
                Case Else
                    Text(i) = ""
            End Select
        End If
    Next
    Periode = Format(Now(), "mm") & "-" & Format(Now(), "yyyy")
    refreshGrid
    cboAdjKabeng.ListIndex = 0
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    Text(0).Enabled = True
    Periode.Enabled = True
    Text(2).Enabled = True
    Text(1).Enabled = True
    cboAdjKabeng.Enabled = Flg
    
   ' Text(3).Enabled = False
    Text(6).Enabled = False
End Sub

Private Sub refreshGrid()
    'GridPart.ClearFields
    createRsGrid
    
    Set rsTemp = objO.FetchDataHO("select KdKPI,NamaKPI,TargetKPICbg,ActKpiCbg,Kategori,SetAchv,AchvKpiCbg,b.fs_kd_dealer,b.fs_nm_dealer from insentif..grKPIResultCbg a with (nolock) inner join bengkel99..t_dealer b on a.fs_kd_dealer = b.fs_kd_dealer where Periode = '" & Format(Periode, "mm-yyyy") & "'")
    If Not rsTemp.EOF Then
        rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            With rsGrid
                .AddNew
                !KdKPI = rsTemp!KdKPI
                !namaKPI = rsTemp!namaKPI
                !kdCabang = rsTemp!FS_KD_DEALER
                !nmCabang = rsTemp!fs_nm_dealer
                !TargetKPICbg = rsTemp!TargetKPICbg
                !ActKpiCbg = rsTemp!ActKpiCbg
                !AchvKpiCbg = k2t(rsTemp!AchvKpiCbg)
                !kategori = rsTemp!kategori
                !setAchv = rsTemp!setAchv
                .Update
            End With
            rsTemp.MoveNext
        Loop
    End If
    
    Set rsTemp = objO.FetchDataHO("SELECT * FROM INSENTIF..GRAdtParameter WHERE periode = '" & Format(Periode.value, "MM-YYYY") & "'")
    If Not rsTemp.EOF Then
        Text(1) = rsTemp!jmlharikerja
        Text(2) = Format(rsTemp!porsiInsentif, "00.00")
    End If
     
End Sub

Private Sub createRsGrid()

   On Error GoTo createRsGrid_Error

    If rsGrid.State Then rsGrid.Close
    Set rsGrid = Nothing
    With rsGrid
        .Fields.Append "KdKPI", adVarChar, 30
        .Fields.Append "NamaKPI", adVarChar, 100
        .Fields.Append "kdCabang", adVarChar, 100
        .Fields.Append "nmCabang", adVarChar, 100
        .Fields.Append "TargetKPICbg", adVarChar, 25
        .Fields.Append "ActKpiCbg", adVarChar, 25
        .Fields.Append "AchvKpiCbg", adSingle
        .Fields.Append "Kategori", adInteger
        .Fields.Append "SetAchv", adInteger
        .Open
    End With
    Set GridPart.DataSource = rsGrid

   On Error GoTo 0
   Exit Sub

createRsGrid_Error:

    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createRsGrid of Form frmHOInsentifSetResultKPICabang"
End Sub


Private Sub HeaderColom()
    GridPart.Columns(0).Caption = "Kode Tipe"
    GridPart.Columns(0).Width = 2000
    GridPart.Columns(1).Caption = "Tipe"
    GridPart.Columns(1).Width = 3500
    GridPart.Columns(2).Caption = "Bobot Posisi"
    GridPart.Columns(2).Width = 2000
    GridPart.Columns(3).Caption = "Bobot Result"
    GridPart.Columns(3).Width = 2000
    GridPart.Columns(4).Caption = "Bobot Process"
    GridPart.Columns(4).Width = 2000
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
        
5       .Sheets("Sheet1").Cells(1, 1).value = "KODE KPI"
6       .Sheets("Sheet1").Cells(1, 2).value = "NAMA KPI"
7       .Sheets("Sheet1").Cells(1, 3).value = "KODE DEALER"
8       .Sheets("Sheet1").Cells(1, 4).value = "NAMA DEALER"
9       .Sheets("Sheet1").Cells(1, 5).value = "TARGET (gunakan koma untuk desimal)"
10      .Sheets("Sheet1").Cells(1, 6).value = "ACHIEVEMENT(gunakan koma untuk desimal)"
11      .Sheets("Sheet1").Cells(1, 7).value = "PERSENTASE (gunakan koma untuk desimal)"
        '.Sheets("Sheet1").Cells(1, 5).value = "KETERANGAN (max 30 kar)"
          

        '.Sheets("Sheet1").Cells(2, 1).CopyFromRecordset rsDownload
12      .Sheets("Sheet1").Range("1:1").Font.Bold = True
13      .Sheets("Sheet1").Select
14      .Close SaveChanges:=True, FileName:=onV
15        End With
16        Set oWB = Nothing
17        oApp.Quit
18        Set oApp = Nothing
End Sub
 

Private Sub cmdCekData_Click()
    Call refreshGrid
End Sub

Private Sub cmdDownloadExcel_Click()
    
   On Error GoTo cmdDownloadExcel_Click_Error
    Dim Er As New ErrInfo
      
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

cmdDownloadExcel_Click_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdDownloadExcel_Click of Form frmHOInsentifSetResultKPICabang"


End Sub

Private Sub cmdGanti_Click()
    Text(8) = ""
    Text(4) = ""
    Text(10) = ""
    Text(9) = ""
    frmeGanti.Visible = True
End Sub

Private Sub cmdHapus_Click()
          Dim isBegin As Boolean
1        On Error GoTo cmdHapus_Click_Error
2         isBegin = False
3         Set rsTemp = objO.FetchDataHO("SELECT * FROM INSENTIF..GRMstKPIPosisi WITH (NOLOCK) WHERE KdTipe = '" & Text(0) & "' and KdKPI = '" & Text(2) & "'")
4         If Not rsTemp.EOF Then
5             temp = MsgBox("Yakin akan menghapus data ini ?", vbYesNo, "HAPUS DATA")
6             If temp = vbYes Then
7                 Call objHT.OpenTransactionHO
8                 isBegin = True
9                 Call objHT.InsertData_Ho("UPDATE INSENTIF..GRMstKPIPosisi SET TglDel = getdate(), UserDel = '" & gSA & "' WHERE KdTipe = '" & Text(0) & "' and KdKPI = '" & Text(2) & "'")
                  
10                Call objHT.CloseTransactionHO
11            End If
12            Call ber_sih(True)
13            Call refreshGrid
14        Else
15            MsgBox "Pilih data yang ingin dihapus"
16            Exit Sub
17        End If

18       On Error GoTo 0
19       Exit Sub

cmdHapus_Click_Error:
20        If isBegin Then
21            Call objHT.CloseTransactionHO_err
22        End If
23        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdHapus_Click of Form frmHOInsentifSetResultKPICabang"


End Sub

Private Sub cmdImport_Click()
      Dim ErrBegin As Boolean, konter As Long
1       On Error GoTo vb_error
         Dim CnExel As New ADODB.Connection
         Dim RcExcel As New ADODB.Recordset
         Dim rsTempExcel As ADODB.Recordset
         Dim Jml As Integer
         Dim itemPart As String
         Dim isBegin As Boolean
2        isBegin = False
         Dim AdjKabeng As Integer
         
        ' Dim konter As Integer
         
3       konter = 0
4       ProgBar.Min = 0
5       CommonDialog1.FileName = ""
6       CommonDialog1.InitDir = App.Path
7       CommonDialog1.filter = "Exel File (*.xls)|*.xls"
8       CommonDialog1.ShowOpen
9       If CommonDialog1.FileName <> "" Then
10         If CnExel.State Then CnExel.Close
            'CnExel.Open "Provider=Microsoft.ACE.OLEDB.12.0;" &
                 ' "Data Source=" & CommonDialog1.FileName & ";" & _
                  '"Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
11         CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
12         GoSub Proses
13     Else
14        Exit Sub
15     End If
16     Exit Sub
       
Proses:
17    ErrBegin = True
         Dim VarNmFile As String
         Dim namaKPI As String
         Dim kategori   As Integer
         Dim setAchv   As Integer
         Dim namaDealer, target, pencapaian, nilaiTarget, nilaiPencapaian, ipSentralisasi, kdDatabase, queryAct As String
         Dim persen As String
         
         Dim rsModel  As ADODB.Recordset
18       Call objHT.OpenTransactionHO
19       isBegin = True
20        gBantuan = "SBE"
21        If cboAdjKabeng.Text = "Single" Then
22            AdjKabeng = 0
23        Else
24            AdjKabeng = 1
25        End If
26       If RcExcel.State Then RcExcel.Close
27       RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
28       If RcExcel.RecordCount > 0 Then
29          ProgBar.Visible = True
30          ProgBar.Max = RcExcel.RecordCount
31          RcExcel.MoveFirst
            
32          sSQL = "DELETE INSENTIF..grKPIResultCbg where periode = '" & Format(Periode.value, "MM-YYYY") & "'"
               
33          Call objHT.InsertData_Ho(sSQL)
34          Do While Not RcExcel.EOF
            
35             namaKPI = ""
36             namaDealer = ""
37             kategori = 0
38             setAchv = 0
39             target = ""
40             pencapaian = ""
41             ipSentralisasi = ""
              
42             If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
43             If NulSID(RcExcel.Fields(0).value) = NulSID("Contoh") Then
44              GoTo lanjutloop
45             End If
               
               ''joborder
46             If objO.FetchDataHO_Count("SELECT * FROM insentif..GRMstKPI where KdKPI = '" & Trim(RcExcel.Fields(0).value) & "' and isnull(userdel,'') = ''") = 0 Then
47                Call objHT.CloseTransactionHO_err
48                ProgBar.Visible = False
49                MsgBox "Kode KPI = " & Trim(RcExcel.Fields(0).value) & " baris ke-" & konter & " tidak ditemukan!", vbCritical, "Error"
50                Exit Sub
51             Else
52                Set rsTemp = objHT.FetchDataHO("SELECT * FROM insentif..GRMstKPI where KdKPI = '" & Trim(RcExcel.Fields(0).value) & "' and isnull(userdel,'') = ''")
53                If Not rsTemp.EOF Then
54                    namaKPI = NulSID(rsTemp!namaKPI)
55                    kategori = rsTemp!kategori
56                    setAchv = rsTemp!setAchv
57                End If
58             End If
               
59             If objO.FetchDataHO_Count("select * from T_DEALER where INTERNAL = 1 and FS_KD_DEALER = '" & Trim(RcExcel.Fields(2).value) & "' ") = 0 Then
60                Call objHT.CloseTransactionHO_err
61                ProgBar.Visible = False
62                MsgBox "Kode Dealer = " & Trim(RcExcel.Fields(2).value) & " baris ke-" & konter & " tidak ditemukan!", vbCritical, "Error"
63                Exit Sub
64             Else
65                Set rsTemp = objHT.FetchDataHO("select * from T_DEALER where INTERNAL = 1 and FS_KD_DEALER = '" & Trim(RcExcel.Fields(2).value) & "'")
66                If Not rsTemp.EOF Then
67                    namaDealer = NulSID(rsTemp!fs_nm_dealer)
68                End If
69             End If
               
70             sSQL = "SELECT sentralisasi FROM profilspk..dealerbd WHERE kdnr = '" & Trim(RcExcel.Fields(2).value) & "'"
71             ipSentralisasi = objHT.FetchDataHO_String(sSQL)
               
72             sSQL = "SELECT kdcbg FROM profilspk..dealerbd WHERE kdnr = '" & Trim(RcExcel.Fields(2).value) & "'"
73             kdDatabase = objHT.FetchDataHO_String(sSQL)
               
74             If NulSID(RcExcel.Fields(5).value) = "X" Then
75                sSQL = "SELECT queryAct FROM insentif..grmstkpi WHERE kdkpi = '" & Trim(RcExcel.Fields(0).value) & "'"
76                queryAct = objHT.FetchDataHO_String(sSQL)
77                sSQL = "SELECT query FROM insentif..grqueryinsentif WHERE kode = '" & Trim(queryAct) & "'"
78                queryAct = objHT.FetchDataHO_String(sSQL)
                  
                  ''queryAct = replaceQuery(queryAct, ipSentralisasi, kdDatabase, Format(Periode.value, "YYYY"), Format(Periode.value, "MM"), Format(Periode.value, "YYYY-MM"), Trim(RcExcel.Fields(2).value), Trim(RcExcel.Fields(0).value), "")
                  
79                queryAct = NulSID(queryAct)
80                queryAct = Replace(queryAct, NulSID("[ipsentralisasi]"), "[" & ipSentralisasi & "]")
81                queryAct = Replace(queryAct, NulSID("bengkelxx"), "BENGKEL" & kdDatabase)
82                queryAct = Replace(queryAct, NulSID("~bulan~"), "'" & Format(Periode.value, "M") & "'")
83                queryAct = Replace(queryAct, NulSID("~tahun~"), "'" & Format(Periode.value, "YYYY") & "'")
84                queryAct = Replace(queryAct, NulSID("accxx"), "ACC" & kdDatabase)
85                queryAct = Replace(queryAct, NulSID("periodeacc"), Format(Periode.value, "YYYYMM"))
86                queryAct = Replace(queryAct, NulSID("~periode~"), "'" & Format(Periode.value, "YYYYMM") & "'")
87                queryAct = Replace(queryAct, NulSID("~dealer~"), "'" & Trim(RcExcel.Fields(2).value) & "'")
88                queryAct = Replace(queryAct, NulSID("~kdkpi~"), "'" & Trim(RcExcel.Fields(0).value) & "'")
89                queryAct = Replace(queryAct, NulSID("~kdDatabase~"), "'" & kdDatabase & "'")
90                queryAct = Replace(queryAct, NulSID("DEPO00-YYYYMM"), NulSID("DEPO00-") & Format(Periode.value, "YYYYMM"))
91                queryAct = Replace(queryAct, NulSID("~"), "'")
92                pencapaian = objHT.FetchDataHO_String(queryAct)
                  
93                If Val(pencapaian) > 0 Then
94                    pencapaian = Round(pencapaian, 0)
95                End If
                  
96                If NulSID(RcExcel.Fields(6).value) <> "" Then
97                    If Right(NulSID(RcExcel.Fields(4).value), 1) = "%" Then
98                        target = Mid(RcExcel.Fields(4).value, 1, Len(RcExcel.Fields(4).value) - 1)
99                    Else
100                       target = k2t(RcExcel.Fields(4).value * 100)
101                   End If
102               Else
103                  target = Trim(RcExcel.Fields(4).value)
104               End If
                  
105               persen = Round((CDbl(pencapaian) * 100) / CDbl(target), 0)
                  'pencapaian = Format(pencapaian, "#,###.00")
                  
106            Else
107               If NulSID(RcExcel.Fields(6).value) <> "" Then
108                   If Right(NulSID(RcExcel.Fields(4).value), 1) = "%" Then
109                       target = Mid(RcExcel.Fields(4).value, 1, Len(RcExcel.Fields(4).value) - 1)
110                   Else
111                       target = k2t(RcExcel.Fields(4).value * 100)
112                   End If
                     
113                   If Right(NulSID(RcExcel.Fields(5).value), 1) = "%" Then
114                       pencapaian = Mid(RcExcel.Fields(5).value, 1, Len(RcExcel.Fields(5).value) - 1)
115                   Else
116                       pencapaian = k2t(RcExcel.Fields(5).value * 100)
117                   End If
                     
118                   persen = Round((CDbl(pencapaian) * 100) / CDbl(target), 0)
119               Else
120                   If NulSID(RcExcel.Fields(0).value) = "KPI-015" Then
121                       target = Trim(RcExcel.Fields(4).value)
122                       pencapaian = Trim(RcExcel.Fields(5).value)
123                       If NulSID(RcExcel.Fields(4).value) = NulSID("Mature") Then
124                           nilaiTarget = 120
125                       ElseIf NulSID(RcExcel.Fields(4).value) = NulSID("Stable") Then
126                           nilaiTarget = 110
127                       ElseIf NulSID(RcExcel.Fields(4).value) = NulSID("Practitioner") Then
128                           nilaiTarget = 100
129                       ElseIf NulSID(RcExcel.Fields(4).value) = NulSID("Beginer") Then
130                           nilaiTarget = 90
131                       End If
                          
132                       If NulSID(RcExcel.Fields(5).value) = NulSID("Mature") Then
133                           nilaiPencapaian = 120
134                       ElseIf NulSID(RcExcel.Fields(5).value) = NulSID("Stable") Then
135                           nilaiPencapaian = 110
136                       ElseIf NulSID(RcExcel.Fields(5).value) = NulSID("Practitioner") Then
137                           nilaiPencapaian = 100
138                       ElseIf NulSID(RcExcel.Fields(5).value) = NulSID("Beginer") Then
139                           nilaiPencapaian = 90
140                       End If
                          
141                       persen = Round((CDbl(nilaiPencapaian) * 100) / CDbl(nilaiTarget), 0)
                      
142                   ElseIf NulSID(RcExcel.Fields(0).value) = "KPI-016" Then
143                       target = Trim(RcExcel.Fields(4).value)
144                       pencapaian = Trim(RcExcel.Fields(5).value)
145                       If NulSID(RcExcel.Fields(4).value) = NulSID("Certified") Then
146                           nilaiTarget = 100
147                       Else
148                           nilaiTarget = 1
149                       End If
                          
150                       If NulSID(RcExcel.Fields(5).value) = NulSID("Certified") Then
151                           nilaiPencapaian = 100
152                       Else
153                           nilaiPencapaian = 0
154                       End If
                          
155                       persen = Round((CDbl(nilaiPencapaian) * 100) / CDbl(nilaiTarget), 0)
                          
156                   Else
157                       target = Trim(RcExcel.Fields(4).value)
158                       pencapaian = Trim(RcExcel.Fields(5).value)
159                       persen = Round((CDbl(pencapaian) * 100) / CDbl(target), 0)
160                   End If
161               End If
162            End If
163            If persen > 150 Then
164               persen = 150
165            End If
166               sSQL = "INSERT INTO INSENTIF..grKPIResultCbg (periode,FS_KD_DEALER,KdKPI,NamaKPI,TargetKPICbg,ActKpiCbg,Kategori,SetAchv,AchvKpiCbg,UserEntry,tglentry) VALUES "
167               sSQL = sSQL & " ('" & Format(Periode.value, "MM-YYYY") & "','" & Trim(RcExcel.Fields(2).value) & "','" & Trim(RcExcel.Fields(0).value) & "','" & namaKPI & "', '" & target & "','" & pencapaian & "','" & kategori & "','" & setAchv & "', " & persen & ",'" & gSA & "',getdate())"
               
168           Call objHT.InsertData_Ho(sSQL)
                
169             konter = konter + 1
lanjutloop:
170             RcExcel.MoveNext
171         Loop
172      End If
173       Call objHT.CloseTransactionHO
174      MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
175      ProgBar.Visible = False
176      Call buka_Tutup(False, False)
177      cmdUbah.Caption = "&Ubah"
178      cmdSimpan.Caption = "&Tambah"
179   refreshGrid
180      Return
181      Exit Sub
vb_error:
182      If ErrBegin = True Then
           ' DbCon.RollbackTrans
183      End If
184      If isBegin Then
185       Call objHT.CloseTransactionHO_err
186      End If
187      ProgBar.Visible = False
188      MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")

End Sub

Private Function replaceQuery(query As String, ipSentralisasiCabang As String, kdDatabase As String, Tahun As String, Bulan As String, Periode As String, FS_KD_DEALER As String, KdKPI As String, Posisi As String) As String
    Dim hasil As String
    query = NulSID(query)
    query = Replace(query, NulSID("[ipsentralisasi]"), "[" & ipSentralisasiCabang & "]")
    query = Replace(query, NulSID("bengkelxx"), "BENGKEL" & kdDatabase)
    query = Replace(query, NulSID("~bulan~"), "'" & Bulan & "'")
    query = Replace(query, NulSID("~tahun~"), "'" & Tahun & "'")
    query = Replace(query, NulSID("accxx"), "ACC" & kdDatabase)
    query = Replace(query, NulSID("periodeacc"), Periode)
    query = Replace(query, NulSID("~periode~"), "'" & Periode & "'")
    query = Replace(query, NulSID("~dealer~"), "'" & FS_KD_DEALER & "'")
    query = Replace(query, NulSID("~kdkpi~"), "'" & KdKPI & "'")
    query = Replace(query, NulSID("~kdDatabase~"), "'" & kdDatabase & "'")
    query = Replace(query, NulSID("DEPO00-YYYYMM"), NulSID("DEPO00-") & Periode)
    
    
    query = Replace(query, NulSID("~"), "'")
    replaceQuery = query
End Function

Private Sub CmdKeluar_Click()
1         Unload Me
End Sub

Private Sub cmdProses_Click()
    Call objO.InsertDataHO("INSERT INTO HO_PARTPACK  (JOB,FS_ITEM,MDL,MODEL,KDSUBLET,NAMA_PART,QTY,KETERANGAN,FS_LOG_DEL,FD_TGL_DEL) SELECT DISTINCT JOB, FS_ITEM, MDL, MODEL, '" & Trim(Text(10)) & "', '" & Trim(Text(9)) & "', QTY, KETERANGAN, FS_LOG_DEL, FD_TGL_DEL FROM HO_PARTPACK WITH (NOLOCK) WHERE KDSUBLET = '" & Trim(Text(8)) & "'")
    Call objO.UpdateDataHO("UPDATE HO_PARTPACK SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE KDSUBLET = '" & Trim(Text(8)) & "'")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(Text(10)) & "', 'S')")
    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('GANTIPART','" & Trim(Text(8)) & "', 'D')")
    frmeGanti.Visible = False
End Sub

Private Sub cmdSimpan_Click()
          Dim isBegin As Boolean
          Dim kode As String
          Dim perhitungan As Integer
          
1        On Error GoTo cmdSimpan_Click_Error
2         isBegin = False
3             boleh = True
4             pesan = "Lengkapi data berikut "
              
5             If setHurufAngka(Text(1)) <= 0 Then
6                 pesan = pesan + "Hari Kerja "
7                 Text(1) = 0
8                 Text(1).SetFocus
9                 boleh = False
10            End If
11            If setHurufAngka(Text(2)) <= 0 Then
12                boleh = False
13                pesan = pesan + "Bobot Insentif. "
               
14            End If
               
15            If boleh Then
16                objHT.OpenTransactionHO
17                isBegin = True
                  
18                sSQL = "SELECT * FROM INSENTIF..GRAdtParameter WHERE periode = '" & Format(Periode.value, "MM-YYYY") & "' "
19                If objHT.FetchDataHO_String(sSQL) <> "" Then
20                    sSQL = "UPDATE INSENTIF..GRAdtParameter set JmlHariKerja = '" & setHurufAngka(Text(1)) & "',PorsiInsentif = '" & setHurufAngka(Text(2)) & "' WHERE periode = '" & Format(Periode.value, "MM-YYYY") & "' "
21                Else
22                    sSQL = "INSERT INTO INSENTIF..GRAdtParameter (periode,FS_KD_DEALER,JmlHariKerja,AdjKabeng,PorsiInsentif,UserEntry,tglentry) VALUES "
23                    sSQL = sSQL & " ('" & Format(Periode.value, "MM-YYYY") & "','','" & setHurufAngka(Text(1)) & "','0', " & setHurufAngka(Text(2)) & ",'" & gSA & "',getdate())"
               
24                End If
                  
                  
25                Call objHT.InsertData_Ho(sSQL)
26                Call objHT.CloseTransactionHO
              
                 ' Call ber_sih(True)
27                Call buka_Tutup(False, False)
28                Call refreshGrid
29                Exit Sub
                 ' cmdUbah.Caption = "&Ubah"
                 ' cmdSimpan.Caption = "&Tambah"
30            Else
31                MsgBox pesan, vbCritical, "VALID DATA"
32                Exit Sub
33            End If

34       On Error GoTo 0
35       Exit Sub

cmdSimpan_Click_Error:
36        If isBegin Then
37            Call objHT.CloseTransactionHO_err
38        End If
39        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure cmdSimpan_Click of Form frmHOInsentifSetResultKPICabang"


End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        newRecord = False
        Call buka_Tutup(False, True)
        KDPART_LAMA = Text(1)
        cmdUbah.Caption = "&Batal"
        cmdSimpan.Caption = "&Simpan"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdUbah.Caption = "&Ubah"
        cmdSimpan.Caption = "&Tambah"
    End If
End Sub

Private Sub Form_Load()
'    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridPart.Width = Frame1.Width - 200
    GridPart.Height = Frame1.Height - cmdSimpan.Top - Text(2).Height - 200
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
   ' cmdSimpan.Top = cmdKeluar.Top
    cmdGanti.Top = cmdKeluar.Top
   ' cmdImport.Top = cmdKeluar.Top
   ' cmdDownloadExcel.Top = cmdKeluar.Top
    ProgBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = Text(2).Left
  '  cmdDownloadExcel.Left = cmdSimpan.Left - cmdDownloadExcel.Width - 100
  '  cmdImport.Left = cmdDownloadExcel.Left - cmdImport.Width - 100
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
    If Not rsGrid.EOF Then
'        Text(6) = rsGrid.Fields(1)       'GridPart.Columns(0)
'        Text(0) = rsGrid.Fields(0)      'GridPart.Columns(1)    'GridPart.Columns(2)
'        Text(2) = rsGrid.Fields(2)       'GridPart.Columns(3)
'        Text(3) = rsGrid.Fields(3)       'GridPart.Columns(5)
'        Text(1) = Format(rsGrid.Fields(4), "00.00")
'        cboPerhitungan.Text = rsGrid.Fields(5)
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
                gBantuan = "INTERNAL"
                gSearchOption = "DEALER"
                frmCariHO.show (vbModal)
            Case 2
'                gMainKeyCari = Text(Index)
'                gSearchOption = "DATAMSTKPI"
'                frmCariHO.show (vbModal)
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
                Set oRs = objO.FetchDataHO("SELECT * FROM T_Dealer WITH (NOLOCK) WHERE FS_KD_Dealer = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gBantuan = "INTERNAL"
                    gSearchOption = "DEALER"
                    frmCariHO.show (vbModal)
                Else
                    Text(6) = oRs("FS_NM_Dealer")
                End If
                oRs.Close
            End If
        Case 2
'            If Text(Index) <> "" Then
'                'Set oRs = objO.FetchDataHO("SELECT * FROM jobOrder WITH (NOLOCK) WHERE FS_Item = '" & Text(Index) & "' and FS_LOG_DEL = ''")
'                Set oRs = objO.FetchDataHO("SELECT * FROM INSENTIF..GRMstKPI where ISNULL(UserDel,'') = '' and KdKPI = '" & Text(Index) & "'")
'                If oRs.BOF Then
'                    Cancel = True
'                    gMainKeyCari = Text(Index)
'                    gSearchOption = "DATAMSTKPI"
'                    gBantuan = "SBE"
'                    frmCariHO.show (vbModal)
'                Else
'                    Text(3) = NulSID(oRs!namaKPI)
'                End If
'                oRs.Close
'            End If
'
    End Select
End Sub

Private Sub isiId()
    Dim noUrut As Integer
    Set rsTemp = objO.FetchDataHO("SELECT TOP 1 * FROM INSENTIF..GRMstKPIPosisi WITH (NOLOCK) ORDER BY RIGHT(KdTipe,3) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KdTipe"), 3))
        noUrut = noUrut + 1
        Text(6) = "GR-" & Format(noUrut, "00#")
    Else
        noUrut = 1
        Text(6) = "GR-" & Format(noUrut, "00#")
    End If
End Sub

Private Function fungsiIsiId() As String
    Dim noUrut As Integer
    Dim hasil As String
    
    Set rsTemp = objO.FetchDataHO("SELECT TOP 1 * FROM INSENTIF..GRMstKPIPosisi WITH (NOLOCK) ORDER BY RIGHT(KdTipe,3) DESC")
    If Not rsTemp.EOF Then
        noUrut = CDbl(Right(rsTemp("KdTipe"), 3))
        noUrut = noUrut + 1
        hasil = "GR-" & Format(noUrut, "00#")
    Else
        noUrut = 1
        hasil = "GR-" & Format(noUrut, "00#")
    End If
    fungsiIsiId = hasil
End Function
