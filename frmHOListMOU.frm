VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmHOListMOU 
   BackColor       =   &H00FFFFFF&
   Caption         =   "LIST MOU"
   ClientHeight    =   10950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   18960
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   18960
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   9120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   7095
      Left            =   480
      TabIndex        =   26
      Top             =   480
      Width           =   12735
      Begin VB.TextBox txtLinkMou 
         Height          =   375
         Left            =   6360
         TabIndex        =   84
         Top             =   3000
         Width           =   6135
      End
      Begin VB.CheckBox ckNotifikasi 
         BackColor       =   &H0068ECE6&
         Caption         =   "Payment Notifications"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1260
         TabIndex        =   78
         Top             =   3600
         Width           =   3285
      End
      Begin VB.CheckBox ckPenagihan 
         BackColor       =   &H0068ECE6&
         Caption         =   "Tagihan Atas ID per Unit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1275
         TabIndex        =   77
         Top             =   840
         Width           =   3285
      End
      Begin VB.Frame frmeDealer 
         BackColor       =   &H0068ECE6&
         Height          =   3735
         Left            =   5160
         TabIndex        =   47
         Top             =   2880
         Width           =   7455
         Begin TrueOleDBGrid70.TDBGrid GridDealer 
            Height          =   3015
            Left            =   120
            TabIndex        =   17
            Top             =   480
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   5318
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "KODE DEALER"
            Columns(0).DataField=   "KODEDEALER"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "NAMA DEALER"
            Columns(1).DataField=   "NAMADEALER"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "JASA"
            Columns(2).DataField=   "DISCJASA"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "PART"
            Columns(3).DataField=   "DISCPART"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   4
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "BERLAKU"
            Columns(4).DataField=   "TERLIBAT"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
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
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bgcolor=&H808080&"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF00&"
            _StyleDefs(17)  =   ":id=9,.borderColor=&H80000001&"
            _StyleDefs(18)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(19)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(20)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(21)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(22)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(41)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(42)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(43)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(44)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(45)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(46)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(47)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(48)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(49)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
            _StyleDefs(50)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(51)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(52)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(53)  =   "Named:id=33:Normal"
            _StyleDefs(54)  =   ":id=33,.parent=0"
            _StyleDefs(55)  =   "Named:id=34:Heading"
            _StyleDefs(56)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(57)  =   ":id=34,.wraptext=-1"
            _StyleDefs(58)  =   "Named:id=35:Footing"
            _StyleDefs(59)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(60)  =   "Named:id=36:Selected"
            _StyleDefs(61)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(62)  =   "Named:id=37:Caption"
            _StyleDefs(63)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(64)  =   "Named:id=38:HighlightRow"
            _StyleDefs(65)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(66)  =   "Named:id=39:EvenRow"
            _StyleDefs(67)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(68)  =   "Named:id=40:OddRow"
            _StyleDefs(69)  =   ":id=40,.parent=33"
            _StyleDefs(70)  =   "Named:id=41:RecordSelector"
            _StyleDefs(71)  =   ":id=41,.parent=34"
            _StyleDefs(72)  =   "Named:id=42:FilterBar"
            _StyleDefs(73)  =   ":id=42,.parent=33"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   14
            Left            =   1200
            TabIndex        =   18
            Top             =   120
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
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
            Mask            =   "999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   15
            Left            =   2520
            TabIndex        =   19
            Top             =   120
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
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
            Mask            =   "999"
            PromptChar      =   ">"
         End
         Begin AIFCmp1.asxPowerButton cmdTutup 
            Height          =   240
            Left            =   120
            TabIndex        =   51
            Top             =   3480
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   423
            BackColor       =   12648447
            BorderStyle     =   3
            Caption         =   "Tut&up"
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
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DISC"
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
            TabIndex        =   50
            Top             =   165
            Width           =   360
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jasa               %"
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
            TabIndex        =   49
            Top             =   165
            Width           =   1170
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part              %"
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
            Left            =   2070
            TabIndex        =   48
            Top             =   165
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H0068ECE6&
         Height          =   2895
         Left            =   9480
         TabIndex        =   45
         Top             =   120
         Width           =   3135
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Status"
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
            Index           =   17
            Left            =   1680
            TabIndex        =   72
            Top             =   1560
            Width           =   1335
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Nama Dealer"
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
            Index           =   16
            Left            =   1680
            TabIndex        =   71
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Dealer Daftar"
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
            Index           =   15
            Left            =   1680
            TabIndex        =   70
            Top             =   1080
            Width           =   1335
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Disc Part"
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
            Index           =   14
            Left            =   1680
            TabIndex        =   69
            Top             =   840
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Disc Jasa"
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
            Index           =   13
            Left            =   1680
            TabIndex        =   68
            Top             =   600
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "TOP"
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
            Index           =   12
            Left            =   1680
            TabIndex        =   67
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tgl Selesai"
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
            Index           =   11
            Left            =   1680
            TabIndex        =   66
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Tgl Mulai"
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
            Index           =   10
            Left            =   120
            TabIndex        =   65
            Top             =   2520
            Width           =   1215
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "FAX"
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
            Index           =   9
            Left            =   120
            TabIndex        =   64
            Top             =   2280
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "PSTN"
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
            Index           =   8
            Left            =   120
            TabIndex        =   63
            Top             =   2040
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "CDMA"
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
            Index           =   7
            Left            =   120
            TabIndex        =   62
            Top             =   1800
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "GSM"
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
            Index           =   6
            Left            =   120
            TabIndex        =   61
            Top             =   1560
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "NPWP"
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
            Index           =   5
            Left            =   120
            TabIndex        =   60
            Top             =   1320
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "Kota"
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
            Index           =   4
            Left            =   120
            TabIndex        =   59
            Top             =   1080
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
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
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   58
            Top             =   840
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "PIC"
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
            Index           =   2
            Left            =   120
            TabIndex        =   57
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "PT"
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
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox chPilih 
            BackColor       =   &H0068ECE6&
            Caption         =   "No MOU"
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
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.ComboBox cboTOP 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   500
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0068ECE6&
         Caption         =   "                          DEALER                             "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   5160
         TabIndex        =   27
         Top             =   1200
         Width           =   4215
         Begin VB.CheckBox chAllDealer 
            BackColor       =   &H0068ECE6&
            Caption         =   "Semua Dealer"
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
            Left            =   960
            TabIndex        =   16
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   12
            Left            =   930
            TabIndex        =   15
            Top             =   360
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   503
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
            Height          =   285
            Index           =   13
            Left            =   930
            TabIndex        =   22
            Top             =   720
            Width           =   2985
            _ExtentX        =   5265
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
         Begin AIFCmp1.asxPowerButton cmdUnit 
            Height          =   360
            Left            =   2940
            TabIndex        =   52
            Top             =   240
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            BackColor       =   12632256
            BorderStyle     =   3
            Caption         =   "Isi Unit"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HotTracking     =   -1  'True
            Appearance      =   1
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   10
            Left            =   3240
            TabIndex        =   20
            Top             =   1080
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
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
            Mask            =   "999"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   285
            Index           =   11
            Left            =   3240
            TabIndex        =   21
            Top             =   1395
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   503
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
            Mask            =   "999"
            PromptChar      =   ">"
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Jasa               %"
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
            Left            =   2715
            TabIndex        =   54
            Top             =   1125
            Width           =   1170
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Part              %"
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
            Left            =   2790
            TabIndex        =   53
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Shape Shape3 
            BackStyle       =   1  'Opaque
            Height          =   270
            Left            =   1800
            Top             =   360
            Width           =   60
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Dealer Pendaftar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   -120
            TabIndex        =   29
            Top             =   405
            Width           =   960
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Berlaku"
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
            Left            =   255
            TabIndex        =   28
            Top             =   1245
            Width           =   525
         End
      End
      Begin MSComCtl2.DTPicker TglMulai 
         Height          =   255
         Left            =   6000
         TabIndex        =   11
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   64618497
         CurrentDate     =   39851
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1260
         TabIndex        =   3
         Top             =   1410
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin TrueOleDBGrid70.TDBGrid GridMOU 
         Height          =   3135
         Left            =   120
         TabIndex        =   30
         Top             =   3840
         Width           =   12495
         _ExtentX        =   22040
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   2
         Left            =   1260
         TabIndex        =   4
         Top             =   1710
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   0
         Left            =   1260
         TabIndex        =   0
         Top             =   240
         Width           =   3345
         _ExtentX        =   5900
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
         Left            =   1260
         TabIndex        =   5
         Top             =   2370
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   6
         Left            =   1260
         TabIndex        =   8
         Top             =   3270
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   3480
         TabIndex        =   9
         Top             =   3270
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   1260
         TabIndex        =   7
         Top             =   2970
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   9
         Left            =   3480
         TabIndex        =   10
         Top             =   2970
         Width           =   1665
         _ExtentX        =   2937
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
         Mask            =   "9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1260
         TabIndex        =   1
         Top             =   1110
         Width           =   3345
         _ExtentX        =   5900
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
      Begin MSComCtl2.DTPicker TglSelesai 
         Height          =   255
         Left            =   7800
         TabIndex        =   12
         Top             =   195
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   109707265
         CurrentDate     =   39851
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   4
         Left            =   1260
         TabIndex        =   6
         Top             =   2670
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   16
         Left            =   6000
         TabIndex        =   14
         Top             =   840
         Width           =   1665
         _ExtentX        =   2937
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
         Mask            =   "999999999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   17
         Left            =   1260
         TabIndex        =   2
         Top             =   555
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Height          =   285
         Index           =   18
         Left            =   1260
         TabIndex        =   79
         Top             =   2040
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin AIFCmp1.asxPowerButton btnPerforma 
         Height          =   360
         Left            =   7800
         TabIndex        =   81
         Top             =   765
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         BackColor       =   65280
         BorderStyle     =   3
         Caption         =   "Lihat Performa"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         HotTracking     =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Link Mou"
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
         Left            =   5505
         TabIndex        =   83
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Telp PIC"
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
         Left            =   375
         TabIndex        =   80
         Top             =   2085
         Width           =   840
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ID Pelanggan"
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
         Left            =   255
         TabIndex        =   76
         Top             =   600
         Width           =   960
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4665
         Top             =   570
         Width           =   60
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Limit"
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
         Left            =   5460
         TabIndex        =   73
         Top             =   885
         Width           =   315
      End
      Begin VB.Shape Shape5 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3240
         Top             =   2670
         Width           =   60
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
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
         Left            =   885
         TabIndex        =   44
         Top             =   2715
         Width           =   330
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
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
         Left            =   7455
         TabIndex        =   43
         Top             =   225
         Width           =   225
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. MOU"
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
         Left            =   540
         TabIndex        =   41
         Top             =   285
         Width           =   645
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4665
         Top             =   255
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perusahaan"
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
         Left            =   330
         TabIndex        =   40
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIC"
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
         Left            =   960
         TabIndex        =   39
         Top             =   1755
         Width           =   255
      End
      Begin VB.Label Label2 
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
         Left            =   720
         TabIndex        =   38
         Top             =   2415
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM"
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
         Left            =   885
         TabIndex        =   37
         Top             =   3315
         Width           =   315
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA"
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
         Left            =   3030
         TabIndex        =   36
         Top             =   3330
         Width           =   435
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN"
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
         Left            =   840
         TabIndex        =   35
         Top             =   3015
         Width           =   375
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax"
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
         Left            =   3165
         TabIndex        =   34
         Top             =   3015
         Width           =   270
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4665
         Top             =   1125
         Width           =   60
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
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
         Left            =   780
         TabIndex        =   33
         Top             =   1155
         Width           =   435
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal"
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
         Left            =   5175
         TabIndex        =   32
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOP"
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
         Left            =   5475
         TabIndex        =   31
         Top             =   555
         Width           =   300
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7995
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   14102
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "   Pendaftaran MOU Dealer"
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
      Width           =   12780
      _ExtentX        =   22543
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   8640
      TabIndex        =   23
      Top             =   7680
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
      Left            =   9600
      TabIndex        =   24
      Top             =   7680
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
      Left            =   10560
      TabIndex        =   25
      Top             =   7680
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
      Left            =   11520
      TabIndex        =   42
      Top             =   7680
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
      Left            =   7560
      TabIndex        =   55
      Top             =   7680
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
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
   Begin AIFCmp1.asxPowerButton cmdImport 
      Height          =   360
      Left            =   6000
      TabIndex        =   74
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
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
      Height          =   360
      Left            =   1920
      TabIndex        =   75
      Top             =   7680
      Visible         =   0   'False
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   635
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
   Begin AIFCmp1.asxPowerButton cmdExport 
      Height          =   360
      Left            =   4440
      TabIndex        =   82
      Top             =   7680
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   635
      BackColor       =   12648447
      Caption         =   "Export Excel"
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
End
Attribute VB_Name = "frmHOListMOU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 06 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsGrid As New ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim rsDealer As New ADODB.Recordset
Dim newRecord As Boolean
Dim tKdDealer As String
Dim cols As TrueOleDBGrid70.Columns

Private Sub ber_sih()
    
    For i = 0 To 16
        Select Case i
            Case 6, 7, 8, 9
                Text(i) = ""
                Text(i).Mask = "9999-9999999999"
            Case 10, 11, 14, 15, 16
                Text(i) = 0
            Case Else
                Text(i) = ""
        End Select
    Next
    Text(18) = ""
    txtLinkMou.Text = ""
    TglMulai = isnow
    TglSelesai = isnow
    chAllDealer.value = 1
    ckNotifikasi.value = vbUnchecked
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    
    Text(0).Enabled = Flg
    For i = 1 To 13
        Text(i).Enabled = flg1
    Next
    txtLinkMou.Enabled = flg1
    Text(18).Enabled = flg1
    Text(17).Enabled = flg1
    Text(16).Enabled = flg1
    TglMulai.Enabled = flg1
    TglSelesai.Enabled = flg1
    cboTOP.Enabled = flg1
    cmdUnit.Enabled = flg1
End Sub

Private Sub refreshGrid()
    GridMOU.ClearFields
        If rsGrid.State Then rsGrid.Close
        Set rsGrid = objO.FetchDataHO("SELECT NO_MOU as [NO MOU],cASe when TglMulai = TglSelesai then '9999' WHEN TglSelesai  < GETDATE() THEN '0' ELSE DATEDIFF(day,GETDATE(),tglselesai) end sisa , PERUSAHAAN, PIC AS [CONTACT PERSON], ALAMAT, KOTA, NPWP, GSM, CDMA, PSTN, FAX, " _
                                   & "TglMulai AS [MULAI BERLAKU], TglSelesai AS [SELESAI BERLAKU], [TOP], DiscJasa AS [DISC JASA], DiscPart AS [DISC PART], " _
                                   & "DealerDaft AS [DEALER YANG MENGAJUKAN], NamaDealer AS [NAMA DEALER], FlagDealer AS [STATUS KESERTAAN], LIMITAR AS [LIMIT],flagNotifWA_VA,telpPic,linkMou FROM MOUhdr WITH (NOLOCK) WHERE FS_LOG_DEL = '' ORDER BY no_MOU")
        If Not rsGrid.EOF Then
            GridMOU.DataSource = rsGrid
            Call headerGrid
        End If
    'End Ifx
End Sub

Private Sub headerGrid()
    
    GridMOU.Columns(0).Caption = "NO MOU"
    GridMOU.Columns(0).Width = IIf(chPilih(0) = Checked, 2000, 0)
    
    GridMOU.Columns(1).Caption = "Sisa Mou"
    GridMOU.Columns(1).Width = "500"
    GridMOU.Columns(2).Caption = "PERUSAHAAN"
    GridMOU.Columns(2).Width = IIf(chPilih(1) = Checked, 3500, 0)
    GridMOU.Columns(3).Caption = "CONTACT PERSON"
    GridMOU.Columns(3).Width = IIf(chPilih(2) = Checked, 2000, 0)
    GridMOU.Columns(4).Caption = "ALAMAT"
    GridMOU.Columns(4).Width = IIf(chPilih(3) = Checked, 3000, 0)
    GridMOU.Columns(5).Caption = "KOTA"
    GridMOU.Columns(5).Width = IIf(chPilih(4) = Checked, 1500, 0)
    GridMOU.Columns(6).Caption = "NPWP"
    GridMOU.Columns(6).Width = IIf(chPilih(5) = Checked, 2000, 0)
    GridMOU.Columns(7).Caption = "GSM"
    GridMOU.Columns(7).Width = IIf(chPilih(6) = Checked, 2000, 0)
    GridMOU.Columns(8).Caption = "CDMA"
    GridMOU.Columns(8).Width = IIf(chPilih(7) = Checked, 2000, 0)
    GridMOU.Columns(9).Caption = "PSTN"
    GridMOU.Columns(9).Width = IIf(chPilih(8) = Checked, 2000, 0)
    GridMOU.Columns(10).Caption = "FAX"
    GridMOU.Columns(10).Width = IIf(chPilih(9) = Checked, 2000, 0)
    GridMOU.Columns(11).Caption = "MULAI BERLAKU"
    GridMOU.Columns(11).Width = IIf(chPilih(10) = Checked, 2000, 0)
    GridMOU.Columns(12).Caption = "SELESAI BERLAKU"
    GridMOU.Columns(12).Width = IIf(chPilih(11) = Checked, 2000, 0)
    GridMOU.Columns(13).Caption = "TOP"
    GridMOU.Columns(13).Width = IIf(chPilih(12) = Checked, 1000, 0)
    GridMOU.Columns(14).Caption = "JASA"
    GridMOU.Columns(14).Width = IIf(chPilih(13) = Checked, 800, 0)
    GridMOU.Columns(15).Caption = "PART"
    GridMOU.Columns(15).Width = IIf(chPilih(14) = Checked, 800, 0)
    GridMOU.Columns(16).Caption = "DEALER YANG MENGAJUKAN"
    GridMOU.Columns(16).Width = IIf(chPilih(15) = Checked, 1000, 0)
    GridMOU.Columns(17).Caption = "NAMA DEALER"
    GridMOU.Columns(17).Width = IIf(chPilih(16) = Checked, 1000, 0)
    GridMOU.Columns(18).Caption = "STATUS"
    GridMOU.Columns(18).Width = IIf(chPilih(17) = Checked, 800, 0)
    GridMOU.Columns(19).Caption = "Notifikasi"
    GridMOU.Columns(20).Caption = "telpPic"
End Sub

Private Sub createRS()
    
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    With rsDealer
        .Fields.Append "KODEDEALER", adVarChar, 15
        .Fields.Append "NAMADEALER", adVarChar, 25
        .Fields.Append "DISCJASA", adInteger
        .Fields.Append "DISCPART", adInteger
        .Fields.Append "TERLIBAT", adBoolean
        .Open
    End With
    
    Set GridDealer.DataSource = rsDealer
End Sub

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 2000
    GridDealer.Columns(2).Width = 800
    GridDealer.Columns(3).Width = 800
    GridDealer.Columns(4).Width = 1000
    GridDealer.Columns(4).Alignment = dbgCenter
End Sub

Private Sub isiDealer()
Dim rsTmp As ADODB.Recordset
    Call HeaderCol
    Set rsTmp = objO.FetchDataHO("SELECT fs_kd_dealer, fs_nm_dealer, isnull(discJasa,0) as disJasa, isnull(discpart,0) as disPart, ISNULL(KODEDEALER,'') AS KD FROM t_dealer a WITH (NOLOCK) " _
                            & "LEFT JOIN moucbg b WITH (NOLOCK) ON a.fs_kd_dealer = b.kodedealer and A.FS_LOG_DEL = '' AND b.no_mou = '" & Text(0) & "' WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
    If Not rsTmp.EOF Then
        Call createRS
        Do While Not rsTmp.EOF
            rsDealer.AddNew
            rsDealer!KodeDealer = rsTmp("FS_KD_DEALER")
            rsDealer!namaDealer = rsTmp("FS_NM_DEALER")
            rsDealer!discJasa = rsTmp("DISJASA")
            rsDealer!discPart = rsTmp("DISPART")
            If Trim(rsTmp("KD")) = "" Then
                rsDealer!terlibat = False
            Else
                rsDealer!terlibat = True
            End If
            rsDealer.Update
            rsTmp.MoveNext
        Loop
        rsDealer.MoveFirst
        GridDealer.DataSource = rsDealer
        GridDealer.EvenRowStyle.BackColor = vbYellow
        GridDealer.OddRowStyle.BackColor = vbWhite
        GridDealer.AlternatingRowStyle = True
    End If
End Sub

Private Sub isiCombo()
        Set rsTemp = objO.FetchDataHO("SELECT [TOP] FROM t_TOP WITH (NOLOCK) ORDER BY [TOP]")
        If Not rsTemp.EOF Then
            cboTOP.Clear
            Do While Not rsTemp.EOF
                cboTOP.AddItem (rsTemp!Top)
                rsTemp.MoveNext
            Loop
        End If
End Sub

Private Sub btnPerforma_Click()
    If Text(0) <> "" Then
        gBantuan = Text(0)
        If TglMulai = TglSelesai Then
            frmHOPencapaianMOU.Text(6) = "999999"
        Else
            frmHOPencapaianMOU.Text(6) = DateDiff("d", Format(TglMulai, "yyyy-mm-dd"), Format(TglSelesai, "yyyy-mm-dd"))
            
        End If
        
        frmHOPencapaianMOU.Text(0) = Text(0)
        frmHOPencapaianMOU.Text(1) = Text(1)
        frmHOPencapaianMOU.show
    Else
        MsgBox "Isi data MOU"
        Exit Sub
    End If
End Sub

Private Sub cboTOP_KeyPress(KeyAscii As Integer)
    Enter KeyAscii
End Sub

Private Sub chAllDealer_Click()
    If chAllDealer.value = Checked Then
        frmeDealer.Visible = False
        Label20.Visible = True
        Label13.Visible = True
        Text(10).Visible = True
        Text(11).Visible = True
    ElseIf chAllDealer.value = Unchecked Then
        Label20.Visible = False
        Label13.Visible = False
        Text(10).Visible = False
        Text(11).Visible = False
        If cmdSimpan.Caption = "&Simpan" Then
            frmeDealer.Visible = True
            Call isiDealer
        End If
    End If
End Sub

Private Sub chPilih_Click(Index As Integer)
    If Not rsGrid.EOF Then
    Select Case Index
        Case 0
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 1
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 3500, 0)
        Case 2
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 3
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 3000, 0)
        Case 4
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1500, 0)
        Case 5
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 6
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 7
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 8
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 9
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 10
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 11
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 2000, 0)
        Case 12
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1000, 0)
        Case 13
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 800, 0)
        Case 14
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 800, 0)
        Case 15
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1000, 0)
        Case 16
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 1000, 0)
        Case 17
            GridMOU.Columns(Index).Width = IIf(chPilih(Index) = Checked, 800, 0)
    End Select
    End If
End Sub

Private Sub ckPenagihan_Click()
    Dim i As Integer
    If ckPenagihan.value = vbChecked Then
        Text(17).BackColor = &HC0E0FF
        For i = 1 To 9
            If i = 2 Then
                Text(i).BackColor = &HC0E0FF
            Else
                Text(i).BackColor = &HFFFFFF
            End If
'            Text(i).BackColor = &HFFFFFF
        Next
    Else
        Text(17).BackColor = &HFFFFFF
        For i = 1 To 9
            If i = 2 Then
                Text(i).BackColor = &HFFFFFF
            Else
                Text(i).BackColor = &HC0E0FF
            End If
            
        Next
    End If
End Sub

Private Sub cmdExport_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
2         If rsGrid.RecordCount = 0 Then
3             Er.HeadCap = "Export To Excel !"
4             Er.InfoCap = "List Data Reminder Service Kosong !"
5             Beep
6             Er.show (vbModal)
7             Exit Sub
8         End If
          Dim MyPath As String
          Dim onVal As String
9         MyPath = ""
          Dim f As New FormExplorer
10        f.show (vbModal)
11        MyPath = f.PathSaving
12        If MyPath = "" Then Exit Sub
13        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
14        If f.FileName = "" Then Exit Sub
15        onVal = MyPath & f.FileName & ".XLS"
16       ' GridMOU.Columns(7).Visible = False
17         GridMOU.ExportToFile onVal, True
18       ' GridMOU.Columns(7).Visible = True
19                Er.HeadCap = "Export to excel !"
20                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
21                Beep
22                Er.show (vbModal)
23         Exit Sub
Err_Export:
24          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
25          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:
End Sub

Private Sub cmdHapus_Click()
    If Text(0) <> "" Then
        Set rsTemp = objO.FetchDataHO("SELECT no_MOU FROM mouhdr WITH (NOLOCK) WHERE no_MOU = '" & Text(0) & "' and FS_LOG_DEL = ''")
        If Not rsTemp.EOF Then
            temp = MsgBox("Yakin akan menghapus MOU : " & Text(0) & vbCr & "Atas Nama : " & Text(1), vbYesNo, "HAPUS DATA")
            If temp = vbYes Then
                Call objO.UpdateDataHO("UPDATE MouHdr SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' WHERE no_MOU = '" & Text(0) & "' and FS_LOG_DEL = ''")
                Call objO.DeleteDataHO("DELETE MOUdtl WHERE no_MOU = '" & Text(0) & "'")
                Call objO.DeleteDataHO("DELETE MOUCBG WHERE no_MOU = '" & Text(0) & "'")
                Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('MOU', '" & Text(0) & "','D')")
            End If
        End If
    Else
        MsgBox "Pilih Data yang akan dihapus"
        Exit Sub
    End If
    Call ber_sih
    Call refreshGrid
End Sub

Private Sub cmdImport_Click()
    Dim ErrBegin As Boolean, konter As Long
  On Error GoTo vb_error
   Dim CnExel As New ADODB.Connection
   Dim RcExcel As New ADODB.Recordset
   Dim rsTempExcel As ADODB.Recordset
   Dim Jml As Integer
   Dim noMou As String
   
  ' Dim konter As Integer
   
  konter = 0
  progBar.Min = 0
  CommonDialog1.FileName = ""
  CommonDialog1.InitDir = App.Path
  CommonDialog1.filter = "Exel File (*.xls)|*.xls"
  CommonDialog1.ShowOpen
  If CommonDialog1.FileName <> "" Then
     If CnExel.State Then CnExel.Close
     CnExel.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & "Data Source=" & CommonDialog1.FileName & ";" & "Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1"""
     GoSub Proses
 Else
    Exit Sub
 End If
 Exit Sub
 
Proses:
ErrBegin = True
   Dim VarNmFile As String
   Dim namaDealer As String
   Dim namaDealerCbg As String
   Dim namaModel As String
   
   Call objHT.OpenTransactionHO
    gBantuan = "SBE"
   If RcExcel.State Then RcExcel.Close
   RcExcel.Open "Select * from [sheet1$]", CnExel, adOpenKeyset
   If RcExcel.RecordCount > 0 Then
      progBar.Visible = True
      progBar.Max = RcExcel.RecordCount
      RcExcel.MoveFirst
      Do Until RcExcel.EOF
         DoEvents
         
         progBar.value = konter
         model = ""
         namaModel = ""
         If Nul(Trim(RcExcel.Fields(0).value)) = "" Then Exit Do
         noMou = Nul(Trim(RcExcel.Fields(0).value))
        If Nul(Trim(RcExcel.Fields(0).value)) = "" Then
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "No MOU baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
        End If
        If Nul(Trim(RcExcel.Fields(1).value)) = "" Then
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "No NPWP baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
        End If
        If Nul(Trim(RcExcel.Fields(2).value)) = "" Then
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "Nama Perusahaan baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
        End If
        If Nul(Trim(RcExcel.Fields(3).value)) = "" Then
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "PIC baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
        End If
        If Nul(Trim(RcExcel.Fields(4).value)) = "" Then
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "Alamat baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
        End If
        If Nul(Trim(RcExcel.Fields(5).value)) = "" Then
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "Kota baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
        End If
        If Nul(Trim(RcExcel.Fields(6).value)) = "" And Nul(Trim(RcExcel.Fields(7).value)) = "" And Nul(Trim(RcExcel.Fields(8).value)) = "" And Nul(Trim(RcExcel.Fields(9).value)) = "" Then
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "No Telepon baris ke-" & konter & " salah / kosong !", vbCritical, "Error"
            Exit Sub
        End If
        Set rsTempExcel = objO.FetchDataHO("select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1 AND FS_KD_DEALER = '" & Trim(RcExcel.Fields(14).value) & "'")
        If rsTempExcel.RecordCount > 0 Then
            namaDealer = Trim(rsTempExcel("FS_NM_DEALER"))
        Else
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "Kode dealer " & RcExcel.Fields(14).value & " baris ke-" & konter & " salah / tidak ada data !", vbCritical, "Error"
            Exit Sub
        End If
        
        If Trim(RcExcel.Fields(18).value) <> "" Then
            If Nul(Trim(RcExcel.Fields(19).value)) = "" And Nul(Trim(RcExcel.Fields(20).value)) = "" Then
                Call objHT.CloseTransactionHO_err
                progBar.Visible = False
                MsgBox "No rangka / no polisi baris ke-" & konter & " salah / tidak ada data !", vbCritical, "Error"
                Exit Sub
            Else
                Set rsTempExcel = objO.FetchDataHO("select MDL,[DESC],NO_CHASSIS from model where FS_LOG_DEL = '' AND mdl = '" & Trim(RcExcel.Fields(18).value) & "'")
                If rsTempExcel.RecordCount > 0 Then
                    namaModel = Trim(rsTempExcel("desc"))
                Else
                    Call objHT.CloseTransactionHO_err
                    progBar.Visible = False
                    MsgBox "Kode model " & RcExcel.Fields(18).value & " baris ke-" & konter & " salah / tidak ada data !", vbCritical, "Error"
                    Exit Sub
                End If
            End If
        End If
        
        ''2015-12-01
        If Nul(RcExcel.Fields(17).value) = "" Then
            Call objHT.CloseTransactionHO_err
            progBar.Visible = False
            MsgBox "Cabang dealer kolom 18 baris ke-" & konter & " kosong !" & vbCrLf & "Silahkan isi 1 untuk kesemua dealer, 0 untuk dealer pendaftar, atau isikan kode dealer lainnya.", vbCritical, "Error"
            Exit Sub
        Else
            If Nul(RcExcel.Fields(17).value) <> "1" Or Nul(RcExcel.Fields(17).value) <> "0" Then
                Set rsTempExcel = objO.FetchDataHO("select FS_KD_DEALER, FS_NM_DEALER from t_dealer WITH(NOLOCK) WHERE FS_LOG_DEL = '' AND INTERNAL = 1 AND FS_KD_DEALER = '" & Trim(RcExcel.Fields(17).value) & "'")
                If Not rsTempExcel.EOF Then
                    namaDealerCbg = Trim(rsTempExcel("FS_NM_DEALER"))
                Else
                    Call objHT.CloseTransactionHO_err
                    progBar.Visible = False
                    MsgBox "Kode dealer " & RcExcel.Fields(14).value & " baris ke-" & konter & " tidak terdaftar !", vbCritical, "Error"
                    Exit Sub
                End If
            End If
        End If
         If Nul(RcExcel.Fields(0).value) = "" Then Exit Do
            Call ClearsSql
            gsSql1(0) = "NO_MOU"
            gsSql2(0) = Nul(Trim(RcExcel.Fields(0).value))
            gsSql1(1) = "PERUSAHAAN"
            gsSql2(1) = Nul(Trim(RcExcel.Fields(2).value))
            gsSql1(2) = "PIC"
            gsSql2(2) = Nul(Trim(RcExcel.Fields(3).value))
            gsSql1(3) = "ALAMAT"
            gsSql2(3) = Nul(Trim(RcExcel.Fields(4).value))
            gsSql1(4) = "KOTA"
            gsSql2(4) = Nul(Trim(RcExcel.Fields(5).value))
            gsSql1(5) = "NPWP"
            gsSql2(5) = Nul(Trim(RcExcel.Fields(1).value))
            gsSql1(6) = "GSM"
            gsSql2(6) = Replace(Replace(Nul(Trim(RcExcel.Fields(7).value)), ">", ""), " ", "")
            gsSql1(7) = "CDMA"
            gsSql2(7) = Replace(Replace(Nul(Trim(RcExcel.Fields(8).value)), ">", ""), " ", "")
            gsSql1(8) = "PSTN"
            gsSql2(8) = Replace(Replace(Nul(Trim(RcExcel.Fields(6).value)), ">", ""), " ", "")
            gsSql1(9) = "FAX"
            gsSql2(9) = Replace(Replace(Nul(Trim(RcExcel.Fields(9).value)), ">", ""), " ", "")
            gsSql1(10) = "TGLMULAI"
            gsSql2(10) = Format(TglYYYYMMDD(Format(Nul(Trim(RcExcel.Fields(10).value)), "DD/MM/YYYY")), "MM/DD/YYYY")
            gsSql1(11) = "TGLSELESAI"
            gsSql2(11) = Format(TglYYYYMMDD(Format(Nul(Trim(RcExcel.Fields(11).value)), "DD/MM/YYYY")), "MM/DD/YYYY")
            gsSql1(12) = "[TOP]"
            gsSql2(12) = Nul(Trim(RcExcel.Fields(12).value))
            gsSql1(13) = "DISCJASA"
            gsSql2(13) = Replace(Nul(Trim(RcExcel.Fields(15).value)), ",", ".")
            gsSql1(14) = "DISCPART"
            gsSql2(14) = Nul(Trim(RcExcel.Fields(16).value))
            gsSql1(15) = "DEALERDAFT"
            gsSql2(15) = Nul(Trim(RcExcel.Fields(14).value))
            gsSql1(17) = "NAMADEALER"
            gsSql2(17) = namaDealer
            
            gsSql1(16) = "FLAGDEALER"
            If Nul(Trim(RcExcel.Fields(17).value)) <> "" Then
                gsSql2(16) = "A"
            Else
                gsSql2(16) = "V"
            End If
            gsSql1(18) = "USERLOG"
            gsSql2(18) = gSA
            gsSql1(18) = "INPUTDATE"
            gsSql2(18) = Format(TglYYYYMMDD(Format(isnow, "DD/MM/YYYY")), "MM/DD/YYYY")   'Now()
            
            ' SIMPAN DATA
            Set rsTemp = objO.FetchDataHO("SELECT * FROM MOUHdr WITH (NOLOCK) WHERE no_MOU = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' and FS_LOG_DEL = ''")
            If Not rsTemp.EOF Then
                sSQL = GensSql2("MOUHDR", True) + " where NO_MOU = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' and FS_LOG_DEL = ''"
            Else
                ' DATA BARU
                sSQL = GensSql2("MOUHDR", False)
            End If
            If rsTemp.State <> adStateClosed Then rsTemp.Close
            Call objHT.InsertData_Ho(sSQL)
            Call objHT.InsertData_Ho("UPDATE MOUHDR SET LIMITAR = " & CDbl(Nul(Trim(RcExcel.Fields(13).value))) & " WHERE no_MOU = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' and FS_LOG_DEL = ''")
        
            ' SIMPAN CABANG
            If Nul(Trim(RcExcel.Fields(17).value)) = "1" Then
                Set rsTemp = objO.FetchData("SELECT * FROM T_DEALER WITH (NOLOCK) WHERE INTERNAL = 1 ORDER BY FS_KD_DEALER")
                If Not rsTemp.EOF Then
                    Do While Not rsTemp.EOF
                        Set rsHlp = objO.FetchDataHO("SELECT * FROM moucbg WITH (NOLOCK) WHERE no_mou = '" & Nul(Trim(RcExcel.Fields(0).value) & "' AND kodeDealer = '" & Nul(Trim(rsTemp("FS_KD_DEALER")))) & "'")
                        If Not rsHlp.EOF Then
                            Call objHT.InsertData_Ho("UPDATE moucbg SET discJasa = '" & CDbl(Nul(Trim(RcExcel.Fields(15).value))) & "', discPart = '" & CDbl(Nul(Trim(RcExcel.Fields(16).value))) & "' WHERE no_mou = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' AND kodeDealer = '" & Nul(Trim(rsTemp("FS_KD_DEALER"))) & "'")
                        Else
                            Call objHT.InsertData_Ho("INSERT INTO MOUCBG (NO_MOU,KODEDEALER,NAMADEALER,DISCJASA,DISCPART) VALUES ('" & Nul(Trim(RcExcel.Fields(0).value)) & "','" & Nul(Trim(rsTemp("FS_KD_DEALER"))) & "','" & Nul(Trim(rsTemp("FS_NM_DEALER"))) & "','" & CDbl(Nul(Trim(RcExcel.Fields(15).value))) & "','" & CDbl(Nul(Trim(RcExcel.Fields(16).value))) & "')")
                        End If
                         
                        rsTemp.MoveNext
                    Loop
                End If
                If rsTemp.State <> adStateClosed Then rsTemp.Close
                
            Else
                If Nul(Trim(RcExcel.Fields(17).value)) = "0" Then
                    Set rsTemp = objO.FetchDataHO("SELECT * FROM moucbg WITH (NOLOCK) WHERE no_mou = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' AND kodeDealer = '" & Nul(Trim(RcExcel.Fields(14).value)) & "'")
                Else
                    Set rsTemp = objO.FetchDataHO("SELECT * FROM moucbg WITH (NOLOCK) WHERE no_mou = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' AND kodeDealer = '" & Nul(Trim(RcExcel.Fields(17).value)) & "'")
                End If
                
                If Not rsTemp.EOF Then
                    Call objHT.InsertData_Ho("UPDATE moucbg SET discJasa = '" & CDbl(Nul(Trim(RcExcel.Fields(15).value))) & "', discPart = '" & CDbl(Nul(Trim(RcExcel.Fields(16).value))) & "' WHERE no_mou = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' AND kodeDealer = '" & Nul(rsTemp!KodeDealer) & "'")
                Else
                    If Nul(Trim(RcExcel.Fields(17).value)) = "0" Then
                    
                        Call objHT.InsertData_Ho("INSERT INTO moucbg (NO_MOU,KODEDEALER,NAMADEALER,DISCJASA,DISCPART) VALUES ('" & Nul(Trim(RcExcel.Fields(0).value)) & "', '" & Nul(RcExcel.Fields(14).value) & "','" & namaDealer & "','" & CDbl(Nul(Trim(RcExcel.Fields(15).value))) & "','" & CDbl(Nul(Trim(RcExcel.Fields(16).value))) & "')")
                    Else
                        Call objHT.InsertData_Ho("INSERT INTO moucbg (NO_MOU,KODEDEALER,NAMADEALER,DISCJASA,DISCPART) VALUES ('" & Nul(Trim(RcExcel.Fields(0).value)) & "', '" & Nul(RcExcel.Fields(17).value) & "','" & namaDealerCbg & "','" & CDbl(Nul(Trim(RcExcel.Fields(15).value))) & "','" & CDbl(Nul(Trim(RcExcel.Fields(16).value))) & "')")
                    End If
                End If
            End If
            'simpan model
            ClearsSql
            
            gsSql1(0) = "NO_MOU"
            gsSql2(0) = Nul(Trim(RcExcel.Fields(0).value))
            gsSql1(1) = "MDL"
            gsSql2(1) = Nul(Trim(RcExcel.Fields(18).value))
            gsSql1(2) = "MODEL"
            gsSql2(2) = namaModel
            gsSql1(3) = "NO_RANGKA"
            gsSql2(3) = Nul(Trim(RcExcel.Fields(19).value))
            gsSql1(4) = "NO_POLISI"
            gsSql2(4) = Nul(Trim(RcExcel.Fields(20).value))
            
            Set rsTemp = objO.FetchDataHO("SELECT * FROM MOUDTL WITH (NOLOCK) WHERE NO_MOU = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' AND MDL = '" & Nul(Trim(RcExcel.Fields(18).value)) & "' AND NO_RANGKA = '" & Nul(Trim(RcExcel.Fields(19).value)) & "' ORDER BY MDL, NO_RANGKA")
            If Not rsTemp.EOF Then
                sSQL = GensSql2("MOUDTL", True) + "WHERE NO_MOU = '" & Nul(Trim(RcExcel.Fields(0).value)) & "' AND MDL = '" & Nul(Trim(RcExcel.Fields(18).value)) & "' AND NO_RANGKA = '" & Nul(Trim(RcExcel.Fields(19).value)) & "'"
            Else
                sSQL = GensSql2("MOUDTL", False)
            End If
            Call objHT.InsertData_Ho(sSQL)
            If Nul(Trim(RcExcel.Fields(0).value)) <> "" And noMou <> Nul(Trim(RcExcel.Fields(0).value)) Then
                Call objHT.InsertData_Ho("INSERT INTO TEMPKIRIM VALUES ('MOU', '" & Nul(Trim(RcExcel.Fields(0).value)) & "','S')")
            End If
         RcExcel.MoveNext
         konter = konter + 1
      Loop
   End If
   RcExcel.Close
   CnExel.Close
   progBar.Visible = False

    Call objHT.CloseTransactionHO
    Call refreshGrid
   MsgBox "Proses Import telah selesai " & konter & " data!", vbInformation
'   GridAFI.Refresh
   Return
   Exit Sub
vb_error:
   If ErrBegin = True Then
     ' DbCon.RollbackTrans
   End If
   Call objHT.CloseTransactionHO_err
   progBar.Visible = False
   MsgBox ErrMessage(Erl, err.Number, "Procedure : " & Me.Name & ".CmdImport_Click")
End Sub

Private Sub CmdKeluar_Click()
    Call objO.DeleteDataHO("DELETE tempMOUdtl WHERE no_MOU = '" & Trim(Text(0)) & "'")
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
      Dim boleh As Boolean
      Dim pesan As String
      Dim rsHlp As ADODB.Recordset

1         If cmdSimpan.Caption = "&Tambah" Then
2             Call ber_sih
3             Call buka_Tutup
4             cmdSimpan.Caption = "&Simpan"
5             cmdUbah.Caption = "&Batal"
6             newRecord = True
7             Text(0).SetFocus
8         ElseIf cmdSimpan.Caption = "&Simpan" Then
              
9             boleh = True
10            pesan = "Lengkapi Data Berikut "
11            If Text(0) = "" Then
12                boleh = False
13                pesan = pesan + "No. MOU "
14            End If
15            If Text(5) = "" Then
16                boleh = False
17                pesan = pesan + "NPWP. "
18            End If
19            If Text(1) = "" Then
20                boleh = False
21                pesan = pesan + "Perusahaan. "
22            End If
23            If Text(2) = "" Then
24                boleh = False
25                pesan = pesan + "PIC. "
26            End If
27            If Text(3) = "" Then
28                boleh = False
29                pesan = pesan + "Alamat. "
30            End If
31            If Text(4) = "" Then
32                boleh = False
33                pesan = pesan + "Kota. "
34            End If
35            If Text(6) = "" And Text(7) = "" And Text(8) = "" And Text(9) = "" Then
36                boleh = False
37                pesan = pesan + "Telpon. "
38            End If
39            If Text(12) = "" Then
40                boleh = False
41                pesan = pesan + "Dealer Pendaftar. "
42            End If
              
43            If CDbl(Text(16)) = 0 Then
44                If Trim(cboTOP) <> "" Then
45                    Set rsHlp = objO.FetchData("SELECT * FROM T_TOP WITH (NOLOCK) WHERE [TOP] = '" & Trim(cboTOP) & "'")
46                    If Not rsHlp.EOF Then
47                        If CDbl(rsHlp("RUMUS")) > 0 Then
48                            boleh = False
49                            pesan = pesan + "Limit Hutang. "
50                        End If
51                    End If
                       
52                End If
53            End If
              
54            If boleh Then
55                Call ClearsSql
56                gsSql1(0) = "NO_MOU"
57                gsSql2(0) = Text(0)
58                gsSql1(1) = "PERUSAHAAN"
59                gsSql2(1) = Text(1)
60                gsSql1(2) = "PIC"
61                gsSql2(2) = Text(2)
62                gsSql1(3) = "ALAMAT"
63                gsSql2(3) = Text(3)
64                gsSql1(4) = "KOTA"
65                gsSql2(4) = Text(4)
66                gsSql1(5) = "NPWP"
67                gsSql2(5) = Text(5)
68                gsSql1(6) = "GSM"
69                gsSql2(6) = Replace(Replace(Text(6), ">", ""), " ", "")
70                gsSql1(7) = "CDMA"
71                gsSql2(7) = Replace(Replace(Text(7), ">", ""), " ", "")
72                gsSql1(8) = "PSTN"
73                gsSql2(8) = Replace(Replace(Text(8), ">", ""), " ", "")
74                gsSql1(9) = "FAX"
75                gsSql2(9) = Replace(Replace(Text(9), ">", ""), " ", "")
76                gsSql1(10) = "TGLMULAI"
77                gsSql2(10) = Format(TglYYYYMMDD(Format(TglMulai, "DD/MM/YYYY")), "MM/DD/YYYY")
78                gsSql1(11) = "TGLSELESAI"
79                gsSql2(11) = Format(TglYYYYMMDD(Format(TglSelesai, "DD/MM/YYYY")), "MM/DD/YYYY")
80                gsSql1(12) = "[TOP]"
81                gsSql2(12) = Trim(cboTOP)
82                gsSql1(13) = "DISCJASA"
83                gsSql2(13) = Text(10)
84                gsSql1(14) = "DISCPART"
85                gsSql2(14) = Text(11)
86                gsSql1(15) = "DEALERDAFT"
87                gsSql2(15) = Text(12)
88                gsSql1(17) = "NAMADEALER"
89                gsSql2(17) = Text(13)
                  
90                gsSql1(16) = "FLAGDEALER"
91                If chAllDealer = Checked Then
92                    gsSql2(16) = "A"
93                Else
94                    gsSql2(16) = "V"
95                End If
96                gsSql1(18) = "USERLOG"
97                gsSql2(18) = gSA
98                gsSql1(19) = "INPUTDATE"
99                gsSql2(19) = Format(TglYYYYMMDD(Format(isnow, "DD/MM/YYYY")), "MM/DD/YYYY")   'Now()
                  
100               gsSql1(20) = "IDCUST"
101               gsSql2(20) = Nul(Text(17)) '"IDCUST"
                  
102               gsSql1(21) = "flagNotifWA_VA"
103               gsSql2(21) = IIf(ckNotifikasi = Checked, 1, 0)
                  
104               gsSql1(22) = "telpPic"
105               gsSql2(22) = Nul(Text(18))
                  
106               gsSql1(23) = "linkMou"
107               gsSql2(23) = Trim(txtLinkMou.Text)
                  
                  ' SIMPAN DATA
108               Set rsTemp = objO.FetchDataHO("SELECT * FROM MOUHdr WITH (NOLOCK) WHERE no_MOU = '" & Text(0) & "' and FS_LOG_DEL = ''")
109               If Not rsTemp.EOF Then
                      ' DATA SUDAH ADA
110                   If newRecord Then
111                       MsgBox "DATA SUDAH ADA. Jika ingin merubah pilih UBAH", vbCritical, "DUPLIKASI DATA"
112                   Else
113                       temp = MsgBox("Yakin akan menyimpan perubahan Data ?", vbYesNo, "UBAH DATA")
114                       If temp = vbYes Then
115                           sSQL = GensSql2("MOUHDR", True) + " where NO_MOU = '" & Text(0) & "' and FS_LOG_DEL = ''"
116                       End If
117                   End If
118               Else
                      ' DATA BARU
119                   sSQL = GensSql2("MOUHDR", False)
120               End If
121               If rsTemp.State <> adStateClosed Then rsTemp.Close
122               Call objO.UpdateDataHO(sSQL)
123               Call objO.UpdateDataHO("UPDATE MOUHDR SET LIMITAR = " & CDbl(Text(16)) & " WHERE no_MOU = '" & Text(0) & "' and FS_LOG_DEL = ''")
                  
                  ' SIMPAN DETAIL
124               Set rsTemp = objO.FetchDataHO("SELECT * FROM tempMOUdtl WITH (NOLOCK) WHERE no_mou = '" & Text(0) & "' ORDER BY MDL, No_Rangka")
125               If Not rsTemp.EOF Then
126                   Call objO.DeleteDataHO("DELETE MOUDTL WHERE no_Mou = '" & Text(0) & "'")
127                   Call objO.InsertDataHO("INSERT INTO MOUDTL (NO_MOU,MDL,MODEL,No_Polisi,No_Rangka,idcust) SELECT no_mou,mdl,model,no_polisi,no_rangka,idcust FROM tempMOUdtl WITH (NOLOCK) WHERE no_mou = '" & Text(0) & "' ORDER BY MDL, No_Rangka")
128                   Call objO.DeleteDataHO("DELETE tempMOUDTL WHERE no_Mou = '" & Text(0) & "'")
129               Else
130                   Call objO.DeleteDataHO("DELETE MOUDTL WHERE no_Mou = '" & Text(0) & "'")
131               End If
              
                  ' SIMPAN CABANG
132               If chAllDealer = Unchecked Then
133                   If rsDealer.State <> adStateClosed Then
134                       rsDealer.MoveFirst
135                       Do While Not rsDealer.EOF
136                           If rsDealer("TERLIBAT") Then
137                               Set rsTemp = objO.FetchDataHO("SELECT * FROM moucbg WITH (NOLOCK) WHERE no_mou = '" & Text(0) & "' AND kodeDealer = '" & rsDealer("KODEDEALER") & "'")
138                               If Not rsTemp.EOF Then
139                                   Call objO.UpdateDataHO("UPDATE moucbg SET discJasa = '" & CDbl(rsDealer("DISCJASA")) & "', discPart = '" & CDbl(rsDealer("DISCPART")) & "' WHERE no_mou = '" & Text(0) & "' AND kodeDealer = '" & rsDealer("KODEDEALER") & "'")
140                               Else
141                                   Call objO.InsertDataHO("INSERT INTO moucbg (NO_MOU,KODEDEALER,NAMADEALER,DISCJASA,DISCPART) VALUES ('" & Text(0) & "', '" & rsDealer("KODEDEALER") & "','" & rsDealer("NAMADEALER") & "','" & CDbl(rsDealer("DISCJASA")) & "','" & CDbl(rsDealer("DISCPART")) & "')")
142                               End If
143                               If rsTemp.State <> adStateClosed Then rsTemp.Close
144                           End If
145                           rsDealer.MoveNext
146                       Loop
147                   End If
148               Else
149                   Set rsTemp = objO.FetchData("SELECT * FROM T_DEALER WITH (NOLOCK) WHERE INTERNAL = 1 ORDER BY FS_KD_DEALER")
150                   If Not rsTemp.EOF Then
151                       Do While Not rsTemp.EOF
152                           Set rsHlp = objO.FetchDataHO("SELECT * FROM moucbg WITH (NOLOCK) WHERE no_mou = '" & Text(0) & "' AND kodeDealer = '" & Trim(rsTemp("FS_KD_DEALER")) & "'")
153                           If Not rsHlp.EOF Then
154                               Call objO.UpdateDataHO("UPDATE moucbg SET discJasa = '" & CDbl(Text(10)) & "', discPart = '" & CDbl(Text(11)) & "' WHERE no_mou = '" & Text(0) & "' AND kodeDealer = '" & Trim(rsTemp("FS_KD_DEALER")) & "'")
155                           Else
156                               Call objO.InsertDataHO("INSERT INTO MOUCBG (NO_MOU,KODEDEALER,NAMADEALER,DISCJASA,DISCPART) VALUES ('" & Trim(Text(0)) & "','" & Trim(rsTemp("FS_KD_DEALER")) & "','" & Trim(rsTemp("FS_NM_DEALER")) & "','" & CDbl(Text(10)) & "','" & CDbl(Text(11)) & "')")
157                           End If
                               
158                           rsTemp.MoveNext
159                       Loop
160                   End If
161                   If rsTemp.State <> adStateClosed Then rsTemp.Close
162               End If
                  
163               If Trim(Text(0)) <> "" Then
164                   Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('MOU', '" & Text(0) & "','S')")
165               End If
                  
166               Call ber_sih
167               Call buka_Tutup(False, False)
168               Call refreshGrid
                  
169               cmdSimpan.Caption = "&Tambah"
170               cmdUbah.Caption = "&Ubah"
171           Else
172               MsgBox pesan, vbCritical, "KELENGKAPAN"
173               Exit Sub
174           End If
              
175       End If
End Sub

Private Sub cmdTutup_Click()
    frmeDealer.Visible = False
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call buka_Tutup(False)
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
        newRecord = False
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call ber_sih
        Call buka_Tutup(False, False)
        cmdSimpan.Caption = "&Tambah"
        cmdUbah.Caption = "&Ubah"
    End If
End Sub

Private Sub cmdUnit_Click()
1         If Text(0) <> "" Then
2             If ckPenagihan.value = vbUnchecked Then
3                 If Nul(Text(17)) = "" Then
                  
4                     MsgBox "Isi data ID Customer"
5                     Exit Sub
6                 End If
7             End If
8                 If Nul(Text(17)) <> "" Then
9                     Set rsTemp = objO.FetchDataHO("SELECT * FROM [SingleIDOri]..NewCustomer WITH (NOLOCK) WHERE idCust = '" & Nul(Text(17)) & "'")
10                    If rsTemp.EOF Then
11                        MsgBox "ID Customer tidak terdaftar"
12                        Exit Sub
13                    End If
14                End If
                  
15                    If ckPenagihan.value = True Then
16                        Call objO.InsertDataHO("INSERT INTO TEMPMOUDTL SELECT no_mou,mdl,model,no_polisi,no_rangka,idcust FROM MOUdtl WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM tempMOUdtl WITH (NOLOCK) WHERE no_MOU = '" & Trim(Text(0)) & "') AND no_mou = '" & Trim(Text(0)) & "' ORDER BY MDL, No_Rangka")
17                    Else
18                        Call objO.InsertDataHO("INSERT INTO TEMPMOUDTL SELECT no_mou,mdl,model,no_polisi,no_rangka,'" & Nul(Text(17)) & "' as idcust FROM MOUdtl WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM tempMOUdtl WITH (NOLOCK) WHERE no_MOU = '" & Trim(Text(0)) & "') AND no_mou = '" & Trim(Text(0)) & "' ORDER BY MDL, No_Rangka")
19                    End If
                      
20                    gBantuan = Text(0)
21                    frmHOMOU.Text(0) = Text(0)
22                    frmHOMOU.Text(1) = Text(1)
23                    If ckPenagihan.value = vbChecked Then
24                        frmHOMOU.Text(17) = ""
25                        frmHOMOU.Text(17).Enabled = True
26                    Else
27                        frmHOMOU.Text(17) = Nul(Text(17))
28                        frmHOMOU.Text(17).Enabled = False
29                    End If
30                    frmHOMOU.show
                  
              
31        Else
32            MsgBox "Isi data MOU"
33            Exit Sub
34        End If
End Sub

Private Sub Form_Activate()
    If Log.LogHOConn = 1 Then
        If ConHO = False And Log.RPLHistori = 0 Then
            MsgBox "Tidak Connect"
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
    Call isiCombo
    
    For i = 0 To 17
        chPilih(i) = Checked
    Next
    
    frmeDealer.Visible = False
    TglMulai = isnow
    TglSelesai = isnow
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdKeluar.Height - 700
    GridMOU.Width = Frame1.Width - 200
    GridMOU.Height = Frame1.Height - Text(9).Top - Text(9).Height - 300
    cmdKeluar.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdKeluar.Top
    cmdUbah.Top = cmdKeluar.Top
    cmdSimpan.Top = cmdKeluar.Top
    cmdCetak.Top = cmdKeluar.Top
    cmdImport.Top = cmdKeluar.Top
    cmdExport.Top = cmdKeluar.Top
    progBar.Top = cmdKeluar.Top
    
    cmdKeluar.Left = Frame1.Left + Frame1.Width - cmdKeluar.Width - 100
    cmdHapus.Left = cmdKeluar.Left - cmdHapus.Width - 100
    cmdUbah.Left = cmdHapus.Left - cmdUbah.Width - 100
    cmdSimpan.Left = cmdUbah.Left - cmdSimpan.Width - 100
    cmdCetak.Left = cmdSimpan.Left - cmdCetak.Width - 100
    
    cmdImport.Left = cmdSimpan.Left - cmdImport.Width - 100
    cmdExport.Left = cmdImport.Left - cmdExport.Width - 100
    progBar.Left = cmdImport.Left - progBar.Width - 100
    
End Sub

Private Sub GridDealer_Click()
    If GridDealer.Columns(4).value = -1 Then
        Text(14) = GridDealer.Columns(2)
        Text(15) = GridDealer.Columns(3)
        tKdDealer = GridDealer.Columns(0)
        Text(14).SetFocus
    End If
End Sub
Private Sub GridDealer_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    Text(14) = GridDealer.Columns(2)
    Text(15) = GridDealer.Columns(3)
    tKdDealer = GridDealer.Columns(0)
End Sub

Private Sub GridMOU_FilterChange()
    On Error GoTo errHandler
    Set cols = GridMOU.Columns
    Dim C As Integer
    
    C = GridMOU.col
    GridMOU.HoldFields
    rsGrid.filter = getFilter()
    GridMOU.col = C
    GridMOU.EditActive = True
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

Private Sub GridMOU_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Dim rsHlp As New ADODB.Recordset

    If rsGrid.State <> adStateClosed Then
        If Not rsGrid.EOF Then
            Text(0) = rsGrid.Fields(0)      'GridMOU.Columns(0)
            Text(1) = rsGrid.Fields(2)      'GridMOU.Columns(1)
            Text(2) = rsGrid.Fields(3)      'GridMOU.Columns(2)
            Text(3) = rsGrid.Fields(4)      'GridMOU.Columns(3)
            Text(4) = rsGrid.Fields(5)      'GridMOU.Columns(4)
            Text(5) = rsGrid.Fields(6)      'GridMOU.Columns(5)
            For i = 6 To 9
                Text(i).Mask = ""
            Next
            Text(6) = rsGrid.Fields(7)      'GridMOU.Columns(6)
            Text(7) = rsGrid.Fields(8)      'GridMOU.Columns(7)
            Text(8) = rsGrid.Fields(9)      'GridMOU.Columns(8)
            Text(9) = rsGrid.Fields(10)      'GridMOU.Columns(9)
            TglMulai = rsGrid.Fields(11)    'GridMOU.Columns(10)
            TglSelesai = rsGrid.Fields(12)  'GridMOU.Columns(11)
            cboTOP.Text = Trim(rsGrid.Fields(13)) 'Trim(GridMOU.Columns(12))
            Text(10) = rsGrid.Fields(14)    'GridMOU.Columns(13)
            Text(11) = rsGrid.Fields(15)    'GridMOU.Columns(14)
            Text(12) = rsGrid.Fields(16)    'GridMOU.Columns(15)
            Text(13) = rsGrid.Fields(17)    'GridMOU.Columns(16)
            If rsGrid.Fields(18) = "A" Then 'GridMOU.Columns(17)
                chAllDealer.value = Checked
            Else
                chAllDealer.value = Unchecked
            End If
            Text(16) = CDbl(rsGrid.Fields(19))  'CDbl(GridMOU.Columns(18))
            ckNotifikasi = IIf(rsGrid!flagNotifWA_VA = 1, vbChecked, vbUnchecked)
            Text(18) = Trim(rsGrid!TelpPIC)
            txtLinkMou.Text = Trim(rsGrid!linkMou)
            Set rsHlp = objO.FetchDataHO("SELECT NO_MOU FROM MOUDTL WITH (NOLOCK) WHERE NO_MOU = '" & Trim(Text(0)) & "'")
            If Not rsHlp.EOF Then
                Call objO.InsertDataHO("INSERT INTO TEMPMOUDTL SELECT no_mou,mdl,model,no_polisi,no_rangka,idcust FROM MOUdtl WITH (NOLOCK) WHERE NOT EXISTS (SELECT * FROM tempMOUdtl WITH (NOLOCK) WHERE no_MOU = '" & Trim(Text(0)) & "') AND no_mou = '" & Trim(Text(0)) & "' ORDER BY MDL, No_Rangka")
            End If
            rsHlp.Close
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
                gSearchOption = "MOU"
                frmCariHO.show (vbModal)
            Case 5
                gMainKeyCari = Text(Index)
                gSearchOption = "PT_MOU"
                frmCariHO.show (vbModal)
            Case 4
                gMainKeyCari = Text(Index)
                gSearchOption = "CITY"
                frmCariHO.show (vbModal)
            Case 12
                gMainKeyCari = Text(Index)
                gBantuan = "INTERNAL"
                gSearchOption = "DEALER"
                frmCariHO.show (vbModal)
            Case 17
                If Text(Index).BackColor <> &HC0E0FF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SID"
                    frmCariHO.show (vbModal)
                End If
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

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 10, 11
            If CDbl(Text(Index)) > 100 Or CDbl(Text(Index)) < 0 Then   'CDbl(text(Index)) > 10 Or
                MsgBox "Discount Jasa salah. Range dalam 0 - 10%"
                If Text(Index) > 10 Then Text(Index) = 10
                If Text(Index) < 0 Then Text(Index) = 0
                Text(Index).SelLength = Len(Text(Index))
                Text(Index).SetFocus
            End If
        Case 14, 15
            If tKdDealer <> "" Then
                If Text(Index) = "" Then
                    Text(Index) = 0
                Else
                If CDbl(Text(14)) > 0 Then
                    rsDealer.Find "KODEDEALER = '" & tKdDealer & "'", , adSearchForward, 1
                    If Not rsDealer.EOF Then
                        rsDealer!discJasa = CDbl(Text(14))
                        rsDealer.Update
                        Text(14) = 0
                    End If
                End If
                If CDbl(Text(15)) > 0 Then
                    rsDealer.Find "KODEDEALER = '" & tKdDealer & "'", , adSearchForward, 1
                    If Not rsDealer.EOF Then
                        rsDealer!discPart = CDbl(Text(15))
                        rsDealer.Update
                        Text(15) = 0
                    End If
                End If
                End If
            End If
    End Select
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
Dim tmp As String

    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                n = Trim(Text(Index))
                Call ber_sih
                Text(Index) = n
                Set oRs = objO.FetchDataHO("Select * from MOUhdr WITH (NOLOCK) where no_MOU ='" & Text(Index) & "' and FS_LOG_DEL = ''")
                If Not oRs.EOF Then
                    Text(1) = oRs("Perusahaan")
                    Text(2) = oRs("PIC")
                    Text(3) = oRs("Alamat")
                    Text(4) = oRs("Kota")
                    Text(5) = oRs("NPWP")
                    Text(6) = oRs("GSM")
                    Text(7) = oRs("CDMA")
                    Text(8) = oRs("PSTN")
                    Text(9) = oRs("FAX")
                End If
                oRs.Close
            End If
        Case 4
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("Select * from T_CITY WITH (NOLOCK) where kota = '" & Text(Index) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "CITY"
                    frmCariHO.show (vbModal)
                End If
                oRs.Close
            End If
        Case 5
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("SELECT * FROM MOUHDR WITH (NOLOCK) WHERE NPWP = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If Not oRs.EOF Then
                    Text(1) = oRs("PERUSAHAAN")
                    Text(2) = oRs("PIC")
                    Text(3) = oRs("ALAMAT")
                    Text(4) = oRs("KOTA")
                    Text(6) = oRs("GSM")
                    Text(7) = oRs("CDMA")
                    Text(8) = oRs("PSTN")
                    Text(9) = oRs("FAX")
                End If
                oRs.Close
            End If
        Case 12
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataHO("SELECT * FROM T_Dealer WITH (NOLOCK) WHERE FS_KD_Dealer = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gBantuan = "INTERNAL"
                    gSearchOption = "DEALER"
                    frmCariHO.show (vbModal)
                Else
                    Text(13) = oRs("FS_NM_Dealer")
                End If
                oRs.Close
            End If
        Case 14
            If Text(Index) = "" Then
                Text(Index) = 0
            ElseIf Text(Index) < 0 Or Text(Index) > 100 Then
                MsgBox "Discount diantara 1 - 100% !!"
                Text(Index).SelLength = Len(Text(Index))
                Text(Index).SetFocus
                Exit Sub
            Else
                With rsDealer
                tmp = GridDealer.Columns(0)
                If .State <> adStateClosed Then
                    .MoveFirst
                    .Find "KODEDEALER = '" & Trim(tmp) & "'", adSearchForward, 1
                    If Not rsDealer.EOF Then
                        !discJasa = CDbl(Text(Index))
                        .Update
                    End If
                End If
                End With
            End If
        Case 15
            If Text(Index) = "" Then
                Text(Index) = 0
            ElseIf Text(Index) < 0 Or Text(Index) > 10 Then
                MsgBox "Discount diantara 1 - 10% !!"
                Text(Index).SelLength = Len(Text(Index))
                Text(Index).SetFocus
                Exit Sub
            Else
                With rsDealer
                tmp = GridDealer.Columns(0)
                If .State <> adStateClosed Then
                    .MoveFirst
                    .Find "KODEDEALER = '" & Trim(tmp) & "'", adSearchForward, 1
                    If Not rsDealer.EOF Then
                        !discPart = CDbl(Text(Index))
                        .Update
                    End If
                End If
                End With
            End If
        Case 17
            If Nul(Text(Index)) = "" Then
                If Text(Index).BackColor <> &HC0E0FF Then
                gMainKeyCari = Text(Index)
                gSearchOption = "SID"
                frmCariHO.show (vbModal)
                End If
                
            Else
                Set oRs = objO.FetchDataHO("select IdTelp + Telp as tlp,HP1,HP2,Alamat,kota,NoNPWP,Title + ' ' + NamaCust1 + ' ' + NamaCust2 + ' ' + NamaCust3 as nama,idcust from [singleidori]..NewCustomer WITH (nolock) where idcust = '" & Nul(Text(Index)) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SID"
                    frmCariHO.show (vbModal)
                Else
                    Text(1) = Nul(oRs!Nama)
                    'Text(2) = Nul(oRs!)
                    Text(3) = Nul(oRs!Alamat)
                    Text(4) = Nul(oRs!Kota)
                    Text(5) = Nul(oRs!noNPWP)
                    Text(6) = Nul(oRs!hp1)
                    Text(7) = Nul(oRs!hp2)
                    Text(8) = Nul(oRs!tlp)
                    'Text(9) = Nul(oRs!)
                End If
            End If
    End Select
End Sub

Private Sub TglMulai_CloseUp()
    TglSelesai.SetFocus
End Sub

Private Sub TglSelesai_CloseUp()
    cboTOP.SetFocus
End Sub
