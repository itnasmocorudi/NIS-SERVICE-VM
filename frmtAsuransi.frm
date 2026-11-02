VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmtAsuransi 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DAFTAR ASURANSI"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   12555
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   7815
      Left            =   480
      TabIndex        =   15
      Top             =   480
      Width           =   12015
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   13
         Left            =   3480
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin VB.CommandButton cmdPelanggan 
         Caption         =   "Tambah Pelanggan"
         Height          =   285
         Left            =   4290
         TabIndex        =   30
         Top             =   600
         Width           =   1695
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   1
         Left            =   1440
         TabIndex        =   1
         Top             =   915
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
      Begin TrueOleDBGrid70.TDBGrid GridAsuransi 
         Height          =   5175
         Left            =   120
         TabIndex        =   16
         Top             =   2520
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   9128
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
         Index           =   2
         Left            =   1440
         TabIndex        =   2
         Top             =   1215
         Width           =   4545
         _ExtentX        =   8017
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
         Index           =   0
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   1440
         TabIndex        =   3
         Top             =   1515
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Index           =   7
         Left            =   7320
         TabIndex        =   8
         Top             =   1095
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
         Index           =   8
         Left            =   7320
         TabIndex        =   9
         Top             =   1395
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
         Index           =   9
         Left            =   7320
         TabIndex        =   7
         Top             =   795
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
         Index           =   10
         Left            =   7320
         TabIndex        =   10
         Top             =   1695
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
         Index           =   6
         Left            =   7320
         TabIndex        =   6
         Top             =   480
         Width           =   3465
         _ExtentX        =   6112
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
         Index           =   4
         Left            =   1440
         TabIndex        =   4
         Top             =   1810
         Width           =   3825
         _ExtentX        =   6747
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
         Index           =   5
         Left            =   1440
         TabIndex        =   5
         Top             =   2115
         Width           =   7545
         _ExtentX        =   13309
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
         Index           =   11
         Left            =   3480
         TabIndex        =   29
         Top             =   600
         Visible         =   0   'False
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox MaskEdBox9 
         Height          =   270
         Left            =   3360
         TabIndex        =   31
         Top             =   600
         Width           =   60
         _ExtentX        =   106
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   1900
         _ExtentX        =   3360
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Lgn"
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
         TabIndex        =   32
         Top             =   645
         Width           =   660
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         Height          =   195
         Left            =   390
         TabIndex        =   28
         Top             =   2160
         Width           =   825
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   5280
         Top             =   1815
         Width           =   60
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NPWP"
         Height          =   195
         Left            =   6720
         TabIndex        =   27
         Top             =   525
         Width           =   495
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   5925
         Top             =   922
         Width           =   60
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fax No."
         Height          =   195
         Left            =   6660
         TabIndex        =   26
         Top             =   1740
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN No."
         Height          =   195
         Left            =   6480
         TabIndex        =   25
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA No."
         Height          =   195
         Left            =   6450
         TabIndex        =   24
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "GSM No."
         Height          =   195
         Left            =   6555
         TabIndex        =   23
         Top             =   1140
         Width           =   660
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alamat"
         Height          =   195
         Left            =   735
         TabIndex        =   22
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PIC"
         Height          =   195
         Left            =   960
         TabIndex        =   21
         Top             =   1260
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asuransi"
         Height          =   195
         Left            =   615
         TabIndex        =   20
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode"
         Height          =   195
         Left            =   840
         TabIndex        =   19
         Top             =   285
         Width           =   375
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kota"
         Height          =   195
         Left            =   885
         TabIndex        =   18
         Top             =   1860
         Width           =   330
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   8715
      Left            =   0
      Top             =   0
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   15372
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "  Table Asuransi"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextColor       =   16777215
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   10020
      _ExtentX        =   17674
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   7920
      TabIndex        =   11
      Top             =   8400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Tambah"
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
   Begin AIFCmp1.asxPowerButton cmdUbah 
      Height          =   360
      Left            =   8880
      TabIndex        =   12
      Top             =   8400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Ubah"
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
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   9840
      TabIndex        =   13
      Top             =   8400
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Hapus"
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
   Begin AIFCmp1.asxPowerButton cmdKeluar 
      Height          =   360
      Left            =   10800
      TabIndex        =   14
      Top             =   8400
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "Kembali Ke &Menu"
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
End
Attribute VB_Name = "frmtAsuransi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Job ID          :
'Tgl. Mulai      : 12 Februari 2009
'Created Coding  : Thania
'Tgl. Modifikasi : Januari-Maret 2012 - Acep
'=======================================

Dim rsGrid As ADODB.Recordset
Dim rsTemp As ADODB.Recordset
Dim newRecord As Boolean
Dim cols As TrueOleDBGrid70.Columns

Private Sub ber_sih()
    For i = 0 To 13
        Select Case i
            Case 7, 8, 9, 10
            Case Else
                Text(i) = ""
        End Select
    Next
    For i = 7 To 10
        n = Text(i).Mask
        Text(i).Mask = ""
        Text(i) = 0
        Text(i).Mask = n
    Next
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True, Optional flg1 As Boolean = True)
    Text(1).Enabled = Flg
    For i = 2 To 13
        Select Case i
            Case 11, 13
            Case Else
                Text(i).Enabled = flg1
        End Select
    Next
End Sub

Private Sub refreshGrid()
    GridAsuransi.ClearFields
    Set rsGrid = objO.FetchData("SELECT KODE, KDPLG_ASR, ASURANSI, PIC as [CONTACT PERSON], ALAMAT, KOTA, NPWP, GSM, CDMA, PSTN, FAX, KETERANGAN FROM T_ASURANSI WITH(NOLOCK) ORDER BY ASURANSI")
    If Not rsGrid.EOF Then
        GridAsuransi.DataSource = rsGrid
    End If
End Sub

Private Function cariCounter(kd As String) As String
    Set rsTemp = objO.FetchData("SELECT kode FROM T_asuransi WITH(NOLOCK) where kode like '" & Trim(kd) & "%' ORDER BY kode DESC")
    If Not rsTemp.EOF Then
        cariCounter = Format(Right(rsTemp("KODE"), 4) + 1, "000#")
    Else
        cariCounter = Format(1, "000#")
    End If
End Function

Private Sub cmdHapus_Click()
    Set rsTemp = objO.FetchData("SELECT * FROM T_asuransi WITH(NOLOCK) WHERE kode = '" & Text(0) & "'")
    If Not rsTemp.EOF Then
        temp = MsgBox("Akan menghapus data Asuransi : " & Text(1), vbYesNo, "HAPUS DATA")
        If temp = vbYes Then
            Call objO.DeleteData("DELETE T_ASURANSI WHERE kode = '" & Text(0) & "'")
        End If
    End If
    Call ber_sih
    Call refreshGrid
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdPelanggan_Click()
    With frmtPlgn
        .Text(57) = Text(12)
        .Text(1) = Text(1)
        .Text(2) = Text(3)
        .Text(3) = Text(4)
        .Text(4) = Text(9) + "," + Trim(Text(7)) + "," + Trim(Text(8))
        .Text(6) = Trim(Text(6))
        gBantuan2 = "kosong"
        frmtPlgn.Show vbModal
    End With
End Sub

Private Sub cmdSimpan_Click()
    If cmdSimpan.Caption = "&Tambah" Then
        newRecord = True
        Call ber_sih
        Call buka_Tutup(True, True)
        Text(12).SetFocus
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
    ElseIf cmdSimpan.Caption = "&Simpan" Then
        
        ClearsSql
        If Text(0) = "" Then
            Text(0) = Right(year(isnow), 2) & AlfaBulan(Month(isnow)) & AmbilInisial(Text(1), 5) & Text(11) & cariCounter(Right(year(isnow), 2) & AlfaBulan(Month(isnow)) & AmbilInisial(Text(1), 5) & Text(11))
        End If
        gsSql1(0) = "KODE"
        gsSql2(0) = Text(0)
        gsSql1(1) = "ASURANSI"
        gsSql2(1) = Text(1)
        gsSql1(2) = "PIC"
        gsSql2(2) = Text(2)
        gsSql1(3) = "ALAMAT"
        gsSql2(3) = Text(3)
        gsSql1(4) = "KOTA"
        gsSql2(4) = Text(4)
        gsSql1(5) = "NPWP"
        gsSql2(5) = Replace(Text(6), ">", "")
        gsSql1(6) = "GSM"
        gsSql2(6) = Replace(Text(7), ">", "")
        gsSql1(7) = "CDMA"
        gsSql2(7) = Replace(Text(8), ">", "")
        gsSql1(8) = "PSTN"
        gsSql2(8) = Replace(Text(9), ">", "")
        gsSql1(9) = "FAX"
        gsSql2(9) = Replace(Text(10), ">", "")
        gsSql1(10) = "KETERANGAN"
        gsSql2(10) = Text(5)
        gsSql1(11) = "KDPLG_ASR"
        gsSql2(11) = Trim(Text(12))
        
        
        Set rsTemp = objO.FetchData("SELECT * FROM T_ASURANSI WITH(NOLOCK) WHERE kode = '" & Text(0) & "'")
        If Not rsTemp.EOF Then
            If newRecord Then
                MsgBox "Data sudah ada. Jika ingin mengubah pilih ubah"
                Exit Sub
            Else
                sSQL = GensSql2("T_ASURANSI", True) + "WHERE kode = '" & Text(0) & "'"
            End If
        Else
            sSQL = GensSql2("T_ASURANSI", False)
        End If
        Call objO.UpdateData(sSQL)
        
        Call ber_sih
        Call buka_Tutup(False, False)
        Call refreshGrid
        cmdSimpan.Caption = "&Tambah"
        cmdUbah.Caption = "&Ubah"
    End If
End Sub

Private Sub cmdUbah_Click()
    If cmdUbah.Caption = "&Ubah" Then
        Call buka_Tutup(False, True)
        newRecord = False
        cmdSimpan.Caption = "&Simpan"
        cmdUbah.Caption = "&Batal"
    ElseIf cmdUbah.Caption = "&Batal" Then
        Call buka_Tutup(False, False)
        cmdSimpan.Caption = "&Tambah"
        cmdUbah.Caption = "&Ubah"
    End If
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(False, False)
    Call refreshGrid
End Sub

Private Sub GridAsuransi_FilterChange()
    On Error GoTo errHandler
    Set cols = GridAsuransi.Columns
    Dim C As Integer
    
    C = GridAsuransi.col
    GridAsuransi.HoldFields
    rsGrid.filter = getFilter()
    GridAsuransi.col = C
    GridAsuransi.EditActive = True
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

Private Sub GridAsuransi_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    If Not rsGrid.EOF Then
        Text(0) = Trim(rsGrid.Fields(0))
        Text(1) = Trim(rsGrid.Fields(2))
        Text(2) = Trim(rsGrid.Fields(3))
        Text(3) = Trim(rsGrid.Fields(4))
        Text(4) = Trim(rsGrid.Fields(5))
        Text(6) = Trim(rsGrid.Fields(6))
        Text(12) = Trim(rsGrid.Fields(1))
        
        For i = 7 To 10
            Text(i) = Trim(rsGrid.Fields(i))
        Next
        Text(5) = Trim(rsGrid.Fields(11))
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
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "ASURANSI"
                frmCari.Show (vbModal)
            Case 4
                gMainKeyCari = Text(Index)
                gSearchOption = "CITY"
                frmCari.Show (vbModal)
            Case 12
                gMainKeyCari = Text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.Show (vbModal)
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
        Case 1
            Set rsSearch = objO.FetchData("Select * from t_asuransi WITH(NOLOCK) where asuransi = '" & Text(Index) & "'")
            If Not rsSearch.EOF Then
                Text(2) = rsSearch("PIC")
                Text(3) = rsSearch("Alamat")
                Text(4) = rsSearch("Kota")
                Text(5) = rsSearch("Keterangan")
                Text(6) = rsSearch("NPWP")
                Text(7) = Replace(rsSearch("GSM"), "-", "")
                Text(8) = Replace(rsSearch("CDMA"), "-", "")
                Text(9) = Replace(rsSearch("PSTN"), "-", "")
                Text(10) = Replace(rsSearch("FAX"), "-", "")
            End If
        Case 4
            Set rsSearch = objO.FetchData("Select * from T_CITY WITH(NOLOCK) where kota = '" & Text(Index) & "'")
            If rsSearch.EOF Then
                gMainKeyCari = Text(Index)
                gSearchOption = "CITY"
                frmCari.Show (vbModal)
            Else
                Text(11) = rsSearch("Inisial")
            End If
        Case 12
            If Trim(Text(Index)) <> "" Then

                    Set rsSearch = objO.FetchDataACC("select KodeLgn, NamaLgn, KodeWil, Alamat1, Alamat2, Alamat3, Telepon, Npwp from Armst WITH(NOLOCK) where KodeLgn = '" & Trim(Text(Index)) & "'")
                    If rsSearch.BOF Then
                        temp = MsgBox("Tambah Kode Pelanggan baru ?", vbYesNo, "KODE PELANGGAN")
                        If temp = vbYes Then
                            cmdPelanggan_Click
                        Else
                            Cancel = True
                            gMainKeyCari = Text(Index)
                            gSearchOption = "KODEPLG"
                            frmCariAcess.Show (vbModal)
                        End If
                    Else
                        If IsNull(rsSearch("NamaLgn")) Then
                            Text(1) = ""
                        Else
                            Text(1) = rsSearch("NamaLgn")
                        End If
                        If IsNull(rsSearch("Alamat1")) Then
                            Text(3) = ""
                        Else
                            Text(3) = rsSearch("Alamat1")
                        End If
                        If IsNull(rsSearch("Alamat2")) Then
                            Text(4) = ""
                        Else
                            Text(4) = rsSearch("Alamat2")
                        End If
                        If IsNull(rsSearch("Telepon")) Then
                            Text(9) = ""
                            Text(7) = ""
                            Text(8) = ""
                        Else
                            Dim nDpn1 As String
                            Dim nTgh1 As String
                            Dim nBlk1 As String
                            nDpn1 = ""
                            nTgh1 = ""
                            nBlk1 = ""
                            Call pecahNoTelp(Trim(rsSearch("Telepon")), nDpn1, nTgh1, nBlk1, ",")
                            Text(9) = Trim(nDpn1)
                            Text(7) = Trim(nTgh1)
                            Text(8) = Trim(nBlk1)
                        End If
                        If IsNull(rsSearch("Npwp")) Then
                            Text(6) = ""
                        Else
                            Text(6) = rsSearch("Npwp")
                        End If
                        If IsNull(rsSearch("KodeWil")) Then
                            Text(13) = ""
                            Text(11) = ""
                        Else
                            Text(13) = rsSearch("KodeWil")
                            Text(11) = rsSearch("KodeWil")
                        End If
                    End If
                    Set rsSearch = objO.FetchDataACC("select KodeWil,NamaWil from Arwilayah WITH(NOLOCK) where KodeWil = '" & Left(Text(13), 3) & "'")
                    If Not rsSearch.BOF Then
                        Text(6) = rsSearch("NamaWil")
                    End If
            End If
    End Select
End Sub
