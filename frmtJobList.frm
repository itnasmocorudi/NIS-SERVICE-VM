VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0D623638-DBA2-11D1-B5DF-0060976089D0}#7.0#0"; "tdbg7.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Begin VB.Form frmtJobList 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Tabel Flate Rate"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   7545
   ScaleWidth      =   11910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   6660
      Left            =   450
      TabIndex        =   9
      Top             =   345
      Width           =   11340
      Begin TrueDBGrid70.TDBGrid TData1 
         Height          =   4755
         Left            =   90
         TabIndex        =   5
         Top             =   1815
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   8387
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
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).SizeMode=   2
         Splits(0).AllowRowSizing=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0).ScrollBars=   0
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=6985"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=6906"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=7170"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=7091"
         Splits(0)._ColumnProps(8)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=979"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=900"
         Splits(0)._ColumnProps(13)=   "Column(2).AllowSizing=0"
         Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
         Splits(1)._UserFlags=   0
         Splits(1).AllowRowSizing=   0   'False
         Splits(1).RecordSelectors=   0   'False
         Splits(1).RecordSelectorWidth=   503
         Splits(1).DividerColor=   12632256
         Splits(1).SpringMode=   0   'False
         Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(1)._ColumnProps(0)=   "Columns.Count=4"
         Splits(1)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(1)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(1)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(1)._ColumnProps(6)=   "Column(1).Width=2170"
         Splits(1)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(1)._ColumnProps(8)=   "Column(1)._WidthInPix=2090"
         Splits(1)._ColumnProps(9)=   "Column(1).AllowSizing=0"
         Splits(1)._ColumnProps(10)=   "Column(1).AllowFocus=0"
         Splits(1)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(1)._ColumnProps(12)=   "Column(2).Width=1191"
         Splits(1)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(1)._ColumnProps(14)=   "Column(2)._WidthInPix=1111"
         Splits(1)._ColumnProps(15)=   "Column(2).AllowSizing=0"
         Splits(1)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(1)._ColumnProps(17)=   "Column(3).Width=8229"
         Splits(1)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(1)._ColumnProps(19)=   "Column(3)._WidthInPix=8149"
         Splits(1)._ColumnProps(20)=   "Column(3).AllowSizing=0"
         Splits(1)._ColumnProps(21)=   "Column(3).Order=4"
         Splits.Count    =   2
         PrintInfos(0)._StateFlags=   0
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         MultiSelect     =   2
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=29,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=48,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=30,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=31,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=32,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=44,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=43,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=45,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=46,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=47,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=49,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=50,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=82,.parent=29"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=30"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=31"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=43"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=54,.parent=29"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=30"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=31"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=43"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=58,.parent=29"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=30"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=31"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=43"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=66,.parent=29"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=63,.parent=30"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=64,.parent=31"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=65,.parent=43"
         _StyleDefs(46)  =   "Splits(1).Style:id=13,.parent=1"
         _StyleDefs(47)  =   "Splits(1).CaptionStyle:id=22,.parent=4"
         _StyleDefs(48)  =   "Splits(1).HeadingStyle:id=14,.parent=2"
         _StyleDefs(49)  =   "Splits(1).FooterStyle:id=15,.parent=3"
         _StyleDefs(50)  =   "Splits(1).InactiveStyle:id=16,.parent=5"
         _StyleDefs(51)  =   "Splits(1).SelectedStyle:id=18,.parent=6"
         _StyleDefs(52)  =   "Splits(1).EditorStyle:id=17,.parent=7"
         _StyleDefs(53)  =   "Splits(1).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(54)  =   "Splits(1).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(55)  =   "Splits(1).OddRowStyle:id=21,.parent=10"
         _StyleDefs(56)  =   "Splits(1).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(57)  =   "Splits(1).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(58)  =   "Splits(1).Columns(0).Style:id=86,.parent=13"
         _StyleDefs(59)  =   "Splits(1).Columns(0).HeadingStyle:id=83,.parent=14"
         _StyleDefs(60)  =   "Splits(1).Columns(0).FooterStyle:id=84,.parent=15"
         _StyleDefs(61)  =   "Splits(1).Columns(0).EditorStyle:id=85,.parent=17"
         _StyleDefs(62)  =   "Splits(1).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(63)  =   "Splits(1).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(64)  =   "Splits(1).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(65)  =   "Splits(1).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(66)  =   "Splits(1).Columns(2).Style:id=62,.parent=13"
         _StyleDefs(67)  =   "Splits(1).Columns(2).HeadingStyle:id=59,.parent=14"
         _StyleDefs(68)  =   "Splits(1).Columns(2).FooterStyle:id=60,.parent=15"
         _StyleDefs(69)  =   "Splits(1).Columns(2).EditorStyle:id=61,.parent=17"
         _StyleDefs(70)  =   "Splits(1).Columns(3).Style:id=70,.parent=13"
         _StyleDefs(71)  =   "Splits(1).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(72)  =   "Splits(1).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(73)  =   "Splits(1).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(74)  =   "Named:id=33:Normal"
         _StyleDefs(75)  =   ":id=33,.parent=0"
         _StyleDefs(76)  =   "Named:id=34:Heading"
         _StyleDefs(77)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(78)  =   ":id=34,.wraptext=-1"
         _StyleDefs(79)  =   "Named:id=35:Footing"
         _StyleDefs(80)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(81)  =   "Named:id=36:Selected"
         _StyleDefs(82)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(83)  =   "Named:id=37:Caption"
         _StyleDefs(84)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(85)  =   "Named:id=38:HighlightRow"
         _StyleDefs(86)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(87)  =   "Named:id=39:EvenRow"
         _StyleDefs(88)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(89)  =   "Named:id=40:OddRow"
         _StyleDefs(90)  =   ":id=40,.parent=33"
         _StyleDefs(91)  =   "Named:id=41:RecordSelector"
         _StyleDefs(92)  =   ":id=41,.parent=34"
         _StyleDefs(93)  =   "Named:id=42:FilterBar"
         _StyleDefs(94)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   2055
         TabIndex        =   1
         Top             =   540
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   476
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   1
         Left            =   2055
         TabIndex        =   0
         Top             =   240
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   476
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   2055
         TabIndex        =   2
         Top             =   840
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   3
         Left            =   2055
         TabIndex        =   4
         Top             =   1440
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   476
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
         Mask            =   ">&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   4
         Left            =   2055
         TabIndex        =   3
         Top             =   1140
         Width           =   540
         _ExtentX        =   953
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&"
         PromptChar      =   ">"
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Jenis Pekerjaan"
         Height          =   195
         Left            =   870
         TabIndex        =   16
         Top             =   1215
         Width           =   1125
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   2625
         Top             =   1140
         Width           =   60
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4830
         Top             =   540
         Width           =   60
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   4305
         Top             =   240
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Keterangan"
         Height          =   195
         Left            =   1170
         TabIndex        =   13
         Top             =   1485
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Rate"
         Height          =   195
         Left            =   1650
         TabIndex        =   12
         Top             =   885
         Width           =   345
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Group Model"
         Height          =   195
         Left            =   1080
         TabIndex        =   11
         Top             =   285
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Item"
         Height          =   195
         Left            =   1695
         TabIndex        =   10
         Top             =   585
         Width           =   300
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7755
      Left            =   0
      Top             =   -225
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13679
      StartColor      =   4210752
      EndColor        =   12632064
      FormatString    =   "Tabel Flate Rate"
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
      Top             =   30
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   12632256
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   6975
      TabIndex        =   7
      Top             =   7095
      Width           =   1215
      _ExtentX        =   2143
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
   Begin AIFCmp1.asxPowerButton cmdExit 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   9405
      TabIndex        =   8
      Top             =   7095
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Kembali ke Menu"
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
   Begin AIFCmp1.asxPowerButton cmdSimpan 
      Height          =   360
      Left            =   5760
      TabIndex        =   6
      Top             =   7095
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Simpan"
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
   Begin AIFCmp1.asxPowerButton cmdTampil 
      Height          =   360
      Left            =   4545
      TabIndex        =   14
      Top             =   7095
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "&Tampil"
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
   Begin AIFCmp1.asxPowerButton cmdHapusGroup 
      Height          =   360
      Left            =   8190
      TabIndex        =   15
      Top             =   7095
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12632256
      BorderStyle     =   3
      Caption         =   "Hapus &Group"
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
Attribute VB_Name = "frmtJobList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Kerja As New XArrayDB
Dim Kolom As Integer
Dim Baris As Integer
Dim KolomAdd As Integer
Dim BarisAdd As Integer

Private Function GensSql(xTabel As String, xUpdate As Boolean) As String
    Dim aa As String
    Dim xHasil As String
    xHasil = ""
    If xUpdate Then
        xHasil = "UPDATE " + xTabel + " SET "
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "FN_RATE"
                xHasil = xHasil + gsSql1(i) + " = " + gsSql2(i)
            Case Else
                xHasil = xHasil + gsSql1(i) + " = '" + gsSql2(i) + "' "
            End Select
            i = i + 1
        Loop
    Else
        xHasil = "INSERT INTO " + xTabel + " ("
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            xHasil = xHasil + gsSql1(i)
            i = i + 1
        Loop
        xHasil = xHasil + ") VALUES ("
        
        i = 0
        Do While gsSql1(i) <> ""
            If i <> 0 Then
                xHasil = xHasil + ","
            End If
            
            aa = UCase(gsSql1(i))
            Select Case aa
            Case "FN_RATE"
                xHasil = xHasil + gsSql2(i)
            Case Else
                xHasil = xHasil + "'" + gsSql2(i) + "'"
            End Select
            i = i + 1
        Loop
        xHasil = xHasil + ")"
    End If
    GensSql = xHasil
End Function

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    If MsgBox("Ingin menghapus data ini ?", vbYesNo, "Konfirmasi") = vbYes Then
        sSQL = "delete T_JOBLIST where FS_ITEM='" & Text(0) & "' and FS_GRP_MDL='" & Text(1) & "'"
        oConn.Execute sSQL
        Call BeriItem
        Call TampilGrid
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub cmdHapusGroup_Click()
    If MsgBox("Ingin menghapus Group Model ini ?", vbYesNo, "Konfirmasi") = vbYes Then
        sSQL = "delete T_JOBLIST where FS_GRP_MDL='" & Trim(Text(1)) & "'"
        oConn.Execute sSQL
        Text(0) = ""
        Text(1) = ""
        Call BeriItem
        Call TampilGrid
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub cmdSimpan_Click()
    If Text(1) <> "" And Text(0) <> "" And Text(2) <> "" Then
        sSQL = "Select * from T_JOBLIST where FS_ITEM='" & Text(0) & "' and FS_GRP_MDL='" & Text(1) & "'"
        oRs.Open sSQL, oConn, adOpenStatic
        TData1.Columns(1).Caption = Text(1)
        Call ClearsSql
        gsSql1(0) = "FS_ITEM"
        gsSql2(0) = Text(0)
        gsSql1(1) = "FS_GRP_MDL"
        gsSql2(1) = Text(1)
        gsSql1(2) = "FN_RATE"
        gsSql2(2) = SimpanKoma2Titik(Round(Text(2), 2))
        gsSql1(3) = "FS_KETERANGAN"
        gsSql2(3) = Text(3)
        gsSql1(4) = "FS_JNS_KERJA"
        gsSql2(4) = Text(4)
        If Not oRs.BOF Then
            sSQL = GensSql("t_joblist", True) & " where FS_ITEM='" & Text(0) & "' and FS_GRP_MDL='" & Text(1) & "'"
        Else
            sSQL = GensSql("t_joblist", False)
        End If
        oConn.Execute sSQL
        oRs.Close
        Call hapusForm
        Text(0) = ""
'        Text(1) = ""
'        TData1.Columns(1).Caption = Text(1)
'        For i = 0 To Baris
'            If Trim(Kerja(i, 0)) = Trim(Text(0)) Then
'                Kerja(i, 1) = Koma2Titik(Text(2))
'                Kerja(i, 2) = Trim(Text(3))
'                TData1.ReBind
'                Exit For
'            End If
'        Next
        Call BeriItem
        Call TampilGrid
        Text(0).SetFocus
    Else
        MsgBox "Masukkan Item dan MDL !"
    End If
End Sub

Private Sub cmdTampil_Click()
    Call BeriItem
    Call TampilGrid
    Text(0).SetFocus
End Sub

Private Sub Form_Activate()
    Call hapusForm
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    KolomAdd = 0
    BarisAdd = 0
    Kolom = 0
    Baris = 0
    TData1.Array = Kerja
    

'    sSQL = "select count(distinct FS_MDL) as JmlKol from T_JOBLIST"
'    oRs.Open sSQL, oConn, adOpenStatic
'    If Not oRs.BOF Then
'        Kolom = oRs("JmlKol") + 1
'    End If
'    oRs.Close
    'Isi Item
    Kolom = 3
    
    Call BeriItem
'    For i = 2 To Kolom
'        Call TambahKolom
'    Next
    
    TData1.Columns(0).Caption = "ITEM"
    TData1.Columns(2).Caption = "JP"
    TData1.Columns(3).Caption = "KETERANGAN"
    TData1.Columns(1).Alignment = dbgRight
    TData1.Columns(1).HeadAlignment = dbgLeft
    
End Sub

Private Sub BeriItem()
    sSQL = "select count(distinct FS_ITEM) as JmlBrs from T_JOBLIST"
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        Baris = oRs("JmlBrs")
    End If
    oRs.Close
    Kerja.ReDim 0, Baris, 0, Kolom
    i = 0
    sSQL = "select distinct FS_ITEM from T_JOBLIST order by FS_ITEM"
    oRs.Open sSQL, oConn, adOpenStatic
    If Not oRs.BOF Then
        Do While Not oRs.EOF
            Kerja(i, 0) = Trim(oRs("FS_ITEM"))
            oRs.MoveNext
            i = i + 1
        Loop
    End If
    oRs.Close
End Sub

Private Sub TampilGrid()
    For i = 0 To Baris
        Kerja(i, 1) = 0
        Kerja(i, 2) = ""
    Next i
    i = 0
    sSQL = "select FS_ITEM, FS_KETERANGAN, FN_RATE, FS_JNS_KERJA from T_JOBLIST where FS_GRP_MDL='" & Text(1) & "' order by FS_ITEM"
    oRs1.Open sSQL, oConn, adOpenStatic
    If Not oRs1.BOF Then
        TData1.Columns(1).Caption = Text(1)
        Do While Not oRs1.EOF
'            sSQL = "select FN_RATE,FS_KETERANGAN from T_JOBLIST where " _
'                 & " FS_ITEM ='" & Trim(oRs1("FS_ITEM")) & "' " _
'                 & " and FS_MDL='" & Text(1) & "' "
'            oRs2.Open sSQL, oConn, adOpenStatic
'            If oRs2.BOF Then
            For i = 0 To Baris
                If Trim(Kerja(i, 0)) = Trim(oRs1("FS_ITEM")) Then
                    Kerja(i, 1) = Trim(oRs1("FN_RATE"))
                    Kerja(i, 2) = Trim(oRs1("FS_JNS_KERJA"))
                    Kerja(i, 3) = Trim(oRs1("FS_KETERANGAN"))
                    Exit For
                End If
            Next
'            oRs2.Close
            oRs1.MoveNext
            i = i + 1
        Loop
    Else
        TData1.Columns(1).Caption = Text(1)
    End If
    oRs1.Close
    TData1.ReBind
End Sub

Private Sub KurangKolom()
    If TData1.Columns.Count <> 1 Then
        TData1.Columns.Remove 1
        KolomAdd = KolomAdd - 1
        Kerja.ReDim 0, Baris, 0, Kolom + KolomAdd
        TData1.ReBind
    End If
End Sub

Private Sub TambahKolom()
    Dim C As TrueDBGrid70.Column
    Set C = TData1.Columns.Add(TData1.Columns.Count)
    With C
        .Visible = True         ' Make it visible
        .Caption = Text(1)      ' Set the column's caption
        .Width = 500
    End With
    KolomAdd = KolomAdd + 1
    Kerja.ReDim 0, Baris, 0, Kolom + KolomAdd
    TData1.ReBind
End Sub

Private Sub TData1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'    If TData1.Col = 1 Then
        Text(0) = Kerja(TData1.FirstRow + TData1.row, 0)
        Text(1) = Trim(TData1.Columns(1).Caption)
        Text(2) = Kerja(TData1.FirstRow + TData1.row, 1)
        Text(3) = Kerja(TData1.FirstRow + TData1.row, 3)
        Text(4) = Kerja(TData1.FirstRow + TData1.row, 2)
        
'        Text(1) = Trim(TData1.Columns(TData1.Col).Caption)
'        Text(2) = Kerja(TData1.FirstRow + TData1.Row, TData1.Col)

'    Else
'        Text(0) = ""
'        Text(1) = ""
'        Text(2) = ""
'    End If
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
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = Text(Index)
                gSearchOption = "JOBLIST2"
                frmCari.Show (vbModal)
            Case 1
                gMainKeyCari = Text(Index)
                gSearchOption = "GROUP_MODEL"
                frmCari.Show (vbModal)
            Case 4
                gMainKeyCari = Text(Index)
                gSearchOption = "JENIS_KERJA"
                frmCari.Show (vbModal)
        End Select
    End If
End Sub

Private Sub Text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
    Select Case Index
        Case 2
            KeyAscii = CekNumerik(KeyAscii)
    End Select
End Sub


Private Sub Text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 0
            sSQL = "Select * from T_JOBLIST where FS_ITEM='" & Trim(Text(0)) & "' AND FS_GRP_MDL='" & Trim(Text(1)) & "'"
            oRs.Open sSQL, oConn, adOpenStatic
            If oRs.BOF Then
                Call hapusForm
            Else
                Text(2) = oRs("FN_RATE")
                Text(3) = oRs("FS_KETERANGAN")
                Text(4) = oRs("FS_JNS_KERJA")
            End If
            oRs.Close
        Case 1
            sSQL = "select MDL, [DESC], MODEL from model where MDL='" & Trim(Text(1)) & "' order by MDL"
            oRs.Open sSQL, oConn, adOpenStatic
            If Not oRs.BOF Then
                sSQL = "Select * from T_JOBLIST where FS_ITEM='" & Trim(Text(0)) & "' AND FS_GRP_MDL='" & Trim(Text(1)) & "'"
                oRs1.Open sSQL, oConn, adOpenStatic
                If oRs1.BOF Then
                    Call hapusForm
                Else
                    Text(2) = oRs1("FN_RATE")
                    Text(3) = oRs1("FS_KETERANGAN")
                    Text(4) = oRs1("FS_JNS_KERJA")
                End If
                oRs1.Close
            End If
            oRs.Close
        Case 4
            If Trim(Text(4)) <> "" Then
                sSQL = "Select * from JOB where JOB='" & Text(4) & "'"
                oRs.Open sSQL, oConn, adOpenStatic
                If oRs.BOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.Show (vbModal)
                End If
                oRs.Close
            End If
    End Select
End Sub

Private Sub hapusForm()
    Dim i As Integer
    
    For i = 2 To 4
        Text(i) = ""
    Next
End Sub
