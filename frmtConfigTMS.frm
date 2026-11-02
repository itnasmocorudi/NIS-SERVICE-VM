VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmtConfigTMS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "KONFIGURASI UNIT TMS"
   ClientHeight    =   9375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9375
   ScaleWidth      =   11280
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   600
      TabIndex        =   13
      Top             =   0
      Width           =   5895
      Begin VB.OptionButton Option1 
         BackColor       =   &H008080FF&
         Caption         =   "FALSE"
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080FF80&
         Caption         =   "TRUE"
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Available Mobile Service"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3255
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8805
      Left            =   600
      TabIndex        =   0
      Top             =   -360
      Width           =   11295
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8400
         TabIndex        =   37
         Top             =   6360
         Width           =   2175
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Aktif"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   39
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Non Aktif"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   960
            TabIndex        =   38
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8400
         TabIndex        =   33
         Top             =   5880
         Width           =   2175
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Aktif"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Non Aktif"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   960
            TabIndex        =   34
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   8400
         TabIndex        =   29
         Top             =   5400
         Width           =   2175
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Aktif"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Non Aktif"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   960
            TabIndex        =   30
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   1680
         TabIndex        =   25
         Top             =   6600
         Width           =   2175
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Non Aktif"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   27
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Aktif"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   120
            Width           =   855
         End
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   21
         Top             =   5880
         Width           =   2295
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1665
         TabIndex        =   19
         Top             =   5520
         Width           =   4815
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Left            =   6000
         TabIndex        =   17
         Top             =   360
         Width           =   5175
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Available Mobile Service"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   4935
         End
      End
      Begin VB.TextBox Text 
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   5160
         Width           =   2295
      End
      Begin TrueOleDBGrid70.TDBGrid Grid 
         Height          =   3375
         Left            =   0
         TabIndex        =   1
         Top             =   1200
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5953
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Squad ID"
         Columns(0).DataField=   "Squad ID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Deskripsi"
         Columns(1).DataField=   "Deskripsi"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "NoRangka"
         Columns(2).DataField=   "NoRangka"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Status Internal"
         Columns(3).DataField=   "Status Internal"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Booking Online"
         Columns(4).DataField=   "Booking Online"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Servis GR"
         Columns(5).DataField=   "Servis GR"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Servis PM"
         Columns(6).DataField=   "Servis PM"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "NIK"
         Columns(7).DataField=   "NIK"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Nama TA"
         Columns(8).DataField=   "Nama TA"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   9
         Splits(0)._UserFlags=   0
         Splits(0).Locked=   -1  'True
         Splits(0).AllowSizing=   -1  'True
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AllowColSelect=   0   'False
         Splits(0).FetchRowStyle=   -1  'True
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14215660
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
         Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(33)=   "Column(8).Width=3281"
         Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
         Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Verdana"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         WrapCellPointer =   -1  'True
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=Verdana"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=Verdana"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H80000016&,.bold=-1"
         _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=Verdana"
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
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9,.bgcolor=&HC1EDFF&"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=46,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=43,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=44,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=45,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=32,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=29,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=30,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=31,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=28,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=17"
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
         _StyleDefs(86)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(87)  =   "Named:id=40:OddRow"
         _StyleDefs(88)  =   ":id=40,.parent=33"
         _StyleDefs(89)  =   "Named:id=41:RecordSelector"
         _StyleDefs(90)  =   ":id=41,.parent=34"
         _StyleDefs(91)  =   "Named:id=42:FilterBar"
         _StyleDefs(92)  =   ":id=42,.parent=33"
      End
      Begin ucXPButton.XPButton cmdDelete 
         Height          =   495
         Left            =   5520
         TabIndex        =   5
         Top             =   8040
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&d"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtConfigTMS.frx":0000
         PICN            =   "frmtConfigTMS.frx":001C
         PICH            =   "frmtConfigTMS.frx":05A4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSave 
         Height          =   495
         Left            =   4320
         TabIndex        =   6
         Top             =   8040
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&s"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtConfigTMS.frx":0BE0
         PICN            =   "frmtConfigTMS.frx":0BFC
         PICH            =   "frmtConfigTMS.frx":110F
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdAdd 
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Top             =   8040
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&a"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtConfigTMS.frx":16B1
         PICN            =   "frmtConfigTMS.frx":16CD
         PICH            =   "frmtConfigTMS.frx":1B19
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdEdit 
         Height          =   495
         Left            =   3120
         TabIndex        =   8
         Top             =   8040
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&e"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtConfigTMS.frx":1FE1
         PICN            =   "frmtConfigTMS.frx":1FFD
         PICH            =   "frmtConfigTMS.frx":24A8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdExit 
         Height          =   495
         Left            =   9840
         TabIndex        =   9
         Top             =   8040
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&x"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmtConfigTMS.frx":29CD
         PICN            =   "frmtConfigTMS.frx":29E9
         PICH            =   "frmtConfigTMS.frx":2EBD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc AdoAs 
         Height          =   330
         Index           =   1
         Left            =   4440
         Top             =   5160
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   582
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
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo DCPetugas 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Tag             =   "Kode Supplier TAM"
         Top             =   6240
         Width           =   3090
         BevelType       =   0
         _Version        =   196616
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox3D      =   0   'False
         ForeColorEven   =   0
         BackColorEven   =   12648447
         BackColorOdd    =   16777215
         RowHeight       =   423
         Columns(0).Width=   3200
         Columns(0).DataType=   8
         Columns(0).FieldLen=   4096
         _ExtentX        =   5450
         _ExtentY        =   556
         _StockProps     =   93
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service PM :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7125
         TabIndex        =   36
         Top             =   6480
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Service GR :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7140
         TabIndex        =   32
         Top             =   6000
         Width           =   1170
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Booking Online :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6735
         TabIndex        =   28
         Top             =   5520
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NIK :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1140
         TabIndex        =   23
         Top             =   6285
         Width           =   465
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Rangka :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   450
         TabIndex        =   22
         Top             =   5925
         Width           =   1155
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Deskripsi :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   570
         TabIndex        =   20
         Top             =   5565
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status Internal :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   0
         TabIndex        =   11
         Top             =   6720
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Squad ID :"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   615
         TabIndex        =   10
         Top             =   5160
         Width           =   1005
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: LIST PERIODE ::."
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
         Left            =   240
         TabIndex        =   4
         Top             =   60
         Width           =   10800
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000080FF&
         X1              =   -120
         X2              =   11160
         Y1              =   8640
         Y2              =   8640
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".:: DATA UNIT TMS ::."
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
         Left            =   4530
         TabIndex        =   3
         Top             =   4740
         Width           =   2235
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000080FF&
         X1              =   -120
         X2              =   11160
         Y1              =   7920
         Y2              =   7920
      End
      Begin VB.Label lblJml 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Left            =   7680
         TabIndex        =   2
         Top             =   60
         Width           =   1380
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11055
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   4680
         Width           =   11055
      End
   End
End
Attribute VB_Name = "frmtConfigTMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rsM As ADODB.Recordset
Private col As TrueOleDBGrid70.Column
Private cols As TrueOleDBGrid70.Columns
Private Urut As String
Private C As Integer
Private isOnWar As Boolean
Private isTambah As Boolean

Private Sub CmdAdd_Click()
    isTambah = True
    Text(1) = ""
    Text(2) = ""
    Text(0) = objO.FetchData_String("SELECT 'TMS' + CAST(ISNULL(MAX(CAST(REPLACE(KodeTMS,'TMS','') AS INT)),0) + 1 AS VARCHAR(10)) AS [Kode TMS] FROM M_UnitTMS")
   Call BindToControl(False, True, True)
End Sub

Private Sub cmdEdit_Click()
    isTambah = False
   Call BindToControl(False, True, True)
End Sub

Private Sub Form_Load()
    Call SetControlForm
    Call FillGridClass
    Call BindToControl(True, False, False)
End Sub
'
Private Sub SetControlForm()

End Sub
'    With Me
'        .cmdAdd.Enabled = True
'       ' .cmdCreate.Visible = False
'        .cmdEdit.Enabled = True
'        .cmdSave.Visible = True
'        .cmdSave.Enabled = False
'        .cmdCancel.Visible = False
'        .cmdDelete.Visible = True
'        .cmdDelete.Enabled = True
'        .Grid.Enabled = True
'        .txt(0).BackColor = fColor
'        .txt(1).BackColor = fColor
'        .txt(2).BackColor = fColor
'        .txtRp(10).BackColor = fColor
'        .txtRp(11).BackColor = fColor
'        .txtRp(12).BackColor = fColor
'        .txtRp(13).BackColor = fColor
'        .cboWar.BackColor = fColor
'        .cboWar.Clear
'        .cboWar.AddItem "NO"
'        .cboWar.AddItem "YES"
'        .cboWar.Text = "NO"
'        .cmdExit.Enabled = True
'    End With
''End Sub

Private Sub FillGridClass()
    If objO.FetchData_String("select enable_mobile_service from ReplBengkel99..p2_squad_config where outlet_ID='" & Log.kdDealerDidx & "'") = "True" Then
        Option1(0).value = True
        Option1(1).value = False
        Label4.Caption = "AVAILABLE FOR ONLINE BOOKING"
    Else
        Option1(1).value = True
        Option1(0).value = False
        Label4.Caption = "NOT AVAILABLE FOR ONLINE BOOKING"
    End If
    
    
    AdoAs(1).ConnectionString = SQLCONBengkel
    AdoAs(1).RecordSource = "SELECT FS_NM_PETUGAS + ' | ' + nik as petugas,FS_KD_PETUGAS,nik FROM T_PETUGAS WITH (NOLOCK) WHERE TIPE = 'TA'"
    DCPetugas.DataSourceList = AdoAs(1)
    DCPetugas.DataFieldList = "petugas"
    DCPetugas.Columns(0).Width = 3000
    DCPetugas.Columns(1).Width = 1500
    DCPetugas.Columns(0).Caption = "Petugas"
    DCPetugas.Columns(1).Caption = "Kode Petugas"
    DCPetugas.Columns(2).Visible = False
    DCPetugas.value = ""
    
    sSQL = "select a.outlet_squad_ID as [Squad ID],a.squad_name as Deskripsi,b.NoRangka,"
    sSQL = sSQL & " case when b.[Status]=1 then 'True' else 'False' end as [Status Internal],"
    sSQL = sSQL & " case when a.squad_status=1 then 'True' else 'False' end as [Booking Online],"
    sSQL = sSQL & " case when a.avail_gr=1 then 'True' else 'False' end as [Servis GR],"
    sSQL = sSQL & " case when a.avail_pm=1 then 'True' else 'False' end as [Servis PM],"
    sSQL = sSQL & " b.NIK,c.FS_NM_PETUGAS as [Nama TA]"
    sSQL = sSQL & " from ReplBengkel99..p2_squad_config_data a"
    sSQL = sSQL & " left join M_UnitTMS b on b.KodeTMS=a.outlet_squad_ID"
    sSQL = sSQL & " left join T_PETUGAS c on c.NIK=b.nik"
    sSQL = sSQL & " where a.outlet_ID='" & Log.kdDealerDidx & "' and c.TIPE='TA'"
    Set rsM = objO.FetchData(sSQL)
    If rsM.RecordCount = 0 Then
        Exit Sub
    End If
    Set Grid.DataSource = rsM
    lblJml.Caption = rsM.RecordCount
 
End Sub

Private Sub Grid_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
     
    Text(0).Text = Grid.Columns("Squad ID").value
    Text(1).Text = Grid.Columns("Deskripsi").value
    Text(2).Text = Grid.Columns("NoRangka").value
    DCPetugas.value = Grid.Columns("NIK").value
    If Grid.Columns("Status Internal").value = "True" Then
        Option2(0).value = True
        Option2(1).value = False
    Else
        Option2(1).value = True
        Option2(2).value = False
    End If
    If Grid.Columns("Booking Online").value = "True" Then
        Option2(3).value = True
        Option2(2).value = False
    Else
        Option2(2).value = True
        Option2(3).value = False
    End If
    If Grid.Columns("Servis GR").value = "True" Then
        Option2(5).value = True
        Option2(4).value = False
    Else
        Option2(4).value = True
        Option2(5).value = False
    End If
    If Grid.Columns("Servis PM").value = "True" Then
        Option2(7).value = True
        Option2(6).value = False
    Else
        Option2(6).value = True
        Option2(7).value = False
    End If
End Sub
Private Sub BindToControl(Edit As Boolean, save As Boolean, batal As Boolean)
On Error GoTo ErrNull
    cmdAdd.Enabled = Edit
    cmdEdit.Enabled = Edit
    cmdSave.Enabled = save
    cmdDelete.Enabled = batal
    
    If Edit = True Then
      Text(1).Enabled = False
      Text(2).Enabled = False
    Else
      Text(1).Enabled = True
      Text(2).Enabled = True
    End If
ErrNull:
End Sub

'===============================================================================================
'========= A D D ======== A D D ======== A D D =================================================
'===============================================================================================
'Private Sub CmdAdd_Click()
'    With Me
'        .cmdAdd.Enabled = False
'       ' .cmdCreate.Visible = True
'       ' .cmdCreate.Enabled = True
'        .cmdEdit.Enabled = False
'        .cmdSave.Visible = False
'        .cmdCancel.Visible = True
'        .cmdCancel.Enabled = True
'        .cmdDelete.Visible = False
'        .Grid.Enabled = False
'        .txt(0).BackColor = tColor
'        .txt(1).BackColor = tColor
'        .txt(2).BackColor = tColor
'        .txtRp(10).BackColor = tColor
'        .txtRp(11).BackColor = tColor
'        .txtRp(12).BackColor = tColor
'        .txtRp(13).BackColor = tColor
'        .cboWar.BackColor = tColor
'        .txt(0).Text = ""
'        .txt(1).Text = ""
'        .txt(2).Text = ""
'        .txtRp(10) = 0
'        .txtRp(11) = 0
'        .txtRp(12) = 0
'        .txtRp(13) = 0
'        .cboWar.Text = "NO"
'        .cmdExit.Enabled = False
'        .txt(0).SetFocus
'
'        If isOnWar = True Then
'            .cboWar.BackColor = fColor
'        Else
'            .cboWar.BackColor = tColor
'        End If
'    End With
'End Sub
'
'Private Sub cmdCreate_Click()
'    If FillNotNull = True Then
'        rsM.Find "Kode = '" & Trim(txt(0).Text) & "'", , adSearchForward, 1
'        If Not rsM.EOF Then
'            MsgBox "Data Kelas " & txt(0).Text & " sudah terdaftar, Ganti Kode !", vbInformation + vbOKOnly, "Car Class"
'            Exit Sub
'        End If
'        If MsgBox("Simpan Data Kelas Baru ?", vbQuestion + vbYesNo, "Car Class") = vbYes Then
'            Call objO.InsertDataHO("INSERT INTO CarClass(ClassCode, ClassName, Price,priceTAM, Note, DiscNRM, DiscNON, war, FS_LOG_DEL, FD_TGL_DEL) " & _
'                                   "VALUES('" & Trim(txt(0).Text) & "','" & Trim(txt(1).Text) & "'," & RepRp(txtRp(10)) & "," & RepRp(txtRp(11)) & ",'" & Trim(txt(2).Text) & "'," & RepRp(txtRp(12)) & "," & RepRp(txtRp(13)) & "," & IIf(cboWar.Text = "YES", 1, 0) & ",'','01/01/3000')")
'            Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('KELAS', '" & Trim(txt(0).Text) & "','S')")
'            Call SetControlForm
'            Call FillGridClass
'        End If
'    Else
'        Call NotValid_Input
'    End If
'End Sub

'===============================================================================================
'======= E D I T =========== E D I T =========== E D I T =======================================
'===============================================================================================
'Private Sub cmdEdit_Click()
'    With Me
'        .cmdAdd.Enabled = False
'    '    .cmdCreate.Visible = False
'        .cmdEdit.Enabled = False
'        .cmdSave.Visible = True
'        .cmdSave.Enabled = True
'        .cmdCancel.Visible = True
'        .cmdCancel.Enabled = True
'        .cmdDelete.Visible = False
'        .Grid.Enabled = False
'        .txt(0).BackColor = fColor
'        .txt(1).BackColor = tColor
'        .txt(2).BackColor = tColor
'        .txtRp(10).BackColor = tColor
'        .txtRp(11).BackColor = tColor
'        .txtRp(12).BackColor = tColor
'        .txtRp(13).BackColor = tColor
'        .cboWar.BackColor = tColor
'        .cmdExit.Enabled = False
'        .txt(1).SetFocus
'
'        If isOnWar = True Then
'            If .cboWar.Text = "YES" Then
'                .cboWar.BackColor = tColor
'            Else
'                .cboWar.BackColor = fColor
'            End If
'        Else
'            .cboWar.BackColor = tColor
'        End If
'    End With
'End Sub

Private Sub cmdSave_Click()
1         If isTambah Then
2             sSQL = "INSERT INTO M_UnitTMS (KodeTMS,NamaTMS,NoRangka,Status,NIK,FRONT,UserEntry,TglEntry)"
3             sSQL = sSQL & " VALUES ('" & NulSID(Text(0)) & "','" & NulSID(Text(1)) & "','" & NulSID(Text(2)) & "','" & IIf(Option2(0).value = True, 1, 0) & "','" & DCPetugas.Columns(2).value & "','" & DCPetugas.Columns(1).value & "','" & gSA & "',GETDATE())"
4             Call objO.InsertData(sSQL)
              
5             If objO.FetchData_Count("select * from ReplBengkel99..p2_squad_config where outlet_id = '" & Log.kdDealerDidx & "'") = 0 Then
6             sSQL = "INSERT INTO ReplBengkel99..p2_squad_config (outlet_ID,outlet_name,enable_mobile_service,event_ID,timestamp,tgl_entry)"
7             sSQL = sSQL & " VALUES ('" & Log.kdDealerDidx & "','" & gNmCbg & "','" & IIf(Option1(0).value = True, 1, 0) & "',NEWID(),GETDATE(),GETDATE())"
              
8             Else
9             sSQL = "update ReplBengkel99..p2_squad_config set enable_mobile_service='" & IIf(Option1(0).value = True, 1, 0) & "' where outlet_ID='" & Log.kdDealerDidx & "'"
              
10            End If
11            Call objO.InsertData(sSQL)
              
12            sSQL = "INSERT INTO ReplBengkel99.dbo.p2_squad_config_data (outlet_ID,outlet_squad_ID,squad_name,squad_description,squad_status,avail_gr,avail_pm)"
13            sSQL = sSQL & " VALUES ('" & Log.kdDealerDidx & "','" & Text(0) & "','" & Text(1) & "','" & Text(1) & "','" & IIf(Option2(3).value = True, 1, 0) & "','" & IIf(Option2(5).value = True, 1, 0) & "','" & IIf(Option2(7).value = True, 1, 0) & "')"
14            Call objO.InsertData(sSQL)
              
15            sSQL = "INSERT INTO ReplBengkel99..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,status,proses,event_no,tglentry,asal_cabang)"
16            sSQL = sSQL & " VALUES ('CONFIG','','" & Log.kdDealerDidx & "','','" & Right(gDb, 2) & "','SQUADCONFIG',0,'',GETDATE(),'" & Right(gDb, 2) & "')"
17            Call objO.InsertData(sSQL)
              
18        Else
19            sSQL = "update M_UnitTMS set NamaTMS='" & NulSID(Text(1)) & "', NoRangka='" & NulSID(Text(2)) & "', [Status]='" & IIf(Option2(0).value = True, 1, 0) & "',NIK='" & DCPetugas.Columns(2).value & "' where KodeTMS='" & Text(0) & "' "
20            Call objO.InsertData(sSQL)
              
21            sSQL = "update ReplBengkel99..p2_squad_config set enable_mobile_service='" & IIf(Option1(0).value = True, 1, 0) & "' where outlet_ID='" & Log.kdDealerDidx & "'"
22            Call objO.InsertData(sSQL)
              
23            sSQL = "update ReplBengkel99..p2_squad_config_data set squad_name='" & NulSID(Text(1)) & "',squad_description='" & Text(1) & "',squad_status='" & IIf(Option2(3).value = True, 1, 0) & "',avail_pm='" & IIf(Option2(7).value = True, 1, 0) & "',avail_gr='" & IIf(Option2(5).value = True, 1, 0) & "' where outlet_ID='" & Log.kdDealerDidx & "' and outlet_squad_ID='" & Text(0) & "'"
24            Call objO.InsertData(sSQL)
              
25            sSQL = "INSERT INTO ReplBengkel99..kirimDataDidx (tipe,nobookingapps,nobooking,noSPK,kodedealer,status,proses,event_no,tglentry,asal_cabang)"
26            sSQL = sSQL & " VALUES ('CONFIG','','" & Log.kdDealerDidx & "','','" & Right(gDb, 2) & "','SQUADCONFIG',0,'',GETDATE(),'" & Right(gDb, 2) & "')"
27            Call objO.InsertData(sSQL)

28        End If
         ''Call objO.InsertData("UPDATE periode SET tglAwal = '" & Format(tglAwal.value, "mm/dd/yyyy") & "',tglAkhir = '" & Format(tglAkhir.value, "mm/dd/yyyy") & "' WHERE id = '" & Trim(Text1.Text) & "' ")
29       MsgBox "Data berhasil disimpan", vbInformation
30       Call SetControlForm
31        Call FillGridClass
32        Call BindToControl(True, False, False)
End Sub

'===============================================================================================
'===============================================================================================
'Private Function FillNotNull() As Boolean
'    FillNotNull = True
'    If txt(0).Text = "" Or txt(1).Text = "" Then
'        FillNotNull = False
'    End If
'End Function
Private Sub NotValid_Input()
    MsgBox "Data Yang Anda Inputkan Belum Lengkap !", vbInformation + vbOKOnly, "Car Class"
End Sub
'===============================================================================================
'===============================================================================================

'===============================================================================================
'========= C A N C E L ========= C A N C E L ========= C A N C E L =============================
'===============================================================================================
Private Sub cmdCancel_Click()
    Call SetControlForm
    Call FillGridClass
End Sub

'===============================================================================================
'========= D E L E T E ========= D E L E T E ========= D E L E T E =============================
'===============================================================================================
Private Sub CmdDelete_Click()
Call SetControlForm
    Call FillGridClass
    Call BindToControl(True, False, False)
End Sub

Private Sub grid_FilterChange()
On Error GoTo err
   Set cols = Grid.Columns
   
   C = Grid.col
   Grid.HoldFields
   rsM.filter = getFilter()
   If rsM.EOF = True Then
      For Each col In Grid.Columns
         col.FilterText = ""
      Next col
      rsM.filter = ""
   End If
   Grid.col = C
   Grid.EditActive = True
   lblJml.Caption = rsM.RecordCount
   Exit Sub
err:
   Resume Next
End Sub

Private Function getFilter() As String
   On Error GoTo err

On Error GoTo err
Dim tmp As String
Dim n As Integer

   For Each col In cols
      If Trim(col.FilterText) <> "" Then
         n = n + 1
         If n > 1 Then
            tmp = tmp & " AND "
         End If
         tmp = tmp & col.DataField & " LIKE '*" & col.FilterText & "*'"
      End If
   Next col
   getFilter = tmp
   
   Exit Function
err:
   Resume Next
End Function

Private Sub grid_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 6 Then rsM.Sort = "[" & Grid.Columns(ColIndex).DataField & "]" & " " & Urut
4        Grid.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6        lblJml.Caption = rsM.RecordCount
7     Exit Sub
err:
8        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter ")
End Sub
 

Private Sub cmdExit_Click()
    Unload Me
End Sub
 
Private Sub Option2_Click(Index As Integer)
    If Index = 1 Then
        If Option2(1).value = True Then
            Option2(2).value = True
            Option2(4).value = True
            Option2(6).value = True
        End If
    End If
End Sub
