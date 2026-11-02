VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmtModel 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tabel Model Kendaraan"
   ClientHeight    =   7980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11820
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7980
   ScaleWidth      =   11820
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H0068ECE6&
      Height          =   6900
      Left            =   450
      TabIndex        =   5
      Top             =   345
      Width           =   11340
      Begin VB.Frame framSync 
         BackColor       =   &H00FFFFFF&
         Caption         =   "  Syncronize Model  "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   11055
         Begin VB.CommandButton cmdBatal 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Batal"
            Height          =   375
            Left            =   9360
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   6000
            Width           =   1575
         End
         Begin VB.CommandButton Proses 
            BackColor       =   &H00C0FFFF&
            Caption         =   "&Proses"
            Height          =   375
            Left            =   7680
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   6000
            Width           =   1575
         End
         Begin TrueOleDBGrid70.TDBDropDown TDBDropDown1 
            Bindings        =   "frmtModel.frx":0000
            Height          =   1695
            Left            =   600
            TabIndex        =   13
            Top             =   2280
            Width           =   8895
            _ExtentX        =   15690
            _ExtentY        =   2990
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Kode Kelas"
            Columns(0).DataField=   "classcode"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nama Kelas"
            Columns(1).DataField=   "CLASSNAME"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Note"
            Columns(2).DataField=   "NOTE"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Harga NRM"
            Columns(3).DataField=   "PRICE"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Harga TAM"
            Columns(4).DataField=   "PRICETAM"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   5
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).MarqueeStyle=   3
            Splits(0).AllowRowSizing=   0   'False
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=5"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2117"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2037"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=3519"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3440"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=4419"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=4339"
            Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2117"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2037"
            Splits(0)._ColumnProps(19)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2117"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2037"
            Splits(0)._ColumnProps(24)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits.Count    =   1
            AllowRowSizing  =   0   'False
            Appearance      =   1
            BorderStyle     =   1
            ColumnHeaders   =   -1  'True
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            RowDividerStyle =   2
            LayoutName      =   ""
            LayoutFileName  =   ""
            LayoutURL       =   ""
            EmptyRows       =   0   'False
            ListField       =   ""
            DataField       =   ""
            IntegralHeight  =   0   'False
            FetchRowStyle   =   0   'False
            AlternatingRowStyle=   0   'False
            DataMember      =   ""
            ColumnFooters   =   0   'False
            FootLines       =   1
            DeadAreaBackColor=   12648447
            ValueTranslate  =   0   'False
            RowDividerColor =   15790320
            RowSubDividerColor=   15790320
            _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=128,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=17"
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
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid gridModel 
            Height          =   5535
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   9763
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Full Model"
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   1
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Kelas"
            Columns(1).DataField=   ""
            Columns(1).DropDown=   "TDBDropDown1"
            Columns(1).DropDown.vt=   8
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Model"
            Columns(2).DataField=   ""
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Engine"
            Columns(3).DataField=   ""
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Serial Chassis"
            Columns(4).DataField=   ""
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
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=8196"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1).Button=1"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(15)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(16)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(17)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(18)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(19)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(20)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(21)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(22)=   "Column(4).Order=5"
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
            DeadAreaBackColor=   6876390
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=212,.bold=0,.fontsize=825,.italic=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13,.locked=-1"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13,.locked=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.locked=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.locked=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
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
            _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(71)  =   "Named:id=40:OddRow"
            _StyleDefs(72)  =   ":id=40,.parent=33"
            _StyleDefs(73)  =   "Named:id=41:RecordSelector"
            _StyleDefs(74)  =   ":id=41,.parent=34"
            _StyleDefs(75)  =   "Named:id=42:FilterBar"
            _StyleDefs(76)  =   ":id=42,.parent=33"
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   495
            Left            =   9120
            Top             =   5160
            Visible         =   0   'False
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
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
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   0
         Left            =   1590
         TabIndex        =   0
         Top             =   465
         Width           =   2055
         _ExtentX        =   3625
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
         Index           =   1
         Left            =   1590
         TabIndex        =   1
         Top             =   765
         Width           =   4020
         _ExtentX        =   7091
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
         Index           =   3
         Left            =   1590
         TabIndex        =   3
         Top             =   1770
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
         Index           =   4
         Left            =   1590
         TabIndex        =   4
         Top             =   2070
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   270
         Index           =   2
         Left            =   1590
         TabIndex        =   2
         Top             =   1080
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">&&&&&&"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   5
         Left            =   1590
         TabIndex        =   11
         Top             =   1395
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
      Begin VB.Image tblSync 
         Height          =   3015
         Left            =   5880
         Picture         =   "frmtModel.frx":0015
         Top             =   240
         Width           =   2385
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   2520
         Top             =   1080
         Width           =   60
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Frame"
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
         TabIndex        =   10
         Top             =   2130
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Group Model"
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
         TabIndex        =   9
         Top             =   1830
         Width           =   900
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3840
         Top             =   1770
         Width           =   60
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Kelas"
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
         TabIndex        =   8
         Top             =   1140
         Width           =   375
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   270
         Left            =   3720
         Top             =   480
         Width           =   60
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   570
         TabIndex        =   7
         Top             =   840
         Width           =   840
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "MDL"
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
         Left            =   1110
         TabIndex        =   6
         Top             =   540
         Width           =   300
      End
   End
   Begin AIFCmp1.asxPowerBanner asxPowerBanner1 
      Height          =   7395
      Left            =   0
      Top             =   -60
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   13044
      StartColor      =   49344
      EndColor        =   12648447
      FormatString    =   "     Tabel Model Kendaraan"
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
      Top             =   30
      Width           =   11340
      _ExtentX        =   20003
      _ExtentY        =   503
      BackColor       =   6876390
      ShadowColor     =   0
   End
   Begin AIFCmp1.asxPowerButton cmdHapus 
      Height          =   360
      Left            =   8070
      TabIndex        =   17
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12648447
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
      Left            =   9285
      TabIndex        =   18
      Top             =   7320
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   635
      BackColor       =   12648447
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
      Left            =   6840
      TabIndex        =   19
      Top             =   7320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   635
      BackColor       =   12648447
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
End
Attribute VB_Name = "frmtModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMdlSync As New ADODB.Recordset

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
            Case "PPN", "DISC"
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
            Case "PPN", "DISC"
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

Private Function createMdlSync() As Boolean
1         On Error GoTo vb_error
          
2         If rsMdlSync.State Then rsMdlSync.Close
3         Set rsMdlSync = Nothing
4         With rsMdlSync
5             .Fields.Append "MDL", adVarChar, 50
6             .Fields.Append "KLAS", adVarChar, 10
7             .Fields.Append "KET", adVarChar, 50
8             .Fields.Append "EG", adVarChar, 20
9             .Fields.Append "CHS", adVarChar, 25
10            .Open
11        End With
12        If rsMdlSync.State Then createMdlSync = True
13        Exit Function
          
vb_error:
14        MsgBox ErrMessage(Erl, err.Number, "Procedure: " & Me.Name & " .CreateTbl"), vbCritical
End Function

Private Sub asxBars1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub asxPowerBanner1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    FormDrag Me
End Sub

Private Sub cmdBatal_Click()
    framSync.Visible = False
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdHapus_Click()
    If MsgBox("Ingin menghapus data ini ?", vbYesNo, "Konfirmasi") = vbYes Then
        Call objO.UpdateDataHO("UPDATE model SET FD_TGL_DEL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', FS_LOG_DEL = '" & gSA & "' where mdl='" & Text(0) & "' and FS_LOG_DEL = ''")
        Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('MODEL', '" & Text(0) & "', 'D')")
        Text(0) = ""
        Call hapusForm
        Text(0).SetFocus
    End If
End Sub

Private Sub cmdSimpan_Click()
Dim rsSearch As ADODB.Recordset

    If Text(2) <> "" Then
        Set rsSearch = objO.FetchDataHO("Select * from model WITH(NOLOCK) where mdl='" & Text(0) & "'")
        
        Call ClearsSql
        gsSql1(0) = "MDL"
        gsSql2(0) = Text(0)
        gsSql1(1) = "[Desc]"
        gsSql2(1) = Text(1)
        gsSql1(2) = "MODEL"
        gsSql2(2) = Text(2)
        gsSql1(3) = "FS_GRP_MDL"
        gsSql2(3) = Text(3)
        gsSql1(4) = "NO_CHASSIS"
        gsSql2(4) = Text(4)
        
        If Not rsSearch.BOF Then
            sSQL = GensSql("model", True) & " where mdl='" & Text(0) & "'"
        Else
            sSQL = GensSql("model", False)
        End If
        Call objO.UpdateDataHO(sSQL)
        Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('MODEL', '" & Text(0) & "', 'S')")
        Text(0) = ""
        Call hapusForm
        Text(0).SetFocus
    Else
        MsgBox "Masukkan Model !"
    End If
End Sub

Private Sub simpanModel()
      Dim rsModel As ADODB.Recordset

1        On Error GoTo simpanModel_Error

2         If rsMdlSync.State <> adStateClosed Then
3             rsMdlSync.MoveFirst
4             Do While Not rsMdlSync.EOF
5                 If Trim(rsMdlSync("KLAS")) <> "" Then
6                     Set rsModel = objO.FetchDataHO("SELECT * FROM MODEL WITH(NOLOCK) WHERE MDL = '" & Trim(rsMdlSync("MDL")) & "' AND FS_LOG_DEL = ''")
7                     If rsModel.EOF Then
8                         Call objO.InsertDataHO("INSERT INTO MODEL (MDL,MODEL,[DESC],FS_GRP_MDL,NO_CHASSIS,FS_LOG_DEL,FD_TGL_DEL,picFolder,ISICABANG) VALUES ('" & Trim(rsMdlSync("MDL")) & "','" & Trim(rsMdlSync("KLAS")) & "','" & Trim(rsMdlSync("KET")) & "','" & Trim(rsMdlSync("EG")) & "','" & Trim(rsMdlSync("CHS")) & "','','',0,0)")
9                     Else
10                        Call objO.UpdateDataHO("UPDATE MODEL SET MODEL = '" & Trim(rsMdlSync("KLAS")) & "', [DESC] = '" & Trim(rsMdlSync("KET")) & "', FS_GRP_MDL = '" & Trim(rsMdlSync("EG")) & "', NO_CHASSIS = '" & Trim(rsMdlSync("CHS")) & "' WHERE MDL = '" & Trim(rsMdlSync("MDL")) & "'")
11                    End If
12                    Call objO.InsertDataHO("INSERT INTO TEMPKIRIM VALUES ('MODEL', '" & Trim(rsMdlSync("MDL")) & "', 'S')")
13                End If
14                rsMdlSync.MoveNext
15            Loop
16        Else
17            MsgBox "Data Kosong. Tidak ada proses syncronisasi"
18        End If

19       On Error GoTo 0
20       Exit Sub

simpanModel_Error:

21        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanModel of Form frmtModel")
End Sub

Private Sub Form_Activate()
    With Adodc1
        .ConnectionString = adoCon
        .CommandType = adCmdText
        .RecordSource = "SELECT ClassCode, ClassName, Note, Price, PriceTAM FROM CARCLASS WITH(NOLOCK) WHERE PRICETAM > 0 AND FS_LOG_DEL = ''"
        .Refresh
    End With
End Sub

Private Sub Form_Load()
    Call hapusForm
    Call clearForm
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 600
    asxBars1.Width = Me.Width - asxPowerBanner1.Width - 300
    Frame1.Width = asxBars1.Width
    Frame1.Height = asxPowerBanner1.Height - cmdExit.Height - 700
    
    cmdExit.Top = Frame1.Top + Frame1.Height + 100
    cmdHapus.Top = cmdExit.Top
    cmdSimpan.Top = cmdExit.Top
    cmdExit.Left = Frame1.Left + Frame1.Width - cmdExit.Width - 100
    cmdHapus.Left = cmdExit.Left - cmdHapus.Width - 100
    cmdSimpan.Left = cmdHapus.Left - cmdSimpan.Width - 100
    
    framSync.Height = Frame1.Height - 300
    GridModel.Height = framSync.Height - cmdBatal.Height - 500
    cmdBatal.Top = GridModel.Top + GridModel.Height + 100
    Proses.Top = cmdBatal.Top
End Sub

Private Sub gridModel_AfterColUpdate(ByVal ColIndex As Integer)
      Dim rsHlp As ADODB.Recordset
          
1         Select Case ColIndex
              Case 1
2                 With GridModel
3                     Set rsHlp = objO.FetchDataHO("SELECT * FROM CARCLASS WITH(NOLOCK) WHERE CLASSCODE = '" & Trim(.Columns(ColIndex)) & "'")
4                     If rsHlp.EOF Then
5                         rsMdlSync.Find "MDL = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
6                         If Not rsMdlSync.EOF Then
7                             rsMdlSync("KLAS") = ""
8                         End If
9                     Else
10                        rsMdlSync("KLAS") = UCase(Trim(.Columns(ColIndex)))
11                    End If
12                End With
13            Case 2
14                If Len(GridModel.Columns(ColIndex)) > 50 Then
15                    MsgBox "Tidak boleh lebih dari 50 karakter. "
16                    GridModel.Columns(ColIndex) = ""
17                End If
18                rsMdlSync.Find "MDL = '" & Trim(GridModel.Columns(0)) & "'", , adSearchForward, 1
19                If Not rsMdlSync.EOF Then
20                    rsMdlSync("KET") = UCase(Trim(GridModel.Columns(ColIndex)))
21                    GridModel.Columns(ColIndex) = UCase(Trim(GridModel.Columns(ColIndex)))
22                End If
23            Case 3
                  'If Len(GridModel.Columns(ColIndex)) > 10 Then
24                If Len(GridModel.Columns(ColIndex)) > 20 Then
                      'MsgBox "Tidak boleh lebih dari 10 karakter. "
25                    MsgBox "Tidak boleh lebih dari 20 karakter. "
26                    GridModel.Columns(ColIndex) = ""
27                End If
28                rsMdlSync.Find "MDL = '" & Trim(GridModel.Columns(0)) & "'", , adSearchForward, 1
29                If Not rsMdlSync.EOF Then
30                    rsMdlSync("EG") = UCase(Trim(GridModel.Columns(ColIndex)))
31                    GridModel.Columns(ColIndex) = UCase(Trim(GridModel.Columns(ColIndex)))
32                End If
33            Case 4
34                If Len(GridModel.Columns(ColIndex)) > 6 Then
35                    MsgBox "Tidak boleh lebih dari 6 karakter. "
36                    GridModel.Columns(ColIndex) = ""
37                End If
38                rsMdlSync.Find "MDL = '" & Trim(GridModel.Columns(0)) & "'", , adSearchForward, 1
39                If Not rsMdlSync.EOF Then
40                    rsMdlSync("CHS") = UCase(Trim(GridModel.Columns(ColIndex)))
41                    GridModel.Columns(ColIndex) = UCase(Trim(GridModel.Columns(ColIndex)))
42                End If
43        End Select
End Sub

Private Sub GridModel_ButtonClick(ByVal ColIndex As Integer)
    Select Case ColIndex
        Case 2
            With Adodc1
                .ConnectionString = adoCon
                .CommandType = adCmdText
                .RecordSource = "SELECT ClassCode, ClassName, Note, Price, PriceTAM FROM CARCLASS WITH(NOLOCK) WHERE PRICETAM > 0 AND FS_LOG_DEL = ''"
                .Refresh
            End With
    End Select
End Sub

Private Sub gridModel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 18
            With Adodc1
                .ConnectionString = adoCon
                .CommandType = adCmdText
                .RecordSource = "SELECT ClassCode, ClassName, Note, Price, PriceTAM FROM CARCLASS WITH(NOLOCK) WHERE PRICETAM > 0 AND FS_LOG_DEL = ''"
                .Refresh
            End With
    End Select
End Sub

Private Sub Proses_Click()
    Call simpanModel
    framSync.Visible = False
End Sub

Private Sub tblSync_Click()
    Call clearForm
    framSync.Visible = True
End Sub

Private Sub clearForm()

    Call dataSync
    TDBDropDown1.Width = 10000
    Set TDBDropDown1.DataSource = Adodc1
    With Adodc1
        .ConnectionString = adoCon
        .CommandType = adCmdText
        .RecordSource = "SELECT ClassCode, ClassName, Note, Price, PriceTAM FROM CARCLASS WITH(NOLOCK) WHERE PRICETAM > 0 AND FS_LOG_DEL = ''"
        .Refresh
    End With

End Sub

Private Sub dataSync()
Dim rsSearch As ADODB.Recordset

        If rsMdlSync.State <> adStateClosed Then rsMdlSync.Close
'        Set rsSearch = objO.FetchDataHO("SELECT KODEKATEGORY AS MDL, '' AS KELAS, NAMAKATEGORY AS KET, '' AS EG,'' AS CHS " & _
'               "FROM OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',SPK00.DBO.INKATEGORY) " & _
'               "WHERE KODEKATEGORY NOT IN (SELECT MDL FROM BENGKEL99..MODEL WITH(NOLOCK) WHERE FS_LOG_DEL = '')")
        Set rsSearch = objO.FetchDataHO("SELECT KODEKATEGORY AS MDL, '' AS KELAS, NAMAKATEGORY AS KET, '' AS EG,'' AS CHS " & _
               "FROM SPK00.DBO.INKATEGORY " & _
               "WHERE KODEKATEGORY NOT IN (SELECT MDL FROM BENGKEL99..MODEL WITH(NOLOCK) WHERE FS_LOG_DEL = '')")

        If Not rsSearch.EOF Then
            Do While Not rsSearch.EOF
                With rsMdlSync
                    If .State = adStateClosed Then createMdlSync
                    If .State Then
                        .AddNew
                        !MDL = Trim(rsSearch("MDL"))
                        !klas = Trim(rsSearch("KELAS"))
                        !Ket = Trim(rsSearch("KET"))
                        !EG = Trim(rsSearch("EG"))
                        !CHS = Trim(rsSearch("CHS"))
                        .Update
                    End If
                End With
                rsSearch.MoveNext
            Loop
            Set GridModel.DataSource = rsMdlSync
            Call headModel
            GridModel.MoveFirst
        End If
End Sub

Private Sub headModel()

    With GridModel
        .Columns(0).Caption = "Full Model"
        .Columns(0).DataField = "MDL"
        .Columns(0).Width = 2000
        
        .Columns(1).Caption = "Kelas"
        .Columns(1).DataField = "KLAS"
        .Columns(1).Width = 1000
        
        .Columns(2).Caption = "Model"
        .Columns(2).DataField = "KET"
        .Columns(2).Width = 4000
        
        .Columns(3).Caption = "Engine"
        .Columns(3).DataField = "EG"
        .Columns(3).Width = 1200
        
        .Columns(4).Caption = "Serial Chassis"
        .Columns(4).DataField = "CHS"
        .Columns(4).Width = 1200
    End With
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
                gSearchOption = "MDL"
                frmCariHO.show (vbModal)
            Case 2
                gMainKeyCari = Text(Index)
                gSearchOption = "KELAS_MOBIL"
                frmCariHO.show (vbModal)
            Case 3
                gMainKeyCari = Text(Index)
                gSearchOption = "GROUP_MODEL"
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
Dim rsSearch As ADODB.Recordset

    Select Case Index
        Case 0
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchDataHO("Select * from model a WITH(NOLOCK) left join carclass b WITH(NOLOCK) on a.model = b.classcode and a.FS_LOG_DEL = '' where MDL='" & Text(0) & "'")
                If Not rsSearch.EOF Then
                    Text(1) = rsSearch("DESC")
                    Text(2) = rsSearch("MODEL")
                    Text(3) = rsSearch("FS_GRP_MDL")
                    Text(4) = rsSearch("NO_CHASSIS")
                    Text(5) = Nul(rsSearch("ClassName"))
                Else
                    Call hapusForm
                End If
            End If
        Case 2
            If Text(Index) <> "" Then
                Set rsSearch = objO.FetchDataHO("select * from carclass WITH(NOLOCK) WHERE classcode = '" & Text(Index) & "' and FS_LOG_DEL = ''")
                If rsSearch.EOF Then
                    Cancel = True
                    gMainKeyCari = Text(Index)
                    gSearchOption = "KELAS_MOBIL"
                    frmCariHO.show (vbModal)
                Else
                    Text(5) = rsSearch("className")
                End If
            End If
    End Select
End Sub

Private Sub hapusForm()
    Dim i As Integer
    
    For i = 1 To 5
        Text(i) = ""
    Next
End Sub
