VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmRepoHansaRenkei 
   Caption         =   "Laporan Hansa Renkei"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   10935
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoHansaRenkei.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   34
      Top             =   0
      Width           =   445
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   8655
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   11175
      Begin VB.Frame frameDetailCustomer 
         Caption         =   "Data Detail Customer"
         Height          =   6855
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Visible         =   0   'False
         Width           =   10935
         Begin VB.CommandButton cmdTutup 
            BackColor       =   &H000000FF&
            Caption         =   "&Tutup"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   9840
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   0
            Width           =   1050
         End
         Begin TrueOleDBGrid70.TDBGrid gridCustomer 
            Height          =   6375
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   10695
            _ExtentX        =   18865
            _ExtentY        =   11245
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Nama Customer"
            Columns(0).DataField=   "Nama"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "No Rangka"
            Columns(1).DataField=   "Norangka"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Kode Sales"
            Columns(2).DataField=   "kodesales"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Tgl Jatuh Tempo"
            Columns(3).DataField=   "jthtempo"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Tgl Follow"
            Columns(4).DataField=   "tglfollow"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Book"
            Columns(5).DataField=   "Book"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "No Order"
            Columns(6).DataField=   "wo"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   7
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14933984
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=7"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(31)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(32)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(33)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(34)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(35)=   "Column(6).Order=7"
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=-1,.fontsize=825"
            _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&H93C6FD&"
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
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Named:id=33:Normal"
            _StyleDefs(65)  =   ":id=33,.parent=0"
            _StyleDefs(66)  =   "Named:id=34:Heading"
            _StyleDefs(67)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(68)  =   ":id=34,.wraptext=-1"
            _StyleDefs(69)  =   "Named:id=35:Footing"
            _StyleDefs(70)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(71)  =   "Named:id=36:Selected"
            _StyleDefs(72)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(73)  =   "Named:id=37:Caption"
            _StyleDefs(74)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(75)  =   "Named:id=38:HighlightRow"
            _StyleDefs(76)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=39:EvenRow"
            _StyleDefs(78)  =   ":id=39,.parent=33,.bgcolor=&H4080FF&"
            _StyleDefs(79)  =   "Named:id=40:OddRow"
            _StyleDefs(80)  =   ":id=40,.parent=33"
            _StyleDefs(81)  =   "Named:id=41:RecordSelector"
            _StyleDefs(82)  =   ":id=41,.parent=34"
            _StyleDefs(83)  =   "Named:id=42:FilterBar"
            _StyleDefs(84)  =   ":id=42,.parent=33"
            _StyleDefs(85)  =   "Named:id=131:TipeTanpaTipe"
            _StyleDefs(86)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
            _StyleDefs(87)  =   "Named:id=132:TigaDigitKosong"
            _StyleDefs(88)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         End
      End
      Begin ARProgBarCtrl.ARProgressBar prog 
         Height          =   495
         Left            =   480
         TabIndex        =   24
         Top             =   840
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Border          =   2
         Caption         =   "0 %"
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Legend"
         Height          =   2895
         Left            =   6960
         TabIndex        =   18
         Top             =   4320
         Width           =   4095
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            Caption         =   "Show "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Top             =   2280
            Width           =   3015
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Booking "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Top             =   1800
            Width           =   3015
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Caption         =   "Terhubung "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   21
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000080FF&
            Caption         =   "Terfollowup "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   20
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FF0000&
            Caption         =   "Data Reminder By Sales"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.PictureBox grafikHasil 
         Height          =   6015
         Left            =   480
         Picture         =   "frmRepoHansaRenkei.frx":8E51
         ScaleHeight     =   5955
         ScaleWidth      =   6765
         TabIndex        =   10
         Top             =   1560
         Width           =   6830
         Begin VB.Label lblPersenShow 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "100% "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   4200
            TabIndex        =   28
            Top             =   4680
            Width           =   975
         End
         Begin VB.Label lblPersenBooking 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "100% "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   4680
            TabIndex        =   27
            Top             =   3840
            Width           =   975
         End
         Begin VB.Label lblPersenTerhubung 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "100% "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   5040
            TabIndex        =   26
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label lblPersenFollow 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "100% "
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   14.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   5760
            TabIndex        =   25
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label lblShow 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   18
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2280
            TabIndex        =   17
            Top             =   4920
            Width           =   2055
         End
         Begin VB.Label lblBook 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   26.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2400
            TabIndex        =   16
            Top             =   4200
            Width           =   2055
         End
         Begin VB.Label lblTerhubung 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   26.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2400
            TabIndex        =   15
            Top             =   3360
            Width           =   2055
         End
         Begin VB.Label lblTerfollow 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   26.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2400
            TabIndex        =   14
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label lblReminder 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   26.25
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   615
            Left            =   2400
            TabIndex        =   11
            Top             =   960
            Width           =   2055
         End
      End
      Begin MSComCtl2.DTPicker Tgl1 
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   96206849
         CurrentDate     =   39852
      End
      Begin MSComCtl2.DTPicker Tgl2 
         Height          =   255
         Left            =   3480
         TabIndex        =   2
         Top             =   210
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         Format          =   96206849
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   3
         Top             =   750
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox MaskEdBox4 
         Height          =   270
         Left            =   3765
         TabIndex        =   4
         Top             =   750
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   8
         Left            =   1800
         TabIndex        =   5
         Top             =   1065
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Index           =   14
         Left            =   1965
         TabIndex        =   6
         Top             =   2145
         Visible         =   0   'False
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   503
         _Version        =   393216
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
      Begin ucXPButton.XPButton tblSimpan 
         Height          =   495
         Left            =   5760
         TabIndex        =   12
         Top             =   840
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   "&p"
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
         MICON           =   "frmRepoHansaRenkei.frx":13BBA
         PICN            =   "frmRepoHansaRenkei.frx":13BD6
         PICH            =   "frmRepoHansaRenkei.frx":1402D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton tblKeluar 
         Height          =   495
         Left            =   7080
         TabIndex        =   13
         Top             =   840
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
         MICON           =   "frmRepoHansaRenkei.frx":14517
         PICN            =   "frmRepoHansaRenkei.frx":14533
         PICH            =   "frmRepoHansaRenkei.frx":14A07
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TDBDate6Ctl.TDBDate tglShow1 
         Height          =   315
         Left            =   1800
         TabIndex        =   35
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "frmRepoHansaRenkei.frx":14F52
         Caption         =   "frmRepoHansaRenkei.frx":1507E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoHansaRenkei.frx":150EA
         Keys            =   "frmRepoHansaRenkei.frx":15108
         Spin            =   "frmRepoHansaRenkei.frx":15166
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin TDBDate6Ctl.TDBDate tglShow2 
         Height          =   315
         Left            =   3480
         TabIndex        =   36
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "frmRepoHansaRenkei.frx":1518E
         Caption         =   "frmRepoHansaRenkei.frx":152BA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoHansaRenkei.frx":15326
         Keys            =   "frmRepoHansaRenkei.frx":15344
         Spin            =   "frmRepoHansaRenkei.frx":153A2
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
         Text            =   "__/__/____"
         ValidateMode    =   0
         ValueVT         =   2118189057
         Value           =   2.12482692446619E-314
         CenturyMode     =   0
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3360
         TabIndex        =   30
         Top             =   510
         Width           =   60
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Show"
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
         TabIndex        =   29
         Top             =   510
         Width           =   1005
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   480
         X2              =   8280
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode Sales"
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
         Left            =   660
         TabIndex        =   9
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Reminder"
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
         Left            =   270
         TabIndex        =   8
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
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
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   60
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   375
      Left            =   480
      Top             =   120
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   661
   End
End
Attribute VB_Name = "frmRepoHansaRenkei"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp      As ADODB.Recordset
Dim rsCustomer  As ADODB.Recordset

Private Sub cmdTutup_Click()
    frameDetailCustomer.Visible = False
End Sub

Private Sub Form_Load()
    prog.Visible = False
    grafikHasil.Visible = False
    Frame2.Visible = False
    Tgl1.value = "01/" & Month(isnow) & "/" & year(isnow)
    Tgl2.value = DateAdd("D", -1, DateAdd("M", 1, "01/" & Month(isnow) & "/" & year(isnow)))
    'tglShow1.value = "01/01/1900"
    'tglShow2.value = "01/01/1900"
End Sub

Private Sub lblBook_Click()
1         If IsNumeric(NulSID(lblBook.Caption)) Then
2             If CInt(NulSID(lblBook.Caption)) > 0 Then
3                 frameDetailCustomer.Visible = True
                  '' 20200408DEV
4                 If Log.sentralisasiStars = 1 Then
5                     sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
6                     sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
7                     sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
8                     sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
9                     sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.KodeCabang & "'"
10                    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan  and c.kodecabang = '" & Log.KodeCabang & "'"
11                    sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and a.TypePlan = 3 and"
12                Else
13                    sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
14                    sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
15                    sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
16                    sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
17                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
18                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan"
19                    sSQL = sSQL & " where a.TypePlan = 3 and"
20                End If
21                If NulSID(Text(7)) <> "" Then
22                    sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
23                End If
24                sSQL = sSQL & " (a.TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' ) "
                  ''sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and"
25                sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and"
26                sSQL = sSQL & " a.NoRangka in (select no_chassis from BOOK where FS_LOG_DEL = '' and DATE_BOOK between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59') and"
                  '' 20200408DEV
27                If Log.sentralisasiStars = 1 Then
28                    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
29                Else
30                    sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
31                End If
32                Set rsCustomer = objO.FetchData(sSQL)
33                gridCustomer.DataSource = rsCustomer
34            End If
35        End If
End Sub

Private Sub lblReminder_Click()
1         If IsNumeric(NulSID(lblReminder.Caption)) Then
2             If CInt(NulSID(lblReminder.Caption)) > 0 Then
3                 frameDetailCustomer.Visible = True
                  '' 20200408DEV
4                 If Log.sentralisasiStars = 1 Then
5                     sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
6                     sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
7                     sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = ''  and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
8                     sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
9                     sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr AND b.kodecabang = '" & Log.KodeCabang & "'"
10                    sSQL = sSQL & " left join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan  AND c.kodecabang = '" & Log.KodeCabang & "'"
11                    sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and a.TypePlan = 3 and"
12                Else
13                    sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
14                    sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
15                    sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = ''  and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
16                    sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
17                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
18                    sSQL = sSQL & " left join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan"
19                    sSQL = sSQL & " where a.TypePlan = 3 and"
20                End If
21                If NulSID(Text(7)) <> "" Then
22                    sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
23                End If
24                sSQL = sSQL & " (a.TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' ) and"
                  '' 20200408DEV
25                If Log.sentralisasiStars = 1 Then
26                    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
27                Else
28                    sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
29                End If
30                Set rsCustomer = objO.FetchData(sSQL)
31                gridCustomer.DataSource = rsCustomer
32            End If
33        End If
End Sub

Private Sub lblShow_Click()
1         If IsNumeric(NulSID(lblShow.Caption)) Then
2             If CInt(NulSID(lblShow.Caption)) > 0 Then
3                 frameDetailCustomer.Visible = True
                  '' 20200408DEV
4                 If Log.sentralisasiStars = 1 Then
5                     sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
6                     sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
7                     sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
8                     sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
9                     sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr AND b.kodecabang = '" & Log.KodeCabang & "'"
10                    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan AND c.kodecabang = '" & Log.KodeCabang & "'"
11                    sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' a.TypePlan = 3 and"
12                Else
13                    sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
14                    sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
15                    sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
16                    sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
17                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
18                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan"
19                    sSQL = sSQL & " where a.TypePlan = 3 and"
20                End If
21                If NulSID(Text(7)) <> "" Then
22                    sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
23                End If
24                sSQL = sSQL & " (a.TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' ) "
                  ''sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and"
25                sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and"
26                sSQL = sSQL & " a.NoRangka in (select no_chassis from BOOK where FS_LOG_DEL = '' and DATE_BOOK between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59') and"
27                sSQL = sSQL & " a.NoRangka in (select no_chassis from JOBDTA where FS_LOG_DEL = '' and DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' and NOFAK <> '') and"
                  '' 20200408DEV
28                If Log.sentralisasiStars = 1 Then
29                    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
30                Else
31                    sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
32                End If
33                Set rsCustomer = objO.FetchData(sSQL)
34                gridCustomer.DataSource = rsCustomer
35            End If
36        End If
End Sub

Private Sub lblTerfollow_Click()
1         If IsNumeric(NulSID(lblTerfollow.Caption)) Then
2             If CInt(NulSID(lblTerfollow.Caption)) > 0 Then
3                 frameDetailCustomer.Visible = True
                  '' 20200408DEV
4                 If Log.sentralisasiStars = 1 Then
5                     sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
6                     sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
7                     sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
8                     sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
9                     sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.KodeCabang & "'"
10                    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan and c.kodecabang = '" & Log.KodeCabang & "'"
11                    sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and a.TypePlan = 3 and"
12                Else
13                    sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
14                    sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
15                    sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
16                    sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
17                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
18                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan"
19                    sSQL = sSQL & " where a.TypePlan = 3 and"
20                End If
21                If NulSID(Text(7)) <> "" Then
22                    sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
23                End If
24                sSQL = sSQL & " (a.TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' ) and"
                  '' 20200408DEV
25                If Log.sentralisasiStars = 1 Then
26                    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where KetPlan like '%reminder service%'and kodecabang = '" & Log.KodeCabang & "')"
27                Else
28                    sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
29                End If
30                Set rsCustomer = objO.FetchData(sSQL)
31                gridCustomer.DataSource = rsCustomer
32            End If
33        End If
End Sub

Private Sub lblTerhubung_Click()
1         If IsNumeric(NulSID(lblTerhubung.Caption)) Then
2             If CInt(NulSID(lblTerhubung.Caption)) > 0 Then
3                 frameDetailCustomer.Visible = True
                  '' 20200408DEV
4                 If Log.sentralisasiStars = 1 Then
5                     sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
6                     sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
7                     sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
8                     sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
9                     sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.KodeCabang & "'"
10                    sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan and c.kodecabang = '" & Log.KodeCabang & "'"
11                    sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and a.TypePlan = 3 and"
12                Else
13                    sSQL = "select b.Nama,a.NoRangka,a.KodeSales,a.[Tanggal Jatuh Tempo] jthtempo,C.tglfollow,"
14                    sSQL = sSQL & " isnull((SELECT TOP 1 NOFAK FROM BOOK where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as book,"
15                    sSQL = sSQL & " isnull((SELECT TOP 1 NO_WO FROM jobdta where FS_LOG_DEL = '' and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and NO_CHASSIS = a.NoRangka and (DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59')),'') as wo"
16                    sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
17                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
18                    sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = a.kodepr and c.kodeprospect = a.kodeprospect and a.KodePlan = c.KodePlan"
19                    sSQL = sSQL & " where a.TypePlan = 3 and"
20                End If
21                If NulSID(Text(7)) <> "" Then
22                    sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
23                End If
24                sSQL = sSQL & " (a.TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' ) "
25                sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and"
                  '' 20200408DEV
26                If Log.sentralisasiStars = 1 Then
27                    sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where a.kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
28                Else
29                    sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
30                End If
31                Set rsCustomer = objO.FetchData(sSQL)
32                gridCustomer.DataSource = rsCustomer
33            End If
34        End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSimpan_Click()
    Dim reminder, terfollow, terhubung, boc, show As Integer
    
    If DateDiff("D", Format(Tgl1, "yyyy-mm-dd"), Format(Tgl2, "yyyy-mm-dd")) < 0 Then
        MsgBox "Range tanggal tidak valid", vbInformation
        Exit Sub
    End If
    If NulSID(Text(7)) <> "" Then
        Set rsTemp = objO.FetchData("select KodeSales,NamaSales from [SPK" & Right(gDb, 2) & "]..ArSales where TdkAktif = 0 and kodesales = '" & NulSID(Text(7)) & "'")
        If rsTemp.EOF Then
            MsgBox "Kode sales " & NulSID(Text(7)) & " tidak ditemukan.", vbInformation
            Text(7) = ""
            Text(8) = ""
            Exit Sub
        End If
    End If
    prog.Visible = True
    prog.Max = 7
    grafikHasil.Visible = True
    
    ''-- DATA REMINDER SALES
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = "select count(b.SalesPr) as reminder"
        sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and TypePlan = 3 and"
    Else
        sSQL = "select count(b.SalesPr) as reminder"
        sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
        sSQL = sSQL & " where TypePlan = 3 and"
    End If
    If NulSID(Text(7)) <> "" Then
        sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
    End If
    sSQL = sSQL & " (TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' ) and"
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = sSQL & " idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
    Else
        sSQL = sSQL & " idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
    End If
    prog.value = 1
    DoEvents
    prog.Caption = "Ambil data Reminder Sales"
    reminder = objO.FetchData_Double(sSQL)
    lblReminder.Caption = reminder
    DoEvents
    
    ''-- DATA REMINDER TERFOLLOW2 SALES
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = "select count(b.SalesPr) as terfollow"
        sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect and c.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and a.TypePlan = 3 and"
    Else
        sSQL = "select count(b.SalesPr) as terfollow"
        sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect"
        sSQL = sSQL & " where a.TypePlan = 3 and"
    End If
    If NulSID(Text(7)) <> "" Then
        sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
    End If
    sSQL = sSQL & " (TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' ) and"
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
    Else
        sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
    End If
     
    prog.value = 2
    DoEvents
    prog.Caption = "Ambil data Terfollow Sales"
    'reminder = objO.FetchData_Double(sSQL)
    terfollow = objO.FetchData_Double(sSQL)
    If reminder = 0 Then
        lblPersenFollow.Caption = "0 % "
    ElseIf terfollow > reminder Then
        lblPersenFollow.Caption = "100% "
    Else
        lblPersenFollow.Caption = Round((terfollow / reminder) * 100, 0) & "% "
    End If
    lblTerfollow.Caption = terfollow
    DoEvents
     
    ''-- DATA REMINDER TERHUBUNG SALES
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = "select count(b.SalesPr) as terhubung"
        sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect and c.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and a.TypePlan = 3 and"
    Else
        sSQL = "select count(b.SalesPr) as terhubung"
        sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect"
        sSQL = sSQL & " where a.TypePlan = 3 and"
    End If
    If NulSID(Text(7)) <> "" Then
        sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
    End If
    sSQL = sSQL & " (TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' )"
    sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and"
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
    Else
        sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
    End If
    prog.value = 3
    DoEvents
    prog.Caption = "Ambil data Terhubung Sales"
    'reminder = objO.FetchData_Double(sSQL)
    terhubung = objO.FetchData_Double(sSQL)
    If terhubung = 0 Then
        lblPersenTerhubung.Caption = "0 % "
    ElseIf terhubung > terfollow Then
        lblPersenTerhubung.Caption = "100% "
    Else
        lblPersenTerhubung.Caption = Round((terhubung / terfollow) * 100, 0) & "% "
    End If
    lblTerhubung.Caption = terhubung
    DoEvents
    
    ''-- DATA REMINDER TERHUBUNG DAN BOC SALES
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = "select count(b.SalesPr) as BOC"
        sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect and c.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and a.TypePlan = 3 and"
    Else
        sSQL = "select count(b.SalesPr) as BOC"
        sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect"
        sSQL = sSQL & " where a.TypePlan = 3 and"
    End If
    If NulSID(Text(7)) <> "" Then
        sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
    End If
    sSQL = sSQL & " (TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' )"
    sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and"
    sSQL = sSQL & " a.NoRangka in (select no_chassis from BOOK where FS_LOG_DEL = '' and DATE_BOOK between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59') and"
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
    Else
        sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
    End If
    prog.value = 4
    DoEvents
    prog.Caption = "Ambil data BOC Sales"
    'reminder = objO.FetchData_Double(sSQL)
    boc = objO.FetchData_Double(sSQL)
    If terhubung = 0 Then
        lblPersenBooking.Caption = "0 % "
    ElseIf boc > terhubung Then
        lblPersenBooking.Caption = "100% "
    Else
        lblPersenBooking.Caption = Round((boc / terhubung) * 100, 0) & "% "
    End If
    lblBook.Caption = boc
    DoEvents
    
    ''- DATA REMINDER TERHUBUNG DAN BOC SALES SHOW BENGKEL
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = "select count(b.SalesPr) as WO"
        sSQL = sSQL & " from [" & dbReplikasiHO & "]..PlanTrn a"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectMst b on a.KodePr = b.KodePr and b.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " inner join [" & dbReplikasiHO & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect and c.kodecabang = '" & Log.KodeCabang & "'"
        sSQL = sSQL & " where a.kodecabang = '" & Log.KodeCabang & "' and a.TypePlan = 3 and"
    Else
        sSQL = "select count(b.SalesPr) as WO"
        sSQL = sSQL & " from [STARS" & Right(gDb, 2) & "]..PlanTrn a"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectMst b on a.KodePr = b.KodePr"
        sSQL = sSQL & " inner join [STARS" & Right(gDb, 2) & "]..ProspectFollow c on c.kodepr = b.KodePr and c.KodePlan = A.KodePlan AND c.KodeProspect = a.KodeProspect"
        sSQL = sSQL & " where a.TypePlan = 3 and"
    End If
    If NulSID(Text(7)) <> "" Then
        sSQL = sSQL & " b.salespr = '" & NulSID(Text(7)) & "' and "
    End If
    sSQL = sSQL & " (TglPlan between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' )"
    sSQL = sSQL & " and c.RetensiKet = 'Terhubung' and c.BoC = 'BOC' and"
    sSQL = sSQL & " a.NoRangka in (select no_chassis from BOOK where FS_LOG_DEL = '' and DATE_BOOK between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59') and"
    sSQL = sSQL & " a.NoRangka in (select no_chassis from JOBDTA where FS_LOG_DEL = '' and DATE between '" & Format(Tgl1.value, "YYYY-MM-DD") & " 00:00:00' AND '" & Format(Tgl2.value, "YYYY-MM-DD") & " 23:59:59' and NOFAK <> '') and"
    '' 20200408DEV
    If Log.sentralisasiStars = 1 Then
        sSQL = sSQL & " a.idketplan in (select idketplan from [" & dbReplikasiHO & "]..Ms_KetPlan where kodecabang = '" & Log.KodeCabang & "' and KetPlan like '%reminder service%')"
    Else
        sSQL = sSQL & " a.idketplan in (select idketplan from [STARS" & Right(gDb, 2) & "]..Ms_KetPlan where KetPlan like '%reminder service%')"
    End If
    
    prog.value = 5
    DoEvents
    prog.Caption = "Ambil data Show Customer Sales"
    'reminder = objO.FetchData_Double(sSQL)
    show = objO.FetchData_Double(sSQL)
    If boc = 0 Then
        lblPersenShow.Caption = "0 % "
    ElseIf show > boc Then
        lblPersenShow.Caption = "100% "
    Else
        lblPersenShow.Caption = Round((show / boc) * 100, 0) & "% "
    End If
    lblShow.Caption = show
    DoEvents
    
    prog.Visible = False
    Frame2.Visible = True
    grafikHasil.Visible = True
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
            Case 7
                gMainKeyCari = Text(Index)
                gSearchOption = "SALES"
                frmCariSPK.show (vbModal)
             
        End Select
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Select Case Index
        Case 7
            If Text(Index) <> "" Then
                Set oRs = objO.FetchDataSPK("select KodeSales,NamaSales from ArSales WITH (NOLOCK) where TdkAktif = 0 AND KODESALES = '" & NulSID(Text(Index)) & "'")
                If oRs.EOF Then
                    gMainKeyCari = Text(Index)
                    gSearchOption = "SALES"
                    frmCariSPK.show (vbModal)
                Else
                    Text(Index + 1) = oRs("NamaSales")
                End If
                oRs.Close
            Else
                Text(Index + 1) = ""
            End If
    End Select
End Sub
