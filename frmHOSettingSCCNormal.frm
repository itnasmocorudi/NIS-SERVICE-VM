VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmHOSettingSCCNormal 
   Caption         =   ".: SETTING SCC NORMAL :."
   ClientHeight    =   9195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   16320
   Begin VB.Frame frmeDealer 
      BackColor       =   &H00FFC0C0&
      Caption         =   $"frmHOSettingSCCNormal.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   2040
      TabIndex        =   2
      Top             =   3840
      Width           =   6615
      Begin TrueOleDBGrid70.TDBGrid GridDealer 
         Height          =   3015
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
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
         Columns(2)._VlistStyle=   4
         Columns(2)._MaxComboItems=   5
         Columns(2).DataField=   "TERLIBAT"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
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
         DeadAreaBackColor=   16761024
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
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39,.bgcolor=&HFFFF00&"
         _StyleDefs(21)  =   ":id=9,.borderColor=&H80000001&"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(49)  =   "Named:id=33:Normal"
         _StyleDefs(50)  =   ":id=33,.parent=0"
         _StyleDefs(51)  =   "Named:id=34:Heading"
         _StyleDefs(52)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   ":id=34,.wraptext=-1"
         _StyleDefs(54)  =   "Named:id=35:Footing"
         _StyleDefs(55)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   "Named:id=36:Selected"
         _StyleDefs(57)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(58)  =   "Named:id=37:Caption"
         _StyleDefs(59)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(60)  =   "Named:id=38:HighlightRow"
         _StyleDefs(61)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   "Named:id=39:EvenRow"
         _StyleDefs(63)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(64)  =   "Named:id=40:OddRow"
         _StyleDefs(65)  =   ":id=40,.parent=33"
         _StyleDefs(66)  =   "Named:id=41:RecordSelector"
         _StyleDefs(67)  =   ":id=41,.parent=34"
         _StyleDefs(68)  =   "Named:id=42:FilterBar"
         _StyleDefs(69)  =   ":id=42,.parent=33"
      End
      Begin AIFCmp1.asxPowerButton cmdTutupDealer 
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   3240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   635
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   360
      TabIndex        =   5
      Top             =   360
      Width           =   12255
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Minggu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   31
         Top             =   2040
         Width           =   10695
         Begin MSComCtl2.DTPicker dtMAwal1 
            Height          =   375
            Left            =   1440
            TabIndex        =   32
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CalendarForeColor=   0
            CustomFormat    =   "HH:mm"
            Format          =   122486787
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtMAkhir1 
            Height          =   375
            Left            =   3960
            TabIndex        =   33
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   122486787
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtMAwal2 
            Height          =   375
            Left            =   6480
            TabIndex        =   34
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   122486787
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtMAkhir2 
            Height          =   375
            Left            =   9120
            TabIndex        =   35
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   122486787
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai 2 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   39
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai 2 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5040
            TabIndex        =   38
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai 1 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   37
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai 1 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   36
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sabtu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   22
         Top             =   1320
         Width           =   10695
         Begin MSComCtl2.DTPicker dtSAwal1 
            Height          =   375
            Left            =   1440
            TabIndex        =   23
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   122486787
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtSAkhir1 
            Height          =   375
            Left            =   3960
            TabIndex        =   24
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   125370371
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtSAwal2 
            Height          =   375
            Left            =   6480
            TabIndex        =   25
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   125370371
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtSAkhir2 
            Height          =   375
            Left            =   9120
            TabIndex        =   26
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   125370371
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai 1 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai 1 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   29
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai 2 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   28
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai 2 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   27
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Senin s/d Jumat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   13
         Top             =   600
         Width           =   10695
         Begin MSComCtl2.DTPicker dtSJAwal1 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "HH:mm"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1057
               SubFormatType   =   4
            EndProperty
            Height          =   375
            Left            =   1440
            TabIndex        =   14
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   125370371
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtSJAkhir1 
            Height          =   375
            Left            =   3960
            TabIndex        =   16
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   125370371
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtSJAwal2 
            Height          =   375
            Left            =   6480
            TabIndex        =   18
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   125370371
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin MSComCtl2.DTPicker dtSJAkhir2 
            Height          =   375
            Left            =   9120
            TabIndex        =   20
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   125370371
            UpDown          =   -1  'True
            CurrentDate     =   43609
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai 2 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   7680
            TabIndex        =   21
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai 2 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5280
            TabIndex        =   19
            Top             =   300
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Selesai 1 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2640
            TabIndex        =   17
            Top             =   300
            Width           =   1305
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jam Mulai 1 : "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   0
            TabIndex        =   15
            Top             =   300
            Width           =   1305
         End
      End
      Begin VB.ComboBox cbDealer 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   240
         Width           =   4815
      End
      Begin ucXPButton.XPButton cmdKeluar 
         Height          =   495
         Left            =   10920
         TabIndex        =   6
         Top             =   7440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOSettingSCCNormal.frx":009C
         PICN            =   "frmHOSettingSCCNormal.frx":00B8
         PICH            =   "frmHOSettingSCCNormal.frx":69EB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin TrueOleDBGrid70.TDBGrid GridSupplier 
         Height          =   4455
         Left            =   120
         TabIndex        =   7
         Top             =   2880
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   7858
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Nama Dealer"
         Columns(0).DataField=   "FS_NM_DEALER"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Hari"
         Columns(1).DataField=   "hari"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "kodedealer"
         Columns(2).DataField=   "kodedealer"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Jam Mulai 1"
         Columns(3).DataField=   "jammulai1"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Jam Selesai 1"
         Columns(4).DataField=   "jamselesai1"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Jam Mulai 2"
         Columns(5).DataField=   "jammulai2"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Jam Selesai 2"
         Columns(6).DataField=   "jamselesai2"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=5292"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5212"
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
         Splits(0)._ColumnProps(15)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(17)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=516"
         Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(22)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=516"
         Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(27)=   "Column(5).Width=2725"
         Splits(0)._ColumnProps(28)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(29)=   "Column(5)._WidthInPix=2646"
         Splits(0)._ColumnProps(30)=   "Column(5)._ColStyle=516"
         Splits(0)._ColumnProps(31)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(32)=   "Column(6).Width=2725"
         Splits(0)._ColumnProps(33)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(6)._WidthInPix=2646"
         Splits(0)._ColumnProps(35)=   "Column(6)._ColStyle=516"
         Splits(0)._ColumnProps(36)=   "Column(6).Order=7"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=62,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=59,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=60,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=61,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=54,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=51,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=52,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=53,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
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
      Begin ucXPButton.XPButton cmdTambah 
         Height          =   495
         Left            =   6120
         TabIndex        =   8
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOSettingSCCNormal.frx":D6A2
         PICN            =   "frmHOSettingSCCNormal.frx":D6BE
         PICH            =   "frmHOSettingSCCNormal.frx":13E61
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
         Left            =   6120
         TabIndex        =   9
         Top             =   7440
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOSettingSCCNormal.frx":1A98A
         PICN            =   "frmHOSettingSCCNormal.frx":1A9A6
         PICH            =   "frmHOSettingSCCNormal.frx":20F08
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdBatal 
         Height          =   495
         Left            =   9720
         TabIndex        =   1
         Top             =   7440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOSettingSCCNormal.frx":2772B
         PICN            =   "frmHOSettingSCCNormal.frx":27747
         PICH            =   "frmHOSettingSCCNormal.frx":2DD40
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdSimpan 
         Height          =   495
         Left            =   8520
         TabIndex        =   0
         Top             =   7440
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         BTYPE           =   6
         TX              =   ""
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
         BCOL            =   15790320
         BCOLO           =   15790320
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmHOSettingSCCNormal.frx":345EF
         PICN            =   "frmHOSettingSCCNormal.frx":3460B
         PICH            =   "frmHOSettingSCCNormal.frx":3AE0E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ucXPButton.XPButton cmdHapus 
         Height          =   495
         Left            =   7350
         TabIndex        =   10
         Top             =   7440
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
         MICON           =   "frmHOSettingSCCNormal.frx":419B4
         PICN            =   "frmHOSettingSCCNormal.frx":419D0
         PICH            =   "frmHOSettingSCCNormal.frx":48543
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   315
         TabIndex        =   11
         Top             =   255
         Width           =   1305
      End
   End
   Begin AIFCmp1.asxBars asxBars1 
      Height          =   285
      Left            =   435
      Top             =   90
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   503
      BackColor       =   16777215
      ShadowColor     =   0
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   8775
      Left            =   0
      Picture         =   "frmHOSettingSCCNormal.frx":4F2DD
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmHOSettingSCCNormal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : frmHOSettingSCCNormal
' DateTime  : 24/05/2019 10:50
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------

 
Dim rsGrid      As New ADODB.Recordset
Dim rsGridMdl   As New ADODB.Recordset
Dim rsTemp      As ADODB.Recordset
Dim rsDealer    As New ADODB.Recordset
Dim cols        As TrueOleDBGrid70.Columns
Dim Urut        As String
Dim isTambah    As Boolean
Dim isBegin     As Boolean


Private Sub ber_sih()
    dtSJAkhir1.value = "07:00"
    dtSJAwal1.value = "00:00"
    
    dtSJAwal2.value = "17:00"
    dtSJAkhir2.value = "23:59"
    
    dtSAkhir1.value = "07:00"
    dtSAwal1.value = "00:00"
    
    dtSAwal2.value = "17:00"
    dtSAkhir2.value = "23:59"
    
    dtMAkhir1.value = "07:00"
    dtMAwal1.value = "00:00"
    
    dtMAwal2.value = "17:00"
    dtMAkhir2.value = "23:59"
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    frmeDealer.Visible = False
    isTambah = True
End Sub

Private Sub buka_Tutup(Optional Flg As Boolean = True)
 
    cmdSimpan.Enabled = Flg
    cmdBatal.Enabled = Flg
    
    cmdTambah.Enabled = Not Flg
    cmdEdit.Enabled = Not Flg
    cmdHapus.Enabled = Flg
End Sub

Private Sub cbDealer_DropDown()
    Call Combo1_Click
End Sub

  

Private Sub cmdBatal_Click()
    Call buka_Tutup(False)
    GridSupplier.Enabled = True
    Call refreshGrid
End Sub

Private Sub cmdEdit_Click()
'    isTambah = False
'    Call buka_Tutup(True)
'    Me.MousePointer = vbHourglass
'    If chModel.value = Unchecked And cmdSimpan.Enabled Then
'        Call isiModel
'    End If
'    If chAllDealer = Unchecked And cmdSimpan.Enabled Then
'        If rsDealer.State = adStateClosed Then
'            Call isiDealer
'        End If
'    End If
'    Me.MousePointer = vbDefault
End Sub

Private Sub cmdHapus_Click()
              
1        On Error GoTo cmdHapus_Click_Error
2         isBegin = False
3         If MsgBox("Apakah yakin menghapus data Dealer " & NulSID(rsGrid!fs_nm_dealer) & " ?", vbYesNo, "Konfirmasi Hapus Data") = vbYes Then
               
4             Call objHT.OpenTransactionHO
5             isBegin = True
6             erTranAll = True
7             sSQL = "UPDATE opLockNormal SET userDel = '" & gSA & "',tglDel = GETDATE(),flagdel = '1' WHERE kodeDealer = '" & NulSID(rsGrid!KodeDealer) & "' and flagdel = '0'"
8             Call objHT.InsertData_Ho(sSQL)
9             If erTranAll Then
10                Call objHT.CloseTransactionHO
11                MsgBox "Data berhasil dihapus", vbInformation
12                Call refreshGrid
13                Call buka_Tutup(True)
14            Else
15                Call objHT.CloseTransactionHO_err
16                MsgBox "Gagal hapus data", vbInformation
17                Exit Sub
18            End If
19        End If

20       On Error GoTo 0
21       Exit Sub

cmdHapus_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
22        varErlError = Erl
23        varErrNumber = err.Number
24        If isBegin Then
25            Call objHT.CloseTransactionHO_err
26        End If
27        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdHapus_Click of Form frmHOSettingSCCNormal")
End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

Private Sub cmdSimpan_Click()
          
1        On Error GoTo cmdSimpan_Click_Error
          Dim isBegin     As Boolean
          Dim boleh       As Boolean
          Dim pesan       As String
          Dim allDealer   As Boolean
2         allDealer = False
3         boleh = True
4         isBegin = False
          
5         rsDealer.filter = "TERLIBAT = " & True
6         If rsDealer.EOF Then
7             allDealer = True
8         End If
9         rsDealer.filter = ""
10        If boleh Then
11            Screen.MousePointer = vbHourglass
12            If allDealer Then
                  ''smua dealer
13                GoTo loopingDealer
14            Else
15                rsDealer.filter = "TERLIBAT = " & True
16                If Not rsDealer.EOF Then
17                    GoTo loopingDealer
18                Else
19                    rsDealer.filter = ""
20                End If
21            End If
              
              
loopingDealer:
22            Call objHT.OpenTransactionHO
23            isBegin = True
24            erTranAll = False
25            rsDealer.MoveFirst
26            Do While Not rsDealer.EOF
                  '' -------------------------------------
                  '' SENIN
                  '' -------------------------------------
27                Call ClearsSql
28                gsSql1(0) = "hari"
29                gsSql2(0) = "0SENIN"
30                gsSql1(1) = "kodeDealer"
31                gsSql2(1) = NulSID(rsDealer!KodeDealer)
32                gsSql1(2) = "jamMulai1"
33                gsSql2(2) = Right("0" & Left(NulSID(dtSJAwal1.value), Len(NulSID(dtSJAwal1.value)) - 3), 5)
34                gsSql1(3) = "jamSelesai1"
35                gsSql2(3) = Right("0" & Left(NulSID(dtSJAkhir1.value), Len(NulSID(dtSJAkhir1.value)) - 3), 5) 'NulSID(dtSJAkhir1.value)
36                gsSql1(4) = "jamMulai2"
37                gsSql2(4) = Right("0" & Left(NulSID(dtSJAwal2.value), Len(NulSID(dtSJAwal2.value)) - 3), 5) 'NulSID(dtSJAwal2.value)
38                gsSql1(5) = "jamSelesai2"
39                gsSql2(5) = Right("0" & Left(NulSID(dtSJAkhir2.value), Len(NulSID(dtSJAkhir2.value)) - 3), 5) 'NulSID(dtSJAkhir2.value)
40                gsSql1(6) = "UserEntry"
41                gsSql2(6) = NulSID(gSA)
42                gsSql1(7) = "tglEntry"
43                gsSql2(7) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
                  
44                Set rsTemp = objHT.FetchDataHO("SELECT * FROM opLockNormal WITH (NOLOCK) WHERE FLAGDEL = 0 AND HARI = '0SENIN' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'")
45                If Not rsTemp.EOF Then
46                    sSQL = GensSql2("opLockNormal", True) & " WHERE FLAGDEL = 0 AND HARI = '0SENIN' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'"
47                Else
48                    sSQL = GensSql2("opLockNormal", False)
49                End If
50                If erTranAll = True Then
51                    GoTo errorRollback
52                Else
53                    Call objHT.InsertData_Ho(sSQL)
54                End If
                  '' ---- AKHIR SENIN ----------------------------------
                  '' ---------------------------------------------------
                  '' -------------------------------------
                  '' SELASA
                  '' -------------------------------------
55                Call ClearsSql
56                gsSql1(0) = "hari"
57                gsSql2(0) = "1SELASA"
58                gsSql1(1) = "kodeDealer"
59                gsSql2(1) = NulSID(rsDealer!KodeDealer)
60                gsSql1(2) = "jamMulai1"
61                gsSql2(2) = Right("0" & Left(NulSID(dtSJAwal1.value), Len(NulSID(dtSJAwal1.value)) - 3), 5) 'NulSID(dtSJAwal1.value)
62                gsSql1(3) = "jamSelesai1"
63                gsSql2(3) = Right("0" & Left(NulSID(dtSJAkhir1.value), Len(NulSID(dtSJAkhir1.value)) - 3), 5) 'NulSID(dtSJAkhir1.value)
64                gsSql1(4) = "jamMulai2"
65                gsSql2(4) = Right("0" & Left(NulSID(dtSJAwal2.value), Len(NulSID(dtSJAwal2.value)) - 3), 5) 'NulSID(dtSJAwal2.value)
66                gsSql1(5) = "jamSelesai2"
67                gsSql2(5) = Right("0" & Left(NulSID(dtSJAkhir2.value), Len(NulSID(dtSJAkhir2.value)) - 3), 5) 'NulSID(dtSJAkhir2.value)
68                gsSql1(6) = "UserEntry"
69                gsSql2(6) = NulSID(gSA)
70                gsSql1(7) = "tglEntry"
71                gsSql2(7) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
                  
72                Set rsTemp = objHT.FetchDataHO("SELECT * FROM opLockNormal WITH (NOLOCK) WHERE FLAGDEL = 0 AND HARI = '1SELASA' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'")
73                If Not rsTemp.EOF Then
74                    sSQL = GensSql2("opLockNormal", True) & " WHERE FLAGDEL = 0 AND HARI = '1SELASA' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'"
75                Else
76                    sSQL = GensSql2("opLockNormal", False)
77                End If
78                If erTranAll = True Then
79                    GoTo errorRollback
80                Else
81                    Call objHT.InsertData_Ho(sSQL)
82                End If
                  '' ---- AKHIR SELASA ----------------------------------
                  '' ---------------------------------------------------
                  
                  '' -------------------------------------
                  '' RABU
                  '' -------------------------------------
83                Call ClearsSql
84                gsSql1(0) = "hari"
85                gsSql2(0) = "2RABU"
86                gsSql1(1) = "kodeDealer"
87                gsSql2(1) = NulSID(rsDealer!KodeDealer)
88                gsSql1(2) = "jamMulai1"
89                gsSql2(2) = Right("0" & Left(NulSID(dtSJAwal1.value), Len(NulSID(dtSJAwal1.value)) - 3), 5) 'NulSID(dtSJAwal1.value)
90                gsSql1(3) = "jamSelesai1"
91                gsSql2(3) = Right("0" & Left(NulSID(dtSJAkhir1.value), Len(NulSID(dtSJAkhir1.value)) - 3), 5) 'NulSID(dtSJAkhir1.value)
92                gsSql1(4) = "jamMulai2"
93                gsSql2(4) = Right("0" & Left(NulSID(dtSJAwal2.value), Len(NulSID(dtSJAwal2.value)) - 3), 5) 'NulSID(dtSJAwal2.value)
94                gsSql1(5) = "jamSelesai2"
95                gsSql2(5) = Right("0" & Left(NulSID(dtSJAkhir2.value), Len(NulSID(dtSJAkhir2.value)) - 3), 5) 'NulSID(dtSJAkhir2.value)
96                gsSql1(6) = "UserEntry"
97                gsSql2(6) = NulSID(gSA)
98                gsSql1(7) = "tglEntry"
99                gsSql2(7) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
                  
100               Set rsTemp = objHT.FetchDataHO("SELECT * FROM opLockNormal WITH (NOLOCK) WHERE FLAGDEL = 0 AND HARI = '2RABU' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'")
101               If Not rsTemp.EOF Then
102                   sSQL = GensSql2("opLockNormal", True) & " WHERE FLAGDEL = 0 AND HARI = '2RABU' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'"
103               Else
104                   sSQL = GensSql2("opLockNormal", False)
105               End If
106               If erTranAll = True Then
107                   GoTo errorRollback
108               Else
109                   Call objHT.InsertData_Ho(sSQL)
110               End If
                  '' ---- AKHIR RABU ----------------------------------
                  '' ---------------------------------------------------
                  
                  '' -------------------------------------
                  '' KAMIS
                  '' -------------------------------------
111               Call ClearsSql
112               gsSql1(0) = "hari"
113               gsSql2(0) = "3KAMIS"
114               gsSql1(1) = "kodeDealer"
115               gsSql2(1) = NulSID(rsDealer!KodeDealer)
116               gsSql1(2) = "jamMulai1"
117               gsSql2(2) = Right("0" & Left(NulSID(dtSJAwal1.value), Len(NulSID(dtSJAwal1.value)) - 3), 5) 'NulSID(dtSJAwal1.value)
118               gsSql1(3) = "jamSelesai1"
119               gsSql2(3) = Right("0" & Left(NulSID(dtSJAkhir1.value), Len(NulSID(dtSJAkhir1.value)) - 3), 5) 'NulSID(dtSJAkhir1.value)
120               gsSql1(4) = "jamMulai2"
121               gsSql2(4) = Right("0" & Left(NulSID(dtSJAwal2.value), Len(NulSID(dtSJAwal2.value)) - 3), 5) 'NulSID(dtSJAwal2.value)
122               gsSql1(5) = "jamSelesai2"
123               gsSql2(5) = Right("0" & Left(NulSID(dtSJAkhir2.value), Len(NulSID(dtSJAkhir2.value)) - 3), 5) 'NulSID(dtSJAkhir2.value)
124               gsSql1(6) = "UserEntry"
125               gsSql2(6) = NulSID(gSA)
126               gsSql1(7) = "tglEntry"
127               gsSql2(7) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
                  
128               Set rsTemp = objHT.FetchDataHO("SELECT * FROM opLockNormal WITH (NOLOCK) WHERE FLAGDEL = 0 AND HARI = '3KAMIS' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'")
129               If Not rsTemp.EOF Then
130                   sSQL = GensSql2("opLockNormal", True) & " WHERE FLAGDEL = 0 AND HARI = '3KAMIS' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'"
131               Else
132                   sSQL = GensSql2("opLockNormal", False)
133               End If
134               If erTranAll = True Then
135                   GoTo errorRollback
136               Else
137                   Call objHT.InsertData_Ho(sSQL)
138               End If
                  '' ---- AKHIR KAMIS ----------------------------------
                  '' ---------------------------------------------------
                  
                  '' -------------------------------------
                  '' JUMAT
                  '' -------------------------------------
139               Call ClearsSql
140               gsSql1(0) = "hari"
141               gsSql2(0) = "4JUMAT"
142               gsSql1(1) = "kodeDealer"
143               gsSql2(1) = NulSID(rsDealer!KodeDealer)
144               gsSql1(2) = "jamMulai1"
145               gsSql2(2) = Right("0" & Left(NulSID(dtSJAwal1.value), Len(NulSID(dtSJAwal1.value)) - 3), 5) 'NulSID(dtSJAwal1.value)
146               gsSql1(3) = "jamSelesai1"
147               gsSql2(3) = Right("0" & Left(NulSID(dtSJAkhir1.value), Len(NulSID(dtSJAkhir1.value)) - 3), 5) 'NulSID(dtSJAkhir1.value)
148               gsSql1(4) = "jamMulai2"
149               gsSql2(4) = Right("0" & Left(NulSID(dtSJAwal2.value), Len(NulSID(dtSJAwal2.value)) - 3), 5) 'NulSID(dtSJAwal2.value)
150               gsSql1(5) = "jamSelesai2"
151               gsSql2(5) = Right("0" & Left(NulSID(dtSJAkhir2.value), Len(NulSID(dtSJAkhir2.value)) - 3), 5) 'NulSID(dtSJAkhir2.value)
152               gsSql1(6) = "UserEntry"
153               gsSql2(6) = NulSID(gSA)
154               gsSql1(7) = "tglEntry"
155               gsSql2(7) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
                  
156               Set rsTemp = objHT.FetchDataHO("SELECT * FROM opLockNormal WITH (NOLOCK) WHERE FLAGDEL = 0 AND HARI = '4JUMAT' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'")
157               If Not rsTemp.EOF Then
158                   sSQL = GensSql2("opLockNormal", True) & " WHERE FLAGDEL = 0 AND HARI = '4JUMAT' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'"
159               Else
160                   sSQL = GensSql2("opLockNormal", False)
161               End If
162               If erTranAll = True Then
163                   GoTo errorRollback
164               Else
165                   Call objHT.InsertData_Ho(sSQL)
166               End If
                  '' ---- AKHIR JUMAT ----------------------------------
                  '' ---------------------------------------------------
                  
                  '' -------------------------------------
                  '' SABTU
                  '' -------------------------------------
167               Call ClearsSql
168               gsSql1(0) = "hari"
169               gsSql2(0) = "5SABTU"
170               gsSql1(1) = "kodeDealer"
171               gsSql2(1) = NulSID(rsDealer!KodeDealer)
172               gsSql1(2) = "jamMulai1"
173               gsSql2(2) = Right("0" & Left(NulSID(dtSAwal1.value), Len(NulSID(dtSAwal1.value)) - 3), 5) 'NulSID(dtSAwal1.value)
174               gsSql1(3) = "jamSelesai1"
175               gsSql2(3) = Right("0" & Left(NulSID(dtSAkhir1.value), Len(NulSID(dtSAkhir1.value)) - 3), 5) 'NulSID(dtSAkhir1.value)
176               gsSql1(4) = "jamMulai2"
177               gsSql2(4) = Right("0" & Left(NulSID(dtSAwal2.value), Len(NulSID(dtSAwal2.value)) - 3), 5) 'NulSID(dtSAwal2.value)
178               gsSql1(5) = "jamSelesai2"
179               gsSql2(5) = Right("0" & Left(NulSID(dtSAkhir2.value), Len(NulSID(dtSAkhir2.value)) - 3), 5) 'NulSID(dtSAkhir2.value)
180               gsSql1(6) = "UserEntry"
181               gsSql2(6) = NulSID(gSA)
182               gsSql1(7) = "tglEntry"
183               gsSql2(7) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
                  
184               Set rsTemp = objHT.FetchDataHO("SELECT * FROM opLockNormal WITH (NOLOCK) WHERE FLAGDEL = 0 AND HARI = '5SABTU' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'")
185               If Not rsTemp.EOF Then
186                   sSQL = GensSql2("opLockNormal", True) & " WHERE FLAGDEL = 0 AND HARI = '5SABTU' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'"
187               Else
188                   sSQL = GensSql2("opLockNormal", False)
189               End If
190               If erTranAll = True Then
191                   GoTo errorRollback
192               Else
193                   Call objHT.InsertData_Ho(sSQL)
194               End If
                  '' ---- AKHIR SABTU ----------------------------------
                  '' ---------------------------------------------------
                  
                  '' -------------------------------------
                  '' MINGGU
                  '' -------------------------------------
195               Call ClearsSql
196               gsSql1(0) = "hari"
197               gsSql2(0) = "6MINGGU"
198               gsSql1(1) = "kodeDealer"
199               gsSql2(1) = NulSID(rsDealer!KodeDealer)
200               gsSql1(2) = "jamMulai1"
201               gsSql2(2) = Right("0" & Left(NulSID(dtMAwal1.value), Len(NulSID(dtMAwal1.value)) - 3), 5) 'NulSID(dtMAwal1.value)
202               gsSql1(3) = "jamSelesai1"
203               gsSql2(3) = Right("0" & Left(NulSID(dtMAkhir1.value), Len(NulSID(dtMAkhir1.value)) - 3), 5) 'NulSID(dtMAkhir1.value)
204               gsSql1(4) = "jamMulai2"
205               gsSql2(4) = Right("0" & Left(NulSID(dtMAwal2.value), Len(NulSID(dtMAwal2.value)) - 3), 5) 'NulSID(dtMAwal2.value)
206               gsSql1(5) = "jamSelesai2"
207               gsSql2(5) = Right("0" & Left(NulSID(dtMAkhir2.value), Len(NulSID(dtMAkhir2.value)) - 3), 5) 'NulSID(dtMAkhir2.value)
208               gsSql1(6) = "UserEntry"
209               gsSql2(6) = NulSID(gSA)
210               gsSql1(7) = "tglEntry"
211               gsSql2(7) = Format(isnow, "yyyy-mm-dd hh:mm:ss")
                  
212               Set rsTemp = objHT.FetchDataHO("SELECT * FROM opLockNormal WITH (NOLOCK) WHERE FLAGDEL = 0 AND HARI = '6MINGGU' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'")
213               If Not rsTemp.EOF Then
214                   sSQL = GensSql2("opLockNormal", True) & " WHERE FLAGDEL = 0 AND HARI = '6MINGGU' AND kodeDealer = '" & NulSID(rsDealer!KodeDealer) & "'"
215               Else
216                   sSQL = GensSql2("opLockNormal", False)
217               End If
218               If erTranAll = True Then
219                   GoTo errorRollback
220               Else
221                   Call objHT.InsertData_Ho(sSQL)
222               End If
                  '' ---- AKHIR MINGGU ----------------------------------
                  '' ---------------------------------------------------
223               rsDealer.MoveNext
224           Loop
               
              
225           If erTranAll = False Then
226               Call objHT.CloseTransactionHO
227               Screen.MousePointer = vbDefault
228               MsgBox "Berhasil menyimpan data.", vbInformation
229               Call ber_sih
230               Call buka_Tutup(True)
231               Call refreshGrid
232           Else
errorRollback:
233               Call objHT.CloseTransactionHO_err
234               Screen.MousePointer = vbDefault
235               MsgBox "Gagal menyimpan data.", vbInformation
236               Exit Sub
237           End If
238       Else
239           Screen.MousePointer = vbDefault
240           MsgBox "Silahkan lengkapi data sebagai berikut : " & vbCrLf & pesan, vbInformation
241       End If

242      On Error GoTo 0
243      Exit Sub

cmdSimpan_Click_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
244       varErlError = Erl
245       varErrNumber = err.Number
246       Screen.MousePointer = vbDefault
247       If isBegin = True Then
248           Call objHT.CloseTransactionHO_err
249       End If
250       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure cmdSimpan_Click of Form frmHOSettingSCCNormal")
End Sub

Private Sub cmdTambah_Click()
    isTambah = True
    Call ber_sih
    Call buka_Tutup(True)
    GridSupplier.Enabled = False
End Sub

 
Private Sub cmdTutupDealer_Click()
    frmeDealer.Visible = False
    
End Sub

Private Sub Combo1_Click()
    frmeDealer.Visible = True
    frmeDealer.Top = 600
    If rsDealer.State = adStateClosed Then
        Call isiDealer
    End If
    frmeDealer.Visible = True
    GridDealer.SetFocus
End Sub

Private Sub Combo1_DropDown()
    Call Combo1_Click
End Sub

Private Sub Form_Load()
    Call ber_sih
    Call buka_Tutup(True)
    Call refreshGrid
    Call isiDealer
End Sub
 
Private Sub GridDealer_HeadClick(ByVal ColIndex As Integer)
    On Error GoTo err
   
    If ColIndex = 2 Then
        If rsDealer.State Then
            rsDealer.MoveFirst
            Do While Not rsDealer.EOF
                rsDealer!terlibat = IIf(rsDealer!terlibat = True, False, True)
                rsDealer.MoveNext
            Loop
            rsDealer.MoveFirst
        End If
    Else
        Urut = IIf(Urut = "DESC", "ASC", "DESC")
        If ColIndex <= 1 Then rsDealer.Sort = "[" & GridDealer.Columns(ColIndex).DataField & "]" & " " & Urut
        GridDealer.SetFocus
        Sendkeys "{Right " & Val(ColIndex) & "}"
    End If
    Exit Sub
err:
   MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure HeadClick ")

End Sub

 

Private Sub GridSupplier_FilterChange()
    On Error GoTo errHandler
    Set cols = GridSupplier.Columns
    Dim C As Integer
    
    C = GridSupplier.col
    GridSupplier.HoldFields
    rsGrid.filter = getFilter()
    GridSupplier.col = C
    GridSupplier.EditActive = True
    Exit Sub
errHandler:
    MsgBox err.Source & ":" & vbCrLf & err.Description
End Sub

Private Sub GridSupplier_HeadClick(ByVal ColIndex As Integer)
          
1        On Error GoTo GridSupplier_HeadClick_Error
2         Urut = IIf(Urut = "DESC", "ASC", "DESC")
3         rsGrid.Sort = GridSupplier.Columns(ColIndex).DataField & " " & Urut
4         GridSupplier.SetFocus
5         Sendkeys "{Right " & Val(ColIndex) & "}"
          

6        On Error GoTo 0
7        Exit Sub

GridSupplier_HeadClick_Error:
          '' 20181002REV2 ' MENCATAT ERROR LINE NUMBER SEBELUM DITAMPILKAN
8         varErlError = Erl
9         varErrNumber = err.Number
10        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(varErlError, varErrNumber, " (" & err.Description & ") in procedure GridSupplier_HeadClick of Form frmHOSettingSCCNormal")
End Sub

Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
    Set rsTmp = objO.FetchDataHO("SELECT A.fs_kd_dealer, A.fs_nm_dealer, '' AS KODE FROM t_dealer a WITH (NOLOCK) WHERE A.INTERNAL = 1 ORDER BY a.fs_kd_dealer")
    If Not rsTmp.EOF Then
        Call createRS
        Do While Not rsTmp.EOF
            rsDealer.AddNew
            rsDealer!KodeDealer = rsTmp("FS_KD_DEALER")
            rsDealer!namaDealer = rsTmp("FS_NM_DEALER")
            If rsTmp("KODE") <> "" Then
                rsDealer!terlibat = True
            Else
                rsDealer!terlibat = False
            End If
            rsDealer.Update
            rsTmp.MoveNext
        Loop
        
        GridDealer.DataSource = rsDealer
        GridDealer.EvenRowStyle.BackColor = vbYellow
        GridDealer.OddRowStyle.BackColor = vbWhite
        GridDealer.AlternatingRowStyle = True
    End If
End Sub

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 3500
    GridDealer.Columns(2).Width = 1000
    GridDealer.Columns(2).Alignment = dbgCenter
End Sub

Private Sub createRS()
    
    If rsDealer.State Then rsDealer.Close
    Set rsDealer = Nothing
    With rsDealer
        .Fields.Append "KODEDEALER", adVarChar, 15
        .Fields.Append "NAMADEALER", adVarChar, 25
        .Fields.Append "TERLIBAT", adBoolean
        .Open
    End With
    
    Set GridDealer.DataSource = rsDealer
End Sub

 
Private Sub createTbl()
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    With rsGridMdl
        .Fields.Append "GRPMDL", adVarChar, 15
        .Fields.Append "KODEMDL", adVarChar, 20
        .Fields.Append "NAMAMDL", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridModel.DataSource = rsGridMdl
    GridModel.AlternatingRowStyle = True
End Sub

Private Sub refreshGrid()
'    gridSupplier.ClearFields
    sSQL = "select b.FS_NM_DEALER,a.hari,a.kodedealer,a.jammulai1,a.jamselesai1,a.jammulai2,a.jamselesai2 from opLockNormal a inner join T_DEALER b on a.kodedealer = b.FS_KD_DEALER and b.FS_LOG_DEL = '' where a.flagdel = 0"
    Set rsGrid = objO.FetchDataHO(sSQL)
    Set GridSupplier.DataSource = rsGrid
End Sub

Private Sub hargaBeli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub hargaJual_KeyPress(KeyAscii As Integer)
    If KeyAscii = 45 Then
        KeyAscii = 0
    End If
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
