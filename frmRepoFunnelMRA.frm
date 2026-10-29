VERSION 5.00
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{EE757A1F-B0AC-40BC-9E72-B8651740F53E}#1.0#0"; "ARProgBar.OCX"
Begin VB.Form frmRepoFunnelMRA 
   Caption         =   "Laporan Funnel MRA"
   ClientHeight    =   9075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10950
   ScaleWidth      =   15120
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   0
      Picture         =   "frmRepoFunnelMRA.frx":0000
      ScaleHeight     =   10455
      ScaleWidth      =   450
      TabIndex        =   30
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
         TabIndex        =   27
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
            TabIndex        =   28
            Top             =   0
            Width           =   1050
         End
         Begin TrueOleDBGrid70.TDBGrid gridCustomer 
            Height          =   6375
            Left            =   120
            TabIndex        =   29
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
            Columns(2).Caption=   "Alamat"
            Columns(2).DataField=   "AlamatCust"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "No Polisi"
            Columns(3).DataField=   "nopolisi"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   4
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AlternatingRowStyle=   -1  'True
            Splits(0).DividerColor=   14933984
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=4"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=5292"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=5212"
            Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=5292"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5212"
            Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(11)=   "Column(2).Width=5292"
            Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=5212"
            Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
            Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
            Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
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
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
            _StyleDefs(52)  =   "Named:id=33:Normal"
            _StyleDefs(53)  =   ":id=33,.parent=0"
            _StyleDefs(54)  =   "Named:id=34:Heading"
            _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(56)  =   ":id=34,.wraptext=-1"
            _StyleDefs(57)  =   "Named:id=35:Footing"
            _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(59)  =   "Named:id=36:Selected"
            _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(61)  =   "Named:id=37:Caption"
            _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(63)  =   "Named:id=38:HighlightRow"
            _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(65)  =   "Named:id=39:EvenRow"
            _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&H4080FF&"
            _StyleDefs(67)  =   "Named:id=40:OddRow"
            _StyleDefs(68)  =   ":id=40,.parent=33"
            _StyleDefs(69)  =   "Named:id=41:RecordSelector"
            _StyleDefs(70)  =   ":id=41,.parent=34"
            _StyleDefs(71)  =   "Named:id=42:FilterBar"
            _StyleDefs(72)  =   ":id=42,.parent=33"
            _StyleDefs(73)  =   "Named:id=131:TipeTanpaTipe"
            _StyleDefs(74)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
            _StyleDefs(75)  =   "Named:id=132:TigaDigitKosong"
            _StyleDefs(76)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         End
      End
      Begin ARProgBarCtrl.ARProgressBar prog 
         Height          =   495
         Left            =   480
         TabIndex        =   20
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
         TabIndex        =   14
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
            TabIndex        =   19
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   16
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FF0000&
            Caption         =   "Data Reminder By MRA"
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
            TabIndex        =   15
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.PictureBox grafikHasil 
         Height          =   6015
         Left            =   480
         Picture         =   "frmRepoFunnelMRA.frx":931E
         ScaleHeight     =   5955
         ScaleWidth      =   6765
         TabIndex        =   6
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   7
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
         Format          =   98041857
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
         Format          =   98041857
         CurrentDate     =   39852
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   285
         Index           =   14
         Left            =   1965
         TabIndex        =   3
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
         Left            =   1800
         TabIndex        =   8
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
         MICON           =   "frmRepoFunnelMRA.frx":14087
         PICN            =   "frmRepoFunnelMRA.frx":140A3
         PICH            =   "frmRepoFunnelMRA.frx":144FA
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
         Left            =   3120
         TabIndex        =   9
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
         MICON           =   "frmRepoFunnelMRA.frx":149E4
         PICN            =   "frmRepoFunnelMRA.frx":14A00
         PICH            =   "frmRepoFunnelMRA.frx":14ED4
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
         TabIndex        =   31
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "frmRepoFunnelMRA.frx":1541F
         Caption         =   "frmRepoFunnelMRA.frx":1554B
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoFunnelMRA.frx":155B7
         Keys            =   "frmRepoFunnelMRA.frx":155D5
         Spin            =   "frmRepoFunnelMRA.frx":15633
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
         TabIndex        =   32
         Top             =   480
         Width           =   1455
         _Version        =   65536
         _ExtentX        =   2566
         _ExtentY        =   556
         Calendar        =   "frmRepoFunnelMRA.frx":1565B
         Caption         =   "frmRepoFunnelMRA.frx":15787
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmRepoFunnelMRA.frx":157F3
         Keys            =   "frmRepoFunnelMRA.frx":15811
         Spin            =   "frmRepoFunnelMRA.frx":1586F
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   510
         Width           =   1005
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000080FF&
         X1              =   480
         X2              =   8280
         Y1              =   3000
         Y2              =   3000
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
         TabIndex        =   5
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
         TabIndex        =   4
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
Attribute VB_Name = "frmRepoFunnelMRA"
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
    If IsNumeric(NulSID(lblBook.Caption)) Then
        If CInt(NulSID(lblBook.Caption)) > 0 Then
            frameDetailCustomer.Visible = True
            sSQL = "select b.customer as nama,a.norangka as NoRangka,b.CUST_ADRES aS AlamatCust,b.NO_POLISI as NoPolisi"
            sSQL = sSQL & " from M_ListReminderCall a with (nolock)"
            sSQL = sSQL & " inner join BOOK b with (nolock) on a.NoRangka = b.NO_CHASSIS"
            sSQL = sSQL & " left join CUSTDAT c with (nolock) on a.NoRangka = c.NO_CHASSIS"
            sSQL = sSQL & " where  K_Telp = 1 AND DATEDIFF(d,TGLFOLLOW, DATE_BOOK) = 0 and ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
            sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
            sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
            sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
            sSQL = sSQL & " )"
            Set rsCustomer = objO.FetchData(sSQL)
            gridCustomer.DataSource = rsCustomer
        End If
    End If
End Sub

Private Sub lblReminder_Click()
    If IsNumeric(NulSID(lblReminder.Caption)) Then
        If CInt(NulSID(lblReminder.Caption)) > 0 Then
            frameDetailCustomer.Visible = True
            sSQL = "select distinct nama,norangka,alamatcust,nopolisi from ( "
            sSQL = sSQL & "SELECT NamaCust as nama,NoRangka,AlamatCust,NoPolisi"
            sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
            sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
            sSQL = sSQL & " UNION "
            sSQL = sSQL & " Select DISTINCT b.SNama AS nama,b.NoRangka as NoRangka,b.SAlamat as AlamatCust,a.NoPol AS NoPolisi"
            sSQL = sSQL & " From [SPK" & Right(gDb, 2) & "]..AFI a WITH(NOLOCK) INNER JOIN  [SPK" & Right(gDb, 2) & "]..SPK b WITH(NOLOCK) ON a.SPK_No = b.SPK_No"
            sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..INMerek c WITH(NOLOCK) ON b.Kode_Merk = c.KodeMerek"
            sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..InKategory d WITH(NOLOCK) ON b.Kode_Type = d.KodeKategory"
            sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..InWarnaExt e WITH(NOLOCK) ON b.Kode_Warna = e.KodeWarna"
            sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..BPK f WITH(NOLOCK) ON a.SPK_No = f.SPK_No"
            sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..ArSales g WITH(NOLOCK) ON b.Kode_Sales = g.KodeSales"
            sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..INMst h WITH(NOLOCK) ON b.NoRangka = h.kodeItem"
            sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..T_MRS i WITH(NOLOCK) ON b.SPK_No = i.NoSPK"
            sSQL = sSQL & " LEFT JOIN [BENGKEL" & Right(gDb, 2) & "]..History y WITH(NOLOCK) ON b.NoRangka = y.No_Chassis And b.norangka not in (SELECT NO_CHASSIS FROM [BENGKEL" & Right(gDb, 2) & "]..HISTORY WHERE FS_RC = 'R001')"
            sSQL = sSQL & " LEFT JOIN [BENGKEL" & Right(gDb, 2) & "]..M_ListReminderCall l WITH(NOLOCK) ON b.NoRangka = l.NoRangka AND K_Ulang = 1 AND KMNEXTSVC = '1000'"
            sSQL = sSQL & " Where i.Km1 = '1000' And a.batal = 0 AND"
            sSQL = sSQL & " (i.Tgl1 between '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "') OR (l.Tanggal between '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "')"
            sSQL = sSQL & " ) A "
            
            Set rsCustomer = objO.FetchData(sSQL)
            gridCustomer.DataSource = rsCustomer
        End If
    End If
End Sub

Private Sub lblShow_Click()
    If IsNumeric(NulSID(lblShow.Caption)) Then
        If CInt(NulSID(lblShow.Caption)) > 0 Then
            frameDetailCustomer.Visible = True
            sSQL = "select c.customer as nama,a.norangka as NoRangka,c.CUST_ADRES aS AlamatCust,c.NO_POLISI as NoPolisi"
            sSQL = sSQL & " from M_ListReminderCall a with (nolock)"
            sSQL = sSQL & " inner join BOOK b with (nolock) on a.NoRangka = b.NO_CHASSIS and b.FS_LOG_DEL = ''"
            sSQL = sSQL & " inner join JOBDTA c with (nolock) on c.NOFAK = b.NOFAK and c.NO_CHASSIS = b.NO_CHASSIS and c.FS_LOG_DEL = ''"
            sSQL = sSQL & " left join CUSTDAT d with (nolock) on a.NoRangka = d.NO_CHASSIS"
            sSQL = sSQL & " where  K_Telp = 1 AND DATEDIFF(d,TGLFOLLOW, DATE_BOOK) = 0 and ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
            sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
            sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
            sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
            sSQL = sSQL & " )            "
            Set rsCustomer = objO.FetchData(sSQL)
            gridCustomer.DataSource = rsCustomer
        End If
    End If
End Sub

Private Sub lblTerfollow_Click()
    If IsNumeric(NulSID(lblTerfollow.Caption)) Then
        If CInt(NulSID(lblTerfollow.Caption)) > 0 Then
            frameDetailCustomer.Visible = True
            sSQL = "select distinct b.CUSTOMER as nama,a.Terfollowup as NoRangka,b.CUST_ADRES aS AlamatCust,b.NO_POLISI as NoPolisi"
            sSQL = sSQL & " from ("
            sSQL = sSQL & " select norangka Terfollowup"
            sSQL = sSQL & " from M_ListReminderCall with (nolock)"
            sSQL = sSQL & " where ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
            sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
            sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
            sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
            sSQL = sSQL & " )"
            sSQL = sSQL & " Union"
            sSQL = sSQL & " select norangka Terfollowup"
            sSQL = sSQL & " from M_ListReminder with (nolock)"
            sSQL = sSQL & " where ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
            sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
            sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
            sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
            sSQL = sSQL & " )) a"
            sSQL = sSQL & " left join CUSTDAT b on a.Terfollowup = b.NO_CHASSIS"

            Set rsCustomer = objO.FetchData(sSQL)
            gridCustomer.DataSource = rsCustomer
        End If
    End If
End Sub

Private Sub lblTerhubung_Click()
    If IsNumeric(NulSID(lblTerhubung.Caption)) Then
        If CInt(NulSID(lblTerhubung.Caption)) > 0 Then
            frameDetailCustomer.Visible = True
            sSQL = "select b.customer as nama,a.norangka as NoRangka,b.CUST_ADRES aS AlamatCust,b.NO_POLISI as NoPolisi"
            sSQL = sSQL & " from M_ListReminderCall a with (nolock)"
            sSQL = sSQL & " left join CUSTDAT b with (nolock) on a.NoRangka = b.NO_CHASSIS"
            sSQL = sSQL & " where K_Telp = 1 and ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
            sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
            sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
            sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
            sSQL = sSQL & " )"
            Set rsCustomer = objO.FetchData(sSQL)
            gridCustomer.DataSource = rsCustomer
        End If
    End If
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Sub tblSimpan_Click()
          Dim reminder, terfollow, terhubung, boc, show As Integer
          
1         If DateDiff("D", Format(Tgl1, "yyyy-mm-dd"), Format(Tgl2, "yyyy-mm-dd")) < 0 Then
2             MsgBox "Range tanggal tidak valid", vbInformation
3             Exit Sub
4         End If
          
5         prog.Visible = True
6         prog.Max = 7
7         grafikHasil.Visible = True
          
          ''-- DATA REMINDER MRA
8         sSQL = "SELECT count(NoRangka) Reminder from ( "
9         sSQL = sSQL & " SELECT NORANGKA "
10        sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
11        sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
12        sSQL = sSQL & " UNION "
13        sSQL = sSQL & " Select DISTINCT  b.NoRangka as NoRangka "
14        sSQL = sSQL & " From [SPK" & Right(gDb, 2) & "]..AFI a WITH(NOLOCK) INNER JOIN  [SPK" & Right(gDb, 2) & "]..SPK b WITH(NOLOCK) ON a.SPK_No = b.SPK_No"
15        sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..INMerek c WITH(NOLOCK) ON b.Kode_Merk = c.KodeMerek"
16        sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..InKategory d WITH(NOLOCK) ON b.Kode_Type = d.KodeKategory"
17        sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..InWarnaExt e WITH(NOLOCK) ON b.Kode_Warna = e.KodeWarna"
18        sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..BPK f WITH(NOLOCK) ON a.SPK_No = f.SPK_No"
19        sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..ArSales g WITH(NOLOCK) ON b.Kode_Sales = g.KodeSales"
20        sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..INMst h WITH(NOLOCK) ON b.NoRangka = h.kodeItem"
21        sSQL = sSQL & " INNER JOIN [SPK" & Right(gDb, 2) & "]..T_MRS i WITH(NOLOCK) ON b.SPK_No = i.NoSPK"
22        sSQL = sSQL & " LEFT JOIN [BENGKEL" & Right(gDb, 2) & "]..History y WITH(NOLOCK) ON b.NoRangka = y.No_Chassis And b.norangka not in (SELECT NO_CHASSIS FROM [BENGKEL" & Right(gDb, 2) & "]..HISTORY WHERE FS_RC = 'R001')"
23        sSQL = sSQL & " LEFT JOIN [BENGKEL" & Right(gDb, 2) & "]..M_ListReminderCall l WITH(NOLOCK) ON b.NoRangka = l.NoRangka AND K_Ulang = 1 AND KMNEXTSVC = '1000'"
24        sSQL = sSQL & " Where i.Km1 = '1000' And a.batal = 0 AND"
25        sSQL = sSQL & " (i.Tgl1 between '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "') OR (l.Tanggal between '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "')"
26        sSQL = sSQL & " ) A "
                  
27        prog.value = 1
28        DoEvents
29        prog.Caption = "Ambil data Reminder MRA"
30        reminder = objO.FetchData_Double(sSQL)
31        lblReminder.Caption = reminder
32        DoEvents
          
          ''-- DATA REMINDER TERFOLLOW2 MRA
33        sSQL = "select count(Terfollowup) Terfollowup from ("
34        sSQL = sSQL & " select norangka Terfollowup"
35        sSQL = sSQL & " from M_ListReminderCall with (nolock)"
36        sSQL = sSQL & " where ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
37        sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
38        sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
39        sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
40        sSQL = sSQL & " )"
41        sSQL = sSQL & " Union"
42        sSQL = sSQL & " select norangka Terfollowup"
43        sSQL = sSQL & " from M_ListReminder with (nolock)"
44        sSQL = sSQL & " where ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
45        sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
46        sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
47        sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 8, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 8, Tgl2.value), "YYYY-MM-DD") & "'"
48        sSQL = sSQL & " )) a"
           
49        prog.value = 2
50        DoEvents
51        prog.Caption = "Ambil data Terfollow MRA"
52        terfollow = objO.FetchData_Double(sSQL)
53        If reminder = 0 Then
54            lblPersenFollow.Caption = "0 % "
55        ElseIf terfollow > reminder Then
56            lblPersenFollow.Caption = "100% "
57        Else
58            lblPersenFollow.Caption = Round((terfollow / reminder) * 100, 0) & "% "
59        End If
60        lblTerfollow.Caption = terfollow
61        DoEvents
           
          ''-- DATA REMINDER TERHUBUNG MRA
62        sSQL = "select COUNT(norangka) Terhubung"
63        sSQL = sSQL & " from M_ListReminderCall with (nolock)"
64        sSQL = sSQL & " where K_Telp = 1 and ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
65        sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
66        sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
67        sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
68        sSQL = sSQL & " )"
          
69        prog.value = 3
70        DoEvents
71        prog.Caption = "Ambil data Terhubung MRA"
          'reminder = objO.FetchData_Double(sSQL)
72        terhubung = objO.FetchData_Double(sSQL)
73        If terhubung = 0 Then
74            lblPersenTerhubung.Caption = "0 % "
75        ElseIf terhubung > terfollow Then
76            lblPersenTerhubung.Caption = "100% "
77        Else
78            lblPersenTerhubung.Caption = Round((terhubung / terfollow) * 100, 0) & "% "
79        End If
80        lblTerhubung.Caption = terhubung
81        DoEvents
          
          ''-- DATA REMINDER TERHUBUNG DAN BOC MRA
82        sSQL = "select COUNT(norangka) book"
83        sSQL = sSQL & " from M_ListReminderCall a with (nolock)"
84        sSQL = sSQL & " inner join BOOK b with (nolock) on a.NoRangka = b.NO_CHASSIS AND b.fs_log_del = ''"
85        sSQL = sSQL & " where  K_Telp = 1 AND DATEDIFF(d,TGLFOLLOW, DATE_BOOK) = 0 and ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
86        sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
87        sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
88        sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
89        sSQL = sSQL & " )"

90        prog.value = 4
91        DoEvents
92        prog.Caption = "Ambil data BOC MRA"
          'reminder = objO.FetchData_Double(sSQL)
93        boc = objO.FetchData_Double(sSQL)
94        If terhubung = 0 Then
95            lblPersenBooking.Caption = "0 % "
96        ElseIf boc > terhubung Then
97            lblPersenBooking.Caption = "100% "
98        Else
99            lblPersenBooking.Caption = Round((boc / terhubung) * 100, 0) & "% "
100       End If
101       lblBook.Caption = boc
102       DoEvents
          
          ''- DATA REMINDER TERHUBUNG DAN BOC MRA SHOW BENGKEL
103       sSQL = "select COUNT(norangka) wo"
104       sSQL = sSQL & " from M_ListReminderCall a with (nolock)"
105       sSQL = sSQL & " inner join BOOK b with (nolock) on a.NoRangka = b.NO_CHASSIS and b.FS_LOG_DEL = ''"
106       sSQL = sSQL & " inner join JOBDTA c with (nolock) on c.NOFAK = b.NOFAK and c.NO_CHASSIS = b.NO_CHASSIS and c.FS_LOG_DEL = ''"
107       sSQL = sSQL & " where  K_Telp = 1 AND DATEDIFF(d,TGLFOLLOW, DATE_BOOK) = 0 and ltrim(rtrim(NoRangka))+convert(varchar(20),KmNextSvc) in ("
108       sSQL = sSQL & " SELECT ltrim(rtrim(NoRangka))+convert(varchar(20),KmNext)"
109       sSQL = sSQL & " from M_ESTIMASI a with (nolock)"
110       sSQL = sSQL & " where (CASE WHEN A.JT_Tgl > A.JT_Bln then  A.JT_Bln WHEN A.JT_Tgl < A.JT_Bln then  A.JT_Tgl WHEN A.JT_Tgl = A.JT_Bln then  A.Jt_Bln END) BETWEEN '" & Format(DateAdd("D", 7, Tgl1.value), "YYYY-MM-DD") & "' and '" & Format(DateAdd("D", 7, Tgl2.value), "YYYY-MM-DD") & "'"
111       sSQL = sSQL & " )"
          
112       prog.value = 5
113       DoEvents
114       prog.Caption = "Ambil data Show Customer MRA"
          'reminder = objO.FetchData_Double(sSQL)
115       show = objO.FetchData_Double(sSQL)
116       If boc = 0 Then
117           lblPersenShow.Caption = "0 % "
118       ElseIf show > boc Then
119           lblPersenShow.Caption = "100% "
120       Else
121           lblPersenShow.Caption = Round((show / boc) * 100, 0) & "% "
122       End If
123       lblShow.Caption = show
124       DoEvents
          
125       prog.Visible = False
126       Frame2.Visible = True
127       grafikHasil.Visible = True
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
