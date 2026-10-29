VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{A49CE0E0-C0F9-11D2-B0EA-00A024695830}#1.0#0"; "tidate6.ocx"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{FBC672E3-F04D-11D2-AFA5-E82C878FD532}#5.0#0"; "AS-IFce1.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Object = "{4A4AA691-3E6F-11D2-822F-00104B9E07A1}#3.0#0"; "ssdw3bo.ocx"
Begin VB.Form frmHOMappingCustomer 
   BackColor       =   &H00FFFFFF&
   Caption         =   ".:: Mapping Customer HO ::."
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14850
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frameKolom 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Kolom"
      Height          =   5655
      Left            =   8640
      TabIndex        =   149
      Top             =   2520
      Width           =   4335
      Begin VB.CheckBox chCentangSemua 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Centang Semua"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   178
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alamat SPK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   21
         Left            =   2160
         TabIndex        =   175
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama SPK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   20
         Left            =   2160
         TabIndex        =   174
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cara Bayar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   19
         Left            =   2160
         TabIndex        =   173
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wiraniaga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   18
         Left            =   2160
         TabIndex        =   172
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kategori Aktif / Pasif"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   2160
         TabIndex        =   171
         Top             =   1680
         Width           =   2055
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dealer Terakhir SV"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   2160
         TabIndex        =   170
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Dealer Penjual"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   2160
         TabIndex        =   169
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Next Service"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   2160
         TabIndex        =   168
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "History SBI / SBE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   2160
         TabIndex        =   167
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usename"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   120
         TabIndex        =   166
         Top             =   4560
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Telp"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   120
         TabIndex        =   165
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alamat Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   120
         TabIndex        =   164
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Model No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   163
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Retail Sales Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   120
         TabIndex        =   162
         Top             =   3120
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delivery Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   120
         TabIndex        =   161
         Top             =   2760
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tahun"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   160
         Top             =   2400
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Rangka"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   159
         Top             =   2040
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Polisi"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   158
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Model"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   157
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   156
         Top             =   960
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Last SVC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   155
         Top             =   600
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Last Reminder"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   154
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "No Telp SPK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   22
         Left            =   2160
         TabIndex        =   153
         Top             =   3480
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nama STNK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   23
         Left            =   2160
         TabIndex        =   152
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Alamat STNK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   24
         Left            =   2160
         TabIndex        =   151
         Top             =   4200
         Width           =   1815
      End
      Begin VB.CheckBox kolom 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VOC CR7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   25
         Left            =   2160
         TabIndex        =   150
         Top             =   4560
         Width           =   1815
      End
      Begin AIFCmp1.asxPowerButton cmdTutupKolom 
         Height          =   360
         Left            =   0
         TabIndex        =   176
         Top             =   5280
         Width           =   4335
         _ExtentX        =   7646
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
   Begin VB.Frame frmeModel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Model Kendaraan"
      Height          =   3615
      Left            =   8640
      TabIndex        =   139
      Top             =   1920
      Visible         =   0   'False
      Width           =   9135
      Begin TrueOleDBGrid70.TDBGrid GridModel 
         Height          =   2895
         Left            =   120
         TabIndex        =   140
         Top             =   360
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   5106
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "GROUP"
         Columns(0).DataField=   "FS_GRP_MDL"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   4
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "OK"
         Columns(1).DataField=   "IKUT"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   2
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   688
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   14215660
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=2"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
         Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
         Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=112,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Named:id=33:Normal"
         _StyleDefs(45)  =   ":id=33,.parent=0"
         _StyleDefs(46)  =   "Named:id=34:Heading"
         _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(48)  =   ":id=34,.wraptext=-1"
         _StyleDefs(49)  =   "Named:id=35:Footing"
         _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(51)  =   "Named:id=36:Selected"
         _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(53)  =   "Named:id=37:Caption"
         _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(55)  =   "Named:id=38:HighlightRow"
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin AIFCmp1.asxPowerButton cmdTutupModel 
         Height          =   360
         Left            =   120
         TabIndex        =   141
         Top             =   3240
         Width           =   8895
         _ExtentX        =   15690
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
   Begin VB.Frame frameDealerNon 
      Caption         =   "Daftar Dealer Non NCM"
      Height          =   3855
      Left            =   8640
      TabIndex        =   146
      Top             =   1560
      Visible         =   0   'False
      Width           =   6615
      Begin TrueOleDBGrid70.TDBGrid gridDealerNon 
         Height          =   3015
         Left            =   120
         TabIndex        =   147
         Top             =   360
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
      Begin AIFCmp1.asxPowerButton cmdTutupNon 
         Height          =   360
         Left            =   120
         TabIndex        =   148
         Top             =   3360
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
   Begin VB.Frame frmeDealer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Daftar Dealer NCM"
      Height          =   3735
      Left            =   8640
      TabIndex        =   123
      Top             =   1560
      Visible         =   0   'False
      Width           =   6615
      Begin TrueOleDBGrid70.TDBGrid GridDealer 
         Height          =   3015
         Left            =   120
         TabIndex        =   124
         Top             =   360
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
         TabIndex        =   125
         Top             =   3360
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
   Begin VB.CheckBox ckKolom 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Semua Kolom"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   177
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox chNonNCM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Non NCM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9600
      TabIndex        =   145
      Top             =   1560
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.CheckBox ckNCM 
      BackColor       =   &H00FFFFFF&
      Caption         =   "NCM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8790
      TabIndex        =   144
      Top             =   1560
      Value           =   1  'Checked
      Width           =   735
   End
   Begin VB.TextBox txtCR7 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8790
      TabIndex        =   138
      Top             =   3000
      Width           =   2895
   End
   Begin VB.ComboBox cbDisplay 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmHOMappingCustomer.frx":0000
      Left            =   8790
      List            =   "frmHOMappingCustomer.frx":0010
      TabIndex        =   135
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8790
      TabIndex        =   130
      Top             =   1200
      Width           =   1695
      Begin TDBNumber6Ctl.TDBNumber sbAwal 
         Height          =   330
         Left            =   0
         TabIndex        =   131
         Top             =   0
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   582
         Calculator      =   "frmHOMappingCustomer.frx":0026
         Caption         =   "frmHOMappingCustomer.frx":0046
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":00B2
         Keys            =   "frmHOMappingCustomer.frx":00D0
         Spin            =   "frmHOMappingCustomer.frx":011A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   100
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber sbAkhir 
         Height          =   330
         Left            =   1080
         TabIndex        =   132
         Top             =   0
         Width           =   495
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   582
         Calculator      =   "frmHOMappingCustomer.frx":0142
         Caption         =   "frmHOMappingCustomer.frx":0162
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":01CE
         Keys            =   "frmHOMappingCustomer.frx":01EC
         Spin            =   "frmHOMappingCustomer.frx":0236
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   200
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   50
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
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
         Index           =   26
         Left            =   600
         TabIndex        =   133
         Top             =   0
         Width           =   345
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8790
      TabIndex        =   129
      Top             =   840
      Width           =   2625
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   10080
      TabIndex        =   72
      Top             =   4080
      Visible         =   0   'False
      Width           =   4575
      Begin VB.ComboBox cbBulanRet 
         Height          =   405
         ItemData        =   "frmHOMappingCustomer.frx":025E
         Left            =   1560
         List            =   "frmHOMappingCustomer.frx":0286
         Style           =   2  'Dropdown List
         TabIndex        =   80
         Top             =   600
         Width           =   975
      End
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Top             =   1560
         Visible         =   0   'False
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   873
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H000000FF&
         Caption         =   "BATAL"
         Height          =   495
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   1560
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H0080FFFF&
         Caption         =   "PRINT"
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   1560
         Width           =   1215
      End
      Begin TDBNumber6Ctl.TDBNumber TDBNumber1 
         Height          =   375
         Left            =   1560
         TabIndex        =   81
         Top             =   1080
         Width           =   975
         _Version        =   65536
         _ExtentX        =   1720
         _ExtentY        =   661
         Calculator      =   "frmHOMappingCustomer.frx":02B1
         Caption         =   "frmHOMappingCustomer.frx":02D1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":033D
         Keys            =   "frmHOMappingCustomer.frx":035B
         Spin            =   "frmHOMappingCustomer.frx":03A5
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   2500
         MinValue        =   1900
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1245189
         Value           =   2015
         MaxValueVT      =   5
         MinValueVT      =   5
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "KRITERIA SB RETENTION"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   0
         TabIndex        =   78
         Top             =   60
         Width           =   4575
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   495
         Left            =   0
         Top             =   0
         Width           =   4575
      End
      Begin VB.Label lblPengambilanDataRet 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   ".:: PROSES PENGAMBILAN DATA ::."
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   120
         TabIndex        =   77
         Top             =   2160
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "TAHUN :"
         Height          =   375
         Left            =   120
         TabIndex        =   74
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "BULAN :"
         Height          =   375
         Left            =   120
         TabIndex        =   73
         Top             =   600
         Width           =   1335
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   360
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox chModel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Semua Model"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8790
      TabIndex        =   127
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.CheckBox ckDealer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Semua Dealer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   11400
      TabIndex        =   126
      Top             =   1080
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CheckBox ckPilihSemua 
      BackColor       =   &H000080FF&
      Caption         =   "Centang Semua"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   480
      TabIndex        =   120
      Top             =   4920
      Width           =   1935
   End
   Begin TDBNumber6Ctl.TDBNumber txtUmurLastSB 
      Height          =   345
      Left            =   2865
      TabIndex        =   118
      Top             =   2190
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   609
      Calculator      =   "frmHOMappingCustomer.frx":03CD
      Caption         =   "frmHOMappingCustomer.frx":03ED
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":0452
      Keys            =   "frmHOMappingCustomer.frx":0470
      Spin            =   "frmHOMappingCustomer.frx":04BA
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   120
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin VB.CommandButton cmdTransfer 
      BackColor       =   &H0000FF00&
      Caption         =   "PROSES REMINDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8070
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   8520
      Visible         =   0   'False
      Width           =   2325
   End
   Begin MSAdodcLib.Adodc AdoDealerPenjual 
      Height          =   375
      Left            =   14265
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox cbLastTipeSVC 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmHOMappingCustomer.frx":04E2
      Left            =   11640
      List            =   "frmHOMappingCustomer.frx":04E4
      Style           =   2  'Dropdown List
      TabIndex        =   103
      Top             =   600
      Visible         =   0   'False
      Width           =   1860
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8790
      TabIndex        =   101
      Top             =   420
      Width           =   2625
   End
   Begin VB.ComboBox cbTahunAkhir 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmHOMappingCustomer.frx":04E6
      Left            =   4890
      List            =   "frmHOMappingCustomer.frx":04E8
      Style           =   2  'Dropdown List
      TabIndex        =   91
      Top             =   1470
      Width           =   1480
   End
   Begin VB.ComboBox cbTahunAwal 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmHOMappingCustomer.frx":04EA
      Left            =   2865
      List            =   "frmHOMappingCustomer.frx":04EC
      Style           =   2  'Dropdown List
      TabIndex        =   90
      Top             =   1455
      Width           =   1470
   End
   Begin VB.CommandButton cmdExportCSV 
      Caption         =   "EXPORT CSV"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11880
      TabIndex        =   82
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton cmdTutup 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Tutup"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   12765
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2580
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   16200
      TabIndex        =   1
      Top             =   -120
      Visible         =   0   'False
      Width           =   7095
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.CommandButton cmdKeluar 
      BackColor       =   &H000000FF&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13320
      Style           =   1  'Graphical
      TabIndex        =   71
      Top             =   8520
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      BackColor       =   &H000080FF&
      Caption         =   "EXPORT XLS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   70
      Top             =   8520
      Width           =   1335
   End
   Begin VB.Frame frameWarna 
      Caption         =   "Pilihan Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      TabIndex        =   36
      Top             =   1800
      Visible         =   0   'False
      Width           =   1935
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Default (Putih)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Biru"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   42
         Top             =   1320
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000FF00&
         Caption         =   "Hijau"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Kuning"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   39
         Text            =   "Text1"
         Top             =   2160
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Pilih Warna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H000000FF&
         Caption         =   "Merah"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame FramePasif2Range 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4845
      TabIndex        =   56
      Top             =   3675
      Width           =   2775
      Begin TDBNumber6Ctl.TDBNumber blnPasif2Awal 
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":04EE
         Caption         =   "frmHOMappingCustomer.frx":050E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":057A
         Keys            =   "frmHOMappingCustomer.frx":0598
         Spin            =   "frmHOMappingCustomer.frx":05E2
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   13
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber blnPasif2Akhir 
         Height          =   255
         Left            =   1560
         TabIndex        =   58
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":060A
         Caption         =   "frmHOMappingCustomer.frx":062A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":0696
         Keys            =   "frmHOMappingCustomer.frx":06B4
         Spin            =   "frmHOMappingCustomer.frx":06FE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   18
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan s.d "
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
         Index           =   13
         Left            =   600
         TabIndex        =   60
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Index           =   12
         Left            =   2040
         TabIndex        =   59
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.Frame framePasif1Range 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4845
      TabIndex        =   24
      Top             =   3315
      Width           =   2775
      Begin TDBNumber6Ctl.TDBNumber blnRangePasif1Awal 
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":0726
         Caption         =   "frmHOMappingCustomer.frx":0746
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":07B2
         Keys            =   "frmHOMappingCustomer.frx":07D0
         Spin            =   "frmHOMappingCustomer.frx":081A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   7
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber blnRangePasif1Akhir 
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":0842
         Caption         =   "frmHOMappingCustomer.frx":0862
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":08CE
         Keys            =   "frmHOMappingCustomer.frx":08EC
         Spin            =   "frmHOMappingCustomer.frx":0936
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   12
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan s.d "
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
         Index           =   7
         Left            =   600
         TabIndex        =   28
         Top             =   0
         Width           =   945
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Index           =   3
         Left            =   2040
         TabIndex        =   27
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   68
      Top             =   3735
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   4095
      Width           =   615
   End
   Begin VB.Frame FramePasif2Bulan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4845
      TabIndex        =   63
      Top             =   3675
      Width           =   1335
      Begin TDBNumber6Ctl.TDBNumber blnPasif2Bulan 
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":095E
         Caption         =   "frmHOMappingCustomer.frx":097E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":09EA
         Keys            =   "frmHOMappingCustomer.frx":0A08
         Spin            =   "frmHOMappingCustomer.frx":0A52
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   1
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Index           =   14
         Left            =   600
         TabIndex        =   65
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.ComboBox Combo3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmHOMappingCustomer.frx":0A7A
      Left            =   3765
      List            =   "frmHOMappingCustomer.frx":0A8A
      Style           =   2  'Dropdown List
      TabIndex        =   62
      Top             =   4035
      Width           =   975
   End
   Begin VB.CheckBox ckPasif2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategori Pasif 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   61
      Top             =   3675
      Width           =   1815
   End
   Begin VB.Frame FramePasif3Range 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4845
      TabIndex        =   51
      Top             =   4035
      Width           =   2655
      Begin TDBNumber6Ctl.TDBNumber blnPasif3Awal 
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":0A9E
         Caption         =   "frmHOMappingCustomer.frx":0ABE
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":0B2A
         Keys            =   "frmHOMappingCustomer.frx":0B48
         Spin            =   "frmHOMappingCustomer.frx":0B92
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   6
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber blnPasif3Akhir 
         Height          =   255
         Left            =   1560
         TabIndex        =   53
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":0BBA
         Caption         =   "frmHOMappingCustomer.frx":0BDA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":0C46
         Keys            =   "frmHOMappingCustomer.frx":0C64
         Spin            =   "frmHOMappingCustomer.frx":0CAE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   6
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Index           =   11
         Left            =   2040
         TabIndex        =   55
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan s.d "
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
         Index           =   10
         Left            =   600
         TabIndex        =   54
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.CheckBox ckPasif3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategori Pasif 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1845
      TabIndex        =   50
      Top             =   4035
      Width           =   1815
   End
   Begin VB.ComboBox Combo2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmHOMappingCustomer.frx":0CD6
      Left            =   3765
      List            =   "frmHOMappingCustomer.frx":0CE6
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   3675
      Width           =   975
   End
   Begin VB.Frame FramePasif3Bulan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4845
      TabIndex        =   46
      Top             =   4035
      Width           =   1335
      Begin TDBNumber6Ctl.TDBNumber blnPasif3Bulan 
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":0CFA
         Caption         =   "frmHOMappingCustomer.frx":0D1A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":0D86
         Keys            =   "frmHOMappingCustomer.frx":0DA4
         Spin            =   "frmHOMappingCustomer.frx":0DEE
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   99
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   18
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Index           =   9
         Left            =   600
         TabIndex        =   48
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3315
      Width           =   615
   End
   Begin VB.CommandButton cmdWarna 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Warna"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   345
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2955
      Width           =   615
   End
   Begin VB.Frame framePasifBulan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4845
      TabIndex        =   31
      Top             =   3315
      Width           =   1335
      Begin TDBNumber6Ctl.TDBNumber blnPasif1 
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":0E16
         Caption         =   "frmHOMappingCustomer.frx":0E36
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":0EA2
         Keys            =   "frmHOMappingCustomer.frx":0EC0
         Spin            =   "frmHOMappingCustomer.frx":0F0A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   1376261
         Value           =   1
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Index           =   8
         Left            =   600
         TabIndex        =   33
         Top             =   0
         Width           =   540
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmHOMappingCustomer.frx":0F32
      Left            =   3765
      List            =   "frmHOMappingCustomer.frx":0F42
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   3315
      Width           =   975
   End
   Begin VB.CheckBox ckPasif1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategori Pasif 1"
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
      Left            =   1845
      TabIndex        =   29
      Top             =   3315
      Width           =   1815
   End
   Begin VB.Frame FrameAktifRange 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4845
      TabIndex        =   19
      Top             =   2955
      Width           =   2775
      Begin TDBNumber6Ctl.TDBNumber blnRangeAktifAwal 
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":0F56
         Caption         =   "frmHOMappingCustomer.frx":0F76
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":0FE2
         Keys            =   "frmHOMappingCustomer.frx":1000
         Spin            =   "frmHOMappingCustomer.frx":104A
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   6
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin TDBNumber6Ctl.TDBNumber blnRangeAktifAkhir 
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":1072
         Caption         =   "frmHOMappingCustomer.frx":1092
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":10FE
         Keys            =   "frmHOMappingCustomer.frx":111C
         Spin            =   "frmHOMappingCustomer.frx":1166
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   6
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Index           =   6
         Left            =   2040
         TabIndex        =   23
         Top             =   0
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan s.d "
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
         Index           =   5
         Left            =   600
         TabIndex        =   21
         Top             =   0
         Width           =   945
      End
   End
   Begin VB.CheckBox ckAktif 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Kategori Aktif"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1845
      TabIndex        =   16
      Top             =   2955
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data HO"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5160
      TabIndex        =   14
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.ComboBox cbOperator 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmHOMappingCustomer.frx":118E
      Left            =   3765
      List            =   "frmHOMappingCustomer.frx":119E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2955
      Width           =   975
   End
   Begin VB.TextBox txtAlamat 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   8790
      TabIndex        =   8
      Top             =   45
      Width           =   2625
   End
   Begin TDBDate6Ctl.TDBDate Tgl1 
      Height          =   330
      Left            =   2865
      TabIndex        =   5
      Top             =   15
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":11B2
      Caption         =   "frmHOMappingCustomer.frx":12DE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":134A
      Keys            =   "frmHOMappingCustomer.frx":1368
      Spin            =   "frmHOMappingCustomer.frx":13C6
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate Tgl2 
      Height          =   330
      Left            =   4890
      TabIndex        =   6
      Top             =   15
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":13EE
      Caption         =   "frmHOMappingCustomer.frx":151A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":1586
      Keys            =   "frmHOMappingCustomer.frx":15A4
      Spin            =   "frmHOMappingCustomer.frx":1602
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin ucXPButton.XPButton cmdRef 
      Height          =   495
      Left            =   405
      TabIndex        =   4
      Top             =   4440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   873
      BTYPE           =   6
      TX              =   "&r"
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
      MICON           =   "frmHOMappingCustomer.frx":162A
      PICN            =   "frmHOMappingCustomer.frx":1646
      PICH            =   "frmHOMappingCustomer.frx":1BC2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame FrameAktifBulan 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4845
      TabIndex        =   15
      Top             =   2955
      Width           =   1335
      Begin TDBNumber6Ctl.TDBNumber blnAktif 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   0
         Width           =   375
         _Version        =   65536
         _ExtentX        =   661
         _ExtentY        =   450
         Calculator      =   "frmHOMappingCustomer.frx":21F9
         Caption         =   "frmHOMappingCustomer.frx":2219
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DropDown        =   "frmHOMappingCustomer.frx":2285
         Keys            =   "frmHOMappingCustomer.frx":22A3
         Spin            =   "frmHOMappingCustomer.frx":22ED
         AlignHorizontal =   1
         AlignVertical   =   0
         Appearance      =   1
         BackColor       =   -2147483643
         BorderStyle     =   1
         BtnPositioning  =   0
         ClipMode        =   0
         ClearAction     =   0
         DecimalPoint    =   ","
         DisplayFormat   =   "####0;;Null"
         EditMode        =   0
         Enabled         =   -1
         ErrorBeep       =   0
         ForeColor       =   -2147483640
         Format          =   "####0"
         HighlightText   =   0
         MarginBottom    =   1
         MarginLeft      =   1
         MarginRight     =   1
         MarginTop       =   1
         MaxValue        =   12
         MinValue        =   1
         MousePointer    =   0
         MoveOnLRKey     =   0
         NegativeColor   =   255
         OLEDragMode     =   0
         OLEDropMode     =   0
         ReadOnly        =   0
         Separator       =   "."
         ShowContextMenu =   -1
         ValueVT         =   5
         Value           =   7
         MaxValueVT      =   1229389829
         MinValueVT      =   1431437317
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bulan"
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
         Index           =   4
         Left            =   600
         TabIndex        =   17
         Top             =   0
         Width           =   540
      End
   End
   Begin TDBDate6Ctl.TDBDate tglDDate1 
      Height          =   330
      Left            =   2865
      TabIndex        =   83
      Top             =   735
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":2315
      Caption         =   "frmHOMappingCustomer.frx":2441
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":24AD
      Keys            =   "frmHOMappingCustomer.frx":24CB
      Spin            =   "frmHOMappingCustomer.frx":2529
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate tglDDate2 
      Height          =   330
      Left            =   4890
      TabIndex        =   84
      Top             =   735
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":2551
      Caption         =   "frmHOMappingCustomer.frx":267D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":26E9
      Keys            =   "frmHOMappingCustomer.frx":2707
      Spin            =   "frmHOMappingCustomer.frx":2765
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate TDBDate1 
      Height          =   330
      Left            =   2865
      TabIndex        =   86
      Top             =   1095
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":278D
      Caption         =   "frmHOMappingCustomer.frx":28B9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":2925
      Keys            =   "frmHOMappingCustomer.frx":2943
      Spin            =   "frmHOMappingCustomer.frx":29A1
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate TDBDate2 
      Height          =   330
      Left            =   4890
      TabIndex        =   87
      Top             =   1095
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":29C9
      Caption         =   "frmHOMappingCustomer.frx":2AF5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":2B61
      Keys            =   "frmHOMappingCustomer.frx":2B7F
      Spin            =   "frmHOMappingCustomer.frx":2BDD
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate TDBDate3 
      Height          =   330
      Left            =   2865
      TabIndex        =   92
      Top             =   1785
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":2C05
      Caption         =   "frmHOMappingCustomer.frx":2D31
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":2D9D
      Keys            =   "frmHOMappingCustomer.frx":2DBB
      Spin            =   "frmHOMappingCustomer.frx":2E19
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate TDBDate4 
      Height          =   330
      Left            =   4890
      TabIndex        =   93
      Top             =   1785
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":2E41
      Caption         =   "frmHOMappingCustomer.frx":2F6D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":2FD9
      Keys            =   "frmHOMappingCustomer.frx":2FF7
      Spin            =   "frmHOMappingCustomer.frx":3055
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin MSAdodcLib.Adodc AdoGroupModel 
      Height          =   375
      Left            =   14265
      Top             =   450
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TrueOleDBGrid70.TDBGrid gridTagih 
      Height          =   3225
      Left            =   360
      TabIndex        =   107
      Top             =   5280
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   5689
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   4
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
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).DataField=   ""
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).DataField=   ""
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).DataField=   ""
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).DataField=   ""
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).DataField=   ""
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).DataField=   ""
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(10)._VlistStyle=   0
      Columns(10)._MaxComboItems=   5
      Columns(10).DataField=   ""
      Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(11)._VlistStyle=   0
      Columns(11)._MaxComboItems=   5
      Columns(11).DataField=   ""
      Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(12)._VlistStyle=   0
      Columns(12)._MaxComboItems=   5
      Columns(12).DataField=   ""
      Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(13)._VlistStyle=   0
      Columns(13)._MaxComboItems=   5
      Columns(13).DataField=   ""
      Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(14)._VlistStyle=   0
      Columns(14)._MaxComboItems=   5
      Columns(14).DataField=   ""
      Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(15)._VlistStyle=   0
      Columns(15)._MaxComboItems=   5
      Columns(15).DataField=   ""
      Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(16)._VlistStyle=   0
      Columns(16)._MaxComboItems=   5
      Columns(16).DataField=   ""
      Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(17)._VlistStyle=   0
      Columns(17)._MaxComboItems=   5
      Columns(17).DataField=   ""
      Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(18)._VlistStyle=   0
      Columns(18)._MaxComboItems=   5
      Columns(18).DataField=   ""
      Columns(18)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(19)._VlistStyle=   0
      Columns(19)._MaxComboItems=   5
      Columns(19).DataField=   ""
      Columns(19).NumberFormat=   "Standard"
      Columns(19)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(20)._VlistStyle=   0
      Columns(20)._MaxComboItems=   5
      Columns(20).DataField=   ""
      Columns(20).NumberFormat=   "FormatText Event"
      Columns(20)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(21)._VlistStyle=   0
      Columns(21)._MaxComboItems=   5
      Columns(21).DataField=   ""
      Columns(21).NumberFormat=   "FormatText Event"
      Columns(21)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(22)._VlistStyle=   0
      Columns(22)._MaxComboItems=   5
      Columns(22).DataField=   ""
      Columns(22)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(23)._VlistStyle=   0
      Columns(23)._MaxComboItems=   5
      Columns(23).DataField=   ""
      Columns(23)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(24)._VlistStyle=   0
      Columns(24)._MaxComboItems=   5
      Columns(24).DataField=   ""
      Columns(24)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(25)._VlistStyle=   0
      Columns(25)._MaxComboItems=   5
      Columns(25).DataField=   ""
      Columns(25)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(26)._VlistStyle=   0
      Columns(26)._MaxComboItems=   5
      Columns(26).DataField=   ""
      Columns(26)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(27)._VlistStyle=   0
      Columns(27)._MaxComboItems=   5
      Columns(27).DataField=   ""
      Columns(27)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(28)._VlistStyle=   0
      Columns(28)._MaxComboItems=   5
      Columns(28).DataField=   ""
      Columns(28)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(29)._VlistStyle=   0
      Columns(29)._MaxComboItems=   5
      Columns(29).DataField=   ""
      Columns(29)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(30)._VlistStyle=   0
      Columns(30)._MaxComboItems=   5
      Columns(30).DataField=   ""
      Columns(30)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(31)._VlistStyle=   0
      Columns(31)._MaxComboItems=   5
      Columns(31).DataField=   ""
      Columns(31)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(32)._VlistStyle=   0
      Columns(32)._MaxComboItems=   5
      Columns(32).DataField=   ""
      Columns(32)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(33)._VlistStyle=   0
      Columns(33)._MaxComboItems=   5
      Columns(33).DataField=   ""
      Columns(33)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(34)._VlistStyle=   0
      Columns(34)._MaxComboItems=   5
      Columns(34).DataField=   ""
      Columns(34)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(35)._VlistStyle=   0
      Columns(35)._MaxComboItems=   5
      Columns(35).DataField=   ""
      Columns(35)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(36)._VlistStyle=   0
      Columns(36)._MaxComboItems=   5
      Columns(36).DataField=   ""
      Columns(36)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(37)._VlistStyle=   0
      Columns(37)._MaxComboItems=   5
      Columns(37).DataField=   ""
      Columns(37)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(38)._VlistStyle=   0
      Columns(38)._MaxComboItems=   5
      Columns(38).DataField=   ""
      Columns(38)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(39)._VlistStyle=   0
      Columns(39)._MaxComboItems=   5
      Columns(39).DataField=   ""
      Columns(39)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(40)._VlistStyle=   0
      Columns(40)._MaxComboItems=   5
      Columns(40).DataField=   ""
      Columns(40)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(41)._VlistStyle=   0
      Columns(41)._MaxComboItems=   5
      Columns(41).DataField=   ""
      Columns(41)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(42)._VlistStyle=   0
      Columns(42)._MaxComboItems=   5
      Columns(42).DataField=   ""
      Columns(42)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(43)._VlistStyle=   0
      Columns(43)._MaxComboItems=   5
      Columns(43).DataField=   ""
      Columns(43)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(44)._VlistStyle=   0
      Columns(44)._MaxComboItems=   5
      Columns(44).DataField=   ""
      Columns(44)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(45)._VlistStyle=   0
      Columns(45)._MaxComboItems=   5
      Columns(45).DataField=   ""
      Columns(45)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(46)._VlistStyle=   0
      Columns(46)._MaxComboItems=   5
      Columns(46).DataField=   ""
      Columns(46)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(47)._VlistStyle=   0
      Columns(47)._MaxComboItems=   5
      Columns(47).DataField=   ""
      Columns(47)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(48)._VlistStyle=   0
      Columns(48)._MaxComboItems=   5
      Columns(48).DataField=   ""
      Columns(48)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(49)._VlistStyle=   0
      Columns(49)._MaxComboItems=   5
      Columns(49).DataField=   ""
      Columns(49)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(50)._VlistStyle=   0
      Columns(50)._MaxComboItems=   5
      Columns(50).DataField=   ""
      Columns(50)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(51)._VlistStyle=   0
      Columns(51)._MaxComboItems=   5
      Columns(51).DataField=   ""
      Columns(51)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(52)._VlistStyle=   0
      Columns(52)._MaxComboItems=   5
      Columns(52).DataField=   ""
      Columns(52)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(53)._VlistStyle=   0
      Columns(53)._MaxComboItems=   5
      Columns(53).DataField=   ""
      Columns(53)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(54)._VlistStyle=   0
      Columns(54)._MaxComboItems=   5
      Columns(54).DataField=   ""
      Columns(54)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(55)._VlistStyle=   0
      Columns(55)._MaxComboItems=   5
      Columns(55).DataField=   ""
      Columns(55)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(56)._VlistStyle=   0
      Columns(56)._MaxComboItems=   5
      Columns(56).DataField=   ""
      Columns(56)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(57)._VlistStyle=   0
      Columns(57)._MaxComboItems=   5
      Columns(57).DataField=   ""
      Columns(57)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(58)._VlistStyle=   0
      Columns(58)._MaxComboItems=   5
      Columns(58).DataField=   ""
      Columns(58)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(59)._VlistStyle=   0
      Columns(59)._MaxComboItems=   5
      Columns(59).DataField=   ""
      Columns(59)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(60)._VlistStyle=   0
      Columns(60)._MaxComboItems=   5
      Columns(60).DataField=   ""
      Columns(60)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(61)._VlistStyle=   0
      Columns(61)._MaxComboItems=   5
      Columns(61).DataField=   ""
      Columns(61)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(62)._VlistStyle=   0
      Columns(62)._MaxComboItems=   5
      Columns(62).DataField=   ""
      Columns(62)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(63)._VlistStyle=   0
      Columns(63)._MaxComboItems=   5
      Columns(63).DataField=   ""
      Columns(63)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(64)._VlistStyle=   0
      Columns(64)._MaxComboItems=   5
      Columns(64).DataField=   ""
      Columns(64)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(65)._VlistStyle=   0
      Columns(65)._MaxComboItems=   5
      Columns(65).DataField=   ""
      Columns(65)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(66)._VlistStyle=   0
      Columns(66)._MaxComboItems=   5
      Columns(66).DataField=   ""
      Columns(66)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(67)._VlistStyle=   0
      Columns(67)._MaxComboItems=   5
      Columns(67).DataField=   ""
      Columns(67)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(68)._VlistStyle=   0
      Columns(68)._MaxComboItems=   5
      Columns(68).DataField=   ""
      Columns(68)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(69)._VlistStyle=   0
      Columns(69)._MaxComboItems=   5
      Columns(69).DataField=   ""
      Columns(69)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(70)._VlistStyle=   0
      Columns(70)._MaxComboItems=   5
      Columns(70).DataField=   ""
      Columns(70)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(71)._VlistStyle=   0
      Columns(71)._MaxComboItems=   5
      Columns(71).DataField=   ""
      Columns(71)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(72)._VlistStyle=   0
      Columns(72)._MaxComboItems=   5
      Columns(72).DataField=   ""
      Columns(72)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(73)._VlistStyle=   0
      Columns(73)._MaxComboItems=   5
      Columns(73).DataField=   ""
      Columns(73)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(74)._VlistStyle=   0
      Columns(74)._MaxComboItems=   5
      Columns(74).DataField=   ""
      Columns(74)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(75)._VlistStyle=   0
      Columns(75)._MaxComboItems=   5
      Columns(75).DataField=   ""
      Columns(75)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(76)._VlistStyle=   0
      Columns(76)._MaxComboItems=   5
      Columns(76).DataField=   ""
      Columns(76)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(77)._VlistStyle=   0
      Columns(77)._MaxComboItems=   5
      Columns(77).DataField=   ""
      Columns(77)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(78)._VlistStyle=   0
      Columns(78)._MaxComboItems=   5
      Columns(78).DataField=   ""
      Columns(78)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(79)._VlistStyle=   0
      Columns(79)._MaxComboItems=   5
      Columns(79).DataField=   ""
      Columns(79)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(80)._VlistStyle=   0
      Columns(80)._MaxComboItems=   5
      Columns(80).DataField=   ""
      Columns(80)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(81)._VlistStyle=   0
      Columns(81)._MaxComboItems=   5
      Columns(81).DataField=   ""
      Columns(81)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(82)._VlistStyle=   0
      Columns(82)._MaxComboItems=   5
      Columns(82).DataField=   ""
      Columns(82)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(83)._VlistStyle=   0
      Columns(83)._MaxComboItems=   5
      Columns(83).DataField=   ""
      Columns(83)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(84)._VlistStyle=   0
      Columns(84)._MaxComboItems=   5
      Columns(84).DataField=   ""
      Columns(84)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(85)._VlistStyle=   0
      Columns(85)._MaxComboItems=   5
      Columns(85).DataField=   ""
      Columns(85)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(86)._VlistStyle=   0
      Columns(86)._MaxComboItems=   5
      Columns(86).DataField=   ""
      Columns(86)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(87)._VlistStyle=   0
      Columns(87)._MaxComboItems=   5
      Columns(87).DataField=   ""
      Columns(87)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(88)._VlistStyle=   0
      Columns(88)._MaxComboItems=   5
      Columns(88).DataField=   ""
      Columns(88)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(89)._VlistStyle=   0
      Columns(89)._MaxComboItems=   5
      Columns(89).DataField=   ""
      Columns(89)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(90)._VlistStyle=   0
      Columns(90)._MaxComboItems=   5
      Columns(90).DataField=   ""
      Columns(90)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(91)._VlistStyle=   0
      Columns(91)._MaxComboItems=   5
      Columns(91).DataField=   ""
      Columns(91)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(92)._VlistStyle=   0
      Columns(92)._MaxComboItems=   5
      Columns(92).DataField=   ""
      Columns(92)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(93)._VlistStyle=   0
      Columns(93)._MaxComboItems=   5
      Columns(93).DataField=   ""
      Columns(93)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(94)._VlistStyle=   0
      Columns(94)._MaxComboItems=   5
      Columns(94).DataField=   ""
      Columns(94)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(95)._VlistStyle=   0
      Columns(95)._MaxComboItems=   5
      Columns(95).DataField=   ""
      Columns(95)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(96)._VlistStyle=   0
      Columns(96)._MaxComboItems=   5
      Columns(96).DataField=   ""
      Columns(96)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(97)._VlistStyle=   0
      Columns(97)._MaxComboItems=   5
      Columns(97).DataField=   ""
      Columns(97)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(98)._VlistStyle=   0
      Columns(98)._MaxComboItems=   5
      Columns(98).DataField=   ""
      Columns(98)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(99)._VlistStyle=   0
      Columns(99)._MaxComboItems=   5
      Columns(99).DataField=   ""
      Columns(99)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(100)._VlistStyle=   0
      Columns(100)._MaxComboItems=   5
      Columns(100).DataField=   ""
      Columns(100)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(101)._VlistStyle=   0
      Columns(101)._MaxComboItems=   5
      Columns(101).DataField=   ""
      Columns(101)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(102)._VlistStyle=   0
      Columns(102)._MaxComboItems=   5
      Columns(102).DataField=   ""
      Columns(102)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(103)._VlistStyle=   0
      Columns(103)._MaxComboItems=   5
      Columns(103).DataField=   ""
      Columns(103)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(104)._VlistStyle=   0
      Columns(104)._MaxComboItems=   5
      Columns(104).DataField=   ""
      Columns(104)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(105)._VlistStyle=   0
      Columns(105)._MaxComboItems=   5
      Columns(105).DataField=   ""
      Columns(105)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(106)._VlistStyle=   0
      Columns(106)._MaxComboItems=   5
      Columns(106).DataField=   ""
      Columns(106)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(107)._VlistStyle=   0
      Columns(107)._MaxComboItems=   5
      Columns(107).DataField=   ""
      Columns(107)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(108)._VlistStyle=   0
      Columns(108)._MaxComboItems=   5
      Columns(108).DataField=   ""
      Columns(108)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(109)._VlistStyle=   0
      Columns(109)._MaxComboItems=   5
      Columns(109).DataField=   ""
      Columns(109)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(110)._VlistStyle=   0
      Columns(110)._MaxComboItems=   5
      Columns(110).DataField=   ""
      Columns(110)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(111)._VlistStyle=   0
      Columns(111)._MaxComboItems=   5
      Columns(111).DataField=   ""
      Columns(111)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(112)._VlistStyle=   0
      Columns(112)._MaxComboItems=   5
      Columns(112).DataField=   ""
      Columns(112)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(113)._VlistStyle=   0
      Columns(113)._MaxComboItems=   5
      Columns(113).DataField=   ""
      Columns(113)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(114)._VlistStyle=   0
      Columns(114)._MaxComboItems=   5
      Columns(114).DataField=   ""
      Columns(114)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(115)._VlistStyle=   0
      Columns(115)._MaxComboItems=   5
      Columns(115).DataField=   ""
      Columns(115)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(116)._VlistStyle=   0
      Columns(116)._MaxComboItems=   5
      Columns(116).DataField=   ""
      Columns(116)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(117)._VlistStyle=   0
      Columns(117)._MaxComboItems=   5
      Columns(117).DataField=   ""
      Columns(117)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(118)._VlistStyle=   0
      Columns(118)._MaxComboItems=   5
      Columns(118).DataField=   ""
      Columns(118)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(119)._VlistStyle=   0
      Columns(119)._MaxComboItems=   5
      Columns(119).DataField=   ""
      Columns(119)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(120)._VlistStyle=   0
      Columns(120)._MaxComboItems=   5
      Columns(120).DataField=   ""
      Columns(120)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(121)._VlistStyle=   0
      Columns(121)._MaxComboItems=   5
      Columns(121).DataField=   ""
      Columns(121)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(122)._VlistStyle=   0
      Columns(122)._MaxComboItems=   5
      Columns(122).DataField=   ""
      Columns(122)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(123)._VlistStyle=   0
      Columns(123)._MaxComboItems=   5
      Columns(123).DataField=   ""
      Columns(123)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(124)._VlistStyle=   0
      Columns(124)._MaxComboItems=   5
      Columns(124).DataField=   ""
      Columns(124)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(125)._VlistStyle=   0
      Columns(125)._MaxComboItems=   5
      Columns(125).DataField=   ""
      Columns(125)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(126)._VlistStyle=   0
      Columns(126)._MaxComboItems=   5
      Columns(126).DataField=   ""
      Columns(126)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(127)._VlistStyle=   0
      Columns(127)._MaxComboItems=   5
      Columns(127).DataField=   ""
      Columns(127)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(128)._VlistStyle=   0
      Columns(128)._MaxComboItems=   5
      Columns(128).DataField=   ""
      Columns(128)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(129)._VlistStyle=   0
      Columns(129)._MaxComboItems=   5
      Columns(129).DataField=   ""
      Columns(129)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(130)._VlistStyle=   0
      Columns(130)._MaxComboItems=   5
      Columns(130).DataField=   ""
      Columns(130)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(131)._VlistStyle=   0
      Columns(131)._MaxComboItems=   5
      Columns(131).DataField=   ""
      Columns(131)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(132)._VlistStyle=   0
      Columns(132)._MaxComboItems=   5
      Columns(132).DataField=   ""
      Columns(132)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(133)._VlistStyle=   0
      Columns(133)._MaxComboItems=   5
      Columns(133).DataField=   ""
      Columns(133)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(134)._VlistStyle=   0
      Columns(134)._MaxComboItems=   5
      Columns(134).DataField=   ""
      Columns(134)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(135)._VlistStyle=   0
      Columns(135)._MaxComboItems=   5
      Columns(135).DataField=   ""
      Columns(135)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(136)._VlistStyle=   0
      Columns(136)._MaxComboItems=   5
      Columns(136).DataField=   ""
      Columns(136)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(137)._VlistStyle=   0
      Columns(137)._MaxComboItems=   5
      Columns(137).DataField=   ""
      Columns(137)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   138
      Splits(0)._UserFlags=   0
      Splits(0).SizeMode=   2
      Splits(0).Size  =   3
      Splits(0).Size.vt=   2
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).FetchRowStyle=   -1  'True
      Splits(0).AlternatingRowStyle=   -1  'True
      Splits(0).DividerColor=   14933984
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=138"
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
      Splits(0)._ColumnProps(37)=   "Column(9).Width=3281"
      Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
      Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(0)._ColumnProps(41)=   "Column(10).Width=3281"
      Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
      Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(0)._ColumnProps(45)=   "Column(11).Width=3281"
      Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
      Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(0)._ColumnProps(49)=   "Column(12).Width=3281"
      Splits(0)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(0)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
      Splits(0)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(0)._ColumnProps(53)=   "Column(13).Width=3281"
      Splits(0)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(0)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
      Splits(0)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(0)._ColumnProps(57)=   "Column(14).Width=3281"
      Splits(0)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(0)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
      Splits(0)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(0)._ColumnProps(61)=   "Column(15).Width=3281"
      Splits(0)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(0)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
      Splits(0)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(0)._ColumnProps(65)=   "Column(16).Width=3281"
      Splits(0)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(0)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
      Splits(0)._ColumnProps(68)=   "Column(16).Order=17"
      Splits(0)._ColumnProps(69)=   "Column(17).Width=3281"
      Splits(0)._ColumnProps(70)=   "Column(17).DividerColor=0"
      Splits(0)._ColumnProps(71)=   "Column(17)._WidthInPix=3175"
      Splits(0)._ColumnProps(72)=   "Column(17).Order=18"
      Splits(0)._ColumnProps(73)=   "Column(18).Width=3281"
      Splits(0)._ColumnProps(74)=   "Column(18).DividerColor=0"
      Splits(0)._ColumnProps(75)=   "Column(18)._WidthInPix=3175"
      Splits(0)._ColumnProps(76)=   "Column(18).Order=19"
      Splits(0)._ColumnProps(77)=   "Column(19).Width=3281"
      Splits(0)._ColumnProps(78)=   "Column(19).DividerColor=0"
      Splits(0)._ColumnProps(79)=   "Column(19)._WidthInPix=3175"
      Splits(0)._ColumnProps(80)=   "Column(19).Order=20"
      Splits(0)._ColumnProps(81)=   "Column(20).Width=3281"
      Splits(0)._ColumnProps(82)=   "Column(20).DividerColor=0"
      Splits(0)._ColumnProps(83)=   "Column(20)._WidthInPix=3175"
      Splits(0)._ColumnProps(84)=   "Column(20).Order=21"
      Splits(0)._ColumnProps(85)=   "Column(21).Width=3281"
      Splits(0)._ColumnProps(86)=   "Column(21).DividerColor=0"
      Splits(0)._ColumnProps(87)=   "Column(21)._WidthInPix=3175"
      Splits(0)._ColumnProps(88)=   "Column(21).Order=22"
      Splits(0)._ColumnProps(89)=   "Column(22).Width=3281"
      Splits(0)._ColumnProps(90)=   "Column(22).DividerColor=0"
      Splits(0)._ColumnProps(91)=   "Column(22)._WidthInPix=3175"
      Splits(0)._ColumnProps(92)=   "Column(22).Order=23"
      Splits(0)._ColumnProps(93)=   "Column(23).Width=3281"
      Splits(0)._ColumnProps(94)=   "Column(23).DividerColor=0"
      Splits(0)._ColumnProps(95)=   "Column(23)._WidthInPix=3175"
      Splits(0)._ColumnProps(96)=   "Column(23).Order=24"
      Splits(0)._ColumnProps(97)=   "Column(24).Width=3281"
      Splits(0)._ColumnProps(98)=   "Column(24).DividerColor=0"
      Splits(0)._ColumnProps(99)=   "Column(24)._WidthInPix=3175"
      Splits(0)._ColumnProps(100)=   "Column(24).Order=25"
      Splits(0)._ColumnProps(101)=   "Column(25).Width=3281"
      Splits(0)._ColumnProps(102)=   "Column(25).DividerColor=0"
      Splits(0)._ColumnProps(103)=   "Column(25)._WidthInPix=3175"
      Splits(0)._ColumnProps(104)=   "Column(25).Order=26"
      Splits(0)._ColumnProps(105)=   "Column(26).Width=3281"
      Splits(0)._ColumnProps(106)=   "Column(26).DividerColor=0"
      Splits(0)._ColumnProps(107)=   "Column(26)._WidthInPix=3175"
      Splits(0)._ColumnProps(108)=   "Column(26).Order=27"
      Splits(0)._ColumnProps(109)=   "Column(27).Width=3281"
      Splits(0)._ColumnProps(110)=   "Column(27).DividerColor=0"
      Splits(0)._ColumnProps(111)=   "Column(27)._WidthInPix=3175"
      Splits(0)._ColumnProps(112)=   "Column(27).Order=28"
      Splits(0)._ColumnProps(113)=   "Column(28).Width=3281"
      Splits(0)._ColumnProps(114)=   "Column(28).DividerColor=0"
      Splits(0)._ColumnProps(115)=   "Column(28)._WidthInPix=3175"
      Splits(0)._ColumnProps(116)=   "Column(28).Order=29"
      Splits(0)._ColumnProps(117)=   "Column(29).Width=3281"
      Splits(0)._ColumnProps(118)=   "Column(29).DividerColor=0"
      Splits(0)._ColumnProps(119)=   "Column(29)._WidthInPix=3175"
      Splits(0)._ColumnProps(120)=   "Column(29).Order=30"
      Splits(0)._ColumnProps(121)=   "Column(30).Width=3281"
      Splits(0)._ColumnProps(122)=   "Column(30).DividerColor=0"
      Splits(0)._ColumnProps(123)=   "Column(30)._WidthInPix=3175"
      Splits(0)._ColumnProps(124)=   "Column(30).Order=31"
      Splits(0)._ColumnProps(125)=   "Column(31).Width=3281"
      Splits(0)._ColumnProps(126)=   "Column(31).DividerColor=0"
      Splits(0)._ColumnProps(127)=   "Column(31)._WidthInPix=3175"
      Splits(0)._ColumnProps(128)=   "Column(31).Order=32"
      Splits(0)._ColumnProps(129)=   "Column(32).Width=3281"
      Splits(0)._ColumnProps(130)=   "Column(32).DividerColor=0"
      Splits(0)._ColumnProps(131)=   "Column(32)._WidthInPix=3175"
      Splits(0)._ColumnProps(132)=   "Column(32).Order=33"
      Splits(0)._ColumnProps(133)=   "Column(33).Width=3281"
      Splits(0)._ColumnProps(134)=   "Column(33).DividerColor=0"
      Splits(0)._ColumnProps(135)=   "Column(33)._WidthInPix=3175"
      Splits(0)._ColumnProps(136)=   "Column(33).Order=34"
      Splits(0)._ColumnProps(137)=   "Column(34).Width=3281"
      Splits(0)._ColumnProps(138)=   "Column(34).DividerColor=0"
      Splits(0)._ColumnProps(139)=   "Column(34)._WidthInPix=3175"
      Splits(0)._ColumnProps(140)=   "Column(34).Order=35"
      Splits(0)._ColumnProps(141)=   "Column(35).Width=3281"
      Splits(0)._ColumnProps(142)=   "Column(35).DividerColor=0"
      Splits(0)._ColumnProps(143)=   "Column(35)._WidthInPix=3175"
      Splits(0)._ColumnProps(144)=   "Column(35).Order=36"
      Splits(0)._ColumnProps(145)=   "Column(36).Width=3281"
      Splits(0)._ColumnProps(146)=   "Column(36).DividerColor=0"
      Splits(0)._ColumnProps(147)=   "Column(36)._WidthInPix=3175"
      Splits(0)._ColumnProps(148)=   "Column(36).Order=37"
      Splits(0)._ColumnProps(149)=   "Column(37).Width=3281"
      Splits(0)._ColumnProps(150)=   "Column(37).DividerColor=0"
      Splits(0)._ColumnProps(151)=   "Column(37)._WidthInPix=3175"
      Splits(0)._ColumnProps(152)=   "Column(37).Order=38"
      Splits(0)._ColumnProps(153)=   "Column(38).Width=3281"
      Splits(0)._ColumnProps(154)=   "Column(38).DividerColor=0"
      Splits(0)._ColumnProps(155)=   "Column(38)._WidthInPix=3175"
      Splits(0)._ColumnProps(156)=   "Column(38).Order=39"
      Splits(0)._ColumnProps(157)=   "Column(39).Width=3281"
      Splits(0)._ColumnProps(158)=   "Column(39).DividerColor=0"
      Splits(0)._ColumnProps(159)=   "Column(39)._WidthInPix=3175"
      Splits(0)._ColumnProps(160)=   "Column(39).Order=40"
      Splits(0)._ColumnProps(161)=   "Column(40).Width=3281"
      Splits(0)._ColumnProps(162)=   "Column(40).DividerColor=0"
      Splits(0)._ColumnProps(163)=   "Column(40)._WidthInPix=3175"
      Splits(0)._ColumnProps(164)=   "Column(40).Order=41"
      Splits(0)._ColumnProps(165)=   "Column(41).Width=3281"
      Splits(0)._ColumnProps(166)=   "Column(41).DividerColor=0"
      Splits(0)._ColumnProps(167)=   "Column(41)._WidthInPix=3175"
      Splits(0)._ColumnProps(168)=   "Column(41).Order=42"
      Splits(0)._ColumnProps(169)=   "Column(42).Width=3281"
      Splits(0)._ColumnProps(170)=   "Column(42).DividerColor=0"
      Splits(0)._ColumnProps(171)=   "Column(42)._WidthInPix=3175"
      Splits(0)._ColumnProps(172)=   "Column(42).Order=43"
      Splits(0)._ColumnProps(173)=   "Column(43).Width=3281"
      Splits(0)._ColumnProps(174)=   "Column(43).DividerColor=0"
      Splits(0)._ColumnProps(175)=   "Column(43)._WidthInPix=3175"
      Splits(0)._ColumnProps(176)=   "Column(43).Order=44"
      Splits(0)._ColumnProps(177)=   "Column(44).Width=3281"
      Splits(0)._ColumnProps(178)=   "Column(44).DividerColor=0"
      Splits(0)._ColumnProps(179)=   "Column(44)._WidthInPix=3175"
      Splits(0)._ColumnProps(180)=   "Column(44).Order=45"
      Splits(0)._ColumnProps(181)=   "Column(45).Width=3281"
      Splits(0)._ColumnProps(182)=   "Column(45).DividerColor=0"
      Splits(0)._ColumnProps(183)=   "Column(45)._WidthInPix=3175"
      Splits(0)._ColumnProps(184)=   "Column(45).Order=46"
      Splits(0)._ColumnProps(185)=   "Column(46).Width=3281"
      Splits(0)._ColumnProps(186)=   "Column(46).DividerColor=0"
      Splits(0)._ColumnProps(187)=   "Column(46)._WidthInPix=3175"
      Splits(0)._ColumnProps(188)=   "Column(46).Order=47"
      Splits(0)._ColumnProps(189)=   "Column(47).Width=3281"
      Splits(0)._ColumnProps(190)=   "Column(47).DividerColor=0"
      Splits(0)._ColumnProps(191)=   "Column(47)._WidthInPix=3175"
      Splits(0)._ColumnProps(192)=   "Column(47).Order=48"
      Splits(0)._ColumnProps(193)=   "Column(48).Width=3281"
      Splits(0)._ColumnProps(194)=   "Column(48).DividerColor=0"
      Splits(0)._ColumnProps(195)=   "Column(48)._WidthInPix=3175"
      Splits(0)._ColumnProps(196)=   "Column(48).Order=49"
      Splits(0)._ColumnProps(197)=   "Column(49).Width=3281"
      Splits(0)._ColumnProps(198)=   "Column(49).DividerColor=0"
      Splits(0)._ColumnProps(199)=   "Column(49)._WidthInPix=3175"
      Splits(0)._ColumnProps(200)=   "Column(49).Order=50"
      Splits(0)._ColumnProps(201)=   "Column(50).Width=3281"
      Splits(0)._ColumnProps(202)=   "Column(50).DividerColor=0"
      Splits(0)._ColumnProps(203)=   "Column(50)._WidthInPix=3175"
      Splits(0)._ColumnProps(204)=   "Column(50).Order=51"
      Splits(0)._ColumnProps(205)=   "Column(51).Width=3281"
      Splits(0)._ColumnProps(206)=   "Column(51).DividerColor=0"
      Splits(0)._ColumnProps(207)=   "Column(51)._WidthInPix=3175"
      Splits(0)._ColumnProps(208)=   "Column(51).Order=52"
      Splits(0)._ColumnProps(209)=   "Column(52).Width=3281"
      Splits(0)._ColumnProps(210)=   "Column(52).DividerColor=0"
      Splits(0)._ColumnProps(211)=   "Column(52)._WidthInPix=3175"
      Splits(0)._ColumnProps(212)=   "Column(52).Order=53"
      Splits(0)._ColumnProps(213)=   "Column(53).Width=3281"
      Splits(0)._ColumnProps(214)=   "Column(53).DividerColor=0"
      Splits(0)._ColumnProps(215)=   "Column(53)._WidthInPix=3175"
      Splits(0)._ColumnProps(216)=   "Column(53).Order=54"
      Splits(0)._ColumnProps(217)=   "Column(54).Width=3281"
      Splits(0)._ColumnProps(218)=   "Column(54).DividerColor=0"
      Splits(0)._ColumnProps(219)=   "Column(54)._WidthInPix=3175"
      Splits(0)._ColumnProps(220)=   "Column(54).Order=55"
      Splits(0)._ColumnProps(221)=   "Column(55).Width=3281"
      Splits(0)._ColumnProps(222)=   "Column(55).DividerColor=0"
      Splits(0)._ColumnProps(223)=   "Column(55)._WidthInPix=3175"
      Splits(0)._ColumnProps(224)=   "Column(55).Order=56"
      Splits(0)._ColumnProps(225)=   "Column(56).Width=3281"
      Splits(0)._ColumnProps(226)=   "Column(56).DividerColor=0"
      Splits(0)._ColumnProps(227)=   "Column(56)._WidthInPix=3175"
      Splits(0)._ColumnProps(228)=   "Column(56).Order=57"
      Splits(0)._ColumnProps(229)=   "Column(57).Width=3281"
      Splits(0)._ColumnProps(230)=   "Column(57).DividerColor=0"
      Splits(0)._ColumnProps(231)=   "Column(57)._WidthInPix=3175"
      Splits(0)._ColumnProps(232)=   "Column(57).Order=58"
      Splits(0)._ColumnProps(233)=   "Column(58).Width=3281"
      Splits(0)._ColumnProps(234)=   "Column(58).DividerColor=0"
      Splits(0)._ColumnProps(235)=   "Column(58)._WidthInPix=3175"
      Splits(0)._ColumnProps(236)=   "Column(58).Order=59"
      Splits(0)._ColumnProps(237)=   "Column(59).Width=3281"
      Splits(0)._ColumnProps(238)=   "Column(59).DividerColor=0"
      Splits(0)._ColumnProps(239)=   "Column(59)._WidthInPix=3175"
      Splits(0)._ColumnProps(240)=   "Column(59).Order=60"
      Splits(0)._ColumnProps(241)=   "Column(60).Width=3281"
      Splits(0)._ColumnProps(242)=   "Column(60).DividerColor=0"
      Splits(0)._ColumnProps(243)=   "Column(60)._WidthInPix=3175"
      Splits(0)._ColumnProps(244)=   "Column(60).Order=61"
      Splits(0)._ColumnProps(245)=   "Column(61).Width=3281"
      Splits(0)._ColumnProps(246)=   "Column(61).DividerColor=0"
      Splits(0)._ColumnProps(247)=   "Column(61)._WidthInPix=3175"
      Splits(0)._ColumnProps(248)=   "Column(61).Order=62"
      Splits(0)._ColumnProps(249)=   "Column(62).Width=3281"
      Splits(0)._ColumnProps(250)=   "Column(62).DividerColor=0"
      Splits(0)._ColumnProps(251)=   "Column(62)._WidthInPix=3175"
      Splits(0)._ColumnProps(252)=   "Column(62).Order=63"
      Splits(0)._ColumnProps(253)=   "Column(63).Width=3281"
      Splits(0)._ColumnProps(254)=   "Column(63).DividerColor=0"
      Splits(0)._ColumnProps(255)=   "Column(63)._WidthInPix=3175"
      Splits(0)._ColumnProps(256)=   "Column(63).Order=64"
      Splits(0)._ColumnProps(257)=   "Column(64).Width=3281"
      Splits(0)._ColumnProps(258)=   "Column(64).DividerColor=0"
      Splits(0)._ColumnProps(259)=   "Column(64)._WidthInPix=3175"
      Splits(0)._ColumnProps(260)=   "Column(64).Order=65"
      Splits(0)._ColumnProps(261)=   "Column(65).Width=3281"
      Splits(0)._ColumnProps(262)=   "Column(65).DividerColor=0"
      Splits(0)._ColumnProps(263)=   "Column(65)._WidthInPix=3175"
      Splits(0)._ColumnProps(264)=   "Column(65).Order=66"
      Splits(0)._ColumnProps(265)=   "Column(66).Width=3281"
      Splits(0)._ColumnProps(266)=   "Column(66).DividerColor=0"
      Splits(0)._ColumnProps(267)=   "Column(66)._WidthInPix=3175"
      Splits(0)._ColumnProps(268)=   "Column(66).Order=67"
      Splits(0)._ColumnProps(269)=   "Column(67).Width=3281"
      Splits(0)._ColumnProps(270)=   "Column(67).DividerColor=0"
      Splits(0)._ColumnProps(271)=   "Column(67)._WidthInPix=3175"
      Splits(0)._ColumnProps(272)=   "Column(67).Order=68"
      Splits(0)._ColumnProps(273)=   "Column(68).Width=3281"
      Splits(0)._ColumnProps(274)=   "Column(68).DividerColor=0"
      Splits(0)._ColumnProps(275)=   "Column(68)._WidthInPix=3175"
      Splits(0)._ColumnProps(276)=   "Column(68).Order=69"
      Splits(0)._ColumnProps(277)=   "Column(69).Width=3281"
      Splits(0)._ColumnProps(278)=   "Column(69).DividerColor=0"
      Splits(0)._ColumnProps(279)=   "Column(69)._WidthInPix=3175"
      Splits(0)._ColumnProps(280)=   "Column(69).Order=70"
      Splits(0)._ColumnProps(281)=   "Column(70).Width=3281"
      Splits(0)._ColumnProps(282)=   "Column(70).DividerColor=0"
      Splits(0)._ColumnProps(283)=   "Column(70)._WidthInPix=3175"
      Splits(0)._ColumnProps(284)=   "Column(70).Order=71"
      Splits(0)._ColumnProps(285)=   "Column(71).Width=3281"
      Splits(0)._ColumnProps(286)=   "Column(71).DividerColor=0"
      Splits(0)._ColumnProps(287)=   "Column(71)._WidthInPix=3175"
      Splits(0)._ColumnProps(288)=   "Column(71).Order=72"
      Splits(0)._ColumnProps(289)=   "Column(72).Width=3281"
      Splits(0)._ColumnProps(290)=   "Column(72).DividerColor=0"
      Splits(0)._ColumnProps(291)=   "Column(72)._WidthInPix=3175"
      Splits(0)._ColumnProps(292)=   "Column(72).Order=73"
      Splits(0)._ColumnProps(293)=   "Column(73).Width=3281"
      Splits(0)._ColumnProps(294)=   "Column(73).DividerColor=0"
      Splits(0)._ColumnProps(295)=   "Column(73)._WidthInPix=3175"
      Splits(0)._ColumnProps(296)=   "Column(73).Order=74"
      Splits(0)._ColumnProps(297)=   "Column(74).Width=3281"
      Splits(0)._ColumnProps(298)=   "Column(74).DividerColor=0"
      Splits(0)._ColumnProps(299)=   "Column(74)._WidthInPix=3175"
      Splits(0)._ColumnProps(300)=   "Column(74).Order=75"
      Splits(0)._ColumnProps(301)=   "Column(75).Width=3281"
      Splits(0)._ColumnProps(302)=   "Column(75).DividerColor=0"
      Splits(0)._ColumnProps(303)=   "Column(75)._WidthInPix=3175"
      Splits(0)._ColumnProps(304)=   "Column(75).Order=76"
      Splits(0)._ColumnProps(305)=   "Column(76).Width=3281"
      Splits(0)._ColumnProps(306)=   "Column(76).DividerColor=0"
      Splits(0)._ColumnProps(307)=   "Column(76)._WidthInPix=3175"
      Splits(0)._ColumnProps(308)=   "Column(76).Order=77"
      Splits(0)._ColumnProps(309)=   "Column(77).Width=3281"
      Splits(0)._ColumnProps(310)=   "Column(77).DividerColor=0"
      Splits(0)._ColumnProps(311)=   "Column(77)._WidthInPix=3175"
      Splits(0)._ColumnProps(312)=   "Column(77).Order=78"
      Splits(0)._ColumnProps(313)=   "Column(78).Width=3281"
      Splits(0)._ColumnProps(314)=   "Column(78).DividerColor=0"
      Splits(0)._ColumnProps(315)=   "Column(78)._WidthInPix=3175"
      Splits(0)._ColumnProps(316)=   "Column(78).Order=79"
      Splits(0)._ColumnProps(317)=   "Column(79).Width=3281"
      Splits(0)._ColumnProps(318)=   "Column(79).DividerColor=0"
      Splits(0)._ColumnProps(319)=   "Column(79)._WidthInPix=3175"
      Splits(0)._ColumnProps(320)=   "Column(79).Order=80"
      Splits(0)._ColumnProps(321)=   "Column(80).Width=3281"
      Splits(0)._ColumnProps(322)=   "Column(80).DividerColor=0"
      Splits(0)._ColumnProps(323)=   "Column(80)._WidthInPix=3175"
      Splits(0)._ColumnProps(324)=   "Column(80).Order=81"
      Splits(0)._ColumnProps(325)=   "Column(81).Width=3281"
      Splits(0)._ColumnProps(326)=   "Column(81).DividerColor=0"
      Splits(0)._ColumnProps(327)=   "Column(81)._WidthInPix=3175"
      Splits(0)._ColumnProps(328)=   "Column(81).Order=82"
      Splits(0)._ColumnProps(329)=   "Column(82).Width=3281"
      Splits(0)._ColumnProps(330)=   "Column(82).DividerColor=0"
      Splits(0)._ColumnProps(331)=   "Column(82)._WidthInPix=3175"
      Splits(0)._ColumnProps(332)=   "Column(82).Order=83"
      Splits(0)._ColumnProps(333)=   "Column(83).Width=3281"
      Splits(0)._ColumnProps(334)=   "Column(83).DividerColor=0"
      Splits(0)._ColumnProps(335)=   "Column(83)._WidthInPix=3175"
      Splits(0)._ColumnProps(336)=   "Column(83).Order=84"
      Splits(0)._ColumnProps(337)=   "Column(84).Width=3281"
      Splits(0)._ColumnProps(338)=   "Column(84).DividerColor=0"
      Splits(0)._ColumnProps(339)=   "Column(84)._WidthInPix=3175"
      Splits(0)._ColumnProps(340)=   "Column(84).Order=85"
      Splits(0)._ColumnProps(341)=   "Column(85).Width=3281"
      Splits(0)._ColumnProps(342)=   "Column(85).DividerColor=0"
      Splits(0)._ColumnProps(343)=   "Column(85)._WidthInPix=3175"
      Splits(0)._ColumnProps(344)=   "Column(85).Order=86"
      Splits(0)._ColumnProps(345)=   "Column(86).Width=3281"
      Splits(0)._ColumnProps(346)=   "Column(86).DividerColor=0"
      Splits(0)._ColumnProps(347)=   "Column(86)._WidthInPix=3175"
      Splits(0)._ColumnProps(348)=   "Column(86).Order=87"
      Splits(0)._ColumnProps(349)=   "Column(87).Width=3281"
      Splits(0)._ColumnProps(350)=   "Column(87).DividerColor=0"
      Splits(0)._ColumnProps(351)=   "Column(87)._WidthInPix=3175"
      Splits(0)._ColumnProps(352)=   "Column(87).Order=88"
      Splits(0)._ColumnProps(353)=   "Column(88).Width=3281"
      Splits(0)._ColumnProps(354)=   "Column(88).DividerColor=0"
      Splits(0)._ColumnProps(355)=   "Column(88)._WidthInPix=3175"
      Splits(0)._ColumnProps(356)=   "Column(88).Order=89"
      Splits(0)._ColumnProps(357)=   "Column(89).Width=3281"
      Splits(0)._ColumnProps(358)=   "Column(89).DividerColor=0"
      Splits(0)._ColumnProps(359)=   "Column(89)._WidthInPix=3175"
      Splits(0)._ColumnProps(360)=   "Column(89).Order=90"
      Splits(0)._ColumnProps(361)=   "Column(90).Width=3281"
      Splits(0)._ColumnProps(362)=   "Column(90).DividerColor=0"
      Splits(0)._ColumnProps(363)=   "Column(90)._WidthInPix=3175"
      Splits(0)._ColumnProps(364)=   "Column(90).Order=91"
      Splits(0)._ColumnProps(365)=   "Column(91).Width=3281"
      Splits(0)._ColumnProps(366)=   "Column(91).DividerColor=0"
      Splits(0)._ColumnProps(367)=   "Column(91)._WidthInPix=3175"
      Splits(0)._ColumnProps(368)=   "Column(91).Order=92"
      Splits(0)._ColumnProps(369)=   "Column(92).Width=3281"
      Splits(0)._ColumnProps(370)=   "Column(92).DividerColor=0"
      Splits(0)._ColumnProps(371)=   "Column(92)._WidthInPix=3175"
      Splits(0)._ColumnProps(372)=   "Column(92).Order=93"
      Splits(0)._ColumnProps(373)=   "Column(93).Width=3281"
      Splits(0)._ColumnProps(374)=   "Column(93).DividerColor=0"
      Splits(0)._ColumnProps(375)=   "Column(93)._WidthInPix=3175"
      Splits(0)._ColumnProps(376)=   "Column(93).Order=94"
      Splits(0)._ColumnProps(377)=   "Column(94).Width=3281"
      Splits(0)._ColumnProps(378)=   "Column(94).DividerColor=0"
      Splits(0)._ColumnProps(379)=   "Column(94)._WidthInPix=3175"
      Splits(0)._ColumnProps(380)=   "Column(94).Order=95"
      Splits(0)._ColumnProps(381)=   "Column(95).Width=3281"
      Splits(0)._ColumnProps(382)=   "Column(95).DividerColor=0"
      Splits(0)._ColumnProps(383)=   "Column(95)._WidthInPix=3175"
      Splits(0)._ColumnProps(384)=   "Column(95).Order=96"
      Splits(0)._ColumnProps(385)=   "Column(96).Width=3281"
      Splits(0)._ColumnProps(386)=   "Column(96).DividerColor=0"
      Splits(0)._ColumnProps(387)=   "Column(96)._WidthInPix=3175"
      Splits(0)._ColumnProps(388)=   "Column(96).Order=97"
      Splits(0)._ColumnProps(389)=   "Column(97).Width=3281"
      Splits(0)._ColumnProps(390)=   "Column(97).DividerColor=0"
      Splits(0)._ColumnProps(391)=   "Column(97)._WidthInPix=3175"
      Splits(0)._ColumnProps(392)=   "Column(97).Order=98"
      Splits(0)._ColumnProps(393)=   "Column(98).Width=3281"
      Splits(0)._ColumnProps(394)=   "Column(98).DividerColor=0"
      Splits(0)._ColumnProps(395)=   "Column(98)._WidthInPix=3175"
      Splits(0)._ColumnProps(396)=   "Column(98).Order=99"
      Splits(0)._ColumnProps(397)=   "Column(99).Width=3281"
      Splits(0)._ColumnProps(398)=   "Column(99).DividerColor=0"
      Splits(0)._ColumnProps(399)=   "Column(99)._WidthInPix=3175"
      Splits(0)._ColumnProps(400)=   "Column(99).Order=100"
      Splits(0)._ColumnProps(401)=   "Column(100).Width=3281"
      Splits(0)._ColumnProps(402)=   "Column(100).DividerColor=0"
      Splits(0)._ColumnProps(403)=   "Column(100)._WidthInPix=3175"
      Splits(0)._ColumnProps(404)=   "Column(100).Order=101"
      Splits(0)._ColumnProps(405)=   "Column(101).Width=3281"
      Splits(0)._ColumnProps(406)=   "Column(101).DividerColor=0"
      Splits(0)._ColumnProps(407)=   "Column(101)._WidthInPix=3175"
      Splits(0)._ColumnProps(408)=   "Column(101).Order=102"
      Splits(0)._ColumnProps(409)=   "Column(102).Width=3281"
      Splits(0)._ColumnProps(410)=   "Column(102).DividerColor=0"
      Splits(0)._ColumnProps(411)=   "Column(102)._WidthInPix=3175"
      Splits(0)._ColumnProps(412)=   "Column(102).Order=103"
      Splits(0)._ColumnProps(413)=   "Column(103).Width=3281"
      Splits(0)._ColumnProps(414)=   "Column(103).DividerColor=0"
      Splits(0)._ColumnProps(415)=   "Column(103)._WidthInPix=3175"
      Splits(0)._ColumnProps(416)=   "Column(103).Order=104"
      Splits(0)._ColumnProps(417)=   "Column(104).Width=3281"
      Splits(0)._ColumnProps(418)=   "Column(104).DividerColor=0"
      Splits(0)._ColumnProps(419)=   "Column(104)._WidthInPix=3175"
      Splits(0)._ColumnProps(420)=   "Column(104).Order=105"
      Splits(0)._ColumnProps(421)=   "Column(105).Width=3281"
      Splits(0)._ColumnProps(422)=   "Column(105).DividerColor=0"
      Splits(0)._ColumnProps(423)=   "Column(105)._WidthInPix=3175"
      Splits(0)._ColumnProps(424)=   "Column(105).Order=106"
      Splits(0)._ColumnProps(425)=   "Column(106).Width=3281"
      Splits(0)._ColumnProps(426)=   "Column(106).DividerColor=0"
      Splits(0)._ColumnProps(427)=   "Column(106)._WidthInPix=3175"
      Splits(0)._ColumnProps(428)=   "Column(106).Order=107"
      Splits(0)._ColumnProps(429)=   "Column(107).Width=3281"
      Splits(0)._ColumnProps(430)=   "Column(107).DividerColor=0"
      Splits(0)._ColumnProps(431)=   "Column(107)._WidthInPix=3175"
      Splits(0)._ColumnProps(432)=   "Column(107).Order=108"
      Splits(0)._ColumnProps(433)=   "Column(108).Width=3281"
      Splits(0)._ColumnProps(434)=   "Column(108).DividerColor=0"
      Splits(0)._ColumnProps(435)=   "Column(108)._WidthInPix=3175"
      Splits(0)._ColumnProps(436)=   "Column(108).Order=109"
      Splits(0)._ColumnProps(437)=   "Column(109).Width=3281"
      Splits(0)._ColumnProps(438)=   "Column(109).DividerColor=0"
      Splits(0)._ColumnProps(439)=   "Column(109)._WidthInPix=3175"
      Splits(0)._ColumnProps(440)=   "Column(109).Order=110"
      Splits(0)._ColumnProps(441)=   "Column(110).Width=3281"
      Splits(0)._ColumnProps(442)=   "Column(110).DividerColor=0"
      Splits(0)._ColumnProps(443)=   "Column(110)._WidthInPix=3175"
      Splits(0)._ColumnProps(444)=   "Column(110).Order=111"
      Splits(0)._ColumnProps(445)=   "Column(111).Width=3281"
      Splits(0)._ColumnProps(446)=   "Column(111).DividerColor=0"
      Splits(0)._ColumnProps(447)=   "Column(111)._WidthInPix=3175"
      Splits(0)._ColumnProps(448)=   "Column(111).Order=112"
      Splits(0)._ColumnProps(449)=   "Column(112).Width=3281"
      Splits(0)._ColumnProps(450)=   "Column(112).DividerColor=0"
      Splits(0)._ColumnProps(451)=   "Column(112)._WidthInPix=3175"
      Splits(0)._ColumnProps(452)=   "Column(112).Order=113"
      Splits(0)._ColumnProps(453)=   "Column(113).Width=3281"
      Splits(0)._ColumnProps(454)=   "Column(113).DividerColor=0"
      Splits(0)._ColumnProps(455)=   "Column(113)._WidthInPix=3175"
      Splits(0)._ColumnProps(456)=   "Column(113).Order=114"
      Splits(0)._ColumnProps(457)=   "Column(114).Width=3281"
      Splits(0)._ColumnProps(458)=   "Column(114).DividerColor=0"
      Splits(0)._ColumnProps(459)=   "Column(114)._WidthInPix=3175"
      Splits(0)._ColumnProps(460)=   "Column(114).Order=115"
      Splits(0)._ColumnProps(461)=   "Column(115).Width=3281"
      Splits(0)._ColumnProps(462)=   "Column(115).DividerColor=0"
      Splits(0)._ColumnProps(463)=   "Column(115)._WidthInPix=3175"
      Splits(0)._ColumnProps(464)=   "Column(115).Order=116"
      Splits(0)._ColumnProps(465)=   "Column(116).Width=3281"
      Splits(0)._ColumnProps(466)=   "Column(116).DividerColor=0"
      Splits(0)._ColumnProps(467)=   "Column(116)._WidthInPix=3175"
      Splits(0)._ColumnProps(468)=   "Column(116).Order=117"
      Splits(0)._ColumnProps(469)=   "Column(117).Width=3281"
      Splits(0)._ColumnProps(470)=   "Column(117).DividerColor=0"
      Splits(0)._ColumnProps(471)=   "Column(117)._WidthInPix=3175"
      Splits(0)._ColumnProps(472)=   "Column(117).Order=118"
      Splits(0)._ColumnProps(473)=   "Column(118).Width=3281"
      Splits(0)._ColumnProps(474)=   "Column(118).DividerColor=0"
      Splits(0)._ColumnProps(475)=   "Column(118)._WidthInPix=3175"
      Splits(0)._ColumnProps(476)=   "Column(118).Order=119"
      Splits(0)._ColumnProps(477)=   "Column(119).Width=3281"
      Splits(0)._ColumnProps(478)=   "Column(119).DividerColor=0"
      Splits(0)._ColumnProps(479)=   "Column(119)._WidthInPix=3175"
      Splits(0)._ColumnProps(480)=   "Column(119).Order=120"
      Splits(0)._ColumnProps(481)=   "Column(120).Width=3281"
      Splits(0)._ColumnProps(482)=   "Column(120).DividerColor=0"
      Splits(0)._ColumnProps(483)=   "Column(120)._WidthInPix=3175"
      Splits(0)._ColumnProps(484)=   "Column(120).Order=121"
      Splits(0)._ColumnProps(485)=   "Column(121).Width=3281"
      Splits(0)._ColumnProps(486)=   "Column(121).DividerColor=0"
      Splits(0)._ColumnProps(487)=   "Column(121)._WidthInPix=3175"
      Splits(0)._ColumnProps(488)=   "Column(121).Order=122"
      Splits(0)._ColumnProps(489)=   "Column(122).Width=3281"
      Splits(0)._ColumnProps(490)=   "Column(122).DividerColor=0"
      Splits(0)._ColumnProps(491)=   "Column(122)._WidthInPix=3175"
      Splits(0)._ColumnProps(492)=   "Column(122).Order=123"
      Splits(0)._ColumnProps(493)=   "Column(123).Width=3281"
      Splits(0)._ColumnProps(494)=   "Column(123).DividerColor=0"
      Splits(0)._ColumnProps(495)=   "Column(123)._WidthInPix=3175"
      Splits(0)._ColumnProps(496)=   "Column(123).Order=124"
      Splits(0)._ColumnProps(497)=   "Column(124).Width=3281"
      Splits(0)._ColumnProps(498)=   "Column(124).DividerColor=0"
      Splits(0)._ColumnProps(499)=   "Column(124)._WidthInPix=3175"
      Splits(0)._ColumnProps(500)=   "Column(124).Order=125"
      Splits(0)._ColumnProps(501)=   "Column(125).Width=3281"
      Splits(0)._ColumnProps(502)=   "Column(125).DividerColor=0"
      Splits(0)._ColumnProps(503)=   "Column(125)._WidthInPix=3175"
      Splits(0)._ColumnProps(504)=   "Column(125).Order=126"
      Splits(0)._ColumnProps(505)=   "Column(126).Width=3281"
      Splits(0)._ColumnProps(506)=   "Column(126).DividerColor=0"
      Splits(0)._ColumnProps(507)=   "Column(126)._WidthInPix=3175"
      Splits(0)._ColumnProps(508)=   "Column(126).Order=127"
      Splits(0)._ColumnProps(509)=   "Column(127).Width=3281"
      Splits(0)._ColumnProps(510)=   "Column(127).DividerColor=0"
      Splits(0)._ColumnProps(511)=   "Column(127)._WidthInPix=3175"
      Splits(0)._ColumnProps(512)=   "Column(127).Order=128"
      Splits(0)._ColumnProps(513)=   "Column(128).Width=3281"
      Splits(0)._ColumnProps(514)=   "Column(128).DividerColor=0"
      Splits(0)._ColumnProps(515)=   "Column(128)._WidthInPix=3175"
      Splits(0)._ColumnProps(516)=   "Column(128).Order=129"
      Splits(0)._ColumnProps(517)=   "Column(129).Width=3281"
      Splits(0)._ColumnProps(518)=   "Column(129).DividerColor=0"
      Splits(0)._ColumnProps(519)=   "Column(129)._WidthInPix=3175"
      Splits(0)._ColumnProps(520)=   "Column(129).Order=130"
      Splits(0)._ColumnProps(521)=   "Column(130).Width=3281"
      Splits(0)._ColumnProps(522)=   "Column(130).DividerColor=0"
      Splits(0)._ColumnProps(523)=   "Column(130)._WidthInPix=3175"
      Splits(0)._ColumnProps(524)=   "Column(130).Order=131"
      Splits(0)._ColumnProps(525)=   "Column(131).Width=3281"
      Splits(0)._ColumnProps(526)=   "Column(131).DividerColor=0"
      Splits(0)._ColumnProps(527)=   "Column(131)._WidthInPix=3175"
      Splits(0)._ColumnProps(528)=   "Column(131).Order=132"
      Splits(0)._ColumnProps(529)=   "Column(132).Width=3281"
      Splits(0)._ColumnProps(530)=   "Column(132).DividerColor=0"
      Splits(0)._ColumnProps(531)=   "Column(132)._WidthInPix=3175"
      Splits(0)._ColumnProps(532)=   "Column(132).Order=133"
      Splits(0)._ColumnProps(533)=   "Column(133).Width=3281"
      Splits(0)._ColumnProps(534)=   "Column(133).DividerColor=0"
      Splits(0)._ColumnProps(535)=   "Column(133)._WidthInPix=3175"
      Splits(0)._ColumnProps(536)=   "Column(133).Order=134"
      Splits(0)._ColumnProps(537)=   "Column(134).Width=3281"
      Splits(0)._ColumnProps(538)=   "Column(134).DividerColor=0"
      Splits(0)._ColumnProps(539)=   "Column(134)._WidthInPix=3175"
      Splits(0)._ColumnProps(540)=   "Column(134).Order=135"
      Splits(0)._ColumnProps(541)=   "Column(135).Width=3281"
      Splits(0)._ColumnProps(542)=   "Column(135).DividerColor=0"
      Splits(0)._ColumnProps(543)=   "Column(135)._WidthInPix=3175"
      Splits(0)._ColumnProps(544)=   "Column(135).Order=136"
      Splits(0)._ColumnProps(545)=   "Column(136).Width=3281"
      Splits(0)._ColumnProps(546)=   "Column(136).DividerColor=0"
      Splits(0)._ColumnProps(547)=   "Column(136)._WidthInPix=3175"
      Splits(0)._ColumnProps(548)=   "Column(136).Order=137"
      Splits(0)._ColumnProps(549)=   "Column(137).Width=3281"
      Splits(0)._ColumnProps(550)=   "Column(137).DividerColor=0"
      Splits(0)._ColumnProps(551)=   "Column(137)._WidthInPix=3175"
      Splits(0)._ColumnProps(552)=   "Column(137).Order=138"
      Splits(1)._UserFlags=   0
      Splits(1).RecordSelectorWidth=   503
      Splits(1)._SavedRecordSelectors=   0   'False
      Splits(1).AlternatingRowStyle=   -1  'True
      Splits(1).DividerColor=   14933984
      Splits(1).FilterBar=   -1  'True
      Splits(1).SpringMode=   0   'False
      Splits(1)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(1)._ColumnProps(0)=   "Columns.Count=138"
      Splits(1)._ColumnProps(1)=   "Column(0).Width=3281"
      Splits(1)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(1)._ColumnProps(3)=   "Column(0)._WidthInPix=3175"
      Splits(1)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(1)._ColumnProps(5)=   "Column(1).Width=3281"
      Splits(1)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(1)._ColumnProps(7)=   "Column(1)._WidthInPix=3175"
      Splits(1)._ColumnProps(8)=   "Column(1).Order=2"
      Splits(1)._ColumnProps(9)=   "Column(2).Width=3281"
      Splits(1)._ColumnProps(10)=   "Column(2).DividerColor=0"
      Splits(1)._ColumnProps(11)=   "Column(2)._WidthInPix=3175"
      Splits(1)._ColumnProps(12)=   "Column(2).Order=3"
      Splits(1)._ColumnProps(13)=   "Column(3).Width=3281"
      Splits(1)._ColumnProps(14)=   "Column(3).DividerColor=0"
      Splits(1)._ColumnProps(15)=   "Column(3)._WidthInPix=3175"
      Splits(1)._ColumnProps(16)=   "Column(3).Order=4"
      Splits(1)._ColumnProps(17)=   "Column(4).Width=3281"
      Splits(1)._ColumnProps(18)=   "Column(4).DividerColor=0"
      Splits(1)._ColumnProps(19)=   "Column(4)._WidthInPix=3175"
      Splits(1)._ColumnProps(20)=   "Column(4).Order=5"
      Splits(1)._ColumnProps(21)=   "Column(5).Width=3281"
      Splits(1)._ColumnProps(22)=   "Column(5).DividerColor=0"
      Splits(1)._ColumnProps(23)=   "Column(5)._WidthInPix=3175"
      Splits(1)._ColumnProps(24)=   "Column(5).Order=6"
      Splits(1)._ColumnProps(25)=   "Column(6).Width=3281"
      Splits(1)._ColumnProps(26)=   "Column(6).DividerColor=0"
      Splits(1)._ColumnProps(27)=   "Column(6)._WidthInPix=3175"
      Splits(1)._ColumnProps(28)=   "Column(6).Order=7"
      Splits(1)._ColumnProps(29)=   "Column(7).Width=3281"
      Splits(1)._ColumnProps(30)=   "Column(7).DividerColor=0"
      Splits(1)._ColumnProps(31)=   "Column(7)._WidthInPix=3175"
      Splits(1)._ColumnProps(32)=   "Column(7).Order=8"
      Splits(1)._ColumnProps(33)=   "Column(8).Width=3281"
      Splits(1)._ColumnProps(34)=   "Column(8).DividerColor=0"
      Splits(1)._ColumnProps(35)=   "Column(8)._WidthInPix=3175"
      Splits(1)._ColumnProps(36)=   "Column(8).Order=9"
      Splits(1)._ColumnProps(37)=   "Column(9).Width=3281"
      Splits(1)._ColumnProps(38)=   "Column(9).DividerColor=0"
      Splits(1)._ColumnProps(39)=   "Column(9)._WidthInPix=3175"
      Splits(1)._ColumnProps(40)=   "Column(9).Order=10"
      Splits(1)._ColumnProps(41)=   "Column(10).Width=3281"
      Splits(1)._ColumnProps(42)=   "Column(10).DividerColor=0"
      Splits(1)._ColumnProps(43)=   "Column(10)._WidthInPix=3175"
      Splits(1)._ColumnProps(44)=   "Column(10).Order=11"
      Splits(1)._ColumnProps(45)=   "Column(11).Width=3281"
      Splits(1)._ColumnProps(46)=   "Column(11).DividerColor=0"
      Splits(1)._ColumnProps(47)=   "Column(11)._WidthInPix=3175"
      Splits(1)._ColumnProps(48)=   "Column(11).Order=12"
      Splits(1)._ColumnProps(49)=   "Column(12).Width=3281"
      Splits(1)._ColumnProps(50)=   "Column(12).DividerColor=0"
      Splits(1)._ColumnProps(51)=   "Column(12)._WidthInPix=3175"
      Splits(1)._ColumnProps(52)=   "Column(12).Order=13"
      Splits(1)._ColumnProps(53)=   "Column(13).Width=3281"
      Splits(1)._ColumnProps(54)=   "Column(13).DividerColor=0"
      Splits(1)._ColumnProps(55)=   "Column(13)._WidthInPix=3175"
      Splits(1)._ColumnProps(56)=   "Column(13).Order=14"
      Splits(1)._ColumnProps(57)=   "Column(14).Width=3281"
      Splits(1)._ColumnProps(58)=   "Column(14).DividerColor=0"
      Splits(1)._ColumnProps(59)=   "Column(14)._WidthInPix=3175"
      Splits(1)._ColumnProps(60)=   "Column(14).Order=15"
      Splits(1)._ColumnProps(61)=   "Column(15).Width=3281"
      Splits(1)._ColumnProps(62)=   "Column(15).DividerColor=0"
      Splits(1)._ColumnProps(63)=   "Column(15)._WidthInPix=3175"
      Splits(1)._ColumnProps(64)=   "Column(15).Order=16"
      Splits(1)._ColumnProps(65)=   "Column(16).Width=3281"
      Splits(1)._ColumnProps(66)=   "Column(16).DividerColor=0"
      Splits(1)._ColumnProps(67)=   "Column(16)._WidthInPix=3175"
      Splits(1)._ColumnProps(68)=   "Column(16).Order=17"
      Splits(1)._ColumnProps(69)=   "Column(17).Width=3281"
      Splits(1)._ColumnProps(70)=   "Column(17).DividerColor=0"
      Splits(1)._ColumnProps(71)=   "Column(17)._WidthInPix=3175"
      Splits(1)._ColumnProps(72)=   "Column(17).Order=18"
      Splits(1)._ColumnProps(73)=   "Column(18).Width=3281"
      Splits(1)._ColumnProps(74)=   "Column(18).DividerColor=0"
      Splits(1)._ColumnProps(75)=   "Column(18)._WidthInPix=3175"
      Splits(1)._ColumnProps(76)=   "Column(18).Order=19"
      Splits(1)._ColumnProps(77)=   "Column(19).Width=3281"
      Splits(1)._ColumnProps(78)=   "Column(19).DividerColor=0"
      Splits(1)._ColumnProps(79)=   "Column(19)._WidthInPix=3175"
      Splits(1)._ColumnProps(80)=   "Column(19).Order=20"
      Splits(1)._ColumnProps(81)=   "Column(20).Width=3281"
      Splits(1)._ColumnProps(82)=   "Column(20).DividerColor=0"
      Splits(1)._ColumnProps(83)=   "Column(20)._WidthInPix=3175"
      Splits(1)._ColumnProps(84)=   "Column(20).Order=21"
      Splits(1)._ColumnProps(85)=   "Column(21).Width=3281"
      Splits(1)._ColumnProps(86)=   "Column(21).DividerColor=0"
      Splits(1)._ColumnProps(87)=   "Column(21)._WidthInPix=3175"
      Splits(1)._ColumnProps(88)=   "Column(21).Order=22"
      Splits(1)._ColumnProps(89)=   "Column(22).Width=3281"
      Splits(1)._ColumnProps(90)=   "Column(22).DividerColor=0"
      Splits(1)._ColumnProps(91)=   "Column(22)._WidthInPix=3175"
      Splits(1)._ColumnProps(92)=   "Column(22).Order=23"
      Splits(1)._ColumnProps(93)=   "Column(23).Width=3281"
      Splits(1)._ColumnProps(94)=   "Column(23).DividerColor=0"
      Splits(1)._ColumnProps(95)=   "Column(23)._WidthInPix=3175"
      Splits(1)._ColumnProps(96)=   "Column(23).Order=24"
      Splits(1)._ColumnProps(97)=   "Column(24).Width=3281"
      Splits(1)._ColumnProps(98)=   "Column(24).DividerColor=0"
      Splits(1)._ColumnProps(99)=   "Column(24)._WidthInPix=3175"
      Splits(1)._ColumnProps(100)=   "Column(24).Order=25"
      Splits(1)._ColumnProps(101)=   "Column(25).Width=3281"
      Splits(1)._ColumnProps(102)=   "Column(25).DividerColor=0"
      Splits(1)._ColumnProps(103)=   "Column(25)._WidthInPix=3175"
      Splits(1)._ColumnProps(104)=   "Column(25).Order=26"
      Splits(1)._ColumnProps(105)=   "Column(26).Width=3281"
      Splits(1)._ColumnProps(106)=   "Column(26).DividerColor=0"
      Splits(1)._ColumnProps(107)=   "Column(26)._WidthInPix=3175"
      Splits(1)._ColumnProps(108)=   "Column(26).Order=27"
      Splits(1)._ColumnProps(109)=   "Column(27).Width=3281"
      Splits(1)._ColumnProps(110)=   "Column(27).DividerColor=0"
      Splits(1)._ColumnProps(111)=   "Column(27)._WidthInPix=3175"
      Splits(1)._ColumnProps(112)=   "Column(27).Order=28"
      Splits(1)._ColumnProps(113)=   "Column(28).Width=3281"
      Splits(1)._ColumnProps(114)=   "Column(28).DividerColor=0"
      Splits(1)._ColumnProps(115)=   "Column(28)._WidthInPix=3175"
      Splits(1)._ColumnProps(116)=   "Column(28).Order=29"
      Splits(1)._ColumnProps(117)=   "Column(29).Width=3281"
      Splits(1)._ColumnProps(118)=   "Column(29).DividerColor=0"
      Splits(1)._ColumnProps(119)=   "Column(29)._WidthInPix=3175"
      Splits(1)._ColumnProps(120)=   "Column(29).Order=30"
      Splits(1)._ColumnProps(121)=   "Column(30).Width=3281"
      Splits(1)._ColumnProps(122)=   "Column(30).DividerColor=0"
      Splits(1)._ColumnProps(123)=   "Column(30)._WidthInPix=3175"
      Splits(1)._ColumnProps(124)=   "Column(30).Order=31"
      Splits(1)._ColumnProps(125)=   "Column(31).Width=3281"
      Splits(1)._ColumnProps(126)=   "Column(31).DividerColor=0"
      Splits(1)._ColumnProps(127)=   "Column(31)._WidthInPix=3175"
      Splits(1)._ColumnProps(128)=   "Column(31).Order=32"
      Splits(1)._ColumnProps(129)=   "Column(32).Width=3281"
      Splits(1)._ColumnProps(130)=   "Column(32).DividerColor=0"
      Splits(1)._ColumnProps(131)=   "Column(32)._WidthInPix=3175"
      Splits(1)._ColumnProps(132)=   "Column(32).Order=33"
      Splits(1)._ColumnProps(133)=   "Column(33).Width=3281"
      Splits(1)._ColumnProps(134)=   "Column(33).DividerColor=0"
      Splits(1)._ColumnProps(135)=   "Column(33)._WidthInPix=3175"
      Splits(1)._ColumnProps(136)=   "Column(33).Order=34"
      Splits(1)._ColumnProps(137)=   "Column(34).Width=3281"
      Splits(1)._ColumnProps(138)=   "Column(34).DividerColor=0"
      Splits(1)._ColumnProps(139)=   "Column(34)._WidthInPix=3175"
      Splits(1)._ColumnProps(140)=   "Column(34).Order=35"
      Splits(1)._ColumnProps(141)=   "Column(35).Width=3281"
      Splits(1)._ColumnProps(142)=   "Column(35).DividerColor=0"
      Splits(1)._ColumnProps(143)=   "Column(35)._WidthInPix=3175"
      Splits(1)._ColumnProps(144)=   "Column(35).Order=36"
      Splits(1)._ColumnProps(145)=   "Column(36).Width=3281"
      Splits(1)._ColumnProps(146)=   "Column(36).DividerColor=0"
      Splits(1)._ColumnProps(147)=   "Column(36)._WidthInPix=3175"
      Splits(1)._ColumnProps(148)=   "Column(36).Order=37"
      Splits(1)._ColumnProps(149)=   "Column(37).Width=3281"
      Splits(1)._ColumnProps(150)=   "Column(37).DividerColor=0"
      Splits(1)._ColumnProps(151)=   "Column(37)._WidthInPix=3175"
      Splits(1)._ColumnProps(152)=   "Column(37).Order=38"
      Splits(1)._ColumnProps(153)=   "Column(38).Width=3281"
      Splits(1)._ColumnProps(154)=   "Column(38).DividerColor=0"
      Splits(1)._ColumnProps(155)=   "Column(38)._WidthInPix=3175"
      Splits(1)._ColumnProps(156)=   "Column(38).Order=39"
      Splits(1)._ColumnProps(157)=   "Column(39).Width=3281"
      Splits(1)._ColumnProps(158)=   "Column(39).DividerColor=0"
      Splits(1)._ColumnProps(159)=   "Column(39)._WidthInPix=3175"
      Splits(1)._ColumnProps(160)=   "Column(39).Order=40"
      Splits(1)._ColumnProps(161)=   "Column(40).Width=3281"
      Splits(1)._ColumnProps(162)=   "Column(40).DividerColor=0"
      Splits(1)._ColumnProps(163)=   "Column(40)._WidthInPix=3175"
      Splits(1)._ColumnProps(164)=   "Column(40).Order=41"
      Splits(1)._ColumnProps(165)=   "Column(41).Width=3281"
      Splits(1)._ColumnProps(166)=   "Column(41).DividerColor=0"
      Splits(1)._ColumnProps(167)=   "Column(41)._WidthInPix=3175"
      Splits(1)._ColumnProps(168)=   "Column(41).Order=42"
      Splits(1)._ColumnProps(169)=   "Column(42).Width=3281"
      Splits(1)._ColumnProps(170)=   "Column(42).DividerColor=0"
      Splits(1)._ColumnProps(171)=   "Column(42)._WidthInPix=3175"
      Splits(1)._ColumnProps(172)=   "Column(42).Order=43"
      Splits(1)._ColumnProps(173)=   "Column(43).Width=3281"
      Splits(1)._ColumnProps(174)=   "Column(43).DividerColor=0"
      Splits(1)._ColumnProps(175)=   "Column(43)._WidthInPix=3175"
      Splits(1)._ColumnProps(176)=   "Column(43).Order=44"
      Splits(1)._ColumnProps(177)=   "Column(44).Width=3281"
      Splits(1)._ColumnProps(178)=   "Column(44).DividerColor=0"
      Splits(1)._ColumnProps(179)=   "Column(44)._WidthInPix=3175"
      Splits(1)._ColumnProps(180)=   "Column(44).Order=45"
      Splits(1)._ColumnProps(181)=   "Column(45).Width=3281"
      Splits(1)._ColumnProps(182)=   "Column(45).DividerColor=0"
      Splits(1)._ColumnProps(183)=   "Column(45)._WidthInPix=3175"
      Splits(1)._ColumnProps(184)=   "Column(45).Order=46"
      Splits(1)._ColumnProps(185)=   "Column(46).Width=3281"
      Splits(1)._ColumnProps(186)=   "Column(46).DividerColor=0"
      Splits(1)._ColumnProps(187)=   "Column(46)._WidthInPix=3175"
      Splits(1)._ColumnProps(188)=   "Column(46).Order=47"
      Splits(1)._ColumnProps(189)=   "Column(47).Width=3281"
      Splits(1)._ColumnProps(190)=   "Column(47).DividerColor=0"
      Splits(1)._ColumnProps(191)=   "Column(47)._WidthInPix=3175"
      Splits(1)._ColumnProps(192)=   "Column(47).Order=48"
      Splits(1)._ColumnProps(193)=   "Column(48).Width=3281"
      Splits(1)._ColumnProps(194)=   "Column(48).DividerColor=0"
      Splits(1)._ColumnProps(195)=   "Column(48)._WidthInPix=3175"
      Splits(1)._ColumnProps(196)=   "Column(48).Order=49"
      Splits(1)._ColumnProps(197)=   "Column(49).Width=3281"
      Splits(1)._ColumnProps(198)=   "Column(49).DividerColor=0"
      Splits(1)._ColumnProps(199)=   "Column(49)._WidthInPix=3175"
      Splits(1)._ColumnProps(200)=   "Column(49).Order=50"
      Splits(1)._ColumnProps(201)=   "Column(50).Width=3281"
      Splits(1)._ColumnProps(202)=   "Column(50).DividerColor=0"
      Splits(1)._ColumnProps(203)=   "Column(50)._WidthInPix=3175"
      Splits(1)._ColumnProps(204)=   "Column(50).Order=51"
      Splits(1)._ColumnProps(205)=   "Column(51).Width=3281"
      Splits(1)._ColumnProps(206)=   "Column(51).DividerColor=0"
      Splits(1)._ColumnProps(207)=   "Column(51)._WidthInPix=3175"
      Splits(1)._ColumnProps(208)=   "Column(51).Order=52"
      Splits(1)._ColumnProps(209)=   "Column(52).Width=3281"
      Splits(1)._ColumnProps(210)=   "Column(52).DividerColor=0"
      Splits(1)._ColumnProps(211)=   "Column(52)._WidthInPix=3175"
      Splits(1)._ColumnProps(212)=   "Column(52).Order=53"
      Splits(1)._ColumnProps(213)=   "Column(53).Width=3281"
      Splits(1)._ColumnProps(214)=   "Column(53).DividerColor=0"
      Splits(1)._ColumnProps(215)=   "Column(53)._WidthInPix=3175"
      Splits(1)._ColumnProps(216)=   "Column(53).Order=54"
      Splits(1)._ColumnProps(217)=   "Column(54).Width=3281"
      Splits(1)._ColumnProps(218)=   "Column(54).DividerColor=0"
      Splits(1)._ColumnProps(219)=   "Column(54)._WidthInPix=3175"
      Splits(1)._ColumnProps(220)=   "Column(54).Order=55"
      Splits(1)._ColumnProps(221)=   "Column(55).Width=3281"
      Splits(1)._ColumnProps(222)=   "Column(55).DividerColor=0"
      Splits(1)._ColumnProps(223)=   "Column(55)._WidthInPix=3175"
      Splits(1)._ColumnProps(224)=   "Column(55).Order=56"
      Splits(1)._ColumnProps(225)=   "Column(56).Width=3281"
      Splits(1)._ColumnProps(226)=   "Column(56).DividerColor=0"
      Splits(1)._ColumnProps(227)=   "Column(56)._WidthInPix=3175"
      Splits(1)._ColumnProps(228)=   "Column(56).Order=57"
      Splits(1)._ColumnProps(229)=   "Column(57).Width=3281"
      Splits(1)._ColumnProps(230)=   "Column(57).DividerColor=0"
      Splits(1)._ColumnProps(231)=   "Column(57)._WidthInPix=3175"
      Splits(1)._ColumnProps(232)=   "Column(57).Order=58"
      Splits(1)._ColumnProps(233)=   "Column(58).Width=3281"
      Splits(1)._ColumnProps(234)=   "Column(58).DividerColor=0"
      Splits(1)._ColumnProps(235)=   "Column(58)._WidthInPix=3175"
      Splits(1)._ColumnProps(236)=   "Column(58).Order=59"
      Splits(1)._ColumnProps(237)=   "Column(59).Width=3281"
      Splits(1)._ColumnProps(238)=   "Column(59).DividerColor=0"
      Splits(1)._ColumnProps(239)=   "Column(59)._WidthInPix=3175"
      Splits(1)._ColumnProps(240)=   "Column(59).Order=60"
      Splits(1)._ColumnProps(241)=   "Column(60).Width=3281"
      Splits(1)._ColumnProps(242)=   "Column(60).DividerColor=0"
      Splits(1)._ColumnProps(243)=   "Column(60)._WidthInPix=3175"
      Splits(1)._ColumnProps(244)=   "Column(60).Order=61"
      Splits(1)._ColumnProps(245)=   "Column(61).Width=3281"
      Splits(1)._ColumnProps(246)=   "Column(61).DividerColor=0"
      Splits(1)._ColumnProps(247)=   "Column(61)._WidthInPix=3175"
      Splits(1)._ColumnProps(248)=   "Column(61).Order=62"
      Splits(1)._ColumnProps(249)=   "Column(62).Width=3281"
      Splits(1)._ColumnProps(250)=   "Column(62).DividerColor=0"
      Splits(1)._ColumnProps(251)=   "Column(62)._WidthInPix=3175"
      Splits(1)._ColumnProps(252)=   "Column(62).Order=63"
      Splits(1)._ColumnProps(253)=   "Column(63).Width=3281"
      Splits(1)._ColumnProps(254)=   "Column(63).DividerColor=0"
      Splits(1)._ColumnProps(255)=   "Column(63)._WidthInPix=3175"
      Splits(1)._ColumnProps(256)=   "Column(63).Order=64"
      Splits(1)._ColumnProps(257)=   "Column(64).Width=3281"
      Splits(1)._ColumnProps(258)=   "Column(64).DividerColor=0"
      Splits(1)._ColumnProps(259)=   "Column(64)._WidthInPix=3175"
      Splits(1)._ColumnProps(260)=   "Column(64).Order=65"
      Splits(1)._ColumnProps(261)=   "Column(65).Width=3281"
      Splits(1)._ColumnProps(262)=   "Column(65).DividerColor=0"
      Splits(1)._ColumnProps(263)=   "Column(65)._WidthInPix=3175"
      Splits(1)._ColumnProps(264)=   "Column(65).Order=66"
      Splits(1)._ColumnProps(265)=   "Column(66).Width=3281"
      Splits(1)._ColumnProps(266)=   "Column(66).DividerColor=0"
      Splits(1)._ColumnProps(267)=   "Column(66)._WidthInPix=3175"
      Splits(1)._ColumnProps(268)=   "Column(66).Order=67"
      Splits(1)._ColumnProps(269)=   "Column(67).Width=3281"
      Splits(1)._ColumnProps(270)=   "Column(67).DividerColor=0"
      Splits(1)._ColumnProps(271)=   "Column(67)._WidthInPix=3175"
      Splits(1)._ColumnProps(272)=   "Column(67).Order=68"
      Splits(1)._ColumnProps(273)=   "Column(68).Width=3281"
      Splits(1)._ColumnProps(274)=   "Column(68).DividerColor=0"
      Splits(1)._ColumnProps(275)=   "Column(68)._WidthInPix=3175"
      Splits(1)._ColumnProps(276)=   "Column(68).Order=69"
      Splits(1)._ColumnProps(277)=   "Column(69).Width=3281"
      Splits(1)._ColumnProps(278)=   "Column(69).DividerColor=0"
      Splits(1)._ColumnProps(279)=   "Column(69)._WidthInPix=3175"
      Splits(1)._ColumnProps(280)=   "Column(69).Order=70"
      Splits(1)._ColumnProps(281)=   "Column(70).Width=3281"
      Splits(1)._ColumnProps(282)=   "Column(70).DividerColor=0"
      Splits(1)._ColumnProps(283)=   "Column(70)._WidthInPix=3175"
      Splits(1)._ColumnProps(284)=   "Column(70).Order=71"
      Splits(1)._ColumnProps(285)=   "Column(71).Width=3281"
      Splits(1)._ColumnProps(286)=   "Column(71).DividerColor=0"
      Splits(1)._ColumnProps(287)=   "Column(71)._WidthInPix=3175"
      Splits(1)._ColumnProps(288)=   "Column(71).Order=72"
      Splits(1)._ColumnProps(289)=   "Column(72).Width=3281"
      Splits(1)._ColumnProps(290)=   "Column(72).DividerColor=0"
      Splits(1)._ColumnProps(291)=   "Column(72)._WidthInPix=3175"
      Splits(1)._ColumnProps(292)=   "Column(72).Order=73"
      Splits(1)._ColumnProps(293)=   "Column(73).Width=3281"
      Splits(1)._ColumnProps(294)=   "Column(73).DividerColor=0"
      Splits(1)._ColumnProps(295)=   "Column(73)._WidthInPix=3175"
      Splits(1)._ColumnProps(296)=   "Column(73).Order=74"
      Splits(1)._ColumnProps(297)=   "Column(74).Width=3281"
      Splits(1)._ColumnProps(298)=   "Column(74).DividerColor=0"
      Splits(1)._ColumnProps(299)=   "Column(74)._WidthInPix=3175"
      Splits(1)._ColumnProps(300)=   "Column(74).Order=75"
      Splits(1)._ColumnProps(301)=   "Column(75).Width=3281"
      Splits(1)._ColumnProps(302)=   "Column(75).DividerColor=0"
      Splits(1)._ColumnProps(303)=   "Column(75)._WidthInPix=3175"
      Splits(1)._ColumnProps(304)=   "Column(75).Order=76"
      Splits(1)._ColumnProps(305)=   "Column(76).Width=3281"
      Splits(1)._ColumnProps(306)=   "Column(76).DividerColor=0"
      Splits(1)._ColumnProps(307)=   "Column(76)._WidthInPix=3175"
      Splits(1)._ColumnProps(308)=   "Column(76).Order=77"
      Splits(1)._ColumnProps(309)=   "Column(77).Width=3281"
      Splits(1)._ColumnProps(310)=   "Column(77).DividerColor=0"
      Splits(1)._ColumnProps(311)=   "Column(77)._WidthInPix=3175"
      Splits(1)._ColumnProps(312)=   "Column(77).Order=78"
      Splits(1)._ColumnProps(313)=   "Column(78).Width=3281"
      Splits(1)._ColumnProps(314)=   "Column(78).DividerColor=0"
      Splits(1)._ColumnProps(315)=   "Column(78)._WidthInPix=3175"
      Splits(1)._ColumnProps(316)=   "Column(78).Order=79"
      Splits(1)._ColumnProps(317)=   "Column(79).Width=3281"
      Splits(1)._ColumnProps(318)=   "Column(79).DividerColor=0"
      Splits(1)._ColumnProps(319)=   "Column(79)._WidthInPix=3175"
      Splits(1)._ColumnProps(320)=   "Column(79).Order=80"
      Splits(1)._ColumnProps(321)=   "Column(80).Width=3281"
      Splits(1)._ColumnProps(322)=   "Column(80).DividerColor=0"
      Splits(1)._ColumnProps(323)=   "Column(80)._WidthInPix=3175"
      Splits(1)._ColumnProps(324)=   "Column(80).Order=81"
      Splits(1)._ColumnProps(325)=   "Column(81).Width=3281"
      Splits(1)._ColumnProps(326)=   "Column(81).DividerColor=0"
      Splits(1)._ColumnProps(327)=   "Column(81)._WidthInPix=3175"
      Splits(1)._ColumnProps(328)=   "Column(81).Order=82"
      Splits(1)._ColumnProps(329)=   "Column(82).Width=3281"
      Splits(1)._ColumnProps(330)=   "Column(82).DividerColor=0"
      Splits(1)._ColumnProps(331)=   "Column(82)._WidthInPix=3175"
      Splits(1)._ColumnProps(332)=   "Column(82).Order=83"
      Splits(1)._ColumnProps(333)=   "Column(83).Width=3281"
      Splits(1)._ColumnProps(334)=   "Column(83).DividerColor=0"
      Splits(1)._ColumnProps(335)=   "Column(83)._WidthInPix=3175"
      Splits(1)._ColumnProps(336)=   "Column(83).Order=84"
      Splits(1)._ColumnProps(337)=   "Column(84).Width=3281"
      Splits(1)._ColumnProps(338)=   "Column(84).DividerColor=0"
      Splits(1)._ColumnProps(339)=   "Column(84)._WidthInPix=3175"
      Splits(1)._ColumnProps(340)=   "Column(84).Order=85"
      Splits(1)._ColumnProps(341)=   "Column(85).Width=3281"
      Splits(1)._ColumnProps(342)=   "Column(85).DividerColor=0"
      Splits(1)._ColumnProps(343)=   "Column(85)._WidthInPix=3175"
      Splits(1)._ColumnProps(344)=   "Column(85).Order=86"
      Splits(1)._ColumnProps(345)=   "Column(86).Width=3281"
      Splits(1)._ColumnProps(346)=   "Column(86).DividerColor=0"
      Splits(1)._ColumnProps(347)=   "Column(86)._WidthInPix=3175"
      Splits(1)._ColumnProps(348)=   "Column(86).Order=87"
      Splits(1)._ColumnProps(349)=   "Column(87).Width=3281"
      Splits(1)._ColumnProps(350)=   "Column(87).DividerColor=0"
      Splits(1)._ColumnProps(351)=   "Column(87)._WidthInPix=3175"
      Splits(1)._ColumnProps(352)=   "Column(87).Order=88"
      Splits(1)._ColumnProps(353)=   "Column(88).Width=3281"
      Splits(1)._ColumnProps(354)=   "Column(88).DividerColor=0"
      Splits(1)._ColumnProps(355)=   "Column(88)._WidthInPix=3175"
      Splits(1)._ColumnProps(356)=   "Column(88).Order=89"
      Splits(1)._ColumnProps(357)=   "Column(89).Width=3281"
      Splits(1)._ColumnProps(358)=   "Column(89).DividerColor=0"
      Splits(1)._ColumnProps(359)=   "Column(89)._WidthInPix=3175"
      Splits(1)._ColumnProps(360)=   "Column(89).Order=90"
      Splits(1)._ColumnProps(361)=   "Column(90).Width=3281"
      Splits(1)._ColumnProps(362)=   "Column(90).DividerColor=0"
      Splits(1)._ColumnProps(363)=   "Column(90)._WidthInPix=3175"
      Splits(1)._ColumnProps(364)=   "Column(90).Order=91"
      Splits(1)._ColumnProps(365)=   "Column(91).Width=3281"
      Splits(1)._ColumnProps(366)=   "Column(91).DividerColor=0"
      Splits(1)._ColumnProps(367)=   "Column(91)._WidthInPix=3175"
      Splits(1)._ColumnProps(368)=   "Column(91).Order=92"
      Splits(1)._ColumnProps(369)=   "Column(92).Width=3281"
      Splits(1)._ColumnProps(370)=   "Column(92).DividerColor=0"
      Splits(1)._ColumnProps(371)=   "Column(92)._WidthInPix=3175"
      Splits(1)._ColumnProps(372)=   "Column(92).Order=93"
      Splits(1)._ColumnProps(373)=   "Column(93).Width=3281"
      Splits(1)._ColumnProps(374)=   "Column(93).DividerColor=0"
      Splits(1)._ColumnProps(375)=   "Column(93)._WidthInPix=3175"
      Splits(1)._ColumnProps(376)=   "Column(93).Order=94"
      Splits(1)._ColumnProps(377)=   "Column(94).Width=3281"
      Splits(1)._ColumnProps(378)=   "Column(94).DividerColor=0"
      Splits(1)._ColumnProps(379)=   "Column(94)._WidthInPix=3175"
      Splits(1)._ColumnProps(380)=   "Column(94).Order=95"
      Splits(1)._ColumnProps(381)=   "Column(95).Width=3281"
      Splits(1)._ColumnProps(382)=   "Column(95).DividerColor=0"
      Splits(1)._ColumnProps(383)=   "Column(95)._WidthInPix=3175"
      Splits(1)._ColumnProps(384)=   "Column(95).Order=96"
      Splits(1)._ColumnProps(385)=   "Column(96).Width=3281"
      Splits(1)._ColumnProps(386)=   "Column(96).DividerColor=0"
      Splits(1)._ColumnProps(387)=   "Column(96)._WidthInPix=3175"
      Splits(1)._ColumnProps(388)=   "Column(96).Order=97"
      Splits(1)._ColumnProps(389)=   "Column(97).Width=3281"
      Splits(1)._ColumnProps(390)=   "Column(97).DividerColor=0"
      Splits(1)._ColumnProps(391)=   "Column(97)._WidthInPix=3175"
      Splits(1)._ColumnProps(392)=   "Column(97).Order=98"
      Splits(1)._ColumnProps(393)=   "Column(98).Width=3281"
      Splits(1)._ColumnProps(394)=   "Column(98).DividerColor=0"
      Splits(1)._ColumnProps(395)=   "Column(98)._WidthInPix=3175"
      Splits(1)._ColumnProps(396)=   "Column(98).Order=99"
      Splits(1)._ColumnProps(397)=   "Column(99).Width=3281"
      Splits(1)._ColumnProps(398)=   "Column(99).DividerColor=0"
      Splits(1)._ColumnProps(399)=   "Column(99)._WidthInPix=3175"
      Splits(1)._ColumnProps(400)=   "Column(99).Order=100"
      Splits(1)._ColumnProps(401)=   "Column(100).Width=3281"
      Splits(1)._ColumnProps(402)=   "Column(100).DividerColor=0"
      Splits(1)._ColumnProps(403)=   "Column(100)._WidthInPix=3175"
      Splits(1)._ColumnProps(404)=   "Column(100).Order=101"
      Splits(1)._ColumnProps(405)=   "Column(101).Width=3281"
      Splits(1)._ColumnProps(406)=   "Column(101).DividerColor=0"
      Splits(1)._ColumnProps(407)=   "Column(101)._WidthInPix=3175"
      Splits(1)._ColumnProps(408)=   "Column(101).Order=102"
      Splits(1)._ColumnProps(409)=   "Column(102).Width=3281"
      Splits(1)._ColumnProps(410)=   "Column(102).DividerColor=0"
      Splits(1)._ColumnProps(411)=   "Column(102)._WidthInPix=3175"
      Splits(1)._ColumnProps(412)=   "Column(102).Order=103"
      Splits(1)._ColumnProps(413)=   "Column(103).Width=3281"
      Splits(1)._ColumnProps(414)=   "Column(103).DividerColor=0"
      Splits(1)._ColumnProps(415)=   "Column(103)._WidthInPix=3175"
      Splits(1)._ColumnProps(416)=   "Column(103).Order=104"
      Splits(1)._ColumnProps(417)=   "Column(104).Width=3281"
      Splits(1)._ColumnProps(418)=   "Column(104).DividerColor=0"
      Splits(1)._ColumnProps(419)=   "Column(104)._WidthInPix=3175"
      Splits(1)._ColumnProps(420)=   "Column(104).Order=105"
      Splits(1)._ColumnProps(421)=   "Column(105).Width=3281"
      Splits(1)._ColumnProps(422)=   "Column(105).DividerColor=0"
      Splits(1)._ColumnProps(423)=   "Column(105)._WidthInPix=3175"
      Splits(1)._ColumnProps(424)=   "Column(105).Order=106"
      Splits(1)._ColumnProps(425)=   "Column(106).Width=3281"
      Splits(1)._ColumnProps(426)=   "Column(106).DividerColor=0"
      Splits(1)._ColumnProps(427)=   "Column(106)._WidthInPix=3175"
      Splits(1)._ColumnProps(428)=   "Column(106).Order=107"
      Splits(1)._ColumnProps(429)=   "Column(107).Width=3281"
      Splits(1)._ColumnProps(430)=   "Column(107).DividerColor=0"
      Splits(1)._ColumnProps(431)=   "Column(107)._WidthInPix=3175"
      Splits(1)._ColumnProps(432)=   "Column(107).Order=108"
      Splits(1)._ColumnProps(433)=   "Column(108).Width=3281"
      Splits(1)._ColumnProps(434)=   "Column(108).DividerColor=0"
      Splits(1)._ColumnProps(435)=   "Column(108)._WidthInPix=3175"
      Splits(1)._ColumnProps(436)=   "Column(108).Order=109"
      Splits(1)._ColumnProps(437)=   "Column(109).Width=3281"
      Splits(1)._ColumnProps(438)=   "Column(109).DividerColor=0"
      Splits(1)._ColumnProps(439)=   "Column(109)._WidthInPix=3175"
      Splits(1)._ColumnProps(440)=   "Column(109).Order=110"
      Splits(1)._ColumnProps(441)=   "Column(110).Width=3281"
      Splits(1)._ColumnProps(442)=   "Column(110).DividerColor=0"
      Splits(1)._ColumnProps(443)=   "Column(110)._WidthInPix=3175"
      Splits(1)._ColumnProps(444)=   "Column(110).Order=111"
      Splits(1)._ColumnProps(445)=   "Column(111).Width=3281"
      Splits(1)._ColumnProps(446)=   "Column(111).DividerColor=0"
      Splits(1)._ColumnProps(447)=   "Column(111)._WidthInPix=3175"
      Splits(1)._ColumnProps(448)=   "Column(111).Order=112"
      Splits(1)._ColumnProps(449)=   "Column(112).Width=3281"
      Splits(1)._ColumnProps(450)=   "Column(112).DividerColor=0"
      Splits(1)._ColumnProps(451)=   "Column(112)._WidthInPix=3175"
      Splits(1)._ColumnProps(452)=   "Column(112).Order=113"
      Splits(1)._ColumnProps(453)=   "Column(113).Width=3281"
      Splits(1)._ColumnProps(454)=   "Column(113).DividerColor=0"
      Splits(1)._ColumnProps(455)=   "Column(113)._WidthInPix=3175"
      Splits(1)._ColumnProps(456)=   "Column(113).Order=114"
      Splits(1)._ColumnProps(457)=   "Column(114).Width=3281"
      Splits(1)._ColumnProps(458)=   "Column(114).DividerColor=0"
      Splits(1)._ColumnProps(459)=   "Column(114)._WidthInPix=3175"
      Splits(1)._ColumnProps(460)=   "Column(114).Order=115"
      Splits(1)._ColumnProps(461)=   "Column(115).Width=3281"
      Splits(1)._ColumnProps(462)=   "Column(115).DividerColor=0"
      Splits(1)._ColumnProps(463)=   "Column(115)._WidthInPix=3175"
      Splits(1)._ColumnProps(464)=   "Column(115).Order=116"
      Splits(1)._ColumnProps(465)=   "Column(116).Width=3281"
      Splits(1)._ColumnProps(466)=   "Column(116).DividerColor=0"
      Splits(1)._ColumnProps(467)=   "Column(116)._WidthInPix=3175"
      Splits(1)._ColumnProps(468)=   "Column(116).Order=117"
      Splits(1)._ColumnProps(469)=   "Column(117).Width=3281"
      Splits(1)._ColumnProps(470)=   "Column(117).DividerColor=0"
      Splits(1)._ColumnProps(471)=   "Column(117)._WidthInPix=3175"
      Splits(1)._ColumnProps(472)=   "Column(117).Order=118"
      Splits(1)._ColumnProps(473)=   "Column(118).Width=3281"
      Splits(1)._ColumnProps(474)=   "Column(118).DividerColor=0"
      Splits(1)._ColumnProps(475)=   "Column(118)._WidthInPix=3175"
      Splits(1)._ColumnProps(476)=   "Column(118).Order=119"
      Splits(1)._ColumnProps(477)=   "Column(119).Width=3281"
      Splits(1)._ColumnProps(478)=   "Column(119).DividerColor=0"
      Splits(1)._ColumnProps(479)=   "Column(119)._WidthInPix=3175"
      Splits(1)._ColumnProps(480)=   "Column(119).Order=120"
      Splits(1)._ColumnProps(481)=   "Column(120).Width=3281"
      Splits(1)._ColumnProps(482)=   "Column(120).DividerColor=0"
      Splits(1)._ColumnProps(483)=   "Column(120)._WidthInPix=3175"
      Splits(1)._ColumnProps(484)=   "Column(120).Order=121"
      Splits(1)._ColumnProps(485)=   "Column(121).Width=3281"
      Splits(1)._ColumnProps(486)=   "Column(121).DividerColor=0"
      Splits(1)._ColumnProps(487)=   "Column(121)._WidthInPix=3175"
      Splits(1)._ColumnProps(488)=   "Column(121).Order=122"
      Splits(1)._ColumnProps(489)=   "Column(122).Width=3281"
      Splits(1)._ColumnProps(490)=   "Column(122).DividerColor=0"
      Splits(1)._ColumnProps(491)=   "Column(122)._WidthInPix=3175"
      Splits(1)._ColumnProps(492)=   "Column(122).Order=123"
      Splits(1)._ColumnProps(493)=   "Column(123).Width=3281"
      Splits(1)._ColumnProps(494)=   "Column(123).DividerColor=0"
      Splits(1)._ColumnProps(495)=   "Column(123)._WidthInPix=3175"
      Splits(1)._ColumnProps(496)=   "Column(123).Order=124"
      Splits(1)._ColumnProps(497)=   "Column(124).Width=3281"
      Splits(1)._ColumnProps(498)=   "Column(124).DividerColor=0"
      Splits(1)._ColumnProps(499)=   "Column(124)._WidthInPix=3175"
      Splits(1)._ColumnProps(500)=   "Column(124).Order=125"
      Splits(1)._ColumnProps(501)=   "Column(125).Width=3281"
      Splits(1)._ColumnProps(502)=   "Column(125).DividerColor=0"
      Splits(1)._ColumnProps(503)=   "Column(125)._WidthInPix=3175"
      Splits(1)._ColumnProps(504)=   "Column(125).Order=126"
      Splits(1)._ColumnProps(505)=   "Column(126).Width=3281"
      Splits(1)._ColumnProps(506)=   "Column(126).DividerColor=0"
      Splits(1)._ColumnProps(507)=   "Column(126)._WidthInPix=3175"
      Splits(1)._ColumnProps(508)=   "Column(126).Order=127"
      Splits(1)._ColumnProps(509)=   "Column(127).Width=3281"
      Splits(1)._ColumnProps(510)=   "Column(127).DividerColor=0"
      Splits(1)._ColumnProps(511)=   "Column(127)._WidthInPix=3175"
      Splits(1)._ColumnProps(512)=   "Column(127).Order=128"
      Splits(1)._ColumnProps(513)=   "Column(128).Width=3281"
      Splits(1)._ColumnProps(514)=   "Column(128).DividerColor=0"
      Splits(1)._ColumnProps(515)=   "Column(128)._WidthInPix=3175"
      Splits(1)._ColumnProps(516)=   "Column(128).Order=129"
      Splits(1)._ColumnProps(517)=   "Column(129).Width=3281"
      Splits(1)._ColumnProps(518)=   "Column(129).DividerColor=0"
      Splits(1)._ColumnProps(519)=   "Column(129)._WidthInPix=3175"
      Splits(1)._ColumnProps(520)=   "Column(129).Order=130"
      Splits(1)._ColumnProps(521)=   "Column(130).Width=3281"
      Splits(1)._ColumnProps(522)=   "Column(130).DividerColor=0"
      Splits(1)._ColumnProps(523)=   "Column(130)._WidthInPix=3175"
      Splits(1)._ColumnProps(524)=   "Column(130).Order=131"
      Splits(1)._ColumnProps(525)=   "Column(131).Width=3281"
      Splits(1)._ColumnProps(526)=   "Column(131).DividerColor=0"
      Splits(1)._ColumnProps(527)=   "Column(131)._WidthInPix=3175"
      Splits(1)._ColumnProps(528)=   "Column(131).Order=132"
      Splits(1)._ColumnProps(529)=   "Column(132).Width=3281"
      Splits(1)._ColumnProps(530)=   "Column(132).DividerColor=0"
      Splits(1)._ColumnProps(531)=   "Column(132)._WidthInPix=3175"
      Splits(1)._ColumnProps(532)=   "Column(132).Order=133"
      Splits(1)._ColumnProps(533)=   "Column(133).Width=3281"
      Splits(1)._ColumnProps(534)=   "Column(133).DividerColor=0"
      Splits(1)._ColumnProps(535)=   "Column(133)._WidthInPix=3175"
      Splits(1)._ColumnProps(536)=   "Column(133).Order=134"
      Splits(1)._ColumnProps(537)=   "Column(134).Width=3281"
      Splits(1)._ColumnProps(538)=   "Column(134).DividerColor=0"
      Splits(1)._ColumnProps(539)=   "Column(134)._WidthInPix=3175"
      Splits(1)._ColumnProps(540)=   "Column(134).Order=135"
      Splits(1)._ColumnProps(541)=   "Column(135).Width=3281"
      Splits(1)._ColumnProps(542)=   "Column(135).DividerColor=0"
      Splits(1)._ColumnProps(543)=   "Column(135)._WidthInPix=3175"
      Splits(1)._ColumnProps(544)=   "Column(135).Order=136"
      Splits(1)._ColumnProps(545)=   "Column(136).Width=3281"
      Splits(1)._ColumnProps(546)=   "Column(136).DividerColor=0"
      Splits(1)._ColumnProps(547)=   "Column(136)._WidthInPix=3175"
      Splits(1)._ColumnProps(548)=   "Column(136).Order=137"
      Splits(1)._ColumnProps(549)=   "Column(137).Width=3281"
      Splits(1)._ColumnProps(550)=   "Column(137).DividerColor=0"
      Splits(1)._ColumnProps(551)=   "Column(137)._WidthInPix=3175"
      Splits(1)._ColumnProps(552)=   "Column(137).Order=138"
      Splits.Count    =   2
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
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
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=245,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=242,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=243,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=244,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=237,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=234,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=235,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=236,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=229,.parent=13"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=226,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=227,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=228,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=221,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=218,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=219,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=220,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=213,.parent=13"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=210,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=211,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=212,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=205,.parent=13"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=202,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=203,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=204,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=197,.parent=13"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=194,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=195,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=196,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=189,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=186,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=187,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=188,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=181,.parent=13"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=178,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=179,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=180,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=173,.parent=13"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=170,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=171,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=172,.parent=17"
      _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=165,.parent=13"
      _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=162,.parent=14"
      _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=163,.parent=15"
      _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=164,.parent=17"
      _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=157,.parent=13"
      _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=154,.parent=14"
      _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=155,.parent=15"
      _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=156,.parent=17"
      _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=149,.parent=13"
      _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=146,.parent=14"
      _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=147,.parent=15"
      _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=148,.parent=17"
      _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=141,.parent=13"
      _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=138,.parent=14"
      _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=139,.parent=15"
      _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=140,.parent=17"
      _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=133,.parent=13"
      _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=128,.parent=14"
      _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=129,.parent=15"
      _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=130,.parent=17"
      _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=123,.parent=13"
      _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=120,.parent=14"
      _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=121,.parent=15"
      _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=122,.parent=17"
      _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=115,.parent=13"
      _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=112,.parent=14"
      _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=113,.parent=15"
      _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=114,.parent=17"
      _StyleDefs(104) =   "Splits(0).Columns(17).Style:id=54,.parent=13"
      _StyleDefs(105) =   "Splits(0).Columns(17).HeadingStyle:id=51,.parent=14"
      _StyleDefs(106) =   "Splits(0).Columns(17).FooterStyle:id=52,.parent=15"
      _StyleDefs(107) =   "Splits(0).Columns(17).EditorStyle:id=53,.parent=17"
      _StyleDefs(108) =   "Splits(0).Columns(18).Style:id=50,.parent=13"
      _StyleDefs(109) =   "Splits(0).Columns(18).HeadingStyle:id=47,.parent=14"
      _StyleDefs(110) =   "Splits(0).Columns(18).FooterStyle:id=48,.parent=15"
      _StyleDefs(111) =   "Splits(0).Columns(18).EditorStyle:id=49,.parent=17"
      _StyleDefs(112) =   "Splits(0).Columns(19).Style:id=46,.parent=13"
      _StyleDefs(113) =   "Splits(0).Columns(19).HeadingStyle:id=43,.parent=14"
      _StyleDefs(114) =   "Splits(0).Columns(19).FooterStyle:id=44,.parent=15"
      _StyleDefs(115) =   "Splits(0).Columns(19).EditorStyle:id=45,.parent=17"
      _StyleDefs(116) =   "Splits(0).Columns(20).Style:id=28,.parent=13"
      _StyleDefs(117) =   "Splits(0).Columns(20).HeadingStyle:id=25,.parent=14"
      _StyleDefs(118) =   "Splits(0).Columns(20).FooterStyle:id=26,.parent=15"
      _StyleDefs(119) =   "Splits(0).Columns(20).EditorStyle:id=27,.parent=17"
      _StyleDefs(120) =   "Splits(0).Columns(21).Style:id=86,.parent=13"
      _StyleDefs(121) =   "Splits(0).Columns(21).HeadingStyle:id=83,.parent=14"
      _StyleDefs(122) =   "Splits(0).Columns(21).FooterStyle:id=84,.parent=15"
      _StyleDefs(123) =   "Splits(0).Columns(21).EditorStyle:id=85,.parent=17"
      _StyleDefs(124) =   "Splits(0).Columns(22).Style:id=32,.parent=13"
      _StyleDefs(125) =   "Splits(0).Columns(22).HeadingStyle:id=29,.parent=14"
      _StyleDefs(126) =   "Splits(0).Columns(22).FooterStyle:id=30,.parent=15"
      _StyleDefs(127) =   "Splits(0).Columns(22).EditorStyle:id=31,.parent=17"
      _StyleDefs(128) =   "Splits(0).Columns(23).Style:id=58,.parent=13"
      _StyleDefs(129) =   "Splits(0).Columns(23).HeadingStyle:id=55,.parent=14"
      _StyleDefs(130) =   "Splits(0).Columns(23).FooterStyle:id=56,.parent=15"
      _StyleDefs(131) =   "Splits(0).Columns(23).EditorStyle:id=57,.parent=17"
      _StyleDefs(132) =   "Splits(0).Columns(24).Style:id=91,.parent=13"
      _StyleDefs(133) =   "Splits(0).Columns(24).HeadingStyle:id=88,.parent=14"
      _StyleDefs(134) =   "Splits(0).Columns(24).FooterStyle:id=89,.parent=15"
      _StyleDefs(135) =   "Splits(0).Columns(24).EditorStyle:id=90,.parent=17"
      _StyleDefs(136) =   "Splits(0).Columns(25).Style:id=253,.parent=13"
      _StyleDefs(137) =   "Splits(0).Columns(25).HeadingStyle:id=250,.parent=14"
      _StyleDefs(138) =   "Splits(0).Columns(25).FooterStyle:id=251,.parent=15"
      _StyleDefs(139) =   "Splits(0).Columns(25).EditorStyle:id=252,.parent=17"
      _StyleDefs(140) =   "Splits(0).Columns(26).Style:id=261,.parent=13"
      _StyleDefs(141) =   "Splits(0).Columns(26).HeadingStyle:id=258,.parent=14"
      _StyleDefs(142) =   "Splits(0).Columns(26).FooterStyle:id=259,.parent=15"
      _StyleDefs(143) =   "Splits(0).Columns(26).EditorStyle:id=260,.parent=17"
      _StyleDefs(144) =   "Splits(0).Columns(27).Style:id=269,.parent=13"
      _StyleDefs(145) =   "Splits(0).Columns(27).HeadingStyle:id=266,.parent=14"
      _StyleDefs(146) =   "Splits(0).Columns(27).FooterStyle:id=267,.parent=15"
      _StyleDefs(147) =   "Splits(0).Columns(27).EditorStyle:id=268,.parent=17"
      _StyleDefs(148) =   "Splits(0).Columns(28).Style:id=277,.parent=13"
      _StyleDefs(149) =   "Splits(0).Columns(28).HeadingStyle:id=274,.parent=14"
      _StyleDefs(150) =   "Splits(0).Columns(28).FooterStyle:id=275,.parent=15"
      _StyleDefs(151) =   "Splits(0).Columns(28).EditorStyle:id=276,.parent=17"
      _StyleDefs(152) =   "Splits(0).Columns(29).Style:id=285,.parent=13"
      _StyleDefs(153) =   "Splits(0).Columns(29).HeadingStyle:id=282,.parent=14"
      _StyleDefs(154) =   "Splits(0).Columns(29).FooterStyle:id=283,.parent=15"
      _StyleDefs(155) =   "Splits(0).Columns(29).EditorStyle:id=284,.parent=17"
      _StyleDefs(156) =   "Splits(0).Columns(30).Style:id=293,.parent=13"
      _StyleDefs(157) =   "Splits(0).Columns(30).HeadingStyle:id=290,.parent=14"
      _StyleDefs(158) =   "Splits(0).Columns(30).FooterStyle:id=291,.parent=15"
      _StyleDefs(159) =   "Splits(0).Columns(30).EditorStyle:id=292,.parent=17"
      _StyleDefs(160) =   "Splits(0).Columns(31).Style:id=301,.parent=13"
      _StyleDefs(161) =   "Splits(0).Columns(31).HeadingStyle:id=298,.parent=14"
      _StyleDefs(162) =   "Splits(0).Columns(31).FooterStyle:id=299,.parent=15"
      _StyleDefs(163) =   "Splits(0).Columns(31).EditorStyle:id=300,.parent=17"
      _StyleDefs(164) =   "Splits(0).Columns(32).Style:id=309,.parent=13"
      _StyleDefs(165) =   "Splits(0).Columns(32).HeadingStyle:id=306,.parent=14"
      _StyleDefs(166) =   "Splits(0).Columns(32).FooterStyle:id=307,.parent=15"
      _StyleDefs(167) =   "Splits(0).Columns(32).EditorStyle:id=308,.parent=17"
      _StyleDefs(168) =   "Splits(0).Columns(33).Style:id=317,.parent=13"
      _StyleDefs(169) =   "Splits(0).Columns(33).HeadingStyle:id=314,.parent=14"
      _StyleDefs(170) =   "Splits(0).Columns(33).FooterStyle:id=315,.parent=15"
      _StyleDefs(171) =   "Splits(0).Columns(33).EditorStyle:id=316,.parent=17"
      _StyleDefs(172) =   "Splits(0).Columns(34).Style:id=325,.parent=13"
      _StyleDefs(173) =   "Splits(0).Columns(34).HeadingStyle:id=322,.parent=14"
      _StyleDefs(174) =   "Splits(0).Columns(34).FooterStyle:id=323,.parent=15"
      _StyleDefs(175) =   "Splits(0).Columns(34).EditorStyle:id=324,.parent=17"
      _StyleDefs(176) =   "Splits(0).Columns(35).Style:id=333,.parent=13"
      _StyleDefs(177) =   "Splits(0).Columns(35).HeadingStyle:id=330,.parent=14"
      _StyleDefs(178) =   "Splits(0).Columns(35).FooterStyle:id=331,.parent=15"
      _StyleDefs(179) =   "Splits(0).Columns(35).EditorStyle:id=332,.parent=17"
      _StyleDefs(180) =   "Splits(0).Columns(36).Style:id=341,.parent=13"
      _StyleDefs(181) =   "Splits(0).Columns(36).HeadingStyle:id=338,.parent=14"
      _StyleDefs(182) =   "Splits(0).Columns(36).FooterStyle:id=339,.parent=15"
      _StyleDefs(183) =   "Splits(0).Columns(36).EditorStyle:id=340,.parent=17"
      _StyleDefs(184) =   "Splits(0).Columns(37).Style:id=349,.parent=13"
      _StyleDefs(185) =   "Splits(0).Columns(37).HeadingStyle:id=346,.parent=14"
      _StyleDefs(186) =   "Splits(0).Columns(37).FooterStyle:id=347,.parent=15"
      _StyleDefs(187) =   "Splits(0).Columns(37).EditorStyle:id=348,.parent=17"
      _StyleDefs(188) =   "Splits(0).Columns(38).Style:id=357,.parent=13"
      _StyleDefs(189) =   "Splits(0).Columns(38).HeadingStyle:id=354,.parent=14"
      _StyleDefs(190) =   "Splits(0).Columns(38).FooterStyle:id=355,.parent=15"
      _StyleDefs(191) =   "Splits(0).Columns(38).EditorStyle:id=356,.parent=17"
      _StyleDefs(192) =   "Splits(0).Columns(39).Style:id=365,.parent=13"
      _StyleDefs(193) =   "Splits(0).Columns(39).HeadingStyle:id=362,.parent=14"
      _StyleDefs(194) =   "Splits(0).Columns(39).FooterStyle:id=363,.parent=15"
      _StyleDefs(195) =   "Splits(0).Columns(39).EditorStyle:id=364,.parent=17"
      _StyleDefs(196) =   "Splits(0).Columns(40).Style:id=373,.parent=13"
      _StyleDefs(197) =   "Splits(0).Columns(40).HeadingStyle:id=370,.parent=14"
      _StyleDefs(198) =   "Splits(0).Columns(40).FooterStyle:id=371,.parent=15"
      _StyleDefs(199) =   "Splits(0).Columns(40).EditorStyle:id=372,.parent=17"
      _StyleDefs(200) =   "Splits(0).Columns(41).Style:id=381,.parent=13"
      _StyleDefs(201) =   "Splits(0).Columns(41).HeadingStyle:id=378,.parent=14"
      _StyleDefs(202) =   "Splits(0).Columns(41).FooterStyle:id=379,.parent=15"
      _StyleDefs(203) =   "Splits(0).Columns(41).EditorStyle:id=380,.parent=17"
      _StyleDefs(204) =   "Splits(0).Columns(42).Style:id=389,.parent=13"
      _StyleDefs(205) =   "Splits(0).Columns(42).HeadingStyle:id=386,.parent=14"
      _StyleDefs(206) =   "Splits(0).Columns(42).FooterStyle:id=387,.parent=15"
      _StyleDefs(207) =   "Splits(0).Columns(42).EditorStyle:id=388,.parent=17"
      _StyleDefs(208) =   "Splits(0).Columns(43).Style:id=397,.parent=13"
      _StyleDefs(209) =   "Splits(0).Columns(43).HeadingStyle:id=394,.parent=14"
      _StyleDefs(210) =   "Splits(0).Columns(43).FooterStyle:id=395,.parent=15"
      _StyleDefs(211) =   "Splits(0).Columns(43).EditorStyle:id=396,.parent=17"
      _StyleDefs(212) =   "Splits(0).Columns(44).Style:id=405,.parent=13"
      _StyleDefs(213) =   "Splits(0).Columns(44).HeadingStyle:id=402,.parent=14"
      _StyleDefs(214) =   "Splits(0).Columns(44).FooterStyle:id=403,.parent=15"
      _StyleDefs(215) =   "Splits(0).Columns(44).EditorStyle:id=404,.parent=17"
      _StyleDefs(216) =   "Splits(0).Columns(45).Style:id=413,.parent=13"
      _StyleDefs(217) =   "Splits(0).Columns(45).HeadingStyle:id=410,.parent=14"
      _StyleDefs(218) =   "Splits(0).Columns(45).FooterStyle:id=411,.parent=15"
      _StyleDefs(219) =   "Splits(0).Columns(45).EditorStyle:id=412,.parent=17"
      _StyleDefs(220) =   "Splits(0).Columns(46).Style:id=421,.parent=13"
      _StyleDefs(221) =   "Splits(0).Columns(46).HeadingStyle:id=418,.parent=14"
      _StyleDefs(222) =   "Splits(0).Columns(46).FooterStyle:id=419,.parent=15"
      _StyleDefs(223) =   "Splits(0).Columns(46).EditorStyle:id=420,.parent=17"
      _StyleDefs(224) =   "Splits(0).Columns(47).Style:id=429,.parent=13"
      _StyleDefs(225) =   "Splits(0).Columns(47).HeadingStyle:id=426,.parent=14"
      _StyleDefs(226) =   "Splits(0).Columns(47).FooterStyle:id=427,.parent=15"
      _StyleDefs(227) =   "Splits(0).Columns(47).EditorStyle:id=428,.parent=17"
      _StyleDefs(228) =   "Splits(0).Columns(48).Style:id=437,.parent=13"
      _StyleDefs(229) =   "Splits(0).Columns(48).HeadingStyle:id=434,.parent=14"
      _StyleDefs(230) =   "Splits(0).Columns(48).FooterStyle:id=435,.parent=15"
      _StyleDefs(231) =   "Splits(0).Columns(48).EditorStyle:id=436,.parent=17"
      _StyleDefs(232) =   "Splits(0).Columns(49).Style:id=445,.parent=13"
      _StyleDefs(233) =   "Splits(0).Columns(49).HeadingStyle:id=442,.parent=14"
      _StyleDefs(234) =   "Splits(0).Columns(49).FooterStyle:id=443,.parent=15"
      _StyleDefs(235) =   "Splits(0).Columns(49).EditorStyle:id=444,.parent=17"
      _StyleDefs(236) =   "Splits(0).Columns(50).Style:id=453,.parent=13"
      _StyleDefs(237) =   "Splits(0).Columns(50).HeadingStyle:id=450,.parent=14"
      _StyleDefs(238) =   "Splits(0).Columns(50).FooterStyle:id=451,.parent=15"
      _StyleDefs(239) =   "Splits(0).Columns(50).EditorStyle:id=452,.parent=17"
      _StyleDefs(240) =   "Splits(0).Columns(51).Style:id=461,.parent=13"
      _StyleDefs(241) =   "Splits(0).Columns(51).HeadingStyle:id=458,.parent=14"
      _StyleDefs(242) =   "Splits(0).Columns(51).FooterStyle:id=459,.parent=15"
      _StyleDefs(243) =   "Splits(0).Columns(51).EditorStyle:id=460,.parent=17"
      _StyleDefs(244) =   "Splits(0).Columns(52).Style:id=469,.parent=13"
      _StyleDefs(245) =   "Splits(0).Columns(52).HeadingStyle:id=466,.parent=14"
      _StyleDefs(246) =   "Splits(0).Columns(52).FooterStyle:id=467,.parent=15"
      _StyleDefs(247) =   "Splits(0).Columns(52).EditorStyle:id=468,.parent=17"
      _StyleDefs(248) =   "Splits(0).Columns(53).Style:id=477,.parent=13"
      _StyleDefs(249) =   "Splits(0).Columns(53).HeadingStyle:id=474,.parent=14"
      _StyleDefs(250) =   "Splits(0).Columns(53).FooterStyle:id=475,.parent=15"
      _StyleDefs(251) =   "Splits(0).Columns(53).EditorStyle:id=476,.parent=17"
      _StyleDefs(252) =   "Splits(0).Columns(54).Style:id=485,.parent=13"
      _StyleDefs(253) =   "Splits(0).Columns(54).HeadingStyle:id=482,.parent=14"
      _StyleDefs(254) =   "Splits(0).Columns(54).FooterStyle:id=483,.parent=15"
      _StyleDefs(255) =   "Splits(0).Columns(54).EditorStyle:id=484,.parent=17"
      _StyleDefs(256) =   "Splits(0).Columns(55).Style:id=493,.parent=13"
      _StyleDefs(257) =   "Splits(0).Columns(55).HeadingStyle:id=490,.parent=14"
      _StyleDefs(258) =   "Splits(0).Columns(55).FooterStyle:id=491,.parent=15"
      _StyleDefs(259) =   "Splits(0).Columns(55).EditorStyle:id=492,.parent=17"
      _StyleDefs(260) =   "Splits(0).Columns(56).Style:id=501,.parent=13"
      _StyleDefs(261) =   "Splits(0).Columns(56).HeadingStyle:id=498,.parent=14"
      _StyleDefs(262) =   "Splits(0).Columns(56).FooterStyle:id=499,.parent=15"
      _StyleDefs(263) =   "Splits(0).Columns(56).EditorStyle:id=500,.parent=17"
      _StyleDefs(264) =   "Splits(0).Columns(57).Style:id=509,.parent=13"
      _StyleDefs(265) =   "Splits(0).Columns(57).HeadingStyle:id=506,.parent=14"
      _StyleDefs(266) =   "Splits(0).Columns(57).FooterStyle:id=507,.parent=15"
      _StyleDefs(267) =   "Splits(0).Columns(57).EditorStyle:id=508,.parent=17"
      _StyleDefs(268) =   "Splits(0).Columns(58).Style:id=517,.parent=13"
      _StyleDefs(269) =   "Splits(0).Columns(58).HeadingStyle:id=514,.parent=14"
      _StyleDefs(270) =   "Splits(0).Columns(58).FooterStyle:id=515,.parent=15"
      _StyleDefs(271) =   "Splits(0).Columns(58).EditorStyle:id=516,.parent=17"
      _StyleDefs(272) =   "Splits(0).Columns(59).Style:id=525,.parent=13"
      _StyleDefs(273) =   "Splits(0).Columns(59).HeadingStyle:id=522,.parent=14"
      _StyleDefs(274) =   "Splits(0).Columns(59).FooterStyle:id=523,.parent=15"
      _StyleDefs(275) =   "Splits(0).Columns(59).EditorStyle:id=524,.parent=17"
      _StyleDefs(276) =   "Splits(0).Columns(60).Style:id=533,.parent=13"
      _StyleDefs(277) =   "Splits(0).Columns(60).HeadingStyle:id=530,.parent=14"
      _StyleDefs(278) =   "Splits(0).Columns(60).FooterStyle:id=531,.parent=15"
      _StyleDefs(279) =   "Splits(0).Columns(60).EditorStyle:id=532,.parent=17"
      _StyleDefs(280) =   "Splits(0).Columns(61).Style:id=541,.parent=13"
      _StyleDefs(281) =   "Splits(0).Columns(61).HeadingStyle:id=538,.parent=14"
      _StyleDefs(282) =   "Splits(0).Columns(61).FooterStyle:id=539,.parent=15"
      _StyleDefs(283) =   "Splits(0).Columns(61).EditorStyle:id=540,.parent=17"
      _StyleDefs(284) =   "Splits(0).Columns(62).Style:id=549,.parent=13"
      _StyleDefs(285) =   "Splits(0).Columns(62).HeadingStyle:id=546,.parent=14"
      _StyleDefs(286) =   "Splits(0).Columns(62).FooterStyle:id=547,.parent=15"
      _StyleDefs(287) =   "Splits(0).Columns(62).EditorStyle:id=548,.parent=17"
      _StyleDefs(288) =   "Splits(0).Columns(63).Style:id=557,.parent=13"
      _StyleDefs(289) =   "Splits(0).Columns(63).HeadingStyle:id=554,.parent=14"
      _StyleDefs(290) =   "Splits(0).Columns(63).FooterStyle:id=555,.parent=15"
      _StyleDefs(291) =   "Splits(0).Columns(63).EditorStyle:id=556,.parent=17"
      _StyleDefs(292) =   "Splits(0).Columns(64).Style:id=565,.parent=13"
      _StyleDefs(293) =   "Splits(0).Columns(64).HeadingStyle:id=562,.parent=14"
      _StyleDefs(294) =   "Splits(0).Columns(64).FooterStyle:id=563,.parent=15"
      _StyleDefs(295) =   "Splits(0).Columns(64).EditorStyle:id=564,.parent=17"
      _StyleDefs(296) =   "Splits(0).Columns(65).Style:id=573,.parent=13"
      _StyleDefs(297) =   "Splits(0).Columns(65).HeadingStyle:id=570,.parent=14"
      _StyleDefs(298) =   "Splits(0).Columns(65).FooterStyle:id=571,.parent=15"
      _StyleDefs(299) =   "Splits(0).Columns(65).EditorStyle:id=572,.parent=17"
      _StyleDefs(300) =   "Splits(0).Columns(66).Style:id=581,.parent=13"
      _StyleDefs(301) =   "Splits(0).Columns(66).HeadingStyle:id=578,.parent=14"
      _StyleDefs(302) =   "Splits(0).Columns(66).FooterStyle:id=579,.parent=15"
      _StyleDefs(303) =   "Splits(0).Columns(66).EditorStyle:id=580,.parent=17"
      _StyleDefs(304) =   "Splits(0).Columns(67).Style:id=589,.parent=13"
      _StyleDefs(305) =   "Splits(0).Columns(67).HeadingStyle:id=586,.parent=14"
      _StyleDefs(306) =   "Splits(0).Columns(67).FooterStyle:id=587,.parent=15"
      _StyleDefs(307) =   "Splits(0).Columns(67).EditorStyle:id=588,.parent=17"
      _StyleDefs(308) =   "Splits(0).Columns(68).Style:id=597,.parent=13"
      _StyleDefs(309) =   "Splits(0).Columns(68).HeadingStyle:id=594,.parent=14"
      _StyleDefs(310) =   "Splits(0).Columns(68).FooterStyle:id=595,.parent=15"
      _StyleDefs(311) =   "Splits(0).Columns(68).EditorStyle:id=596,.parent=17"
      _StyleDefs(312) =   "Splits(0).Columns(69).Style:id=605,.parent=13"
      _StyleDefs(313) =   "Splits(0).Columns(69).HeadingStyle:id=602,.parent=14"
      _StyleDefs(314) =   "Splits(0).Columns(69).FooterStyle:id=603,.parent=15"
      _StyleDefs(315) =   "Splits(0).Columns(69).EditorStyle:id=604,.parent=17"
      _StyleDefs(316) =   "Splits(0).Columns(70).Style:id=613,.parent=13"
      _StyleDefs(317) =   "Splits(0).Columns(70).HeadingStyle:id=610,.parent=14"
      _StyleDefs(318) =   "Splits(0).Columns(70).FooterStyle:id=611,.parent=15"
      _StyleDefs(319) =   "Splits(0).Columns(70).EditorStyle:id=612,.parent=17"
      _StyleDefs(320) =   "Splits(0).Columns(71).Style:id=621,.parent=13"
      _StyleDefs(321) =   "Splits(0).Columns(71).HeadingStyle:id=618,.parent=14"
      _StyleDefs(322) =   "Splits(0).Columns(71).FooterStyle:id=619,.parent=15"
      _StyleDefs(323) =   "Splits(0).Columns(71).EditorStyle:id=620,.parent=17"
      _StyleDefs(324) =   "Splits(0).Columns(72).Style:id=629,.parent=13"
      _StyleDefs(325) =   "Splits(0).Columns(72).HeadingStyle:id=626,.parent=14"
      _StyleDefs(326) =   "Splits(0).Columns(72).FooterStyle:id=627,.parent=15"
      _StyleDefs(327) =   "Splits(0).Columns(72).EditorStyle:id=628,.parent=17"
      _StyleDefs(328) =   "Splits(0).Columns(73).Style:id=637,.parent=13"
      _StyleDefs(329) =   "Splits(0).Columns(73).HeadingStyle:id=634,.parent=14"
      _StyleDefs(330) =   "Splits(0).Columns(73).FooterStyle:id=635,.parent=15"
      _StyleDefs(331) =   "Splits(0).Columns(73).EditorStyle:id=636,.parent=17"
      _StyleDefs(332) =   "Splits(0).Columns(74).Style:id=645,.parent=13"
      _StyleDefs(333) =   "Splits(0).Columns(74).HeadingStyle:id=642,.parent=14"
      _StyleDefs(334) =   "Splits(0).Columns(74).FooterStyle:id=643,.parent=15"
      _StyleDefs(335) =   "Splits(0).Columns(74).EditorStyle:id=644,.parent=17"
      _StyleDefs(336) =   "Splits(0).Columns(75).Style:id=653,.parent=13"
      _StyleDefs(337) =   "Splits(0).Columns(75).HeadingStyle:id=650,.parent=14"
      _StyleDefs(338) =   "Splits(0).Columns(75).FooterStyle:id=651,.parent=15"
      _StyleDefs(339) =   "Splits(0).Columns(75).EditorStyle:id=652,.parent=17"
      _StyleDefs(340) =   "Splits(0).Columns(76).Style:id=661,.parent=13"
      _StyleDefs(341) =   "Splits(0).Columns(76).HeadingStyle:id=658,.parent=14"
      _StyleDefs(342) =   "Splits(0).Columns(76).FooterStyle:id=659,.parent=15"
      _StyleDefs(343) =   "Splits(0).Columns(76).EditorStyle:id=660,.parent=17"
      _StyleDefs(344) =   "Splits(0).Columns(77).Style:id=669,.parent=13"
      _StyleDefs(345) =   "Splits(0).Columns(77).HeadingStyle:id=666,.parent=14"
      _StyleDefs(346) =   "Splits(0).Columns(77).FooterStyle:id=667,.parent=15"
      _StyleDefs(347) =   "Splits(0).Columns(77).EditorStyle:id=668,.parent=17"
      _StyleDefs(348) =   "Splits(0).Columns(78).Style:id=677,.parent=13"
      _StyleDefs(349) =   "Splits(0).Columns(78).HeadingStyle:id=674,.parent=14"
      _StyleDefs(350) =   "Splits(0).Columns(78).FooterStyle:id=675,.parent=15"
      _StyleDefs(351) =   "Splits(0).Columns(78).EditorStyle:id=676,.parent=17"
      _StyleDefs(352) =   "Splits(0).Columns(79).Style:id=685,.parent=13"
      _StyleDefs(353) =   "Splits(0).Columns(79).HeadingStyle:id=682,.parent=14"
      _StyleDefs(354) =   "Splits(0).Columns(79).FooterStyle:id=683,.parent=15"
      _StyleDefs(355) =   "Splits(0).Columns(79).EditorStyle:id=684,.parent=17"
      _StyleDefs(356) =   "Splits(0).Columns(80).Style:id=693,.parent=13"
      _StyleDefs(357) =   "Splits(0).Columns(80).HeadingStyle:id=690,.parent=14"
      _StyleDefs(358) =   "Splits(0).Columns(80).FooterStyle:id=691,.parent=15"
      _StyleDefs(359) =   "Splits(0).Columns(80).EditorStyle:id=692,.parent=17"
      _StyleDefs(360) =   "Splits(0).Columns(81).Style:id=701,.parent=13"
      _StyleDefs(361) =   "Splits(0).Columns(81).HeadingStyle:id=698,.parent=14"
      _StyleDefs(362) =   "Splits(0).Columns(81).FooterStyle:id=699,.parent=15"
      _StyleDefs(363) =   "Splits(0).Columns(81).EditorStyle:id=700,.parent=17"
      _StyleDefs(364) =   "Splits(0).Columns(82).Style:id=709,.parent=13"
      _StyleDefs(365) =   "Splits(0).Columns(82).HeadingStyle:id=706,.parent=14"
      _StyleDefs(366) =   "Splits(0).Columns(82).FooterStyle:id=707,.parent=15"
      _StyleDefs(367) =   "Splits(0).Columns(82).EditorStyle:id=708,.parent=17"
      _StyleDefs(368) =   "Splits(0).Columns(83).Style:id=717,.parent=13"
      _StyleDefs(369) =   "Splits(0).Columns(83).HeadingStyle:id=714,.parent=14"
      _StyleDefs(370) =   "Splits(0).Columns(83).FooterStyle:id=715,.parent=15"
      _StyleDefs(371) =   "Splits(0).Columns(83).EditorStyle:id=716,.parent=17"
      _StyleDefs(372) =   "Splits(0).Columns(84).Style:id=725,.parent=13"
      _StyleDefs(373) =   "Splits(0).Columns(84).HeadingStyle:id=722,.parent=14"
      _StyleDefs(374) =   "Splits(0).Columns(84).FooterStyle:id=723,.parent=15"
      _StyleDefs(375) =   "Splits(0).Columns(84).EditorStyle:id=724,.parent=17"
      _StyleDefs(376) =   "Splits(0).Columns(85).Style:id=733,.parent=13"
      _StyleDefs(377) =   "Splits(0).Columns(85).HeadingStyle:id=730,.parent=14"
      _StyleDefs(378) =   "Splits(0).Columns(85).FooterStyle:id=731,.parent=15"
      _StyleDefs(379) =   "Splits(0).Columns(85).EditorStyle:id=732,.parent=17"
      _StyleDefs(380) =   "Splits(0).Columns(86).Style:id=741,.parent=13"
      _StyleDefs(381) =   "Splits(0).Columns(86).HeadingStyle:id=738,.parent=14"
      _StyleDefs(382) =   "Splits(0).Columns(86).FooterStyle:id=739,.parent=15"
      _StyleDefs(383) =   "Splits(0).Columns(86).EditorStyle:id=740,.parent=17"
      _StyleDefs(384) =   "Splits(0).Columns(87).Style:id=749,.parent=13"
      _StyleDefs(385) =   "Splits(0).Columns(87).HeadingStyle:id=746,.parent=14"
      _StyleDefs(386) =   "Splits(0).Columns(87).FooterStyle:id=747,.parent=15"
      _StyleDefs(387) =   "Splits(0).Columns(87).EditorStyle:id=748,.parent=17"
      _StyleDefs(388) =   "Splits(0).Columns(88).Style:id=757,.parent=13"
      _StyleDefs(389) =   "Splits(0).Columns(88).HeadingStyle:id=754,.parent=14"
      _StyleDefs(390) =   "Splits(0).Columns(88).FooterStyle:id=755,.parent=15"
      _StyleDefs(391) =   "Splits(0).Columns(88).EditorStyle:id=756,.parent=17"
      _StyleDefs(392) =   "Splits(0).Columns(89).Style:id=765,.parent=13"
      _StyleDefs(393) =   "Splits(0).Columns(89).HeadingStyle:id=762,.parent=14"
      _StyleDefs(394) =   "Splits(0).Columns(89).FooterStyle:id=763,.parent=15"
      _StyleDefs(395) =   "Splits(0).Columns(89).EditorStyle:id=764,.parent=17"
      _StyleDefs(396) =   "Splits(0).Columns(90).Style:id=773,.parent=13"
      _StyleDefs(397) =   "Splits(0).Columns(90).HeadingStyle:id=770,.parent=14"
      _StyleDefs(398) =   "Splits(0).Columns(90).FooterStyle:id=771,.parent=15"
      _StyleDefs(399) =   "Splits(0).Columns(90).EditorStyle:id=772,.parent=17"
      _StyleDefs(400) =   "Splits(0).Columns(91).Style:id=781,.parent=13"
      _StyleDefs(401) =   "Splits(0).Columns(91).HeadingStyle:id=778,.parent=14"
      _StyleDefs(402) =   "Splits(0).Columns(91).FooterStyle:id=779,.parent=15"
      _StyleDefs(403) =   "Splits(0).Columns(91).EditorStyle:id=780,.parent=17"
      _StyleDefs(404) =   "Splits(0).Columns(92).Style:id=789,.parent=13"
      _StyleDefs(405) =   "Splits(0).Columns(92).HeadingStyle:id=786,.parent=14"
      _StyleDefs(406) =   "Splits(0).Columns(92).FooterStyle:id=787,.parent=15"
      _StyleDefs(407) =   "Splits(0).Columns(92).EditorStyle:id=788,.parent=17"
      _StyleDefs(408) =   "Splits(0).Columns(93).Style:id=797,.parent=13"
      _StyleDefs(409) =   "Splits(0).Columns(93).HeadingStyle:id=794,.parent=14"
      _StyleDefs(410) =   "Splits(0).Columns(93).FooterStyle:id=795,.parent=15"
      _StyleDefs(411) =   "Splits(0).Columns(93).EditorStyle:id=796,.parent=17"
      _StyleDefs(412) =   "Splits(0).Columns(94).Style:id=805,.parent=13"
      _StyleDefs(413) =   "Splits(0).Columns(94).HeadingStyle:id=802,.parent=14"
      _StyleDefs(414) =   "Splits(0).Columns(94).FooterStyle:id=803,.parent=15"
      _StyleDefs(415) =   "Splits(0).Columns(94).EditorStyle:id=804,.parent=17"
      _StyleDefs(416) =   "Splits(0).Columns(95).Style:id=813,.parent=13"
      _StyleDefs(417) =   "Splits(0).Columns(95).HeadingStyle:id=810,.parent=14"
      _StyleDefs(418) =   "Splits(0).Columns(95).FooterStyle:id=811,.parent=15"
      _StyleDefs(419) =   "Splits(0).Columns(95).EditorStyle:id=812,.parent=17"
      _StyleDefs(420) =   "Splits(0).Columns(96).Style:id=821,.parent=13"
      _StyleDefs(421) =   "Splits(0).Columns(96).HeadingStyle:id=818,.parent=14"
      _StyleDefs(422) =   "Splits(0).Columns(96).FooterStyle:id=819,.parent=15"
      _StyleDefs(423) =   "Splits(0).Columns(96).EditorStyle:id=820,.parent=17"
      _StyleDefs(424) =   "Splits(0).Columns(97).Style:id=829,.parent=13"
      _StyleDefs(425) =   "Splits(0).Columns(97).HeadingStyle:id=826,.parent=14"
      _StyleDefs(426) =   "Splits(0).Columns(97).FooterStyle:id=827,.parent=15"
      _StyleDefs(427) =   "Splits(0).Columns(97).EditorStyle:id=828,.parent=17"
      _StyleDefs(428) =   "Splits(0).Columns(98).Style:id=837,.parent=13"
      _StyleDefs(429) =   "Splits(0).Columns(98).HeadingStyle:id=834,.parent=14"
      _StyleDefs(430) =   "Splits(0).Columns(98).FooterStyle:id=835,.parent=15"
      _StyleDefs(431) =   "Splits(0).Columns(98).EditorStyle:id=836,.parent=17"
      _StyleDefs(432) =   "Splits(0).Columns(99).Style:id=845,.parent=13"
      _StyleDefs(433) =   "Splits(0).Columns(99).HeadingStyle:id=842,.parent=14"
      _StyleDefs(434) =   "Splits(0).Columns(99).FooterStyle:id=843,.parent=15"
      _StyleDefs(435) =   "Splits(0).Columns(99).EditorStyle:id=844,.parent=17"
      _StyleDefs(436) =   "Splits(0).Columns(100).Style:id=853,.parent=13"
      _StyleDefs(437) =   "Splits(0).Columns(100).HeadingStyle:id=850,.parent=14"
      _StyleDefs(438) =   "Splits(0).Columns(100).FooterStyle:id=851,.parent=15"
      _StyleDefs(439) =   "Splits(0).Columns(100).EditorStyle:id=852,.parent=17"
      _StyleDefs(440) =   "Splits(0).Columns(101).Style:id=861,.parent=13"
      _StyleDefs(441) =   "Splits(0).Columns(101).HeadingStyle:id=858,.parent=14"
      _StyleDefs(442) =   "Splits(0).Columns(101).FooterStyle:id=859,.parent=15"
      _StyleDefs(443) =   "Splits(0).Columns(101).EditorStyle:id=860,.parent=17"
      _StyleDefs(444) =   "Splits(0).Columns(102).Style:id=869,.parent=13"
      _StyleDefs(445) =   "Splits(0).Columns(102).HeadingStyle:id=866,.parent=14"
      _StyleDefs(446) =   "Splits(0).Columns(102).FooterStyle:id=867,.parent=15"
      _StyleDefs(447) =   "Splits(0).Columns(102).EditorStyle:id=868,.parent=17"
      _StyleDefs(448) =   "Splits(0).Columns(103).Style:id=877,.parent=13"
      _StyleDefs(449) =   "Splits(0).Columns(103).HeadingStyle:id=874,.parent=14"
      _StyleDefs(450) =   "Splits(0).Columns(103).FooterStyle:id=875,.parent=15"
      _StyleDefs(451) =   "Splits(0).Columns(103).EditorStyle:id=876,.parent=17"
      _StyleDefs(452) =   "Splits(0).Columns(104).Style:id=885,.parent=13"
      _StyleDefs(453) =   "Splits(0).Columns(104).HeadingStyle:id=882,.parent=14"
      _StyleDefs(454) =   "Splits(0).Columns(104).FooterStyle:id=883,.parent=15"
      _StyleDefs(455) =   "Splits(0).Columns(104).EditorStyle:id=884,.parent=17"
      _StyleDefs(456) =   "Splits(0).Columns(105).Style:id=893,.parent=13"
      _StyleDefs(457) =   "Splits(0).Columns(105).HeadingStyle:id=890,.parent=14"
      _StyleDefs(458) =   "Splits(0).Columns(105).FooterStyle:id=891,.parent=15"
      _StyleDefs(459) =   "Splits(0).Columns(105).EditorStyle:id=892,.parent=17"
      _StyleDefs(460) =   "Splits(0).Columns(106).Style:id=901,.parent=13"
      _StyleDefs(461) =   "Splits(0).Columns(106).HeadingStyle:id=898,.parent=14"
      _StyleDefs(462) =   "Splits(0).Columns(106).FooterStyle:id=899,.parent=15"
      _StyleDefs(463) =   "Splits(0).Columns(106).EditorStyle:id=900,.parent=17"
      _StyleDefs(464) =   "Splits(0).Columns(107).Style:id=909,.parent=13"
      _StyleDefs(465) =   "Splits(0).Columns(107).HeadingStyle:id=906,.parent=14"
      _StyleDefs(466) =   "Splits(0).Columns(107).FooterStyle:id=907,.parent=15"
      _StyleDefs(467) =   "Splits(0).Columns(107).EditorStyle:id=908,.parent=17"
      _StyleDefs(468) =   "Splits(0).Columns(108).Style:id=917,.parent=13"
      _StyleDefs(469) =   "Splits(0).Columns(108).HeadingStyle:id=914,.parent=14"
      _StyleDefs(470) =   "Splits(0).Columns(108).FooterStyle:id=915,.parent=15"
      _StyleDefs(471) =   "Splits(0).Columns(108).EditorStyle:id=916,.parent=17"
      _StyleDefs(472) =   "Splits(0).Columns(109).Style:id=925,.parent=13"
      _StyleDefs(473) =   "Splits(0).Columns(109).HeadingStyle:id=922,.parent=14"
      _StyleDefs(474) =   "Splits(0).Columns(109).FooterStyle:id=923,.parent=15"
      _StyleDefs(475) =   "Splits(0).Columns(109).EditorStyle:id=924,.parent=17"
      _StyleDefs(476) =   "Splits(0).Columns(110).Style:id=933,.parent=13"
      _StyleDefs(477) =   "Splits(0).Columns(110).HeadingStyle:id=930,.parent=14"
      _StyleDefs(478) =   "Splits(0).Columns(110).FooterStyle:id=931,.parent=15"
      _StyleDefs(479) =   "Splits(0).Columns(110).EditorStyle:id=932,.parent=17"
      _StyleDefs(480) =   "Splits(0).Columns(111).Style:id=941,.parent=13"
      _StyleDefs(481) =   "Splits(0).Columns(111).HeadingStyle:id=938,.parent=14"
      _StyleDefs(482) =   "Splits(0).Columns(111).FooterStyle:id=939,.parent=15"
      _StyleDefs(483) =   "Splits(0).Columns(111).EditorStyle:id=940,.parent=17"
      _StyleDefs(484) =   "Splits(0).Columns(112).Style:id=949,.parent=13"
      _StyleDefs(485) =   "Splits(0).Columns(112).HeadingStyle:id=946,.parent=14"
      _StyleDefs(486) =   "Splits(0).Columns(112).FooterStyle:id=947,.parent=15"
      _StyleDefs(487) =   "Splits(0).Columns(112).EditorStyle:id=948,.parent=17"
      _StyleDefs(488) =   "Splits(0).Columns(113).Style:id=957,.parent=13"
      _StyleDefs(489) =   "Splits(0).Columns(113).HeadingStyle:id=954,.parent=14"
      _StyleDefs(490) =   "Splits(0).Columns(113).FooterStyle:id=955,.parent=15"
      _StyleDefs(491) =   "Splits(0).Columns(113).EditorStyle:id=956,.parent=17"
      _StyleDefs(492) =   "Splits(0).Columns(114).Style:id=965,.parent=13"
      _StyleDefs(493) =   "Splits(0).Columns(114).HeadingStyle:id=962,.parent=14"
      _StyleDefs(494) =   "Splits(0).Columns(114).FooterStyle:id=963,.parent=15"
      _StyleDefs(495) =   "Splits(0).Columns(114).EditorStyle:id=964,.parent=17"
      _StyleDefs(496) =   "Splits(0).Columns(115).Style:id=973,.parent=13"
      _StyleDefs(497) =   "Splits(0).Columns(115).HeadingStyle:id=970,.parent=14"
      _StyleDefs(498) =   "Splits(0).Columns(115).FooterStyle:id=971,.parent=15"
      _StyleDefs(499) =   "Splits(0).Columns(115).EditorStyle:id=972,.parent=17"
      _StyleDefs(500) =   "Splits(0).Columns(116).Style:id=981,.parent=13"
      _StyleDefs(501) =   "Splits(0).Columns(116).HeadingStyle:id=978,.parent=14"
      _StyleDefs(502) =   "Splits(0).Columns(116).FooterStyle:id=979,.parent=15"
      _StyleDefs(503) =   "Splits(0).Columns(116).EditorStyle:id=980,.parent=17"
      _StyleDefs(504) =   "Splits(0).Columns(117).Style:id=989,.parent=13"
      _StyleDefs(505) =   "Splits(0).Columns(117).HeadingStyle:id=986,.parent=14"
      _StyleDefs(506) =   "Splits(0).Columns(117).FooterStyle:id=987,.parent=15"
      _StyleDefs(507) =   "Splits(0).Columns(117).EditorStyle:id=988,.parent=17"
      _StyleDefs(508) =   "Splits(0).Columns(118).Style:id=997,.parent=13"
      _StyleDefs(509) =   "Splits(0).Columns(118).HeadingStyle:id=994,.parent=14"
      _StyleDefs(510) =   "Splits(0).Columns(118).FooterStyle:id=995,.parent=15"
      _StyleDefs(511) =   "Splits(0).Columns(118).EditorStyle:id=996,.parent=17"
      _StyleDefs(512) =   "Splits(0).Columns(119).Style:id=1005,.parent=13"
      _StyleDefs(513) =   "Splits(0).Columns(119).HeadingStyle:id=1002,.parent=14"
      _StyleDefs(514) =   "Splits(0).Columns(119).FooterStyle:id=1003,.parent=15"
      _StyleDefs(515) =   "Splits(0).Columns(119).EditorStyle:id=1004,.parent=17"
      _StyleDefs(516) =   "Splits(0).Columns(120).Style:id=1013,.parent=13"
      _StyleDefs(517) =   "Splits(0).Columns(120).HeadingStyle:id=1010,.parent=14"
      _StyleDefs(518) =   "Splits(0).Columns(120).FooterStyle:id=1011,.parent=15"
      _StyleDefs(519) =   "Splits(0).Columns(120).EditorStyle:id=1012,.parent=17"
      _StyleDefs(520) =   "Splits(0).Columns(121).Style:id=1021,.parent=13"
      _StyleDefs(521) =   "Splits(0).Columns(121).HeadingStyle:id=1018,.parent=14"
      _StyleDefs(522) =   "Splits(0).Columns(121).FooterStyle:id=1019,.parent=15"
      _StyleDefs(523) =   "Splits(0).Columns(121).EditorStyle:id=1020,.parent=17"
      _StyleDefs(524) =   "Splits(0).Columns(122).Style:id=1029,.parent=13"
      _StyleDefs(525) =   "Splits(0).Columns(122).HeadingStyle:id=1026,.parent=14"
      _StyleDefs(526) =   "Splits(0).Columns(122).FooterStyle:id=1027,.parent=15"
      _StyleDefs(527) =   "Splits(0).Columns(122).EditorStyle:id=1028,.parent=17"
      _StyleDefs(528) =   "Splits(0).Columns(123).Style:id=1037,.parent=13"
      _StyleDefs(529) =   "Splits(0).Columns(123).HeadingStyle:id=1034,.parent=14"
      _StyleDefs(530) =   "Splits(0).Columns(123).FooterStyle:id=1035,.parent=15"
      _StyleDefs(531) =   "Splits(0).Columns(123).EditorStyle:id=1036,.parent=17"
      _StyleDefs(532) =   "Splits(0).Columns(124).Style:id=1045,.parent=13"
      _StyleDefs(533) =   "Splits(0).Columns(124).HeadingStyle:id=1042,.parent=14"
      _StyleDefs(534) =   "Splits(0).Columns(124).FooterStyle:id=1043,.parent=15"
      _StyleDefs(535) =   "Splits(0).Columns(124).EditorStyle:id=1044,.parent=17"
      _StyleDefs(536) =   "Splits(0).Columns(125).Style:id=1053,.parent=13"
      _StyleDefs(537) =   "Splits(0).Columns(125).HeadingStyle:id=1050,.parent=14"
      _StyleDefs(538) =   "Splits(0).Columns(125).FooterStyle:id=1051,.parent=15"
      _StyleDefs(539) =   "Splits(0).Columns(125).EditorStyle:id=1052,.parent=17"
      _StyleDefs(540) =   "Splits(0).Columns(126).Style:id=1061,.parent=13"
      _StyleDefs(541) =   "Splits(0).Columns(126).HeadingStyle:id=1058,.parent=14"
      _StyleDefs(542) =   "Splits(0).Columns(126).FooterStyle:id=1059,.parent=15"
      _StyleDefs(543) =   "Splits(0).Columns(126).EditorStyle:id=1060,.parent=17"
      _StyleDefs(544) =   "Splits(0).Columns(127).Style:id=1069,.parent=13"
      _StyleDefs(545) =   "Splits(0).Columns(127).HeadingStyle:id=1066,.parent=14"
      _StyleDefs(546) =   "Splits(0).Columns(127).FooterStyle:id=1067,.parent=15"
      _StyleDefs(547) =   "Splits(0).Columns(127).EditorStyle:id=1068,.parent=17"
      _StyleDefs(548) =   "Splits(0).Columns(128).Style:id=1077,.parent=13"
      _StyleDefs(549) =   "Splits(0).Columns(128).HeadingStyle:id=1074,.parent=14"
      _StyleDefs(550) =   "Splits(0).Columns(128).FooterStyle:id=1075,.parent=15"
      _StyleDefs(551) =   "Splits(0).Columns(128).EditorStyle:id=1076,.parent=17"
      _StyleDefs(552) =   "Splits(0).Columns(129).Style:id=1085,.parent=13"
      _StyleDefs(553) =   "Splits(0).Columns(129).HeadingStyle:id=1082,.parent=14"
      _StyleDefs(554) =   "Splits(0).Columns(129).FooterStyle:id=1083,.parent=15"
      _StyleDefs(555) =   "Splits(0).Columns(129).EditorStyle:id=1084,.parent=17"
      _StyleDefs(556) =   "Splits(0).Columns(130).Style:id=1093,.parent=13"
      _StyleDefs(557) =   "Splits(0).Columns(130).HeadingStyle:id=1090,.parent=14"
      _StyleDefs(558) =   "Splits(0).Columns(130).FooterStyle:id=1091,.parent=15"
      _StyleDefs(559) =   "Splits(0).Columns(130).EditorStyle:id=1092,.parent=17"
      _StyleDefs(560) =   "Splits(0).Columns(131).Style:id=1101,.parent=13"
      _StyleDefs(561) =   "Splits(0).Columns(131).HeadingStyle:id=1098,.parent=14"
      _StyleDefs(562) =   "Splits(0).Columns(131).FooterStyle:id=1099,.parent=15"
      _StyleDefs(563) =   "Splits(0).Columns(131).EditorStyle:id=1100,.parent=17"
      _StyleDefs(564) =   "Splits(0).Columns(132).Style:id=1109,.parent=13"
      _StyleDefs(565) =   "Splits(0).Columns(132).HeadingStyle:id=1106,.parent=14"
      _StyleDefs(566) =   "Splits(0).Columns(132).FooterStyle:id=1107,.parent=15"
      _StyleDefs(567) =   "Splits(0).Columns(132).EditorStyle:id=1108,.parent=17"
      _StyleDefs(568) =   "Splits(0).Columns(133).Style:id=1117,.parent=13"
      _StyleDefs(569) =   "Splits(0).Columns(133).HeadingStyle:id=1114,.parent=14"
      _StyleDefs(570) =   "Splits(0).Columns(133).FooterStyle:id=1115,.parent=15"
      _StyleDefs(571) =   "Splits(0).Columns(133).EditorStyle:id=1116,.parent=17"
      _StyleDefs(572) =   "Splits(0).Columns(134).Style:id=1125,.parent=13"
      _StyleDefs(573) =   "Splits(0).Columns(134).HeadingStyle:id=1122,.parent=14"
      _StyleDefs(574) =   "Splits(0).Columns(134).FooterStyle:id=1123,.parent=15"
      _StyleDefs(575) =   "Splits(0).Columns(134).EditorStyle:id=1124,.parent=17"
      _StyleDefs(576) =   "Splits(0).Columns(135).Style:id=1138,.parent=13"
      _StyleDefs(577) =   "Splits(0).Columns(135).HeadingStyle:id=1135,.parent=14"
      _StyleDefs(578) =   "Splits(0).Columns(135).FooterStyle:id=1136,.parent=15"
      _StyleDefs(579) =   "Splits(0).Columns(135).EditorStyle:id=1137,.parent=17"
      _StyleDefs(580) =   "Splits(0).Columns(136).Style:id=1146,.parent=13"
      _StyleDefs(581) =   "Splits(0).Columns(136).HeadingStyle:id=1143,.parent=14"
      _StyleDefs(582) =   "Splits(0).Columns(136).FooterStyle:id=1144,.parent=15"
      _StyleDefs(583) =   "Splits(0).Columns(136).EditorStyle:id=1145,.parent=17"
      _StyleDefs(584) =   "Splits(0).Columns(137).Style:id=1154,.parent=13"
      _StyleDefs(585) =   "Splits(0).Columns(137).HeadingStyle:id=1151,.parent=14"
      _StyleDefs(586) =   "Splits(0).Columns(137).FooterStyle:id=1152,.parent=15"
      _StyleDefs(587) =   "Splits(0).Columns(137).EditorStyle:id=1153,.parent=17"
      _StyleDefs(588) =   "Splits(1).Style:id=59,.parent=1"
      _StyleDefs(589) =   "Splits(1).CaptionStyle:id=68,.parent=4"
      _StyleDefs(590) =   "Splits(1).HeadingStyle:id=60,.parent=2"
      _StyleDefs(591) =   "Splits(1).FooterStyle:id=61,.parent=3"
      _StyleDefs(592) =   "Splits(1).InactiveStyle:id=62,.parent=5"
      _StyleDefs(593) =   "Splits(1).SelectedStyle:id=64,.parent=6"
      _StyleDefs(594) =   "Splits(1).EditorStyle:id=63,.parent=7"
      _StyleDefs(595) =   "Splits(1).HighlightRowStyle:id=65,.parent=8"
      _StyleDefs(596) =   "Splits(1).EvenRowStyle:id=66,.parent=9,.bgcolor=&H80FFFF&"
      _StyleDefs(597) =   "Splits(1).OddRowStyle:id=67,.parent=10,.bgcolor=&HFFFFFF&"
      _StyleDefs(598) =   "Splits(1).RecordSelectorStyle:id=69,.parent=11"
      _StyleDefs(599) =   "Splits(1).FilterBarStyle:id=70,.parent=12"
      _StyleDefs(600) =   "Splits(1).Columns(0).Style:id=249,.parent=59"
      _StyleDefs(601) =   "Splits(1).Columns(0).HeadingStyle:id=246,.parent=60"
      _StyleDefs(602) =   "Splits(1).Columns(0).FooterStyle:id=247,.parent=61"
      _StyleDefs(603) =   "Splits(1).Columns(0).EditorStyle:id=248,.parent=63"
      _StyleDefs(604) =   "Splits(1).Columns(1).Style:id=241,.parent=59"
      _StyleDefs(605) =   "Splits(1).Columns(1).HeadingStyle:id=238,.parent=60"
      _StyleDefs(606) =   "Splits(1).Columns(1).FooterStyle:id=239,.parent=61"
      _StyleDefs(607) =   "Splits(1).Columns(1).EditorStyle:id=240,.parent=63"
      _StyleDefs(608) =   "Splits(1).Columns(2).Style:id=233,.parent=59"
      _StyleDefs(609) =   "Splits(1).Columns(2).HeadingStyle:id=230,.parent=60"
      _StyleDefs(610) =   "Splits(1).Columns(2).FooterStyle:id=231,.parent=61"
      _StyleDefs(611) =   "Splits(1).Columns(2).EditorStyle:id=232,.parent=63"
      _StyleDefs(612) =   "Splits(1).Columns(3).Style:id=225,.parent=59"
      _StyleDefs(613) =   "Splits(1).Columns(3).HeadingStyle:id=222,.parent=60"
      _StyleDefs(614) =   "Splits(1).Columns(3).FooterStyle:id=223,.parent=61"
      _StyleDefs(615) =   "Splits(1).Columns(3).EditorStyle:id=224,.parent=63"
      _StyleDefs(616) =   "Splits(1).Columns(4).Style:id=217,.parent=59"
      _StyleDefs(617) =   "Splits(1).Columns(4).HeadingStyle:id=214,.parent=60"
      _StyleDefs(618) =   "Splits(1).Columns(4).FooterStyle:id=215,.parent=61"
      _StyleDefs(619) =   "Splits(1).Columns(4).EditorStyle:id=216,.parent=63"
      _StyleDefs(620) =   "Splits(1).Columns(5).Style:id=209,.parent=59"
      _StyleDefs(621) =   "Splits(1).Columns(5).HeadingStyle:id=206,.parent=60"
      _StyleDefs(622) =   "Splits(1).Columns(5).FooterStyle:id=207,.parent=61"
      _StyleDefs(623) =   "Splits(1).Columns(5).EditorStyle:id=208,.parent=63"
      _StyleDefs(624) =   "Splits(1).Columns(6).Style:id=201,.parent=59"
      _StyleDefs(625) =   "Splits(1).Columns(6).HeadingStyle:id=198,.parent=60"
      _StyleDefs(626) =   "Splits(1).Columns(6).FooterStyle:id=199,.parent=61"
      _StyleDefs(627) =   "Splits(1).Columns(6).EditorStyle:id=200,.parent=63"
      _StyleDefs(628) =   "Splits(1).Columns(7).Style:id=193,.parent=59"
      _StyleDefs(629) =   "Splits(1).Columns(7).HeadingStyle:id=190,.parent=60"
      _StyleDefs(630) =   "Splits(1).Columns(7).FooterStyle:id=191,.parent=61"
      _StyleDefs(631) =   "Splits(1).Columns(7).EditorStyle:id=192,.parent=63"
      _StyleDefs(632) =   "Splits(1).Columns(8).Style:id=185,.parent=59"
      _StyleDefs(633) =   "Splits(1).Columns(8).HeadingStyle:id=182,.parent=60"
      _StyleDefs(634) =   "Splits(1).Columns(8).FooterStyle:id=183,.parent=61"
      _StyleDefs(635) =   "Splits(1).Columns(8).EditorStyle:id=184,.parent=63"
      _StyleDefs(636) =   "Splits(1).Columns(9).Style:id=177,.parent=59"
      _StyleDefs(637) =   "Splits(1).Columns(9).HeadingStyle:id=174,.parent=60"
      _StyleDefs(638) =   "Splits(1).Columns(9).FooterStyle:id=175,.parent=61"
      _StyleDefs(639) =   "Splits(1).Columns(9).EditorStyle:id=176,.parent=63"
      _StyleDefs(640) =   "Splits(1).Columns(10).Style:id=169,.parent=59"
      _StyleDefs(641) =   "Splits(1).Columns(10).HeadingStyle:id=166,.parent=60"
      _StyleDefs(642) =   "Splits(1).Columns(10).FooterStyle:id=167,.parent=61"
      _StyleDefs(643) =   "Splits(1).Columns(10).EditorStyle:id=168,.parent=63"
      _StyleDefs(644) =   "Splits(1).Columns(11).Style:id=161,.parent=59"
      _StyleDefs(645) =   "Splits(1).Columns(11).HeadingStyle:id=158,.parent=60"
      _StyleDefs(646) =   "Splits(1).Columns(11).FooterStyle:id=159,.parent=61"
      _StyleDefs(647) =   "Splits(1).Columns(11).EditorStyle:id=160,.parent=63"
      _StyleDefs(648) =   "Splits(1).Columns(12).Style:id=153,.parent=59"
      _StyleDefs(649) =   "Splits(1).Columns(12).HeadingStyle:id=150,.parent=60"
      _StyleDefs(650) =   "Splits(1).Columns(12).FooterStyle:id=151,.parent=61"
      _StyleDefs(651) =   "Splits(1).Columns(12).EditorStyle:id=152,.parent=63"
      _StyleDefs(652) =   "Splits(1).Columns(13).Style:id=145,.parent=59"
      _StyleDefs(653) =   "Splits(1).Columns(13).HeadingStyle:id=142,.parent=60"
      _StyleDefs(654) =   "Splits(1).Columns(13).FooterStyle:id=143,.parent=61"
      _StyleDefs(655) =   "Splits(1).Columns(13).EditorStyle:id=144,.parent=63"
      _StyleDefs(656) =   "Splits(1).Columns(14).Style:id=137,.parent=59"
      _StyleDefs(657) =   "Splits(1).Columns(14).HeadingStyle:id=134,.parent=60"
      _StyleDefs(658) =   "Splits(1).Columns(14).FooterStyle:id=135,.parent=61"
      _StyleDefs(659) =   "Splits(1).Columns(14).EditorStyle:id=136,.parent=63"
      _StyleDefs(660) =   "Splits(1).Columns(15).Style:id=127,.parent=59"
      _StyleDefs(661) =   "Splits(1).Columns(15).HeadingStyle:id=124,.parent=60"
      _StyleDefs(662) =   "Splits(1).Columns(15).FooterStyle:id=125,.parent=61"
      _StyleDefs(663) =   "Splits(1).Columns(15).EditorStyle:id=126,.parent=63"
      _StyleDefs(664) =   "Splits(1).Columns(16).Style:id=119,.parent=59"
      _StyleDefs(665) =   "Splits(1).Columns(16).HeadingStyle:id=116,.parent=60"
      _StyleDefs(666) =   "Splits(1).Columns(16).FooterStyle:id=117,.parent=61"
      _StyleDefs(667) =   "Splits(1).Columns(16).EditorStyle:id=118,.parent=63"
      _StyleDefs(668) =   "Splits(1).Columns(17).Style:id=74,.parent=59"
      _StyleDefs(669) =   "Splits(1).Columns(17).HeadingStyle:id=71,.parent=60"
      _StyleDefs(670) =   "Splits(1).Columns(17).FooterStyle:id=72,.parent=61"
      _StyleDefs(671) =   "Splits(1).Columns(17).EditorStyle:id=73,.parent=63"
      _StyleDefs(672) =   "Splits(1).Columns(18).Style:id=78,.parent=59"
      _StyleDefs(673) =   "Splits(1).Columns(18).HeadingStyle:id=75,.parent=60"
      _StyleDefs(674) =   "Splits(1).Columns(18).FooterStyle:id=76,.parent=61"
      _StyleDefs(675) =   "Splits(1).Columns(18).EditorStyle:id=77,.parent=63"
      _StyleDefs(676) =   "Splits(1).Columns(19).Style:id=82,.parent=59"
      _StyleDefs(677) =   "Splits(1).Columns(19).HeadingStyle:id=79,.parent=60"
      _StyleDefs(678) =   "Splits(1).Columns(19).FooterStyle:id=80,.parent=61"
      _StyleDefs(679) =   "Splits(1).Columns(19).EditorStyle:id=81,.parent=63"
      _StyleDefs(680) =   "Splits(1).Columns(20).Style:id=95,.parent=59"
      _StyleDefs(681) =   "Splits(1).Columns(20).HeadingStyle:id=92,.parent=60"
      _StyleDefs(682) =   "Splits(1).Columns(20).FooterStyle:id=93,.parent=61"
      _StyleDefs(683) =   "Splits(1).Columns(20).EditorStyle:id=94,.parent=63"
      _StyleDefs(684) =   "Splits(1).Columns(21).Style:id=99,.parent=59"
      _StyleDefs(685) =   "Splits(1).Columns(21).HeadingStyle:id=96,.parent=60"
      _StyleDefs(686) =   "Splits(1).Columns(21).FooterStyle:id=97,.parent=61"
      _StyleDefs(687) =   "Splits(1).Columns(21).EditorStyle:id=98,.parent=63"
      _StyleDefs(688) =   "Splits(1).Columns(22).Style:id=103,.parent=59"
      _StyleDefs(689) =   "Splits(1).Columns(22).HeadingStyle:id=100,.parent=60"
      _StyleDefs(690) =   "Splits(1).Columns(22).FooterStyle:id=101,.parent=61"
      _StyleDefs(691) =   "Splits(1).Columns(22).EditorStyle:id=102,.parent=63"
      _StyleDefs(692) =   "Splits(1).Columns(23).Style:id=107,.parent=59"
      _StyleDefs(693) =   "Splits(1).Columns(23).HeadingStyle:id=104,.parent=60"
      _StyleDefs(694) =   "Splits(1).Columns(23).FooterStyle:id=105,.parent=61"
      _StyleDefs(695) =   "Splits(1).Columns(23).EditorStyle:id=106,.parent=63"
      _StyleDefs(696) =   "Splits(1).Columns(24).Style:id=111,.parent=59"
      _StyleDefs(697) =   "Splits(1).Columns(24).HeadingStyle:id=108,.parent=60"
      _StyleDefs(698) =   "Splits(1).Columns(24).FooterStyle:id=109,.parent=61"
      _StyleDefs(699) =   "Splits(1).Columns(24).EditorStyle:id=110,.parent=63"
      _StyleDefs(700) =   "Splits(1).Columns(25).Style:id=257,.parent=59"
      _StyleDefs(701) =   "Splits(1).Columns(25).HeadingStyle:id=254,.parent=60"
      _StyleDefs(702) =   "Splits(1).Columns(25).FooterStyle:id=255,.parent=61"
      _StyleDefs(703) =   "Splits(1).Columns(25).EditorStyle:id=256,.parent=63"
      _StyleDefs(704) =   "Splits(1).Columns(26).Style:id=265,.parent=59"
      _StyleDefs(705) =   "Splits(1).Columns(26).HeadingStyle:id=262,.parent=60"
      _StyleDefs(706) =   "Splits(1).Columns(26).FooterStyle:id=263,.parent=61"
      _StyleDefs(707) =   "Splits(1).Columns(26).EditorStyle:id=264,.parent=63"
      _StyleDefs(708) =   "Splits(1).Columns(27).Style:id=273,.parent=59"
      _StyleDefs(709) =   "Splits(1).Columns(27).HeadingStyle:id=270,.parent=60"
      _StyleDefs(710) =   "Splits(1).Columns(27).FooterStyle:id=271,.parent=61"
      _StyleDefs(711) =   "Splits(1).Columns(27).EditorStyle:id=272,.parent=63"
      _StyleDefs(712) =   "Splits(1).Columns(28).Style:id=281,.parent=59"
      _StyleDefs(713) =   "Splits(1).Columns(28).HeadingStyle:id=278,.parent=60"
      _StyleDefs(714) =   "Splits(1).Columns(28).FooterStyle:id=279,.parent=61"
      _StyleDefs(715) =   "Splits(1).Columns(28).EditorStyle:id=280,.parent=63"
      _StyleDefs(716) =   "Splits(1).Columns(29).Style:id=289,.parent=59"
      _StyleDefs(717) =   "Splits(1).Columns(29).HeadingStyle:id=286,.parent=60"
      _StyleDefs(718) =   "Splits(1).Columns(29).FooterStyle:id=287,.parent=61"
      _StyleDefs(719) =   "Splits(1).Columns(29).EditorStyle:id=288,.parent=63"
      _StyleDefs(720) =   "Splits(1).Columns(30).Style:id=297,.parent=59"
      _StyleDefs(721) =   "Splits(1).Columns(30).HeadingStyle:id=294,.parent=60"
      _StyleDefs(722) =   "Splits(1).Columns(30).FooterStyle:id=295,.parent=61"
      _StyleDefs(723) =   "Splits(1).Columns(30).EditorStyle:id=296,.parent=63"
      _StyleDefs(724) =   "Splits(1).Columns(31).Style:id=305,.parent=59"
      _StyleDefs(725) =   "Splits(1).Columns(31).HeadingStyle:id=302,.parent=60"
      _StyleDefs(726) =   "Splits(1).Columns(31).FooterStyle:id=303,.parent=61"
      _StyleDefs(727) =   "Splits(1).Columns(31).EditorStyle:id=304,.parent=63"
      _StyleDefs(728) =   "Splits(1).Columns(32).Style:id=313,.parent=59"
      _StyleDefs(729) =   "Splits(1).Columns(32).HeadingStyle:id=310,.parent=60"
      _StyleDefs(730) =   "Splits(1).Columns(32).FooterStyle:id=311,.parent=61"
      _StyleDefs(731) =   "Splits(1).Columns(32).EditorStyle:id=312,.parent=63"
      _StyleDefs(732) =   "Splits(1).Columns(33).Style:id=321,.parent=59"
      _StyleDefs(733) =   "Splits(1).Columns(33).HeadingStyle:id=318,.parent=60"
      _StyleDefs(734) =   "Splits(1).Columns(33).FooterStyle:id=319,.parent=61"
      _StyleDefs(735) =   "Splits(1).Columns(33).EditorStyle:id=320,.parent=63"
      _StyleDefs(736) =   "Splits(1).Columns(34).Style:id=329,.parent=59"
      _StyleDefs(737) =   "Splits(1).Columns(34).HeadingStyle:id=326,.parent=60"
      _StyleDefs(738) =   "Splits(1).Columns(34).FooterStyle:id=327,.parent=61"
      _StyleDefs(739) =   "Splits(1).Columns(34).EditorStyle:id=328,.parent=63"
      _StyleDefs(740) =   "Splits(1).Columns(35).Style:id=337,.parent=59"
      _StyleDefs(741) =   "Splits(1).Columns(35).HeadingStyle:id=334,.parent=60"
      _StyleDefs(742) =   "Splits(1).Columns(35).FooterStyle:id=335,.parent=61"
      _StyleDefs(743) =   "Splits(1).Columns(35).EditorStyle:id=336,.parent=63"
      _StyleDefs(744) =   "Splits(1).Columns(36).Style:id=345,.parent=59"
      _StyleDefs(745) =   "Splits(1).Columns(36).HeadingStyle:id=342,.parent=60"
      _StyleDefs(746) =   "Splits(1).Columns(36).FooterStyle:id=343,.parent=61"
      _StyleDefs(747) =   "Splits(1).Columns(36).EditorStyle:id=344,.parent=63"
      _StyleDefs(748) =   "Splits(1).Columns(37).Style:id=353,.parent=59"
      _StyleDefs(749) =   "Splits(1).Columns(37).HeadingStyle:id=350,.parent=60"
      _StyleDefs(750) =   "Splits(1).Columns(37).FooterStyle:id=351,.parent=61"
      _StyleDefs(751) =   "Splits(1).Columns(37).EditorStyle:id=352,.parent=63"
      _StyleDefs(752) =   "Splits(1).Columns(38).Style:id=361,.parent=59"
      _StyleDefs(753) =   "Splits(1).Columns(38).HeadingStyle:id=358,.parent=60"
      _StyleDefs(754) =   "Splits(1).Columns(38).FooterStyle:id=359,.parent=61"
      _StyleDefs(755) =   "Splits(1).Columns(38).EditorStyle:id=360,.parent=63"
      _StyleDefs(756) =   "Splits(1).Columns(39).Style:id=369,.parent=59"
      _StyleDefs(757) =   "Splits(1).Columns(39).HeadingStyle:id=366,.parent=60"
      _StyleDefs(758) =   "Splits(1).Columns(39).FooterStyle:id=367,.parent=61"
      _StyleDefs(759) =   "Splits(1).Columns(39).EditorStyle:id=368,.parent=63"
      _StyleDefs(760) =   "Splits(1).Columns(40).Style:id=377,.parent=59"
      _StyleDefs(761) =   "Splits(1).Columns(40).HeadingStyle:id=374,.parent=60"
      _StyleDefs(762) =   "Splits(1).Columns(40).FooterStyle:id=375,.parent=61"
      _StyleDefs(763) =   "Splits(1).Columns(40).EditorStyle:id=376,.parent=63"
      _StyleDefs(764) =   "Splits(1).Columns(41).Style:id=385,.parent=59"
      _StyleDefs(765) =   "Splits(1).Columns(41).HeadingStyle:id=382,.parent=60"
      _StyleDefs(766) =   "Splits(1).Columns(41).FooterStyle:id=383,.parent=61"
      _StyleDefs(767) =   "Splits(1).Columns(41).EditorStyle:id=384,.parent=63"
      _StyleDefs(768) =   "Splits(1).Columns(42).Style:id=393,.parent=59"
      _StyleDefs(769) =   "Splits(1).Columns(42).HeadingStyle:id=390,.parent=60"
      _StyleDefs(770) =   "Splits(1).Columns(42).FooterStyle:id=391,.parent=61"
      _StyleDefs(771) =   "Splits(1).Columns(42).EditorStyle:id=392,.parent=63"
      _StyleDefs(772) =   "Splits(1).Columns(43).Style:id=401,.parent=59"
      _StyleDefs(773) =   "Splits(1).Columns(43).HeadingStyle:id=398,.parent=60"
      _StyleDefs(774) =   "Splits(1).Columns(43).FooterStyle:id=399,.parent=61"
      _StyleDefs(775) =   "Splits(1).Columns(43).EditorStyle:id=400,.parent=63"
      _StyleDefs(776) =   "Splits(1).Columns(44).Style:id=409,.parent=59"
      _StyleDefs(777) =   "Splits(1).Columns(44).HeadingStyle:id=406,.parent=60"
      _StyleDefs(778) =   "Splits(1).Columns(44).FooterStyle:id=407,.parent=61"
      _StyleDefs(779) =   "Splits(1).Columns(44).EditorStyle:id=408,.parent=63"
      _StyleDefs(780) =   "Splits(1).Columns(45).Style:id=417,.parent=59"
      _StyleDefs(781) =   "Splits(1).Columns(45).HeadingStyle:id=414,.parent=60"
      _StyleDefs(782) =   "Splits(1).Columns(45).FooterStyle:id=415,.parent=61"
      _StyleDefs(783) =   "Splits(1).Columns(45).EditorStyle:id=416,.parent=63"
      _StyleDefs(784) =   "Splits(1).Columns(46).Style:id=425,.parent=59"
      _StyleDefs(785) =   "Splits(1).Columns(46).HeadingStyle:id=422,.parent=60"
      _StyleDefs(786) =   "Splits(1).Columns(46).FooterStyle:id=423,.parent=61"
      _StyleDefs(787) =   "Splits(1).Columns(46).EditorStyle:id=424,.parent=63"
      _StyleDefs(788) =   "Splits(1).Columns(47).Style:id=433,.parent=59"
      _StyleDefs(789) =   "Splits(1).Columns(47).HeadingStyle:id=430,.parent=60"
      _StyleDefs(790) =   "Splits(1).Columns(47).FooterStyle:id=431,.parent=61"
      _StyleDefs(791) =   "Splits(1).Columns(47).EditorStyle:id=432,.parent=63"
      _StyleDefs(792) =   "Splits(1).Columns(48).Style:id=441,.parent=59"
      _StyleDefs(793) =   "Splits(1).Columns(48).HeadingStyle:id=438,.parent=60"
      _StyleDefs(794) =   "Splits(1).Columns(48).FooterStyle:id=439,.parent=61"
      _StyleDefs(795) =   "Splits(1).Columns(48).EditorStyle:id=440,.parent=63"
      _StyleDefs(796) =   "Splits(1).Columns(49).Style:id=449,.parent=59"
      _StyleDefs(797) =   "Splits(1).Columns(49).HeadingStyle:id=446,.parent=60"
      _StyleDefs(798) =   "Splits(1).Columns(49).FooterStyle:id=447,.parent=61"
      _StyleDefs(799) =   "Splits(1).Columns(49).EditorStyle:id=448,.parent=63"
      _StyleDefs(800) =   "Splits(1).Columns(50).Style:id=457,.parent=59"
      _StyleDefs(801) =   "Splits(1).Columns(50).HeadingStyle:id=454,.parent=60"
      _StyleDefs(802) =   "Splits(1).Columns(50).FooterStyle:id=455,.parent=61"
      _StyleDefs(803) =   "Splits(1).Columns(50).EditorStyle:id=456,.parent=63"
      _StyleDefs(804) =   "Splits(1).Columns(51).Style:id=465,.parent=59"
      _StyleDefs(805) =   "Splits(1).Columns(51).HeadingStyle:id=462,.parent=60"
      _StyleDefs(806) =   "Splits(1).Columns(51).FooterStyle:id=463,.parent=61"
      _StyleDefs(807) =   "Splits(1).Columns(51).EditorStyle:id=464,.parent=63"
      _StyleDefs(808) =   "Splits(1).Columns(52).Style:id=473,.parent=59"
      _StyleDefs(809) =   "Splits(1).Columns(52).HeadingStyle:id=470,.parent=60"
      _StyleDefs(810) =   "Splits(1).Columns(52).FooterStyle:id=471,.parent=61"
      _StyleDefs(811) =   "Splits(1).Columns(52).EditorStyle:id=472,.parent=63"
      _StyleDefs(812) =   "Splits(1).Columns(53).Style:id=481,.parent=59"
      _StyleDefs(813) =   "Splits(1).Columns(53).HeadingStyle:id=478,.parent=60"
      _StyleDefs(814) =   "Splits(1).Columns(53).FooterStyle:id=479,.parent=61"
      _StyleDefs(815) =   "Splits(1).Columns(53).EditorStyle:id=480,.parent=63"
      _StyleDefs(816) =   "Splits(1).Columns(54).Style:id=489,.parent=59"
      _StyleDefs(817) =   "Splits(1).Columns(54).HeadingStyle:id=486,.parent=60"
      _StyleDefs(818) =   "Splits(1).Columns(54).FooterStyle:id=487,.parent=61"
      _StyleDefs(819) =   "Splits(1).Columns(54).EditorStyle:id=488,.parent=63"
      _StyleDefs(820) =   "Splits(1).Columns(55).Style:id=497,.parent=59"
      _StyleDefs(821) =   "Splits(1).Columns(55).HeadingStyle:id=494,.parent=60"
      _StyleDefs(822) =   "Splits(1).Columns(55).FooterStyle:id=495,.parent=61"
      _StyleDefs(823) =   "Splits(1).Columns(55).EditorStyle:id=496,.parent=63"
      _StyleDefs(824) =   "Splits(1).Columns(56).Style:id=505,.parent=59"
      _StyleDefs(825) =   "Splits(1).Columns(56).HeadingStyle:id=502,.parent=60"
      _StyleDefs(826) =   "Splits(1).Columns(56).FooterStyle:id=503,.parent=61"
      _StyleDefs(827) =   "Splits(1).Columns(56).EditorStyle:id=504,.parent=63"
      _StyleDefs(828) =   "Splits(1).Columns(57).Style:id=513,.parent=59"
      _StyleDefs(829) =   "Splits(1).Columns(57).HeadingStyle:id=510,.parent=60"
      _StyleDefs(830) =   "Splits(1).Columns(57).FooterStyle:id=511,.parent=61"
      _StyleDefs(831) =   "Splits(1).Columns(57).EditorStyle:id=512,.parent=63"
      _StyleDefs(832) =   "Splits(1).Columns(58).Style:id=521,.parent=59"
      _StyleDefs(833) =   "Splits(1).Columns(58).HeadingStyle:id=518,.parent=60"
      _StyleDefs(834) =   "Splits(1).Columns(58).FooterStyle:id=519,.parent=61"
      _StyleDefs(835) =   "Splits(1).Columns(58).EditorStyle:id=520,.parent=63"
      _StyleDefs(836) =   "Splits(1).Columns(59).Style:id=529,.parent=59"
      _StyleDefs(837) =   "Splits(1).Columns(59).HeadingStyle:id=526,.parent=60"
      _StyleDefs(838) =   "Splits(1).Columns(59).FooterStyle:id=527,.parent=61"
      _StyleDefs(839) =   "Splits(1).Columns(59).EditorStyle:id=528,.parent=63"
      _StyleDefs(840) =   "Splits(1).Columns(60).Style:id=537,.parent=59"
      _StyleDefs(841) =   "Splits(1).Columns(60).HeadingStyle:id=534,.parent=60"
      _StyleDefs(842) =   "Splits(1).Columns(60).FooterStyle:id=535,.parent=61"
      _StyleDefs(843) =   "Splits(1).Columns(60).EditorStyle:id=536,.parent=63"
      _StyleDefs(844) =   "Splits(1).Columns(61).Style:id=545,.parent=59"
      _StyleDefs(845) =   "Splits(1).Columns(61).HeadingStyle:id=542,.parent=60"
      _StyleDefs(846) =   "Splits(1).Columns(61).FooterStyle:id=543,.parent=61"
      _StyleDefs(847) =   "Splits(1).Columns(61).EditorStyle:id=544,.parent=63"
      _StyleDefs(848) =   "Splits(1).Columns(62).Style:id=553,.parent=59"
      _StyleDefs(849) =   "Splits(1).Columns(62).HeadingStyle:id=550,.parent=60"
      _StyleDefs(850) =   "Splits(1).Columns(62).FooterStyle:id=551,.parent=61"
      _StyleDefs(851) =   "Splits(1).Columns(62).EditorStyle:id=552,.parent=63"
      _StyleDefs(852) =   "Splits(1).Columns(63).Style:id=561,.parent=59"
      _StyleDefs(853) =   "Splits(1).Columns(63).HeadingStyle:id=558,.parent=60"
      _StyleDefs(854) =   "Splits(1).Columns(63).FooterStyle:id=559,.parent=61"
      _StyleDefs(855) =   "Splits(1).Columns(63).EditorStyle:id=560,.parent=63"
      _StyleDefs(856) =   "Splits(1).Columns(64).Style:id=569,.parent=59"
      _StyleDefs(857) =   "Splits(1).Columns(64).HeadingStyle:id=566,.parent=60"
      _StyleDefs(858) =   "Splits(1).Columns(64).FooterStyle:id=567,.parent=61"
      _StyleDefs(859) =   "Splits(1).Columns(64).EditorStyle:id=568,.parent=63"
      _StyleDefs(860) =   "Splits(1).Columns(65).Style:id=577,.parent=59"
      _StyleDefs(861) =   "Splits(1).Columns(65).HeadingStyle:id=574,.parent=60"
      _StyleDefs(862) =   "Splits(1).Columns(65).FooterStyle:id=575,.parent=61"
      _StyleDefs(863) =   "Splits(1).Columns(65).EditorStyle:id=576,.parent=63"
      _StyleDefs(864) =   "Splits(1).Columns(66).Style:id=585,.parent=59"
      _StyleDefs(865) =   "Splits(1).Columns(66).HeadingStyle:id=582,.parent=60"
      _StyleDefs(866) =   "Splits(1).Columns(66).FooterStyle:id=583,.parent=61"
      _StyleDefs(867) =   "Splits(1).Columns(66).EditorStyle:id=584,.parent=63"
      _StyleDefs(868) =   "Splits(1).Columns(67).Style:id=593,.parent=59"
      _StyleDefs(869) =   "Splits(1).Columns(67).HeadingStyle:id=590,.parent=60"
      _StyleDefs(870) =   "Splits(1).Columns(67).FooterStyle:id=591,.parent=61"
      _StyleDefs(871) =   "Splits(1).Columns(67).EditorStyle:id=592,.parent=63"
      _StyleDefs(872) =   "Splits(1).Columns(68).Style:id=601,.parent=59"
      _StyleDefs(873) =   "Splits(1).Columns(68).HeadingStyle:id=598,.parent=60"
      _StyleDefs(874) =   "Splits(1).Columns(68).FooterStyle:id=599,.parent=61"
      _StyleDefs(875) =   "Splits(1).Columns(68).EditorStyle:id=600,.parent=63"
      _StyleDefs(876) =   "Splits(1).Columns(69).Style:id=609,.parent=59"
      _StyleDefs(877) =   "Splits(1).Columns(69).HeadingStyle:id=606,.parent=60"
      _StyleDefs(878) =   "Splits(1).Columns(69).FooterStyle:id=607,.parent=61"
      _StyleDefs(879) =   "Splits(1).Columns(69).EditorStyle:id=608,.parent=63"
      _StyleDefs(880) =   "Splits(1).Columns(70).Style:id=617,.parent=59"
      _StyleDefs(881) =   "Splits(1).Columns(70).HeadingStyle:id=614,.parent=60"
      _StyleDefs(882) =   "Splits(1).Columns(70).FooterStyle:id=615,.parent=61"
      _StyleDefs(883) =   "Splits(1).Columns(70).EditorStyle:id=616,.parent=63"
      _StyleDefs(884) =   "Splits(1).Columns(71).Style:id=625,.parent=59"
      _StyleDefs(885) =   "Splits(1).Columns(71).HeadingStyle:id=622,.parent=60"
      _StyleDefs(886) =   "Splits(1).Columns(71).FooterStyle:id=623,.parent=61"
      _StyleDefs(887) =   "Splits(1).Columns(71).EditorStyle:id=624,.parent=63"
      _StyleDefs(888) =   "Splits(1).Columns(72).Style:id=633,.parent=59"
      _StyleDefs(889) =   "Splits(1).Columns(72).HeadingStyle:id=630,.parent=60"
      _StyleDefs(890) =   "Splits(1).Columns(72).FooterStyle:id=631,.parent=61"
      _StyleDefs(891) =   "Splits(1).Columns(72).EditorStyle:id=632,.parent=63"
      _StyleDefs(892) =   "Splits(1).Columns(73).Style:id=641,.parent=59"
      _StyleDefs(893) =   "Splits(1).Columns(73).HeadingStyle:id=638,.parent=60"
      _StyleDefs(894) =   "Splits(1).Columns(73).FooterStyle:id=639,.parent=61"
      _StyleDefs(895) =   "Splits(1).Columns(73).EditorStyle:id=640,.parent=63"
      _StyleDefs(896) =   "Splits(1).Columns(74).Style:id=649,.parent=59"
      _StyleDefs(897) =   "Splits(1).Columns(74).HeadingStyle:id=646,.parent=60"
      _StyleDefs(898) =   "Splits(1).Columns(74).FooterStyle:id=647,.parent=61"
      _StyleDefs(899) =   "Splits(1).Columns(74).EditorStyle:id=648,.parent=63"
      _StyleDefs(900) =   "Splits(1).Columns(75).Style:id=657,.parent=59"
      _StyleDefs(901) =   "Splits(1).Columns(75).HeadingStyle:id=654,.parent=60"
      _StyleDefs(902) =   "Splits(1).Columns(75).FooterStyle:id=655,.parent=61"
      _StyleDefs(903) =   "Splits(1).Columns(75).EditorStyle:id=656,.parent=63"
      _StyleDefs(904) =   "Splits(1).Columns(76).Style:id=665,.parent=59"
      _StyleDefs(905) =   "Splits(1).Columns(76).HeadingStyle:id=662,.parent=60"
      _StyleDefs(906) =   "Splits(1).Columns(76).FooterStyle:id=663,.parent=61"
      _StyleDefs(907) =   "Splits(1).Columns(76).EditorStyle:id=664,.parent=63"
      _StyleDefs(908) =   "Splits(1).Columns(77).Style:id=673,.parent=59"
      _StyleDefs(909) =   "Splits(1).Columns(77).HeadingStyle:id=670,.parent=60"
      _StyleDefs(910) =   "Splits(1).Columns(77).FooterStyle:id=671,.parent=61"
      _StyleDefs(911) =   "Splits(1).Columns(77).EditorStyle:id=672,.parent=63"
      _StyleDefs(912) =   "Splits(1).Columns(78).Style:id=681,.parent=59"
      _StyleDefs(913) =   "Splits(1).Columns(78).HeadingStyle:id=678,.parent=60"
      _StyleDefs(914) =   "Splits(1).Columns(78).FooterStyle:id=679,.parent=61"
      _StyleDefs(915) =   "Splits(1).Columns(78).EditorStyle:id=680,.parent=63"
      _StyleDefs(916) =   "Splits(1).Columns(79).Style:id=689,.parent=59"
      _StyleDefs(917) =   "Splits(1).Columns(79).HeadingStyle:id=686,.parent=60"
      _StyleDefs(918) =   "Splits(1).Columns(79).FooterStyle:id=687,.parent=61"
      _StyleDefs(919) =   "Splits(1).Columns(79).EditorStyle:id=688,.parent=63"
      _StyleDefs(920) =   "Splits(1).Columns(80).Style:id=697,.parent=59"
      _StyleDefs(921) =   "Splits(1).Columns(80).HeadingStyle:id=694,.parent=60"
      _StyleDefs(922) =   "Splits(1).Columns(80).FooterStyle:id=695,.parent=61"
      _StyleDefs(923) =   "Splits(1).Columns(80).EditorStyle:id=696,.parent=63"
      _StyleDefs(924) =   "Splits(1).Columns(81).Style:id=705,.parent=59"
      _StyleDefs(925) =   "Splits(1).Columns(81).HeadingStyle:id=702,.parent=60"
      _StyleDefs(926) =   "Splits(1).Columns(81).FooterStyle:id=703,.parent=61"
      _StyleDefs(927) =   "Splits(1).Columns(81).EditorStyle:id=704,.parent=63"
      _StyleDefs(928) =   "Splits(1).Columns(82).Style:id=713,.parent=59"
      _StyleDefs(929) =   "Splits(1).Columns(82).HeadingStyle:id=710,.parent=60"
      _StyleDefs(930) =   "Splits(1).Columns(82).FooterStyle:id=711,.parent=61"
      _StyleDefs(931) =   "Splits(1).Columns(82).EditorStyle:id=712,.parent=63"
      _StyleDefs(932) =   "Splits(1).Columns(83).Style:id=721,.parent=59"
      _StyleDefs(933) =   "Splits(1).Columns(83).HeadingStyle:id=718,.parent=60"
      _StyleDefs(934) =   "Splits(1).Columns(83).FooterStyle:id=719,.parent=61"
      _StyleDefs(935) =   "Splits(1).Columns(83).EditorStyle:id=720,.parent=63"
      _StyleDefs(936) =   "Splits(1).Columns(84).Style:id=729,.parent=59"
      _StyleDefs(937) =   "Splits(1).Columns(84).HeadingStyle:id=726,.parent=60"
      _StyleDefs(938) =   "Splits(1).Columns(84).FooterStyle:id=727,.parent=61"
      _StyleDefs(939) =   "Splits(1).Columns(84).EditorStyle:id=728,.parent=63"
      _StyleDefs(940) =   "Splits(1).Columns(85).Style:id=737,.parent=59"
      _StyleDefs(941) =   "Splits(1).Columns(85).HeadingStyle:id=734,.parent=60"
      _StyleDefs(942) =   "Splits(1).Columns(85).FooterStyle:id=735,.parent=61"
      _StyleDefs(943) =   "Splits(1).Columns(85).EditorStyle:id=736,.parent=63"
      _StyleDefs(944) =   "Splits(1).Columns(86).Style:id=745,.parent=59"
      _StyleDefs(945) =   "Splits(1).Columns(86).HeadingStyle:id=742,.parent=60"
      _StyleDefs(946) =   "Splits(1).Columns(86).FooterStyle:id=743,.parent=61"
      _StyleDefs(947) =   "Splits(1).Columns(86).EditorStyle:id=744,.parent=63"
      _StyleDefs(948) =   "Splits(1).Columns(87).Style:id=753,.parent=59"
      _StyleDefs(949) =   "Splits(1).Columns(87).HeadingStyle:id=750,.parent=60"
      _StyleDefs(950) =   "Splits(1).Columns(87).FooterStyle:id=751,.parent=61"
      _StyleDefs(951) =   "Splits(1).Columns(87).EditorStyle:id=752,.parent=63"
      _StyleDefs(952) =   "Splits(1).Columns(88).Style:id=761,.parent=59"
      _StyleDefs(953) =   "Splits(1).Columns(88).HeadingStyle:id=758,.parent=60"
      _StyleDefs(954) =   "Splits(1).Columns(88).FooterStyle:id=759,.parent=61"
      _StyleDefs(955) =   "Splits(1).Columns(88).EditorStyle:id=760,.parent=63"
      _StyleDefs(956) =   "Splits(1).Columns(89).Style:id=769,.parent=59"
      _StyleDefs(957) =   "Splits(1).Columns(89).HeadingStyle:id=766,.parent=60"
      _StyleDefs(958) =   "Splits(1).Columns(89).FooterStyle:id=767,.parent=61"
      _StyleDefs(959) =   "Splits(1).Columns(89).EditorStyle:id=768,.parent=63"
      _StyleDefs(960) =   "Splits(1).Columns(90).Style:id=777,.parent=59"
      _StyleDefs(961) =   "Splits(1).Columns(90).HeadingStyle:id=774,.parent=60"
      _StyleDefs(962) =   "Splits(1).Columns(90).FooterStyle:id=775,.parent=61"
      _StyleDefs(963) =   "Splits(1).Columns(90).EditorStyle:id=776,.parent=63"
      _StyleDefs(964) =   "Splits(1).Columns(91).Style:id=785,.parent=59"
      _StyleDefs(965) =   "Splits(1).Columns(91).HeadingStyle:id=782,.parent=60"
      _StyleDefs(966) =   "Splits(1).Columns(91).FooterStyle:id=783,.parent=61"
      _StyleDefs(967) =   "Splits(1).Columns(91).EditorStyle:id=784,.parent=63"
      _StyleDefs(968) =   "Splits(1).Columns(92).Style:id=793,.parent=59"
      _StyleDefs(969) =   "Splits(1).Columns(92).HeadingStyle:id=790,.parent=60"
      _StyleDefs(970) =   "Splits(1).Columns(92).FooterStyle:id=791,.parent=61"
      _StyleDefs(971) =   "Splits(1).Columns(92).EditorStyle:id=792,.parent=63"
      _StyleDefs(972) =   "Splits(1).Columns(93).Style:id=801,.parent=59"
      _StyleDefs(973) =   "Splits(1).Columns(93).HeadingStyle:id=798,.parent=60"
      _StyleDefs(974) =   "Splits(1).Columns(93).FooterStyle:id=799,.parent=61"
      _StyleDefs(975) =   "Splits(1).Columns(93).EditorStyle:id=800,.parent=63"
      _StyleDefs(976) =   "Splits(1).Columns(94).Style:id=809,.parent=59"
      _StyleDefs(977) =   "Splits(1).Columns(94).HeadingStyle:id=806,.parent=60"
      _StyleDefs(978) =   "Splits(1).Columns(94).FooterStyle:id=807,.parent=61"
      _StyleDefs(979) =   "Splits(1).Columns(94).EditorStyle:id=808,.parent=63"
      _StyleDefs(980) =   "Splits(1).Columns(95).Style:id=817,.parent=59"
      _StyleDefs(981) =   "Splits(1).Columns(95).HeadingStyle:id=814,.parent=60"
      _StyleDefs(982) =   "Splits(1).Columns(95).FooterStyle:id=815,.parent=61"
      _StyleDefs(983) =   "Splits(1).Columns(95).EditorStyle:id=816,.parent=63"
      _StyleDefs(984) =   "Splits(1).Columns(96).Style:id=825,.parent=59"
      _StyleDefs(985) =   "Splits(1).Columns(96).HeadingStyle:id=822,.parent=60"
      _StyleDefs(986) =   "Splits(1).Columns(96).FooterStyle:id=823,.parent=61"
      _StyleDefs(987) =   "Splits(1).Columns(96).EditorStyle:id=824,.parent=63"
      _StyleDefs(988) =   "Splits(1).Columns(97).Style:id=833,.parent=59"
      _StyleDefs(989) =   "Splits(1).Columns(97).HeadingStyle:id=830,.parent=60"
      _StyleDefs(990) =   "Splits(1).Columns(97).FooterStyle:id=831,.parent=61"
      _StyleDefs(991) =   "Splits(1).Columns(97).EditorStyle:id=832,.parent=63"
      _StyleDefs(992) =   "Splits(1).Columns(98).Style:id=841,.parent=59"
      _StyleDefs(993) =   "Splits(1).Columns(98).HeadingStyle:id=838,.parent=60"
      _StyleDefs(994) =   "Splits(1).Columns(98).FooterStyle:id=839,.parent=61"
      _StyleDefs(995) =   "Splits(1).Columns(98).EditorStyle:id=840,.parent=63"
      _StyleDefs(996) =   "Splits(1).Columns(99).Style:id=849,.parent=59"
      _StyleDefs(997) =   "Splits(1).Columns(99).HeadingStyle:id=846,.parent=60"
      _StyleDefs(998) =   "Splits(1).Columns(99).FooterStyle:id=847,.parent=61"
      _StyleDefs(999) =   "Splits(1).Columns(99).EditorStyle:id=848,.parent=63"
      _StyleDefs(1000)=   "Splits(1).Columns(100).Style:id=857,.parent=59"
      _StyleDefs(1001)=   "Splits(1).Columns(100).HeadingStyle:id=854,.parent=60"
      _StyleDefs(1002)=   "Splits(1).Columns(100).FooterStyle:id=855,.parent=61"
      _StyleDefs(1003)=   "Splits(1).Columns(100).EditorStyle:id=856,.parent=63"
      _StyleDefs(1004)=   "Splits(1).Columns(101).Style:id=865,.parent=59"
      _StyleDefs(1005)=   "Splits(1).Columns(101).HeadingStyle:id=862,.parent=60"
      _StyleDefs(1006)=   "Splits(1).Columns(101).FooterStyle:id=863,.parent=61"
      _StyleDefs(1007)=   "Splits(1).Columns(101).EditorStyle:id=864,.parent=63"
      _StyleDefs(1008)=   "Splits(1).Columns(102).Style:id=873,.parent=59"
      _StyleDefs(1009)=   "Splits(1).Columns(102).HeadingStyle:id=870,.parent=60"
      _StyleDefs(1010)=   "Splits(1).Columns(102).FooterStyle:id=871,.parent=61"
      _StyleDefs(1011)=   "Splits(1).Columns(102).EditorStyle:id=872,.parent=63"
      _StyleDefs(1012)=   "Splits(1).Columns(103).Style:id=881,.parent=59"
      _StyleDefs(1013)=   "Splits(1).Columns(103).HeadingStyle:id=878,.parent=60"
      _StyleDefs(1014)=   "Splits(1).Columns(103).FooterStyle:id=879,.parent=61"
      _StyleDefs(1015)=   "Splits(1).Columns(103).EditorStyle:id=880,.parent=63"
      _StyleDefs(1016)=   "Splits(1).Columns(104).Style:id=889,.parent=59"
      _StyleDefs(1017)=   "Splits(1).Columns(104).HeadingStyle:id=886,.parent=60"
      _StyleDefs(1018)=   "Splits(1).Columns(104).FooterStyle:id=887,.parent=61"
      _StyleDefs(1019)=   "Splits(1).Columns(104).EditorStyle:id=888,.parent=63"
      _StyleDefs(1020)=   "Splits(1).Columns(105).Style:id=897,.parent=59"
      _StyleDefs(1021)=   "Splits(1).Columns(105).HeadingStyle:id=894,.parent=60"
      _StyleDefs(1022)=   "Splits(1).Columns(105).FooterStyle:id=895,.parent=61"
      _StyleDefs(1023)=   "Splits(1).Columns(105).EditorStyle:id=896,.parent=63"
      _StyleDefs(1024)=   "Splits(1).Columns(106).Style:id=905,.parent=59"
      _StyleDefs(1025)=   "Splits(1).Columns(106).HeadingStyle:id=902,.parent=60"
      _StyleDefs(1026)=   "Splits(1).Columns(106).FooterStyle:id=903,.parent=61"
      _StyleDefs(1027)=   "Splits(1).Columns(106).EditorStyle:id=904,.parent=63"
      _StyleDefs(1028)=   "Splits(1).Columns(107).Style:id=913,.parent=59"
      _StyleDefs(1029)=   "Splits(1).Columns(107).HeadingStyle:id=910,.parent=60"
      _StyleDefs(1030)=   "Splits(1).Columns(107).FooterStyle:id=911,.parent=61"
      _StyleDefs(1031)=   "Splits(1).Columns(107).EditorStyle:id=912,.parent=63"
      _StyleDefs(1032)=   "Splits(1).Columns(108).Style:id=921,.parent=59"
      _StyleDefs(1033)=   "Splits(1).Columns(108).HeadingStyle:id=918,.parent=60"
      _StyleDefs(1034)=   "Splits(1).Columns(108).FooterStyle:id=919,.parent=61"
      _StyleDefs(1035)=   "Splits(1).Columns(108).EditorStyle:id=920,.parent=63"
      _StyleDefs(1036)=   "Splits(1).Columns(109).Style:id=929,.parent=59"
      _StyleDefs(1037)=   "Splits(1).Columns(109).HeadingStyle:id=926,.parent=60"
      _StyleDefs(1038)=   "Splits(1).Columns(109).FooterStyle:id=927,.parent=61"
      _StyleDefs(1039)=   "Splits(1).Columns(109).EditorStyle:id=928,.parent=63"
      _StyleDefs(1040)=   "Splits(1).Columns(110).Style:id=937,.parent=59"
      _StyleDefs(1041)=   "Splits(1).Columns(110).HeadingStyle:id=934,.parent=60"
      _StyleDefs(1042)=   "Splits(1).Columns(110).FooterStyle:id=935,.parent=61"
      _StyleDefs(1043)=   "Splits(1).Columns(110).EditorStyle:id=936,.parent=63"
      _StyleDefs(1044)=   "Splits(1).Columns(111).Style:id=945,.parent=59"
      _StyleDefs(1045)=   "Splits(1).Columns(111).HeadingStyle:id=942,.parent=60"
      _StyleDefs(1046)=   "Splits(1).Columns(111).FooterStyle:id=943,.parent=61"
      _StyleDefs(1047)=   "Splits(1).Columns(111).EditorStyle:id=944,.parent=63"
      _StyleDefs(1048)=   "Splits(1).Columns(112).Style:id=953,.parent=59"
      _StyleDefs(1049)=   "Splits(1).Columns(112).HeadingStyle:id=950,.parent=60"
      _StyleDefs(1050)=   "Splits(1).Columns(112).FooterStyle:id=951,.parent=61"
      _StyleDefs(1051)=   "Splits(1).Columns(112).EditorStyle:id=952,.parent=63"
      _StyleDefs(1052)=   "Splits(1).Columns(113).Style:id=961,.parent=59"
      _StyleDefs(1053)=   "Splits(1).Columns(113).HeadingStyle:id=958,.parent=60"
      _StyleDefs(1054)=   "Splits(1).Columns(113).FooterStyle:id=959,.parent=61"
      _StyleDefs(1055)=   "Splits(1).Columns(113).EditorStyle:id=960,.parent=63"
      _StyleDefs(1056)=   "Splits(1).Columns(114).Style:id=969,.parent=59"
      _StyleDefs(1057)=   "Splits(1).Columns(114).HeadingStyle:id=966,.parent=60"
      _StyleDefs(1058)=   "Splits(1).Columns(114).FooterStyle:id=967,.parent=61"
      _StyleDefs(1059)=   "Splits(1).Columns(114).EditorStyle:id=968,.parent=63"
      _StyleDefs(1060)=   "Splits(1).Columns(115).Style:id=977,.parent=59"
      _StyleDefs(1061)=   "Splits(1).Columns(115).HeadingStyle:id=974,.parent=60"
      _StyleDefs(1062)=   "Splits(1).Columns(115).FooterStyle:id=975,.parent=61"
      _StyleDefs(1063)=   "Splits(1).Columns(115).EditorStyle:id=976,.parent=63"
      _StyleDefs(1064)=   "Splits(1).Columns(116).Style:id=985,.parent=59"
      _StyleDefs(1065)=   "Splits(1).Columns(116).HeadingStyle:id=982,.parent=60"
      _StyleDefs(1066)=   "Splits(1).Columns(116).FooterStyle:id=983,.parent=61"
      _StyleDefs(1067)=   "Splits(1).Columns(116).EditorStyle:id=984,.parent=63"
      _StyleDefs(1068)=   "Splits(1).Columns(117).Style:id=993,.parent=59"
      _StyleDefs(1069)=   "Splits(1).Columns(117).HeadingStyle:id=990,.parent=60"
      _StyleDefs(1070)=   "Splits(1).Columns(117).FooterStyle:id=991,.parent=61"
      _StyleDefs(1071)=   "Splits(1).Columns(117).EditorStyle:id=992,.parent=63"
      _StyleDefs(1072)=   "Splits(1).Columns(118).Style:id=1001,.parent=59"
      _StyleDefs(1073)=   "Splits(1).Columns(118).HeadingStyle:id=998,.parent=60"
      _StyleDefs(1074)=   "Splits(1).Columns(118).FooterStyle:id=999,.parent=61"
      _StyleDefs(1075)=   "Splits(1).Columns(118).EditorStyle:id=1000,.parent=63"
      _StyleDefs(1076)=   "Splits(1).Columns(119).Style:id=1009,.parent=59"
      _StyleDefs(1077)=   "Splits(1).Columns(119).HeadingStyle:id=1006,.parent=60"
      _StyleDefs(1078)=   "Splits(1).Columns(119).FooterStyle:id=1007,.parent=61"
      _StyleDefs(1079)=   "Splits(1).Columns(119).EditorStyle:id=1008,.parent=63"
      _StyleDefs(1080)=   "Splits(1).Columns(120).Style:id=1017,.parent=59"
      _StyleDefs(1081)=   "Splits(1).Columns(120).HeadingStyle:id=1014,.parent=60"
      _StyleDefs(1082)=   "Splits(1).Columns(120).FooterStyle:id=1015,.parent=61"
      _StyleDefs(1083)=   "Splits(1).Columns(120).EditorStyle:id=1016,.parent=63"
      _StyleDefs(1084)=   "Splits(1).Columns(121).Style:id=1025,.parent=59"
      _StyleDefs(1085)=   "Splits(1).Columns(121).HeadingStyle:id=1022,.parent=60"
      _StyleDefs(1086)=   "Splits(1).Columns(121).FooterStyle:id=1023,.parent=61"
      _StyleDefs(1087)=   "Splits(1).Columns(121).EditorStyle:id=1024,.parent=63"
      _StyleDefs(1088)=   "Splits(1).Columns(122).Style:id=1033,.parent=59"
      _StyleDefs(1089)=   "Splits(1).Columns(122).HeadingStyle:id=1030,.parent=60"
      _StyleDefs(1090)=   "Splits(1).Columns(122).FooterStyle:id=1031,.parent=61"
      _StyleDefs(1091)=   "Splits(1).Columns(122).EditorStyle:id=1032,.parent=63"
      _StyleDefs(1092)=   "Splits(1).Columns(123).Style:id=1041,.parent=59"
      _StyleDefs(1093)=   "Splits(1).Columns(123).HeadingStyle:id=1038,.parent=60"
      _StyleDefs(1094)=   "Splits(1).Columns(123).FooterStyle:id=1039,.parent=61"
      _StyleDefs(1095)=   "Splits(1).Columns(123).EditorStyle:id=1040,.parent=63"
      _StyleDefs(1096)=   "Splits(1).Columns(124).Style:id=1049,.parent=59"
      _StyleDefs(1097)=   "Splits(1).Columns(124).HeadingStyle:id=1046,.parent=60"
      _StyleDefs(1098)=   "Splits(1).Columns(124).FooterStyle:id=1047,.parent=61"
      _StyleDefs(1099)=   "Splits(1).Columns(124).EditorStyle:id=1048,.parent=63"
      _StyleDefs(1100)=   "Splits(1).Columns(125).Style:id=1057,.parent=59"
      _StyleDefs(1101)=   "Splits(1).Columns(125).HeadingStyle:id=1054,.parent=60"
      _StyleDefs(1102)=   "Splits(1).Columns(125).FooterStyle:id=1055,.parent=61"
      _StyleDefs(1103)=   "Splits(1).Columns(125).EditorStyle:id=1056,.parent=63"
      _StyleDefs(1104)=   "Splits(1).Columns(126).Style:id=1065,.parent=59"
      _StyleDefs(1105)=   "Splits(1).Columns(126).HeadingStyle:id=1062,.parent=60"
      _StyleDefs(1106)=   "Splits(1).Columns(126).FooterStyle:id=1063,.parent=61"
      _StyleDefs(1107)=   "Splits(1).Columns(126).EditorStyle:id=1064,.parent=63"
      _StyleDefs(1108)=   "Splits(1).Columns(127).Style:id=1073,.parent=59"
      _StyleDefs(1109)=   "Splits(1).Columns(127).HeadingStyle:id=1070,.parent=60"
      _StyleDefs(1110)=   "Splits(1).Columns(127).FooterStyle:id=1071,.parent=61"
      _StyleDefs(1111)=   "Splits(1).Columns(127).EditorStyle:id=1072,.parent=63"
      _StyleDefs(1112)=   "Splits(1).Columns(128).Style:id=1081,.parent=59"
      _StyleDefs(1113)=   "Splits(1).Columns(128).HeadingStyle:id=1078,.parent=60"
      _StyleDefs(1114)=   "Splits(1).Columns(128).FooterStyle:id=1079,.parent=61"
      _StyleDefs(1115)=   "Splits(1).Columns(128).EditorStyle:id=1080,.parent=63"
      _StyleDefs(1116)=   "Splits(1).Columns(129).Style:id=1089,.parent=59"
      _StyleDefs(1117)=   "Splits(1).Columns(129).HeadingStyle:id=1086,.parent=60"
      _StyleDefs(1118)=   "Splits(1).Columns(129).FooterStyle:id=1087,.parent=61"
      _StyleDefs(1119)=   "Splits(1).Columns(129).EditorStyle:id=1088,.parent=63"
      _StyleDefs(1120)=   "Splits(1).Columns(130).Style:id=1097,.parent=59"
      _StyleDefs(1121)=   "Splits(1).Columns(130).HeadingStyle:id=1094,.parent=60"
      _StyleDefs(1122)=   "Splits(1).Columns(130).FooterStyle:id=1095,.parent=61"
      _StyleDefs(1123)=   "Splits(1).Columns(130).EditorStyle:id=1096,.parent=63"
      _StyleDefs(1124)=   "Splits(1).Columns(131).Style:id=1105,.parent=59"
      _StyleDefs(1125)=   "Splits(1).Columns(131).HeadingStyle:id=1102,.parent=60"
      _StyleDefs(1126)=   "Splits(1).Columns(131).FooterStyle:id=1103,.parent=61"
      _StyleDefs(1127)=   "Splits(1).Columns(131).EditorStyle:id=1104,.parent=63"
      _StyleDefs(1128)=   "Splits(1).Columns(132).Style:id=1113,.parent=59"
      _StyleDefs(1129)=   "Splits(1).Columns(132).HeadingStyle:id=1110,.parent=60"
      _StyleDefs(1130)=   "Splits(1).Columns(132).FooterStyle:id=1111,.parent=61"
      _StyleDefs(1131)=   "Splits(1).Columns(132).EditorStyle:id=1112,.parent=63"
      _StyleDefs(1132)=   "Splits(1).Columns(133).Style:id=1121,.parent=59"
      _StyleDefs(1133)=   "Splits(1).Columns(133).HeadingStyle:id=1118,.parent=60"
      _StyleDefs(1134)=   "Splits(1).Columns(133).FooterStyle:id=1119,.parent=61"
      _StyleDefs(1135)=   "Splits(1).Columns(133).EditorStyle:id=1120,.parent=63"
      _StyleDefs(1136)=   "Splits(1).Columns(134).Style:id=1129,.parent=59"
      _StyleDefs(1137)=   "Splits(1).Columns(134).HeadingStyle:id=1126,.parent=60"
      _StyleDefs(1138)=   "Splits(1).Columns(134).FooterStyle:id=1127,.parent=61"
      _StyleDefs(1139)=   "Splits(1).Columns(134).EditorStyle:id=1128,.parent=63"
      _StyleDefs(1140)=   "Splits(1).Columns(135).Style:id=1142,.parent=59"
      _StyleDefs(1141)=   "Splits(1).Columns(135).HeadingStyle:id=1139,.parent=60"
      _StyleDefs(1142)=   "Splits(1).Columns(135).FooterStyle:id=1140,.parent=61"
      _StyleDefs(1143)=   "Splits(1).Columns(135).EditorStyle:id=1141,.parent=63"
      _StyleDefs(1144)=   "Splits(1).Columns(136).Style:id=1150,.parent=59"
      _StyleDefs(1145)=   "Splits(1).Columns(136).HeadingStyle:id=1147,.parent=60"
      _StyleDefs(1146)=   "Splits(1).Columns(136).FooterStyle:id=1148,.parent=61"
      _StyleDefs(1147)=   "Splits(1).Columns(136).EditorStyle:id=1149,.parent=63"
      _StyleDefs(1148)=   "Splits(1).Columns(137).Style:id=1158,.parent=59"
      _StyleDefs(1149)=   "Splits(1).Columns(137).HeadingStyle:id=1155,.parent=60"
      _StyleDefs(1150)=   "Splits(1).Columns(137).FooterStyle:id=1156,.parent=61"
      _StyleDefs(1151)=   "Splits(1).Columns(137).EditorStyle:id=1157,.parent=63"
      _StyleDefs(1152)=   "Named:id=33:Normal"
      _StyleDefs(1153)=   ":id=33,.parent=0"
      _StyleDefs(1154)=   "Named:id=34:Heading"
      _StyleDefs(1155)=   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(1156)=   ":id=34,.wraptext=-1"
      _StyleDefs(1157)=   "Named:id=35:Footing"
      _StyleDefs(1158)=   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(1159)=   "Named:id=36:Selected"
      _StyleDefs(1160)=   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(1161)=   "Named:id=37:Caption"
      _StyleDefs(1162)=   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(1163)=   "Named:id=38:HighlightRow"
      _StyleDefs(1164)=   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(1165)=   "Named:id=39:EvenRow"
      _StyleDefs(1166)=   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
      _StyleDefs(1167)=   "Named:id=40:OddRow"
      _StyleDefs(1168)=   ":id=40,.parent=33"
      _StyleDefs(1169)=   "Named:id=41:RecordSelector"
      _StyleDefs(1170)=   ":id=41,.parent=34"
      _StyleDefs(1171)=   "Named:id=42:FilterBar"
      _StyleDefs(1172)=   ":id=42,.parent=33"
      _StyleDefs(1173)=   "Named:id=131:TipeTanpaTipe"
      _StyleDefs(1174)=   ":id=131,.parent=33,.fgcolor=&HFF&"
      _StyleDefs(1175)=   "Named:id=132:TigaDigitKosong"
      _StyleDefs(1176)=   ":id=132,.parent=33,.fgcolor=&HFF00&"
      _StyleDefs(1177)=   "Named:id=87:SudahDitagih"
      _StyleDefs(1178)=   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
      _StyleDefs(1179)=   ":id=87,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(1180)=   ":id=87,.fontname=Tahoma"
      _StyleDefs(1181)=   "Named:id=1130:Putih"
      _StyleDefs(1182)=   ":id=1130,.parent=33"
      _StyleDefs(1183)=   "Named:id=1131:Merah"
      _StyleDefs(1184)=   ":id=1131,.parent=33,.bgcolor=&HFF&"
      _StyleDefs(1185)=   "Named:id=1132:Biru"
      _StyleDefs(1186)=   ":id=1132,.parent=33,.bgcolor=&HFF0000&"
      _StyleDefs(1187)=   "Named:id=1133:Hijau"
      _StyleDefs(1188)=   ":id=1133,.parent=33,.bgcolor=&HFF00&"
      _StyleDefs(1189)=   "Named:id=1134:Kuning"
      _StyleDefs(1190)=   ":id=1134,.parent=1133,.bgcolor=&HFFFF&"
   End
   Begin TDBDate6Ctl.TDBDate Tgl3 
      Height          =   330
      Left            =   2865
      TabIndex        =   109
      Top             =   370
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":307D
      Caption         =   "frmHOMappingCustomer.frx":31A9
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":3215
      Keys            =   "frmHOMappingCustomer.frx":3233
      Spin            =   "frmHOMappingCustomer.frx":3291
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBDate6Ctl.TDBDate Tgl4 
      Height          =   330
      Left            =   4890
      TabIndex        =   110
      Top             =   390
      Width           =   1455
      _Version        =   65536
      _ExtentX        =   2566
      _ExtentY        =   582
      Calendar        =   "frmHOMappingCustomer.frx":32B9
      Caption         =   "frmHOMappingCustomer.frx":33E5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":3451
      Keys            =   "frmHOMappingCustomer.frx":346F
      Spin            =   "frmHOMappingCustomer.frx":34CD
      AlignHorizontal =   0
      AlignVertical   =   0
      Appearance      =   0
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
   Begin TDBNumber6Ctl.TDBNumber txtUmurLastSV 
      Height          =   345
      Left            =   2865
      TabIndex        =   119
      Top             =   2550
      Width           =   750
      _Version        =   65536
      _ExtentX        =   1323
      _ExtentY        =   609
      Calculator      =   "frmHOMappingCustomer.frx":34F5
      Caption         =   "frmHOMappingCustomer.frx":3515
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      DropDown        =   "frmHOMappingCustomer.frx":357A
      Keys            =   "frmHOMappingCustomer.frx":3598
      Spin            =   "frmHOMappingCustomer.frx":35E2
      AlignHorizontal =   1
      AlignVertical   =   0
      Appearance      =   1
      BackColor       =   -2147483643
      BorderStyle     =   1
      BtnPositioning  =   0
      ClipMode        =   0
      ClearAction     =   0
      DecimalPoint    =   ","
      DisplayFormat   =   "####0;;Null"
      EditMode        =   0
      Enabled         =   -1
      ErrorBeep       =   0
      ForeColor       =   -2147483640
      Format          =   "####0"
      HighlightText   =   0
      MarginBottom    =   1
      MarginLeft      =   1
      MarginRight     =   1
      MarginTop       =   1
      MaxValue        =   120
      MinValue        =   0
      MousePointer    =   0
      MoveOnLRKey     =   0
      NegativeColor   =   255
      OLEDragMode     =   0
      OLEDropMode     =   0
      ReadOnly        =   0
      Separator       =   "."
      ShowContextMenu =   -1
      ValueVT         =   5
      Value           =   0
      MaxValueVT      =   5
      MinValueVT      =   5
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   270
      Left            =   11520
      TabIndex        =   121
      Top             =   45
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   11520
      TabIndex        =   122
      Top             =   420
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8790
      TabIndex        =   142
      Top             =   2640
      Width           =   3615
      Begin SSDataWidgets_B_OLEDB.SSOleDBCombo cbVOC 
         Height          =   330
         Left            =   0
         TabIndex        =   143
         Top             =   0
         Visible         =   0   'False
         Width           =   2925
         AllowInput      =   0   'False
         _Version        =   196616
         Columns(0).Width=   3200
         _ExtentX        =   5159
         _ExtentY        =   582
         _StockProps     =   93
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CR 7 :"
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
      Index           =   29
      Left            =   8115
      TabIndex        =   137
      Top             =   3120
      Width           =   540
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VOC :"
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
      Index           =   28
      Left            =   8160
      TabIndex        =   136
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display Max Display :"
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
      Index           =   27
      Left            =   6630
      TabIndex        =   134
      Top             =   2280
      Width           =   2070
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Tipe SBE :"
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
      Index           =   25
      Left            =   7275
      TabIndex        =   128
      Top             =   1200
      Width           =   1425
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Index           =   0
      Left            =   4425
      TabIndex        =   117
      Top             =   15
      Width           =   345
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Index           =   2
      Left            =   4425
      TabIndex        =   116
      Top             =   870
      Width           =   345
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Index           =   3
      Left            =   4440
      TabIndex        =   115
      Top             =   1230
      Width           =   345
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Index           =   4
      Left            =   4425
      TabIndex        =   114
      Top             =   1590
      Width           =   345
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Index           =   5
      Left            =   4425
      TabIndex        =   113
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "s/d"
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
      Index           =   1
      Left            =   4425
      TabIndex        =   112
      Top             =   390
      Width           =   345
   End
   Begin VB.Label lblPengambilanData 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PROSES PENGAMBILAN DATA"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1710
      TabIndex        =   10
      Top             =   4515
      Visible         =   0   'False
      Width           =   4020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last SBE"
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
      Index           =   1
      Left            =   1410
      TabIndex        =   111
      Top             =   420
      Width           =   825
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Model :"
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
      Index           =   24
      Left            =   7365
      TabIndex        =   106
      Top             =   1980
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Penjual :"
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
      Index           =   23
      Left            =   7140
      TabIndex        =   105
      Top             =   1605
      Width           =   1560
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Tipe SVC :"
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
      Index           =   22
      Left            =   7275
      TabIndex        =   104
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kota/Kec. SPK : "
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
      Index           =   21
      Left            =   6705
      TabIndex        =   102
      Top             =   420
      Width           =   2070
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
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
      Index           =   9
      Left            =   3675
      TabIndex        =   100
      Top             =   2610
      Width           =   540
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bulan"
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
      Index           =   8
      Left            =   3690
      TabIndex        =   99
      Top             =   2295
      Width           =   540
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<="
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
      Index           =   7
      Left            =   2370
      TabIndex        =   98
      Top             =   2625
      Width           =   270
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Umur Service Last SV :  "
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
      Index           =   20
      Left            =   45
      TabIndex        =   97
      Top             =   2640
      Width           =   2325
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<="
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
      Index           =   6
      Left            =   2385
      TabIndex        =   96
      Top             =   2310
      Width           =   270
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Umur Service Last SB :  "
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
      Index           =   19
      Left            =   60
      TabIndex        =   95
      Top             =   2325
      Width           =   2325
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Remainder :"
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
      Index           =   18
      Left            =   645
      TabIndex        =   94
      Top             =   1920
      Width           =   1635
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tahun Kendaraan :"
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
      Index           =   17
      Left            =   435
      TabIndex        =   89
      Top             =   1560
      Width           =   1830
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tanggal Retail :"
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
      Index           =   16
      Left            =   765
      TabIndex        =   88
      Top             =   1185
      Width           =   1515
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Delivery Date :"
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
      Index           =   15
      Left            =   825
      TabIndex        =   85
      Top             =   825
      Width           =   1455
   End
   Begin VB.Label LabelPasif2 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kuning"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1065
      TabIndex        =   69
      Top             =   3735
      Width           =   735
   End
   Begin VB.Label LabelPasif3 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Merah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1065
      TabIndex        =   67
      Top             =   4095
      Width           =   735
   End
   Begin VB.Label LabelPasif 
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Biru"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1065
      TabIndex        =   45
      Top             =   3315
      Width           =   735
   End
   Begin VB.Label LabelAktif 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hijau"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1065
      TabIndex        =   35
      Top             =   2955
      Width           =   735
   End
   Begin VB.Label lblJumlah 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   8520
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JUMLAH DATA = "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   8520
      Width           =   2655
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Kota/Kec. Pemakai : "
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
      Index           =   2
      Left            =   6735
      TabIndex        =   7
      Top             =   45
      Width           =   2040
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ".:: DAFTAR PELANGGAN ::."
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
      Left            =   6480
      TabIndex        =   3
      Top             =   2100
      Width           =   2760
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Last Service :"
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
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   83
      Width           =   2115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000080FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   14055
   End
End
Attribute VB_Name = "frmHOMappingCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRecord        As New ADODB.Recordset
Dim rsCari          As ADODB.Recordset
Dim rsSPK           As ADODB.Recordset
Dim rsReminder      As New ADODB.Recordset
Dim rsDealer        As New ADODB.Recordset
Dim rsDealerNon     As New ADODB.Recordset
Dim rsGridMdl       As New ADODB.Recordset

Dim cols As TrueOleDBGrid70.Columns
Private Urut As String
Private jmlTagih As Integer
Dim jumAktif As Integer
Dim jumPasif1 As Integer
Dim jumPasif2 As Integer
Dim jumPasif3 As Integer

Dim jumlahLastGr As Integer
Dim jumlanLastSBE As Integer
Dim jumlahService As Integer

Private iAlamatSPK      As Integer
Private iAlamatSTNK     As Integer

Private Sub asxPowerButton1_Click()

End Sub

Private Sub cbOperator_Change()
    If cbOperator.Text = "RANGE" Then
        FrameAktifRange.Visible = True
        FrameAktifBulan.Visible = False
    Else
        FrameAktifBulan.Visible = True
       FrameAktifRange.Visible = False
    End If
End Sub

Private Sub cbOperator_Click()
    Call cbOperator_Change
End Sub

Private Sub chCentangSemua_Click()
    If chCentangSemua.value = vbChecked Then
        For i = 0 To 25
            kolom(i).value = vbChecked
        Next
    Else
        For i = 0 To 25
            kolom(i).value = vbUnchecked
        Next
    End If
End Sub

Private Sub Check1_Click()
    If Check1.value = vbChecked Then
        If Log.LogHOConn = 0 Then
            MsgBox "Koneksi ke HO terputus", vbInformation
            Check1.value = vbUnchecked
            
        End If
    End If
End Sub

Private Sub refreshCombo()
    sSQL = "select pesan from m_voiceofcustomer with (nolock)"
    
'    AdoDealerPenjual.ConnectionString = SQLCONBengkel
'    AdoDealerPenjual.RecordSource = sSQL
'    cbVOC.DataSourceList = AdoDealerPenjual
'    cbVOC.DataFieldList = "pesan"
'    cbVOC.Columns(0).Width = 4000
'    cbVOC.Columns(0).Caption = "Voice Of Customer"
    
     
    sSQL = "select 'ALL' as FS_GRP_MDL, '1' urut "
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select distinct FS_GRP_MDL as model,'2' as urut from model"
    sSQL = sSQL & " order by urut,FS_GRP_MDL"
    AdoGroupModel.ConnectionString = SQLCONBengkel
    AdoGroupModel.RecordSource = sSQL
'    cbGroupModel.DataSourceList = AdoGroupModel
'    cbGroupModel.DataFieldList = "FS_GRP_MDL"
'    cbGroupModel.Columns(0).Width = 4000
'    cbGroupModel.Columns(0).Caption = "Group Model"
'    cbGroupModel.Columns(1).Visible = False
    
    Dim Tahun As Integer
    Dim tahunAwal As Integer
    Tahun = year(isnow)
    tahunAwal = Tahun - 50
    For i = tahunAwal To Tahun
        cbTahunAwal.AddItem i
        cbTahunAkhir.AddItem i
    Next
End Sub

Private Sub chModel_Click()
    If chModel.value = Checked Then
        frmeModel.Visible = False
    ElseIf chModel.value = Unchecked Then
        frmeModel.Visible = True
        Call isiModel
    End If
End Sub

Private Sub chNonNCM_Click()
    If chNonNCM = Unchecked Then
        frameDealerNon.Visible = True
        If rsDealerNon.State = adStateClosed Then
            Call isiDealerNon
        End If
        frameDealerNon.Visible = True
        gridDealerNon.SetFocus
    ElseIf chNonNCM = Checked Then
        frameDealerNon.Visible = False
    End If
End Sub

Private Sub ckAktif_Click()
    If ckAktif.value = vbChecked Then
        cbOperator.Enabled = True
        cmdWarna.Enabled = True
        FrameAktifBulan.Enabled = True
        FrameAktifRange.Enabled = True
    Else
        cbOperator.Enabled = False
        cmdWarna.Enabled = False
        FrameAktifBulan.Enabled = False
        FrameAktifRange.Enabled = False
    End If
End Sub

Private Sub ckDealer_Click()
    If ckDealer = Unchecked Then
        frmeDealer.Visible = True
        If rsDealer.State = adStateClosed Then
            Call isiDealer
        End If
        frmeDealer.Visible = True
        GridDealer.SetFocus
    ElseIf ckDealer = Checked Then
        frmeDealer.Visible = False
    End If
End Sub

Private Sub ckNCM_Click()
    If ckNCM = Unchecked Then
        frmeDealer.Visible = True
        If rsDealer.State = adStateClosed Then
            Call isiDealer
        End If
        frmeDealer.Visible = True
        GridDealer.SetFocus
    ElseIf ckNCM = Checked Then
        frmeDealer.Visible = False
    End If
End Sub

Private Sub ckPilihSemua_Click()
    If ckPilihSemua.value = vbChecked Then
        If Not rsRecord.EOF Or rsRecord.RecordCount > 0 Then
            rsRecord.MoveFirst
            Do While Not rsRecord.EOF
                rsRecord!Pilih = True
                rsRecord.MoveNext
            Loop
            rsRecord.MoveFirst
        End If
    Else
        If Not rsRecord.EOF Or rsRecord.RecordCount > 0 Then
            rsRecord.MoveFirst
            Do While Not rsRecord.EOF
                rsRecord!Pilih = False
                rsRecord.MoveNext
            Loop
            rsRecord.MoveFirst
        End If
    End If
End Sub

Private Sub cmdExport_Click()
1     On Error GoTo Err_Export
          Dim Er As New ErrInfo
          Dim rsClone As New ADODB.Recordset
2         Set rsClone = rsRecord.Clone
3         rsClone.filter = "pilih=true"
4         If rsClone.RecordCount = 0 Then
5             Er.HeadCap = "Export To Excel !"
6             Er.InfoCap = "List Data Mapping Service Kosong !"
7             Beep
8             Er.show (vbModal)
9             Exit Sub
10        End If
          Dim MyPath As String
          Dim onVal As String
11        MyPath = ""
          Dim f As New FormExplorer
12        f.show (vbModal)
13        MyPath = f.PathSaving
14        If MyPath = "" Then Exit Sub
15        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
16        If f.FileName = "" Then Exit Sub
17        onVal = MyPath & f.FileName & ".XLS"
          
18        Set gridTagih.DataSource = rsClone
19        gridTagih.Refresh
           
      '     gridTagih.Columns(11).Visible = False
      '     gridTagih.Columns(iAlamatSPK).Visible = False
      '     gridTagih.Columns(iAlamatSTNK).Visible = False
          gridTagih.ExportToFile onVal, True
      '     gridTagih.Columns(11).Visible = True
      '     gridTagih.Columns(iAlamatSPK).Visible = True
      '     gridTagih.Columns(iAlamatSTNK).Visible = True
20                Er.HeadCap = "Export to excel !"
21                Er.InfoCap = "Data sudah di Export ke File Excel." & Chr(13) & _
                               "File : " & onVal
                  
22                Beep
23                Er.show (vbModal)
24        Set gridTagih.DataSource = rsRecord
25        gridTagih.Refresh
          
26         Exit Sub
Err_Export:
27          If err.Number = 32755 Or err.Number = 91 Then GoTo Keluar
28          MsgBox err.Number, "Procedure : " & Me.Name & ".CmdXls_Click"
Keluar:

End Sub

Private Sub cmdExportCSV_Click()
1     On Error GoTo Err_Export
      Dim nmFile As String
      '2     CDialog.filter = "*.csv"
      '3     CDialog.FileName = LblRincian & " " & NamaCabang & " per. " & Format("01/" & Right(Dbs, 2) & "/" & Mid(Dbs, 7, 4), "mm-yyyy")
      '4     CDialog.CancelError = True
      '5     CDialog.ShowSave
      '6         nmFile = CDialog.FileName & ".csv"
      '7         GridGSO.ExportToDelimitedFile nmFile, , ";", , , True
      '8         If MsgBox("File " & nmFile & " telah di Export." + vbCrLf + "Apakah mau dibuka sekarang?", vbQuestion + vbYesNo, "Export File") = vbYes Then
      '9            Shell "Explorer " & nmFile, vbMaximizedFocus
      '10        End If
      '
          Dim Er As New ErrInfo
          Dim rsClone As New ADODB.Recordset
2         Set rsClone = rsRecord.Clone
3         rsClone.filter = "pilih=true"
4         If rsClone.RecordCount = 0 Then
5             Er.HeadCap = "Export To Excel !"
6             Er.InfoCap = "List Data Mapping Service Kosong !"
7             Beep
8             Er.show (vbModal)
9             Exit Sub
10        End If
          Dim MyPath As String
          Dim onVal As String
11        MyPath = ""
          Dim f As New FormExplorer
12        f.show (vbModal)
13        MyPath = f.PathSaving
14        If MyPath = "" Then Exit Sub
15        If Right(MyPath, 1) <> "\" Then MyPath = MyPath & "\"
16        If f.FileName = "" Then Exit Sub
17        onVal = MyPath & f.FileName & ".csv"
18        Set gridTagih.DataSource = rsClone
19        gridTagih.Refresh
      '    gridTagih.Columns(11).Visible = False
      '    gridTagih.Columns(iAlamatSPK).Visible = False
      '    gridTagih.Columns(iAlamatSTNK).Visible = False
          gridTagih.ExportToDelimitedFile onVal, , ";", , , True
      '    gridTagih.Columns(11).Visible = True
      '    gridTagih.Columns(iAlamatSPK).Visible = True
      '    gridTagih.Columns(iAlamatSTNK).Visible = True
20        If MsgBox("File " & nmFile & " telah di Export." + vbCrLf + "Apakah mau dibuka sekarang?", vbQuestion + vbYesNo, "Export File") = vbYes Then
21           Shell "Explorer " & onVal, vbMaximizedFocus
22        End If
23        Set gridTagih.DataSource = rsRecord
24        gridTagih.Refresh
25         Exit Sub
Err_Export:

End Sub

Private Sub CmdKeluar_Click()
    Unload Me
End Sub

 

Private Sub cmdRef_Click()
    lblPengambilanData.Visible = False
    'Call refreshGridBaru
    'Call refreshGrid
    Call RefreshGridMapping
    
End Sub


Private Sub cmdRetension_Click()
    Frame2.Visible = True
End Sub
 
Private Sub cmdTutupDealer_Click()
    frmeDealer.Visible = False
End Sub

Private Sub cmdTutupKolom_Click()
    frameKolom.Visible = False
End Sub

Private Sub cmdTutupModel_Click()
    frmeModel.Visible = False
End Sub

Private Sub cmdTutupNon_Click()
    frameDealerNon.Visible = False
End Sub

Private Sub Command8_Click()

1        On Error GoTo cmdRetension_Click_Error
2          Screen.MousePointer = vbHourglass
3         Set Report = New rptRepoMappingRetention
          
          Dim oRs As ADODB.Recordset
          Dim rsTempAll As ADODB.Recordset
          Dim rsTemp As ADODB.Recordset
          
          Dim konter As Integer
          Dim isTambah As Boolean
          
          Dim total1 As Integer
          Dim total2 As Integer
          Dim total3 As Integer
          Dim total4 As Integer
          Dim total5 As Integer
          Dim total6 As Integer
          Dim total7 As Integer
          Dim total8 As Integer
          Dim total9 As Integer
          Dim total10 As Integer
          
          Dim Cetak As Double
          'Call createRS
          
          Dim i As Integer
4             i = 1
5             konter = 1
6             ProgressBar2.Visible = True
7             ProgressBar2.Min = 0
8             ProgressBar2.Max = CInt(cbBulanRet.Text)
9         Do While i <= CInt(cbBulanRet.Text)
10            lblPengambilanDataRet.Visible = True
11            sSQL = "select COUNT(noinv) total,fs_rc"
12            sSQL = sSQL & " From T_INVOICE"
13            sSQL = sSQL & " where NOINV <> '' AND FS_NO_RETUR = '' "
14            sSQL = sSQL & " and FS_RC <> '' and noinv not like '%-B' and FS_RC in ('R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') "
15            sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "' and dateadd(day,-1,(DATEADD(month,1,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))"
16            sSQL = sSQL & " grOUP BY FS_RC"
17            sSQL = sSQL & " order by FS_RC"
18            Set rsTempAll = objO.FetchData(sSQL)
              
      '        sSQL = "select COUNT(noinv) total,fs_rc"
      '        sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
      '        sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
      '        sSQL = sSQL & " where NOINV <> '' and noinv not like '%-B' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
      '        sSQL = sSQL & " and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') "
      '        sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "' and dateadd(day,-1,(DATEADD(month,1,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))"
      '        sSQL = sSQL & " grOUP BY FS_RC"
      '        sSQL = sSQL & " order by FS_RC"
      '        Set rsTemp = objO.FetchData(sSQL)
              
              ''--------------------------------------------
              '' 2016-02-25 || ganti sales unit
              ''--------------------------------------------
19            If Log.RPLHistori = 0 Then
20                sSQL = "select COUNT(noinv) total,fs_rc"
21                sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
22                sSQL = sSQL & " INNER JOIN SALESUNIT s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
23                sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
24                sSQL = sSQL & " where NOINV <> '' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
25                sSQL = sSQL & " and noinv not like '%-B' and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') "
26                sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "' and dateadd(day,-1,(DATEADD(month,1,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))"
27                sSQL = sSQL & " grOUP BY FS_RC"
28                sSQL = sSQL & " order by FS_RC"
29            Else
30                sSQL = "select COUNT(noinv) total,fs_rc"
31                sSQL = sSQL & " From T_INVOICE T WITH (NOLOCK) "
32                sSQL = sSQL & " INNER JOIN [" & dbReplikasiHO & "]..SUMMARY s with (Nolock) on t.NO_CHASSIS = s.NORANGKA"
33                sSQL = sSQL & " LEFT JOIN CUSTDAT C WITH (NOLOCK) ON T.NO_CHASSIS = C.NO_CHASSIS"
34                sSQL = sSQL & " where NOINV <> '' AND C.FS_KD_DEALER = '" & Log.KodeDealer & "' AND FS_NO_RETUR = '' "
35                sSQL = sSQL & " and noinv not like '%-B' and FS_RC <> '' and FS_RC in ('R010','R020','R030','R040','R050','R060','R070','R080','R090','R100') "
36                sSQL = sSQL & " AND FS_RC <> 'R001' AND INVDATE between '" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "' and dateadd(day,-1,(DATEADD(month,1,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))"
37                sSQL = sSQL & " grOUP BY FS_RC"
38                sSQL = sSQL & " order by FS_RC"
39            End If
40            Set rsTemp = objO.FetchData(sSQL)
              ''--------------------------------------------
              
41            total1 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-1,DATEADD(MONTH,-6,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,1,dateadd(day,-1,(DATEADD(month,-5,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
42            total2 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-2,DATEADD(MONTH,-12,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,2,dateadd(day,-1,(DATEADD(month,-11,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
43            total3 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-3,DATEADD(MONTH,-18,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,3,dateadd(day,-1,(DATEADD(month,-17,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
44            total4 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-4,DATEADD(MONTH,-24,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,4,dateadd(day,-1,(DATEADD(month,-23,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
45            total5 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-5,DATEADD(MONTH,-30,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,5,dateadd(day,-1,(DATEADD(month,-29,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
46            total6 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-6,DATEADD(MONTH,-36,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,6,dateadd(day,-1,(DATEADD(month,-35,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
47            total7 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-7,DATEADD(MONTH,-42,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,7,dateadd(day,-1,(DATEADD(month,-41,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
48            total8 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-8,DATEADD(MONTH,-48,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,8,dateadd(day,-1,(DATEADD(month,-47,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
49            total9 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-9,DATEADD(MONTH,-54,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,9,dateadd(day,-1,(DATEADD(month,-53,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
50            total10 = objO.FetchDataSPK_Double("select isnull(AVG(a.total),0) from (select COUNT(typetrn) as total,MONTH(TglFaktur) as bulan,YEAR(TglFaktur) as tahun from Sihdr where TypeTrn = 'A' and TglFaktur between DATEADD(MONTH,-10,DATEADD(MONTH,-60,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')) and DATEADD(MONTH,10,dateadd(day,-1,(DATEADD(month,-59,'" & IIf(CInt(i) < 10, "0" & i, i) & "/01/" & TDBNumber1.value & "')))) group by MONTH(TglFaktur),YEAR(TglFaktur) ) a")
              
51            lblPengambilanDataRet.Visible = False
52            If Not rsTemp.EOF Then
              
53                rsTemp.filter = "fs_rc = 'R010'"
54                If Not rsTemp.EOF Then
55                    If total1 > 0 Then
56                        Cetak = Round((CInt(rsTemp("TOTAL")) / total1) * 100)
57                    Else
58                        Cetak = 0
59                    End If
60                End If
61                rsTemp.filter = ""
62                If i = 1 Then Report.txtS11.SetText Cetak & " %"
63                If i = 2 Then Report.txtS21.SetText Cetak & " %"
64                If i = 3 Then Report.txtS31.SetText Cetak & " %"
65                If i = 4 Then Report.txtS41.SetText Cetak & " %"
66                If i = 5 Then Report.txtS51.SetText Cetak & " %"
67                If i = 6 Then Report.txtS61.SetText Cetak & " %"
68                If i = 7 Then Report.txtS71.SetText Cetak & " %"
69                If i = 8 Then Report.txtS81.SetText Cetak & " %"
70                If i = 9 Then Report.txtS91.SetText Cetak & " %"
71                If i = 10 Then Report.txtS101.SetText Cetak & " %"
72                If i = 11 Then Report.txtS111.SetText Cetak & " %"
73                If i = 12 Then Report.txtS121.SetText Cetak & " %"
74                DoEvents
                  
75                rsTemp.filter = "fs_rc = 'R020'"
76                If Not rsTemp.EOF Then
77                    If total2 > 0 Then
78                        Cetak = Round((CInt(rsTemp("TOTAL")) / total2) * 100)
79                    Else
80                        Cetak = 0
81                    End If
82                End If
83                rsTemp.filter = ""
84                If i = 1 Then Report.txtS12.SetText Cetak & " %"
85                If i = 2 Then Report.txtS22.SetText Cetak & " %"
86                If i = 3 Then Report.txtS32.SetText Cetak & " %"
87                If i = 4 Then Report.txtS42.SetText Cetak & " %"
88                If i = 5 Then Report.txtS52.SetText Cetak & " %"
89                If i = 6 Then Report.txtS62.SetText Cetak & " %"
90                If i = 7 Then Report.txtS72.SetText Cetak & " %"
91                If i = 8 Then Report.txtS82.SetText Cetak & " %"
92                If i = 9 Then Report.txtS92.SetText Cetak & " %"
93                If i = 10 Then Report.txtS102.SetText Cetak & " %"
94                If i = 11 Then Report.txtS112.SetText Cetak & " %"
95                If i = 12 Then Report.txtS122.SetText Cetak & " %"
96                DoEvents
                  
97                rsTemp.filter = "fs_rc = 'R030'"
98                If Not rsTemp.EOF Then
99                    If total3 > 0 Then
100                       Cetak = Round((CInt(rsTemp("TOTAL")) / total3) * 100)
101                   Else
102                       Cetak = 0
103                   End If
104               End If
105               rsTemp.filter = ""
106               If i = 1 Then Report.txtS13.SetText Cetak & " %"
107               If i = 2 Then Report.txtS23.SetText Cetak & " %"
108               If i = 3 Then Report.txtS33.SetText Cetak & " %"
109               If i = 4 Then Report.txtS43.SetText Cetak & " %"
110               If i = 5 Then Report.txtS53.SetText Cetak & " %"
111               If i = 6 Then Report.txtS63.SetText Cetak & " %"
112               If i = 7 Then Report.txtS73.SetText Cetak & " %"
113               If i = 8 Then Report.txtS83.SetText Cetak & " %"
114               If i = 9 Then Report.txtS93.SetText Cetak & " %"
115               If i = 10 Then Report.txtS103.SetText Cetak & " %"
116               If i = 11 Then Report.txtS113.SetText Cetak & " %"
117               If i = 12 Then Report.txtS123.SetText Cetak & " %"
118               DoEvents
                  
119               rsTemp.filter = "fs_rc = 'R040'"
120               If Not rsTemp.EOF Then
121                   If total4 > 0 Then
122                       Cetak = Round((CInt(rsTemp("TOTAL")) / total4) * 100)
123                   Else
124                       Cetak = 0
125                   End If
126               End If
127               rsTemp.filter = ""
128               If i = 1 Then Report.txtS14.SetText Cetak & " %"
129               If i = 2 Then Report.txtS24.SetText Cetak & " %"
130               If i = 3 Then Report.txtS34.SetText Cetak & " %"
131               If i = 4 Then Report.txtS44.SetText Cetak & " %"
132               If i = 5 Then Report.txtS54.SetText Cetak & " %"
133               If i = 6 Then Report.txtS64.SetText Cetak & " %"
134               If i = 7 Then Report.txtS74.SetText Cetak & " %"
135               If i = 8 Then Report.txtS84.SetText Cetak & " %"
136               If i = 9 Then Report.txtS94.SetText Cetak & " %"
137               If i = 10 Then Report.txtS104.SetText Cetak & " %"
138               If i = 11 Then Report.txtS114.SetText Cetak & " %"
139               If i = 12 Then Report.txtS124.SetText Cetak & " %"
140               DoEvents
                  
                  
141               rsTemp.filter = "fs_rc = 'R050'"
142               If Not rsTemp.EOF Then
143                   If total5 > 0 Then
144                       Cetak = Round((CInt(rsTemp("TOTAL")) / total5) * 100)
145                   Else
146                       Cetak = 0
147                   End If
148               End If
149               rsTemp.filter = ""
150               If i = 1 Then Report.txtS15.SetText Cetak & " %"
151               If i = 2 Then Report.txtS25.SetText Cetak & " %"
152               If i = 3 Then Report.txtS35.SetText Cetak & " %"
153               If i = 4 Then Report.txtS45.SetText Cetak & " %"
154               If i = 5 Then Report.txtS55.SetText Cetak & " %"
155               If i = 6 Then Report.txtS65.SetText Cetak & " %"
156               If i = 7 Then Report.txtS75.SetText Cetak & " %"
157               If i = 8 Then Report.txtS85.SetText Cetak & " %"
158               If i = 9 Then Report.txtS95.SetText Cetak & " %"
159               If i = 10 Then Report.txtS105.SetText Cetak & " %"
160               If i = 11 Then Report.txtS115.SetText Cetak & " %"
161               If i = 12 Then Report.txtS125.SetText Cetak & " %"
162               DoEvents
                  
                  
163               rsTemp.filter = "fs_rc = 'R060'"
164               If Not rsTemp.EOF Then
165                   If total6 > 0 Then
166                       Cetak = Round((CInt(rsTemp("TOTAL")) / total6) * 100)
167                   Else
168                       Cetak = 0
169                   End If
170               End If
171               rsTemp.filter = ""
172               If i = 1 Then Report.txtS16.SetText Cetak & " %"
173               If i = 2 Then Report.txtS26.SetText Cetak & " %"
174               If i = 3 Then Report.txtS36.SetText Cetak & " %"
175               If i = 4 Then Report.txtS46.SetText Cetak & " %"
176               If i = 5 Then Report.txtS56.SetText Cetak & " %"
177               If i = 6 Then Report.txtS66.SetText Cetak & " %"
178               If i = 7 Then Report.txtS76.SetText Cetak & " %"
179               If i = 8 Then Report.txtS86.SetText Cetak & " %"
180               If i = 9 Then Report.txtS96.SetText Cetak & " %"
181               If i = 10 Then Report.txtS106.SetText Cetak & " %"
182               If i = 11 Then Report.txtS116.SetText Cetak & " %"
183               If i = 12 Then Report.txtS126.SetText Cetak & " %"
184               DoEvents
                  
                  
185               rsTemp.filter = "fs_rc = 'R070'"
186               If Not rsTemp.EOF Then
187                   If total7 > 0 Then
188                       Cetak = Round((CInt(rsTemp("TOTAL")) / total7) * 100)
189                   Else
190                       Cetak = 0
191                   End If
192               End If
193               rsTemp.filter = ""
194               If i = 1 Then Report.txtS17.SetText Cetak & " %"
195               If i = 2 Then Report.txtS27.SetText Cetak & " %"
196               If i = 3 Then Report.txtS37.SetText Cetak & " %"
197               If i = 4 Then Report.txtS47.SetText Cetak & " %"
198               If i = 5 Then Report.txtS57.SetText Cetak & " %"
199               If i = 6 Then Report.txtS67.SetText Cetak & " %"
200               If i = 7 Then Report.txtS77.SetText Cetak & " %"
201               If i = 8 Then Report.txtS87.SetText Cetak & " %"
202               If i = 9 Then Report.txtS97.SetText Cetak & " %"
203               If i = 10 Then Report.txtS107.SetText Cetak & " %"
204               If i = 11 Then Report.txtS117.SetText Cetak & " %"
205               If i = 12 Then Report.txtS127.SetText Cetak & " %"
206               DoEvents
                  
207               rsTemp.filter = "fs_rc = 'R080'"
208               If Not rsTemp.EOF Then
209                   If total8 > 0 Then
210                       Cetak = Round((CInt(rsTemp("TOTAL")) / total8) * 100)
211                   Else
212                       Cetak = 0
213                   End If
214               End If
215               rsTemp.filter = ""
216               If i = 1 Then Report.txtS18.SetText Cetak & " %"
217               If i = 2 Then Report.txtS28.SetText Cetak & " %"
218               If i = 3 Then Report.txtS38.SetText Cetak & " %"
219               If i = 4 Then Report.txtS48.SetText Cetak & " %"
220               If i = 5 Then Report.txtS58.SetText Cetak & " %"
221               If i = 6 Then Report.txtS68.SetText Cetak & " %"
222               If i = 7 Then Report.txtS78.SetText Cetak & " %"
223               If i = 8 Then Report.txtS88.SetText Cetak & " %"
224               If i = 9 Then Report.txtS98.SetText Cetak & " %"
225               If i = 10 Then Report.txtS108.SetText Cetak & " %"
226               If i = 11 Then Report.txtS118.SetText Cetak & " %"
227               If i = 12 Then Report.txtS128.SetText Cetak & " %"
228               DoEvents
                  
                  
229               rsTemp.filter = "fs_rc = 'R090'"
230               If Not rsTemp.EOF Then
231                   If total9 > 0 Then
232                       Cetak = Round((CInt(rsTemp("TOTAL")) / total9) * 100)
233                   Else
234                       Cetak = 0
235                   End If
236               End If
237               rsTemp.filter = ""
238               If i = 1 Then Report.txtS19.SetText Cetak & " %"
239               If i = 2 Then Report.txtS29.SetText Cetak & " %"
240               If i = 3 Then Report.txtS39.SetText Cetak & " %"
241               If i = 4 Then Report.txtS49.SetText Cetak & " %"
242               If i = 5 Then Report.txtS59.SetText Cetak & " %"
243               If i = 6 Then Report.txtS69.SetText Cetak & " %"
244               If i = 7 Then Report.txtS79.SetText Cetak & " %"
245               If i = 8 Then Report.txtS89.SetText Cetak & " %"
246               If i = 9 Then Report.txtS99.SetText Cetak & " %"
247               If i = 10 Then Report.txtS109.SetText Cetak & " %"
248               If i = 11 Then Report.txtS119.SetText Cetak & " %"
249               If i = 12 Then Report.txtS129.SetText Cetak & " %"
250               DoEvents
                  
                  
251               rsTemp.filter = "fs_rc = 'R100'"
252               If Not rsTemp.EOF Then
253                   If total10 > 0 Then
254                       Cetak = Round((CInt(rsTemp("TOTAL")) / total10) * 100)
255                   Else
256                       Cetak = 0
257                   End If
258               End If
259               rsTemp.filter = ""
260               If i = 1 Then Report.txtS110.SetText Cetak & " %"
261               If i = 2 Then Report.txtS210.SetText Cetak & " %"
262               If i = 3 Then Report.txtS310.SetText Cetak & " %"
263               If i = 4 Then Report.txtS410.SetText Cetak & " %"
264               If i = 5 Then Report.txtS510.SetText Cetak & " %"
265               If i = 6 Then Report.txtS610.SetText Cetak & " %"
266               If i = 7 Then Report.txtS710.SetText Cetak & " %"
267               If i = 8 Then Report.txtS810.SetText Cetak & " %"
268               If i = 9 Then Report.txtS910.SetText Cetak & " %"
269               If i = 10 Then Report.txtS1010.SetText Cetak & " %"
270               If i = 11 Then Report.txtS1110.SetText Cetak & " %"
271               If i = 12 Then Report.txtS1210.SetText Cetak & " %"
272               DoEvents
273           End If
              
274           If Not rsTempAll.EOF Then
              
275               rsTempAll.filter = "fs_rc = 'R010'"
276               If Not rsTempAll.EOF Then
277                   If total1 > 0 Then
278                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total1) * 100)
279                   Else
280                       Cetak = 0
281                   End If
282               End If
283               rsTempAll.filter = ""
284               If i = 1 Then Report.txtA11.SetText Cetak & " %"
285               If i = 2 Then Report.txtA21.SetText Cetak & " %"
286               If i = 3 Then Report.txtA31.SetText Cetak & " %"
287               If i = 4 Then Report.txtA41.SetText Cetak & " %"
288               If i = 5 Then Report.txtA51.SetText Cetak & " %"
289               If i = 6 Then Report.txtA61.SetText Cetak & " %"
290               If i = 7 Then Report.txtA71.SetText Cetak & " %"
291               If i = 8 Then Report.txtA81.SetText Cetak & " %"
292               If i = 9 Then Report.txtA91.SetText Cetak & " %"
293               If i = 10 Then Report.txtA101.SetText Cetak & " %"
294               If i = 11 Then Report.txtA111.SetText Cetak & " %"
295               If i = 12 Then Report.txtA121.SetText Cetak & " %"
296               DoEvents
                  
297               rsTempAll.filter = "fs_rc = 'R020'"
298               If Not rsTempAll.EOF Then
299                   If total2 > 0 Then
300                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total2) * 100)
301                   Else
302                       Cetak = 0
303                   End If
304               End If
305               rsTempAll.filter = ""
306               If i = 1 Then Report.txtA12.SetText Cetak & " %"
307               If i = 2 Then Report.txtA22.SetText Cetak & " %"
308               If i = 3 Then Report.txtA32.SetText Cetak & " %"
309               If i = 4 Then Report.txtA42.SetText Cetak & " %"
310               If i = 5 Then Report.txtA52.SetText Cetak & " %"
311               If i = 6 Then Report.txtA62.SetText Cetak & " %"
312               If i = 7 Then Report.txtA72.SetText Cetak & " %"
313               If i = 8 Then Report.txtA82.SetText Cetak & " %"
314               If i = 9 Then Report.txtA92.SetText Cetak & " %"
315               If i = 10 Then Report.txtA102.SetText Cetak & " %"
316               If i = 11 Then Report.txtA112.SetText Cetak & " %"
317               If i = 12 Then Report.txtA122.SetText Cetak & " %"
318               DoEvents
                  
319               rsTempAll.filter = "fs_rc = 'R030'"
320               If Not rsTempAll.EOF Then
321                   If total3 > 0 Then
322                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total3) * 100)
323                   Else
324                       Cetak = 0
325                   End If
326               End If
327               rsTempAll.filter = ""
328               If i = 1 Then Report.txtA13.SetText Cetak & " %"
329               If i = 2 Then Report.txtA23.SetText Cetak & " %"
330               If i = 3 Then Report.txtA33.SetText Cetak & " %"
331               If i = 4 Then Report.txtA43.SetText Cetak & " %"
332               If i = 5 Then Report.txtA53.SetText Cetak & " %"
333               If i = 6 Then Report.txtA63.SetText Cetak & " %"
334               If i = 7 Then Report.txtA73.SetText Cetak & " %"
335               If i = 8 Then Report.txtA83.SetText Cetak & " %"
336               If i = 9 Then Report.txtA93.SetText Cetak & " %"
337               If i = 10 Then Report.txtA103.SetText Cetak & " %"
338               If i = 11 Then Report.txtA113.SetText Cetak & " %"
339               If i = 12 Then Report.txtA123.SetText Cetak & " %"
340               DoEvents
                  
341               rsTempAll.filter = "fs_rc = 'R040'"
342               If Not rsTempAll.EOF Then
343                   If total4 > 0 Then
344                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total4) * 100)
345                   Else
346                       Cetak = 0
347                   End If
348               End If
349               rsTempAll.filter = ""
350               If i = 1 Then Report.txtA14.SetText Cetak & " %"
351               If i = 2 Then Report.txtA24.SetText Cetak & " %"
352               If i = 3 Then Report.txtA34.SetText Cetak & " %"
353               If i = 4 Then Report.txtA44.SetText Cetak & " %"
354               If i = 5 Then Report.txtA54.SetText Cetak & " %"
355               If i = 6 Then Report.txtA64.SetText Cetak & " %"
356               If i = 7 Then Report.txtA74.SetText Cetak & " %"
357               If i = 8 Then Report.txtA84.SetText Cetak & " %"
358               If i = 9 Then Report.txtA94.SetText Cetak & " %"
359               If i = 10 Then Report.txtA104.SetText Cetak & " %"
360               If i = 11 Then Report.txtA114.SetText Cetak & " %"
361               If i = 12 Then Report.txtA124.SetText Cetak & " %"
362               DoEvents
                  
                  
363               rsTempAll.filter = "fs_rc = 'R050'"
364               If Not rsTempAll.EOF Then
365                   If total5 > 0 Then
366                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total5) * 100)
367                   Else
368                       Cetak = 0
369                   End If
370               End If
371               rsTempAll.filter = ""
372               If i = 1 Then Report.txtA15.SetText Cetak & " %"
373               If i = 2 Then Report.txtA25.SetText Cetak & " %"
374               If i = 3 Then Report.txtA35.SetText Cetak & " %"
375               If i = 4 Then Report.txtA45.SetText Cetak & " %"
376               If i = 5 Then Report.txtA55.SetText Cetak & " %"
377               If i = 6 Then Report.txtA65.SetText Cetak & " %"
378               If i = 7 Then Report.txtA75.SetText Cetak & " %"
379               If i = 8 Then Report.txtA85.SetText Cetak & " %"
380               If i = 9 Then Report.txtA95.SetText Cetak & " %"
381               If i = 10 Then Report.txtA105.SetText Cetak & " %"
382               If i = 11 Then Report.txtA115.SetText Cetak & " %"
383               If i = 12 Then Report.txtA125.SetText Cetak & " %"
384               DoEvents
                  
                  
385               rsTempAll.filter = "fs_rc = 'R060'"
386               If Not rsTempAll.EOF Then
387                   If total6 > 0 Then
388                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total6) * 100)
389                   Else
390                       Cetak = 0
391                   End If
392               End If
393               rsTempAll.filter = ""
394               If i = 1 Then Report.txtA16.SetText Cetak & " %"
395               If i = 2 Then Report.txtA26.SetText Cetak & " %"
396               If i = 3 Then Report.txtA36.SetText Cetak & " %"
397               If i = 4 Then Report.txtA46.SetText Cetak & " %"
398               If i = 5 Then Report.txtA56.SetText Cetak & " %"
399               If i = 6 Then Report.txtA66.SetText Cetak & " %"
400               If i = 7 Then Report.txtA76.SetText Cetak & " %"
401               If i = 8 Then Report.txtA86.SetText Cetak & " %"
402               If i = 9 Then Report.txtA96.SetText Cetak & " %"
403               If i = 10 Then Report.txtA106.SetText Cetak & " %"
404               If i = 11 Then Report.txtA116.SetText Cetak & " %"
405               If i = 12 Then Report.txtA126.SetText Cetak & " %"
406               DoEvents
                  
                  
407               rsTempAll.filter = "fs_rc = 'R070'"
408               If Not rsTempAll.EOF Then
409                   If total7 > 0 Then
410                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total7) * 100)
411                   Else
412                       Cetak = 0
413                   End If
414               End If
415               rsTempAll.filter = ""
416               If i = 1 Then Report.txtA17.SetText Cetak & " %"
417               If i = 2 Then Report.txtA27.SetText Cetak & " %"
418               If i = 3 Then Report.txtA37.SetText Cetak & " %"
419               If i = 4 Then Report.txtA47.SetText Cetak & " %"
420               If i = 5 Then Report.txtA57.SetText Cetak & " %"
421               If i = 6 Then Report.txtA67.SetText Cetak & " %"
422               If i = 7 Then Report.txtA77.SetText Cetak & " %"
423               If i = 8 Then Report.txtA87.SetText Cetak & " %"
424               If i = 9 Then Report.txtA97.SetText Cetak & " %"
425               If i = 10 Then Report.txtA107.SetText Cetak & " %"
426               If i = 11 Then Report.txtA117.SetText Cetak & " %"
427               If i = 12 Then Report.txtA127.SetText Cetak & " %"
428               DoEvents
                  
429               rsTempAll.filter = "fs_rc = 'R080'"
430               If Not rsTempAll.EOF Then
431                   If total8 > 0 Then
432                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total8) * 100)
433                   Else
434                       Cetak = 0
435                   End If
436               End If
437               rsTempAll.filter = ""
438               If i = 1 Then Report.txtA18.SetText Cetak & " %"
439               If i = 2 Then Report.txtA28.SetText Cetak & " %"
440               If i = 3 Then Report.txtA38.SetText Cetak & " %"
441               If i = 4 Then Report.txtA48.SetText Cetak & " %"
442               If i = 5 Then Report.txtA58.SetText Cetak & " %"
443               If i = 6 Then Report.txtA68.SetText Cetak & " %"
444               If i = 7 Then Report.txtA78.SetText Cetak & " %"
445               If i = 8 Then Report.txtA88.SetText Cetak & " %"
446               If i = 9 Then Report.txtA98.SetText Cetak & " %"
447               If i = 10 Then Report.txtA108.SetText Cetak & " %"
448               If i = 11 Then Report.txtA118.SetText Cetak & " %"
449               If i = 12 Then Report.txtA128.SetText Cetak & " %"
450               DoEvents
                  
                  
451               rsTempAll.filter = "fs_rc = 'R090'"
452               If Not rsTempAll.EOF Then
453                   If total9 > 0 Then
454                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total9) * 100)
455                   Else
456                       Cetak = 0
457                   End If
458               End If
459               rsTempAll.filter = ""
460               If i = 1 Then Report.txtA19.SetText Cetak & " %"
461               If i = 2 Then Report.txtA29.SetText Cetak & " %"
462               If i = 3 Then Report.txtA39.SetText Cetak & " %"
463               If i = 4 Then Report.txtA49.SetText Cetak & " %"
464               If i = 5 Then Report.txtA59.SetText Cetak & " %"
465               If i = 6 Then Report.txtA69.SetText Cetak & " %"
466               If i = 7 Then Report.txtA79.SetText Cetak & " %"
467               If i = 8 Then Report.txtA89.SetText Cetak & " %"
468               If i = 9 Then Report.txtA99.SetText Cetak & " %"
469               If i = 10 Then Report.txtA109.SetText Cetak & " %"
470               If i = 11 Then Report.txtA119.SetText Cetak & " %"
471               If i = 12 Then Report.txtA129.SetText Cetak & " %"
472               DoEvents
                  
                  
473               rsTempAll.filter = "fs_rc = 'R100'"
474               If Not rsTempAll.EOF Then
475                   If total10 > 0 Then
476                       Cetak = Round((CInt(rsTempAll("TOTAL")) / total10) * 100)
477                   Else
478                       Cetak = 0
479                   End If
480               End If
481               rsTempAll.filter = ""
482               If i = 1 Then Report.txtA110.SetText Cetak & " %"
483               If i = 2 Then Report.txtA210.SetText Cetak & " %"
484               If i = 3 Then Report.txtA310.SetText Cetak & " %"
485               If i = 4 Then Report.txtA410.SetText Cetak & " %"
486               If i = 5 Then Report.txtA510.SetText Cetak & " %"
487               If i = 6 Then Report.txtA610.SetText Cetak & " %"
488               If i = 7 Then Report.txtA710.SetText Cetak & " %"
489               If i = 8 Then Report.txtA810.SetText Cetak & " %"
490               If i = 9 Then Report.txtA910.SetText Cetak & " %"
491               If i = 10 Then Report.txtA1010.SetText Cetak & " %"
492               If i = 11 Then Report.txtA1110.SetText Cetak & " %"
493               If i = 12 Then Report.txtA1210.SetText Cetak & " %"
494               DoEvents
495           End If
              
              
496           ProgressBar2.value = i
497           i = i + 1
498       Loop
499       Report.txtTahunSame.SetText TDBNumber1.value
500       Report.txtTahunAll.SetText TDBNumber1.value
          
501       With CRViewer1
502           .ReportSource = Report
503           .ViewReport
504           .Zoom 100
505           .Visible = True
506           .Left = 300
507           .Height = Me.Height - 400
508           .Width = Me.Width - .Left
509       End With
         
510               cmdTutup.Visible = True
511               cmdTutup.Left = CRViewer1.Left + 20
512               cmdTutup.Top = 435
                  
513       ProgressBar2.Visible = False
514       Frame2.Visible = False
515       Screen.MousePointer = vbDefault
          

516      On Error GoTo 0
517      Exit Sub

cmdRetension_Click_Error:

518       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdRetension_Click of Form frmHOMappingCustomer")
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub cmdWarna_Click()
    frameWarna.Visible = True
    Text1.Text = "aktif"
    Option1(4).value = vbChecked
    
End Sub

Private Sub Combo1_Change()
    If Combo1.Text = "RANGE" Then
        framePasif1Range.Visible = True
        framePasifBulan.Visible = False
    Else
        framePasifBulan.Visible = True
        framePasif1Range.Visible = False
    End If
End Sub

Private Sub Combo1_Click()
    Call Combo1_Change
End Sub

Private Sub Combo2_Change()
    If Combo2.Text = "RANGE" Then
        FramePasif2Range.Visible = True
        FramePasif2Bulan.Visible = False
    Else
        FramePasif2Bulan.Visible = True
        FramePasif2Range.Visible = False
    End If
End Sub

Private Sub Combo2_Click()
    Call Combo2_Change
End Sub

Private Sub Combo3_Change()
    If Combo3.Text = "RANGE" Then
        FramePasif3Range.Visible = True
        FramePasif3Bulan.Visible = False
    Else
        FramePasif3Bulan.Visible = True
        FramePasif3Range.Visible = False
    End If
End Sub

Private Sub Combo3_Click()
    Call Combo3_Change
End Sub

Private Sub Command1_Click()
    If Text1.Text = "aktif" Then
        If Option1(0).value Then
            LabelAktif.Caption = Option1(0).Caption
            LabelAktif.BackColor = Option1(0).BackColor
        ElseIf Option1(1).value Then
            LabelAktif.Caption = Option1(1).Caption
            LabelAktif.BackColor = Option1(1).BackColor
        ElseIf Option1(2).value Then
            LabelAktif.Caption = Option1(2).Caption
            LabelAktif.BackColor = Option1(2).BackColor
        ElseIf Option1(3).value Then
            LabelAktif.Caption = Option1(3).Caption
            LabelAktif.BackColor = Option1(3).BackColor
        ElseIf Option1(4).value Then
            LabelAktif.Caption = "Putih" 'Option1(4).Caption
            LabelAktif.BackColor = Option1(4).BackColor
        End If
    ElseIf Text1.Text = "pasif1" Then
        If Option1(0).value Then
            LabelPasif.Caption = Option1(0).Caption
            LabelPasif.BackColor = Option1(0).BackColor
        ElseIf Option1(1).value Then
            LabelPasif.Caption = Option1(1).Caption
            LabelPasif.BackColor = Option1(1).BackColor
        ElseIf Option1(2).value Then
            LabelPasif.Caption = Option1(2).Caption
            LabelPasif.BackColor = Option1(2).BackColor
        ElseIf Option1(3).value Then
            LabelPasif.Caption = Option1(3).Caption
            LabelPasif.BackColor = Option1(3).BackColor
        ElseIf Option1(4).value Then
            LabelPasif.Caption = "Putih" 'Option1(4).Caption
            LabelPasif.BackColor = Option1(4).BackColor
        End If
    ElseIf Text1.Text = "pasif2" Then
        If Option1(0).value Then
            LabelPasif2.Caption = Option1(0).Caption
            LabelPasif2.BackColor = Option1(0).BackColor
        ElseIf Option1(1).value Then
            LabelPasif2.Caption = Option1(1).Caption
            LabelPasif2.BackColor = Option1(1).BackColor
        ElseIf Option1(2).value Then
            LabelPasif2.Caption = Option1(2).Caption
            LabelPasif2.BackColor = Option1(2).BackColor
        ElseIf Option1(3).value Then
            LabelPasif2.Caption = Option1(3).Caption
            LabelPasif2.BackColor = Option1(3).BackColor
        ElseIf Option1(4).value Then
            LabelPasif2.Caption = "Putih" 'Option1(4).Caption
            LabelPasif2.BackColor = Option1(4).BackColor
        End If
    ElseIf Text1.Text = "pasif3" Then
        If Option1(0).value Then
            LabelPasif3.Caption = Option1(0).Caption
            LabelPasif3.BackColor = Option1(0).BackColor
        ElseIf Option1(1).value Then
            LabelPasif3.Caption = Option1(1).Caption
            LabelPasif3.BackColor = Option1(1).BackColor
        ElseIf Option1(2).value Then
            LabelPasif3.Caption = Option1(2).Caption
            LabelPasif3.BackColor = Option1(2).BackColor
        ElseIf Option1(3).value Then
            LabelPasif3.Caption = Option1(3).Caption
            LabelPasif3.BackColor = Option1(3).BackColor
        ElseIf Option1(4).value Then
            LabelPasif3.Caption = "Putih" 'Option1(4).Caption
            LabelPasif3.BackColor = Option1(4).BackColor
        End If
    End If
    frameWarna.Visible = False
End Sub

Private Sub Command2_Click()
    frameWarna.Visible = True
    Text1.Text = "pasif1"
    Option1(4).value = vbChecked
End Sub

Private Sub Command3_Click()
    frameWarna.Visible = True
    Text1.Text = "pasif3"
    Option1(4).value = vbChecked
End Sub

Private Sub Command4_Click()
    frameWarna.Visible = True
    Text1.Text = "pasif2"
    Option1(4).value = vbChecked
End Sub

Private Sub Command5_Click()
    Frame1.Visible = True
    Dim rsChart As New ADODB.Recordset
    If rsChart.State Then rsChart.Close
    Set rsChart = Nothing
    With rsChart
        'If .State <> adStateClosed Then .Close
        .Fields.Append "JUMLAH", adInteger
        .Open
    End With
 '   If ckAktif.value = vbChecked Then
        With rsChart
            .AddNew
            !JUMLAH = jumAktif
            .Update
        End With
   ' End If
   ' If ckPasif1.value = vbChecked Then
        With rsChart
            .AddNew
            !JUMLAH = jumPasif1
            .Update
        End With
   ' End If
   ' If ckPasif2.value = vbChecked Then
        With rsChart
            .AddNew
            !JUMLAH = jumPasif2
            .Update
        End With
   ' End If
   ' If ckPasif3.value = vbChecked Then
        With rsChart
            .AddNew
            !JUMLAH = jumPasif3
            .Update
        End With
   ' End If
    rsChart.MoveFirst
    
    MSChart1.ShowWhatsThis
    MSChart1.SeriesType = VtChSeriesType2dPie
    With MSChart1
        For i = 1 To .Plot.SeriesCollection.Count
            With .Plot.SeriesCollection(i).DataPoints(-1).DataPointLabel
                .LocationType = VtChLabelLocationTypeInside
                
                .Component = VtChLabelComponent.VtChLabelComponentValue
                .PercentFormat = "0%"
                .VtFont.Size = 10
'                .LocationType = VtChLabelLocationTypeInside
'                .Component = VtChLabelComponentValue '& " " & VtChLabelComponentPercent
                
            End With
            If i = 1 Then
                .Plot.SeriesCollection(i).LegendText = "Kategori Aktif"
            ElseIf i = 2 Then
                .Plot.SeriesCollection(i).LegendText = "Kategori Pasif1"
            ElseIf i = 3 Then
                .Plot.SeriesCollection(i).LegendText = "Kategori Pasif2"
            ElseIf i = 4 Then
                .Plot.SeriesCollection(i).LegendText = "Kategori Pasif3"
            End If
            
        Next i
    End With
    Set MSChart1.DataSource = rsChart
End Sub

Public Function getLastGRService(kriteria As String)
    
End Function

Public Function GetService(kriteria As String, Optional isHO As Boolean = False)
    rsCari.filter = ""
    rsCari.filter = "fs_rc = '" & kriteria & "'"
    If Not rsCari.EOF Then
        If isHO Then
            GetService = Format(rsCari("invdate"), "DD/MM/YYYY") '' & vbCrLf & Trim(rsCari("Fs_nm_dealer"))
        Else
            GetService = Format(rsCari("invdate"), "DD/MM/YYYY")
        End If
        
        Exit Function
    End If
    GetService = ""
End Function


Public Function GetServiceDealer(kriteria As String, Optional isHO As Boolean = False)
    rsCari.filter = ""
    rsCari.filter = "fs_rc = '" & kriteria & "'"
    If Not rsCari.EOF Then
        If isHO Then
            GetServiceDealer = Trim(rsCari("Fs_nm_dealer"))
         
        End If
        
        Exit Function
    End If
    GetServiceDealer = ""
End Function

Public Function GetServiceLast(kriteria As String, isHO As Boolean, isRep As Boolean, isLastServ As Boolean)
    rsCari.filter = ""
    rsCari.filter = kriteria
    If isHO Then
        If Not rsCari.EOF Then
            If isRep Then
                GetServiceLast = Trim(rsCari("rep"))
            ElseIf isLastServ Then
                GetServiceLast = Trim(rsCari("fs_nm_dealer"))
            Else
                GetServiceLast = Format(rsCari("invdate"), "DD/MM/YYYY") ' & vbCrLf & Trim(rsCari("fs_nm_dealer"))
            End If
            
            Exit Function
        End If
    Else
        If Not rsCari.EOF Then
            If isRep Then
                GetServiceLast = Trim(rsCari("rep"))
            ElseIf isLastServ Then
                GetServiceLast = Trim(rsCari("fs_nm_dealer"))
            Else
                GetServiceLast = Format(rsCari("invdate"), "DD/MM/YYYY")
            End If
            
            Exit Function
        End If
    End If
    
    GetServiceLast = ""
End Function

Private Sub RefreshGridMapping()
      Dim rsTagih             As ADODB.Recordset
      Dim rsDataHO            As ADODB.Recordset

      Dim varTagih            As String
      Dim kriteriaLastService As String
      Dim havingKriteria      As String
      Dim kriteria            As String
      Dim sSQL                As String
      Dim isSBE               As Boolean
      Dim konter              As Integer
      Dim daftarNOChassis     As String
      Dim warna               As String
      Dim totalAll            As Double
      Dim totalDisplay        As Double
      Dim isBoleh             As Boolean

1        On Error GoTo RefreshGridMapping_Error

2     jumAktif = 0
3     jumPasif1 = 0
4     jumPasif2 = 0
5     jumPasif3 = 0

6     isSBE = False
7     On Error GoTo vb_error
8         Screen.MousePointer = vbHourglass
9         havingKriteria = "HAVING "
10        kriteria = "WHERE c.no_chassis not in (SELECT norangka FROM recoverymapping WITH (NOLOCK) ) and "
11        kriteria = "WHERE 1=1 AND "
          'pengecekan last service berdasarkan tanggal
      '    If (IIf(IsNull(Tgl1), "", Tgl1) <> "") Then
      '        kriteria = kriteria & " (J.[DATE] BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "') AND "
      '    End If
      '    'pengecekan last sbe berdasarkan tanggal
      '    If (IIf(IsNull(Tgl3), "", Tgl3) <> "") Then
      '        isSBE = True
      '        kriteria = kriteria & " ((J.[DATE] BETWEEN '" & Format(Tgl3, "MM/DD/YYYY") & "' AND '" & Format(Tgl4, "MM/DD/YYYY") & "') AND FS_RC <> '') and "
      '    End If
          
12        If (IIf(IsNull(Tgl1), "", Tgl1) <> "") Then
13            havingKriteria = havingKriteria & " (max(J.[INVDATE]) BETWEEN '" & Format(Tgl1, "MM/DD/YYYY") & "' AND '" & Format(Tgl2, "MM/DD/YYYY") & "') AND "
14        End If
          'pengecekan last sbe berdasarkan tanggal
15        If (IIf(IsNull(Tgl3), "", Tgl3) <> "") Then
16            isSBE = True
17            havingKriteria = havingKriteria & " ((max(J.[DATE]) BETWEEN '" & Format(Tgl3, "MM/DD/YYYY") & "' AND '" & Format(Tgl4, "MM/DD/YYYY") & "')) and "
18            kriteria = kriteria & " FS_RC <> '' AND "
19        End If
          
          '' pengecekan ddate berdasarkan tanggal
20        If (IIf(IsNull(tglDDate2), "", tglDDate2) <> "") Then

21            kriteria = kriteria & " ((CONVERT(datetime,RIGHT(c.ddate,4)+SUBSTRING(c.ddate,4,2)+LEFT(c.ddate,2)) BETWEEN '" & Format(tglDDate1, "MM/DD/YYYY") & "' AND '" & Format(tglDDate2, "MM/DD/YYYY") & "')) and "
22        End If
          
          'pengecekan alamat customer
23        If Trim(txtAlamat.Text) <> "" Then
24            kriteria = kriteria & " C.CUST_ADRES + C.KOTA LIKE '%" & Trim(txtAlamat.Text) & "%' AND "
25        End If
          
26        If txtUmurLastSB.Text > 0 Then
27            kriteria = kriteria & " ((DATEDIFF(MONTH,J.DATE,GETDATE())) <= '" & txtUmurLastSB.value & "' and j.fs_rc <> '') AND "
28        End If
          
29        If txtUmurLastSV.Text > 0 Then
30            kriteria = kriteria & " ((DATEDIFF(MONTH,J.DATE,GETDATE())) <= '" & txtUmurLastSV.value & "') AND "
31        End If
          
32        If blnAkhir > 0 And blnAwal > 0 And (blnAkhir >= blnAwal) Then
33            kriteria = kriteria & " (DATEDIFF(MONTH,J.DATE,GETDATE())) BETWEEN " & blnAwal & " AND " & blnAkhir & " AND "
34        End If
35        If blnPasif > 0 Then
36            kriteria = kriteria & " (DATEDIFF(MONTH,J.DATE,GETDATE())) " & IIf(cbOperator.Text = "", "=", cbOperator.Text) & blnPasif & " AND "
37        End If
          
          
          Dim pilihDealer As String
          Dim pilihModel  As String
          
38        If ckNCM.value = vbChecked And chNonNCM.value = vbChecked Then
39            kriteria = kriteria & " (C.KDDLRPENJUAL IN (select FS_KD_DEALER from T_DEALER with (nolock) where FS_LOG_DEL = '' and FS_KD_DEALER <> 'NR 00' )) and "
40        Else
41            If ckNCM.value = vbChecked Then
42                kriteria = kriteria & " (C.KDDLRPENJUAL IN (select FS_KD_DEALER from T_DEALER with (nolock) where FS_LOG_DEL = '' and INTERNAL = 1 and FS_KD_DEALER <> 'NR 00' )) and "
43            Else
44                If rsDealer.State Then
45                    rsDealer.MoveFirst
46                    Do While Not rsDealer.EOF
47                        If rsDealer!terlibat Then
48                            pilihDealer = pilihDealer & "'" & rsDealer!KodeDealer & "',"
49                        End If
50                        rsDealer.MoveNext
51                    Loop
                      
52                End If
53            End If
              
54            If chNonNCM.value = vbChecked Then
55                kriteria = kriteria & " (C.KDDLRPENJUAL IN (select FS_KD_DEALER from T_DEALER with (nolock) where FS_LOG_DEL = '' and INTERNAL = 0 and FS_KD_DEALER <> 'NR 00' )) and "
56            Else
57                If rsDealerNon.State Then
58                    rsDealerNon.MoveFirst
59                    Do While Not rsDealerNon.EOF
60                        If rsDealerNon!terlibat Then
61                            pilihDealer = pilihDealer & "'" & rsDealerNon!KodeDealer & "',"
62                        End If
63                        rsDealerNon.MoveNext
64                    Loop
65                End If
66            End If
67            If pilihDealer <> "" Then
68                pilihDealer = Left(pilihDealer, Len(pilihDealer) - 1)
69                kriteria = kriteria & " (C.KDDLRPENJUAL IN (" & pilihDealer & ")) and "
70            End If
71        End If
          
72        If chModel.value = vbUnchecked Then
73            If rsGridMdl.State Then
74                rsGridMdl.MoveFirst
75                Do While Not rsGridMdl.EOF
76                    If rsGridMdl!ikut Then
77                        pilihModel = pilihModel & "'" & NulSID(rsGridMdl!fs_grp_mdl) & "',"
78                    End If
79                    rsGridMdl.MoveNext
80                Loop
81            End If
82            If pilihModel <> "" Then
83                pilihModel = Left(pilihModel, Len(pilihModel) - 1)
84                kriteria = kriteria & " (C.MDL IN (select MDL from model where fs_log_del = '' and fs_grp_mdl in (" & pilihModel & "))) and "
85            End If
86            sSQL = "select MDL from [" & dbReplikasiHO & "]..model where fs_log_del = '' and fs_grp_mdl in (" & pilihDealer & ")"
87        End If
          
      '    If Trim(kriteria) <> "" Then
88            kriteria = Left(kriteria, Len(kriteria) - 4)
      '        kriteria = " WHERE " & kriteria
      '    End If
89        varTagih = " and isnull(notagih,'') = ''"
90        If nTagih Then
91            varTagih = " or isnull(notagih,'') <> ''"
92        End If
          
          
93        jmlTagih = 0
          'If Check1.value = vbChecked Then
              'MENGAMBIL NO_CHASSIS LOCAL
94            sSQL = "SELECT DISTINCT C.NO_CHASSIS,C.MDL,C.FMDL,C.NO_POLISI,C.FS_TH_MOBIL,C.CUSTOMER,C.CUST_ADRES,"
95            sSQL = sSQL & " C.KOTA,C.TELPON,C.USENAME,C.USEGSM,C.FS_NM_DEALER, C.Ddate,C.CDMA,C.GSM "
96            sSQL = sSQL & " FROM CUSTDAT C WITH (NOLOCK) "
97            sSQL = sSQL & " LEFT OUTER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_CHASSIS = C.NO_CHASSIS " & kriteria
          
98        If havingKriteria = "HAVING " Then
99            havingKriteria = ""
100           sSQL1 = ""
101       Else
102           havingKriteria = Left(havingKriteria, Len(havingKriteria) - 4)
103           sSQL1 = " and j.no_chassis in (SELECT a.norangka NO_CHASSIS"
104           sSQL1 = sSQL1 & " FROM  SingleIDOri..UnitCustomer a"
105           sSQL1 = sSQL1 & " left join SingleIDOri..NewCustomer b on b.IDCust = case when a.IDCustSales <> '' then a.IDCustSales when a.idcustservice <> '' then a.idcustservice   end"
106           sSQL1 = sSQL1 & " inner JOIN T_INVOICENEW J WITH (NOLOCK) ON J.NO_CHASSIS = a.norangka and j.fs_log_del = '' and j.noinv not like '%b' "
107           sSQL1 = sSQL1 & " GROUP BY a.norangka "
108           sSQL1 = sSQL1 & havingKriteria & " ) "
109       End If
          
110       sSQL = "SELECT DISTINCT C.NO_CHASSIS,C.MDL,C.FMDL,C.NO_POLISI,C.FS_TH_MOBIL,C.CUSTOMER,C.CUST_ADRES,"
111           sSQL = sSQL & " C.KOTA,C.TELPON,C.USENAME,C.USEGSM,C.FS_NM_DEALER, C.Ddate,C.CDMA,C.GSM,b.Alamat,b.Kelurahan,b.Propinsi,b.kecamatan as kecamatansing,b.kota as kotasing,b.kodepos as kodepossing "
112           sSQL = sSQL & " FROM CUSTDAT C WITH (NOLOCK) "
113           sSQL = sSQL & " left join SingleIDOri..UnitCustomer a on a.norangka = c.no_chassis "
114           sSQL = sSQL & " left join SingleIDOri..NewCustomer b on b.IDCust = case when a.IDCustSales <> '' then a.IDCustSales when a.idcustservice <> '' then a.idcustservice   end"
115           sSQL = sSQL & " inner JOIN JOBDTA J WITH (NOLOCK) ON J.NO_CHASSIS = C.NO_CHASSIS " & kriteria
116           sSQL = sSQL & sSQL1
              
117       sSQL = "SELECT DISTINCT C.NoRangka AS NO_CHASSIS,C.MDL,C.FMDL,C.NoPolisi NO_POLISI,C.Thmobil FS_TH_MOBIL,C.NamaCust1SV + ' ' + C.NamaCust2SV + ' ' + C.NamaCust3SV CUSTOMER,C.AlamatSV CUST_ADRES,"
118           sSQL = sSQL & " C.KotaSV KOTA,C.TelpSV TELPON,C.NamaCust1STNK + ' ' + C.NamaCust2STNK + ' ' + C.NamaCust3STNK USENAME,C.TelpSV USEGSM,C.NmDlrPenjual FS_NM_DEALER, C.Ddate,C.TelpSV CDMA,C.TelpSV GSM,b.Alamat,b.Kelurahan,b.Propinsi,b.kecamatan as kecamatansing,b.kota as kotasing,b.kodepos as kodepossing "
119           sSQL = sSQL & " FROM [SingleIDOri]..UnitCustomer C WITH (NOLOCK)"
120           sSQL = sSQL & " left join SingleIDOri..NewCustomer b on b.IDCust = case when c.IDCustSales <> '' then c.IDCustSales when c.idcustservice <> '' then c.idcustservice   end"
121           sSQL = sSQL & " LEFT OUTER JOIN JOBDTA J WITH (NOLOCK) ON J.NO_CHASSIS = C.NoRangka " & kriteria
122           sSQL = sSQL & sSQL1
123       lblPengambilanData.Visible = True
124       DoEvents
125       Set rsTagih = objO.FetchDataHO(sSQL)
          
126       If rsTagih.RecordCount > 0 Then
127           Call createRS
128           lblPengambilanData.Visible = False
129           ProgressBar1.Visible = True
130           rsTagih.MoveFirst
131           ProgressBar1.Max = rsTagih.RecordCount + 1
132           ProgressBar1.Min = 0
133           Do While Not rsTagih.EOF
                  'rsDataHO.filter = ""
134               If Check1.value = vbChecked Then
                      ''@@@@ 2016-05-18 || pengambilan data dari t_invoice
                      ''Set rsCari = objO.FetchDataHOUnlimitedTime("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,fs_nm_dealer,KM,INVDATE,NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' order by INVDATE desc")
                      
                      
135                   Set rsCari = objO.FetchDataHO("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,fs_nm_dealer,KM,INVDATE,NO_CHASSIS FROM T_INVOICENEW WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B'  order by INVDATE desc")
                       
                      ''2017-09-20
136                   sSQL = "select top 1 a.Sales,"
137                   sSQL = sSQL & " case when b.Cash = 1 then 'TUNAI' else 'KREDIT' end as syarat,"
138                   sSQL = sSQL & " a.Nama,a.Alamat,'' as provinsi,'' as kecamatan,'' as kodepos,'' as hp1,'' as hp2,"
139                   sSQL = sSQL & " c.NamaFaktur,c.AlamatFaktur,'' as provinsiFaktur,c.Kota as KotaFaktur,'' as kecFaktur,"
140                   sSQL = sSQL & " c.kodepos as kodeposFaktur,c.gsmfaktur,c.cdmafaktur,b.Tanggal as tgl_retail"
141                   sSQL = sSQL & " from [SPK99]..dataspk a"
142                   sSQL = sSQL & " left join [SPK99]..Summary b on a.NoSPK = b.No_bukti and b.TipeTr = 'spk' and b.Status = '1'"
143                   sSQL = sSQL & " left join [SPK99]..DataCust c on a.NoRangka = c.norangka and Batal = 0"
144                   sSQL = sSQL & " where a.norangka = '" & rsTagih("NO_CHASSIS") & "' and a.status = '1'"
                      
145                   sSQL = "select top 1 a.Sales,"
146                   sSQL = sSQL & " case when b.Cash = 1 then 'TUNAI' else 'KREDIT' end as syarat,"
147                   sSQL = sSQL & " a.Nama,a.Alamat,f.Propinsi as provinsi,f.Kecamatan as kecamatan,f.KodePos as kodepos,f.HP1 as hp1,f.HP2 as hp2,"
148                   sSQL = sSQL & " c.NamaFaktur,c.AlamatFaktur,case when isnull(g.IDCustSTNK,'') <> '' then h.Propinsi else '' end as provinsiFaktur,c.Kota as KotaFaktur,case when isnull(g.IDCustSTNK,'') <> '' then h.Kecamatan else '' end as kecFaktur,"
149                   sSQL = sSQL & " c.kodepos as kodeposFaktur,c.gsmfaktur,c.cdmafaktur,b.Tanggal as tgl_retail"
150                   sSQL = sSQL & " from [SPK99]..dataspk a"
151                   sSQL = sSQL & " left join [SPK99]..Summary b on a.NoSPK = b.No_bukti and b.TipeTr = 'spk' and b.Status = '1'"
152                   sSQL = sSQL & " left join [SPK99]..DataCust c on a.NoRangka = c.norangka and Batal = 0"
153                   sSQL = sSQL & " left join SPK00..DoDealer d on d.KodeDealer = left(a.NoSPK,4)"
154                   sSQL = sSQL & " left join SingleIDOri..Bridging e on a.KodeLgn = e.KodeLgn and e.KodeDealer = RIGHT(d.DBName, 2)"
155                   sSQL = sSQL & " left join SingleIDOri..NewCustomer f on e.IDCust = f.IDCust"
156                   sSQL = sSQL & " left join SingleIDOri..UnitCustomer g on a.NoRangka = g.NoRangka"
157                   sSQL = sSQL & " left join SingleIDOri..NewCustomer H on h.IDCust = g.IDCustSTNK"
158                   sSQL = sSQL & " where a.norangka = '" & rsTagih("NO_CHASSIS") & "' and a.status = '1'"
                      
                      
159                   Set rsSPK = objO.FetchDataHO(sSQL)
                      
160                   sSQL = " select top 1 (select MAX(tglreminder) tglreminder from ("
161                   sSQL = sSQL & " select isnull(MAX(tglfollow),'1900-01-01') tglreminder from m_listremindercall where NoRangka = a.NO_CHASSIS"
162                   sSQL = sSQL & " Union"
163                   sSQL = sSQL & " select isnull(MAX(tglproses),'1900-01-01') from m_listreminder where NoRangka = a.NO_CHASSIS"
164                   sSQL = sSQL & " ) a  ) lastReminder , B.kecamatansv kecamatan, "
165                   sSQL = sSQL & " b.kodepossv KODEPOS,(select top 1 fs_rc from jobdta where fs_rc <> '' and no_chassis = a.no_chassis and fs_log_del = '' order by date desc) tipeLastSB,"
166                   sSQL = sSQL & " (select top 1 DATEDIFF(month,date,getdate()) from jobdta where fs_rc <> '' and no_chassis = a.no_chassis and fs_log_del = '' order by date desc) umurLastSB,"
167                   sSQL = sSQL & " a.KM,a.front,a.[DATE],a.FS_RC,DATEDIFF(month,a.date,getdate()) selisihSV,"
168                   sSQL = sSQL & " '' as voc,"
169                   sSQL = sSQL & " isnull(c.SNJ1 + '; ' + c.SNJ2 + '; ' + c.SNJ3 + '; ' + c.SNJ4 + '; ' + c.SNJ5 + '; ' + c.SNJ6 + '; ' + c.SNJ1 + '; ','') as CR7"
170                   sSQL = sSQL & " from JOBDTA a"
171                   sSQL = sSQL & " inner join singleidori..unitcustomer b on a.NO_CHASSIS = b.NOrangka"
                      '' 20180515REV02'INSPECTION
                      ''sSQL = sSQL & " left join INSPECTION c on a.NO_WO = c.NoWo"
172                   sSQL = sSQL & " left join INSPECTION c on a.NO_WO = c.NoWo and c.TglEntry = (select top 1 TglEntry from INSPECTION with(nolock) where NoWo = a.NO_WO order by TglEntry desc ) "
173                   sSQL = sSQL & "  where a.NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' and a.FS_LOG_DEL = ''"
174                   sSQL = sSQL & "  order by a.[date] desc"
175                   If rsReminder.State Then rsReminder.Close
176                   Set rsReminder = Nothing
177                   Set rsReminder = objO.FetchDataHO(sSQL)
                      
178               Else
                      ''---- Set rsCari = objO.FetchData("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,KM,INVDATE,fs_nm_dealer,NO_CHASSIS FROM JOBDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' order by INVDATE desc")
                      ''@@@@ 2016-05-18 || pengambilan data dari t_invoice
179                   Set rsCari = objO.FetchData("SELECT REPLACE(REP1 + ',' + REP2 + ',' + REP3 + ',' + REP4 + ',' + REP5 + ',' + REP6 + ',' + REP7 + ',' + REP8,',,','') REP,DATEDIFF(MONTH,INVDATE,GETDATE()) SELISIH,FS_RC,KM,INVDATE,fs_nm_dealer,NO_CHASSIS FROM T_INVOICE WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' AND FS_LOG_DEL = '' AND RIGHT(NOINV,2) <> '-B' order by INVDATE desc")
180                   sSQL = "select top 1 b.namasales, case when a.Syarat_Bayar = 0 then 'TUNAI' else 'KREDIT Via ' + d.Nama_Leasing end as syarat, a.syarat_bayar,d.Nama_Leasing, a.Nama_pemesan,a.alamat1 + ' ' + a.alamat2 as alamat,"
181                   sSQL = sSQL & " a.kota,isnull(a.propinsi,'') as propinsi,isnull(a.kecamatan,'') as kecamatan,"
182                   sSQL = sSQL & " a.kode_pos,a.hp1,a.hp2,"
183                   sSQL = sSQL & " a.nama_faktur,a.alamat1_faktur + ' ' + a.alamat2_faktur + ' ' + a.alamat3_faktur + ' ' + a.alamat4_faktur + ' ' + a.alamat5_faktur as alamat_faktur,"
184                   sSQL = sSQL & " isnull(a.propinsi,'') AS propinsi_faktur,kotamadya as kota_faktur,isnull(a.kecamatan,'') as kecamatan_faktur,a.tgl_notadebet"
185                   sSQL = sSQL & " from spk a"
186                   sSQL = sSQL & " left join arsales b on a.kode_sales = b.kodesales"
187                   sSQL = sSQL & " left join Leasing c on a.SPK_No = c.SPK_No"
188                   sSQL = sSQL & " left join LeasingMst d on c.Kode_Leasing = d.Kode_Leasing"
189                   sSQL = sSQL & " where a.norangka = '" & rsTagih("NO_CHASSIS") & "' and a.spk_status = '1'"
190                   sSQL = sSQL & " order by a.tgl_notadebet desc"
191                   Set rsSPK = objO.FetchDataSPK(sSQL)
                      
                      
                      ''2017-09-18
192                   sSQL = " select top 1 (select top 1 tglfollow from m_listremindercall where norangka= a.no_chassis order by tglentry desc) lastReminder,b.kecamatan, "
193                   sSQL = sSQL & " b.KODEPOS,(select top 1 fs_rc from jobdta where fs_rc <> '' and no_chassis = a.no_chassis and fs_log_del = '' order by date desc) tipeLastSB,"
194                   sSQL = sSQL & " (select top 1 DATEDIFF(month,date,getdate()) from jobdta where fs_rc <> '' and no_chassis = a.no_chassis and fs_log_del = '' order by date desc) umurLastSB,"
195                   sSQL = sSQL & " a.KM,a.front,a.[DATE],a.FS_RC,DATEDIFF(month,a.date,getdate()) selisihSV,"
196                   sSQL = sSQL & " (select top 1 PESAN from M_ListReminderCall where NoRangka = a.no_chassis order by tglentry desc) as voc,"
197                   sSQL = sSQL & " isnull(c.SNJ1 + '; ' + c.SNJ2 + '; ' + c.SNJ3 + '; ' + c.SNJ4 + '; ' + c.SNJ5 + '; ' + c.SNJ6 + '; ' + c.SNJ1 + '; ','') as CR7"
198                   sSQL = sSQL & " from JOBDTA a"
199                   sSQL = sSQL & " inner join CUSTDAT b on a.NO_CHASSIS = b.NO_CHASSIS"
                      '' 20180515REV02'INSPECTION
                      ''sSQL = sSQL & " left join INSPECTION c on a.NO_WO = c.NoWo"
200                   sSQL = sSQL & " left join INSPECTION c on a.NO_WO = c.NoWo and c.TglEntry = (select top 1 TglEntry from INSPECTION with(nolock) where NoWo = a.NO_WO order by TglEntry desc ) "
201                   sSQL = sSQL & "  where a.NO_CHASSIS = '" & rsTagih("NO_CHASSIS") & "' and a.FS_LOG_DEL = ''"
202                   sSQL = sSQL & "  order by a.[date] desc"
                      
203                   Set rsReminder = objO.FetchData(sSQL)
204               End If
             '     rsCari.filter
205               jumlahLastGr = 0
206               jumlanLastSBE = 0
207               jumlahService = 0
208               If Not rsCari.EOF Then
209                   jumlahService = CInt(rsCari("selisih"))
210                   rsCari.filter = "fs_rc = ''"
211                   If Not rsCari.EOF Then jumlahLastGr = CInt(rsCari("selisih"))
212                   rsCari.filter = ""
213                   rsCari.filter = "fs_rc <> ''"
214                   If Not rsCari.EOF Then jumlanLastSBE = CInt(rsCari("selisih"))
215               End If
216               rsCari.filter = ""
217               With rsRecord
218                   isBoleh = True
219                   .AddNew
220                   !NO_CHASSIS = Trim(rsTagih("NO_CHASSIS"))
221                   !MDL = Trim(rsTagih("MDL"))
222                   !Fmdl = Trim(rsTagih("FMDL"))
223                   !NO_POLISI = Trim(rsTagih("NO_POLISI"))
224                   !Tahun = Trim(rsTagih("FS_TH_MOBIL"))
225                   !CUSTOMER = Trim(rsTagih("CUSTOMER"))
      '                !cust_adres = Trim(rsTagih("CUST_ADRES"))
      '                !Kota = Trim(rsTagih("KOTA"))
                      
226                   !cust_adres = NulSID(rsTagih!Alamat) ''Trim(rsTagih("CUST_ADRES"))
227                   !kelurahan = NulSID(rsTagih!kelurahan)
228                   !kecamatan = Nul(rsTagih!kecamatansing)
229                   !KodePos = NulSID(rsTagih!kodepossing)
230                   !Propinsi = NulSID(rsTagih!Propinsi)
231                   !Kota = NulSID(rsTagih!kotasing) ''Trim(rsTagih("KOTA"))
                      
232                   !telp = Trim(rsTagih("TELPON"))
233                   !UseName = Trim(rsTagih("USENAME"))
234                   !UseGSM = Trim(rsTagih("USEGSM"))
235                   !GSM = Trim(rsTagih("GSM"))
236                   !CDMA = Trim(rsTagih("CDMA"))
237                   !dDate = Trim(rsTagih("ddate"))
238                   !DealerPenjual = IIf(IsNull(Trim((rsTagih("fs_nm_dealer")))), "", Trim(rsTagih("fs_nm_dealer"))) 'Trim(rsTagih("fs_nm_dealer"))
239                   If Check1.value = vbChecked Then
240                       rsCari.filter = ""
                          'sbi
241                       !K1 = GetService("R001", True) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R001' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
242                       !dealerK1 = GetServiceDealer("R001", True)
                          'sbe
243                       !k5 = GetService("R005", True) 'IIf(IsNull(Trim(rsTagih("5K"))), "", Trim(rsTagih("5K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R005' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
244                       !dealerK5 = GetServiceDealer("R005", True)
245                       !K10 = GetService("R010", True) 'IIf(IsNull(Trim((rsTagih("10K")))), "", Trim(rsTagih("10K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R010' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
246                       !dealerK10 = GetServiceDealer("R010", True)
247                       !k15 = GetService("R015", True) 'IIf(IsNull(Trim((rsTagih("15K")))), "", Trim(rsTagih("15K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R015' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
248                       !dealerK15 = GetServiceDealer("R015", True)
249                       !K20 = GetService("R020", True) 'IIf(IsNull(Trim((rsTagih("20K")))), "", Trim(rsTagih("20K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R020' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
250                       !dealerK20 = GetServiceDealer("R020", True)
251                       !k25 = GetService("R025", True) 'IIf(IsNull(Trim((rsTagih("25K")))), "", Trim(rsTagih("25K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R025' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
252                       !dealerK25 = GetServiceDealer("R025", True)
253                       !K30 = GetService("R030", True) 'IIf(IsNull(Trim((rsTagih("30K")))), "", Trim(rsTagih("30K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R030' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
254                       !dealerK30 = GetServiceDealer("R030", True)
255                       !k35 = GetService("R035", True) 'IIf(IsNull(Trim((rsTagih("35K")))), "", Trim(rsTagih("35K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R035' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
256                       !dealerK35 = GetServiceDealer("R035", True)
257                       !K40 = GetService("R040", True) 'IIf(IsNull(Trim((rsTagih("40K")))), "", Trim(rsTagih("40K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R040' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
258                       !dealerK40 = GetServiceDealer("R040", True)
259                       !k45 = GetService("R045", True) 'IIf(IsNull(Trim((rsTagih("45K")))), "", Trim(rsTagih("45K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R045' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
260                       !dealerK45 = GetServiceDealer("R045", True)
261                       !K50 = GetService("R050", True) 'IIf(IsNull(Trim((rsTagih("50K")))), "", Trim(rsTagih("50K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R050' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
262                       !dealerK50 = GetServiceDealer("R050", True)
263                       !k55 = GetService("R055", True) 'IIf(IsNull(Trim((rsTagih("55K")))), "", Trim(rsTagih("55K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R055' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
264                       !dealerK55 = GetServiceDealer("R055", True)
265                       !K60 = GetService("R060", True) 'IIf(IsNull(Trim((rsTagih("60K")))), "", Trim(rsTagih("60K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R060' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
266                       !dealerK60 = GetServiceDealer("R060", True)
267                       !k65 = GetService("R065", True) 'IIf(IsNull(Trim((rsTagih("65K")))), "", Trim(rsTagih("65K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R065' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
268                       !dealerK65 = GetServiceDealer("R065", True)
269                       !K70 = GetService("R070", True) 'IIf(IsNull(Trim((rsTagih("70K")))), "", Trim(rsTagih("70K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R070' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
270                       !dealerK70 = GetServiceDealer("R070", True)
271                       !k75 = GetService("R075", True) 'IIf(IsNull(Trim((rsTagih("75K")))), "", Trim(rsTagih("75K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R075' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
272                       !dealerK75 = GetServiceDealer("R075", True)
273                       !K80 = GetService("R080", True) 'IIf(IsNull(Trim((rsTagih("80K")))), "", Trim(rsTagih("80K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R080' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
274                       !dealerK80 = GetServiceDealer("R080", True)
275                       !k85 = GetService("R085", True) 'IIf(IsNull(Trim((rsTagih("85K")))), "", Trim(rsTagih("85K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R085' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
276                       !dealerK85 = GetServiceDealer("R085", True)
277                       !k90 = GetService("R090", True) 'IIf(IsNull(Trim((rsTagih("90K")))), "", Trim(rsTagih("90K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R090' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
278                       !dealerK90 = GetServiceDealer("R090", True)
279                       !k95 = GetService("R095", True) 'IIf(IsNull(Trim((rsTagih("95K")))), "", Trim(rsTagih("95K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R095' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
280                       !dealerK95 = GetServiceDealer("R095", True)
281                       !k100 = GetService("R100", True) 'IIf(IsNull(Trim((rsTagih("100K")))), "", Trim(rsTagih("100K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R100' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
282                       !dealerK100 = GetServiceDealer("R100", True)
283                       !k105 = GetService("R105", True) 'IIf(IsNull(Trim((rsTagih("105K")))), "", Trim(rsTagih("105K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R105' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
284                       !dealerK105 = GetServiceDealer("R105", True)
285                       !k110 = GetService("R110", True) 'IIf(IsNull(Trim((rsTagih("110K")))), "", Trim(rsTagih("110K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R110' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
286                       !dealerK110 = GetServiceDealer("R110", True)
287                       !k115 = GetService("R115", True) 'IIf(IsNull(Trim((rsTagih("115K")))), "", Trim(rsTagih("115K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R115' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
288                       !dealerK115 = GetServiceDealer("R115", True)
289                       !k120 = GetService("R120", True) 'IIf(IsNull(Trim((rsTagih("120K")))), "", Trim(rsTagih("120K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R120' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
290                       !dealerK120 = GetServiceDealer("R120", True)
291                       !k125 = GetService("R125", True) 'IIf(IsNull(Trim((rsTagih("125K")))), "", Trim(rsTagih("125K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R125' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
292                       !dealerK125 = GetServiceDealer("R125", True)
293                       !k130 = GetService("R130", True) 'IIf(IsNull(Trim((rsTagih("130K")))), "", Trim(rsTagih("130K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R130' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
294                       !dealerK130 = GetServiceDealer("R130", True)
295                       !k135 = GetService("R135", True) 'IIf(IsNull(Trim((rsTagih("135K")))), "", Trim(rsTagih("135K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R135' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
296                       !dealerK135 = GetServiceDealer("R135", True)
297                       !k140 = GetService("R140", True) 'IIf(IsNull(Trim((rsTagih("140K")))), "", Trim(rsTagih("140K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R140' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
298                       !dealerK140 = GetServiceDealer("R140", True)
299                       !k145 = GetService("R145", True) 'IIf(IsNull(Trim((rsTagih("145K")))), "", Trim(rsTagih("145K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R145' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
300                       !dealerK145 = GetServiceDealer("R145", True)
301                       !k150 = GetService("R150", True) 'IIf(IsNull(Trim((rsTagih("150K")))), "", Trim(rsTagih("150K"))) 'ObjO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R150' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
302                       !dealerK150 = GetServiceDealer("R150", True)
303                       !k155 = GetService("R155", True) 'IIf(IsNull(Trim((rsTagih("155K")))), "", Trim(rsTagih("155K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R155' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
304                       !dealerK155 = GetServiceDealer("R155", True)
305                       !k160 = GetService("R160", True) 'IIf(IsNull(Trim((rsTagih("160K")))), "", Trim(rsTagih("160K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R160' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
306                       !dealerK160 = GetServiceDealer("R160", True)
307                       !k165 = GetService("R165", True) 'IIf(IsNull(Trim((rsTagih("165K")))), "", Trim(rsTagih("165K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R165' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
308                       !dealerK165 = GetServiceDealer("R165", True)
309                       !k170 = GetService("R170", True) 'IIf(IsNull(Trim((rsTagih("170K")))), "", Trim(rsTagih("170K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R170' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
310                       !dealerK170 = GetServiceDealer("R170", True)
311                       !k175 = GetService("R175", True) 'IIf(IsNull(Trim((rsTagih("175K")))), "", Trim(rsTagih("175K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R175' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
312                       !dealerK175 = GetServiceDealer("R175", True)
313                       !k180 = GetService("R180", True) 'IIf(IsNull(Trim((rsTagih("180K")))), "", Trim(rsTagih("180K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R180' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
314                       !dealerK180 = GetServiceDealer("R180", True)
315                       !k185 = GetService("R185", True) 'IIf(IsNull(Trim((rsTagih("185K")))), "", Trim(rsTagih("185K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R185' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
316                       !dealerK185 = GetServiceDealer("R185", True)
317                       !k190 = GetService("R190", True) 'IIf(IsNull(Trim((rsTagih("190K")))), "", Trim(rsTagih("190K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R190' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
318                       !dealerK190 = GetServiceDealer("R190", True)
319                       !k195 = GetService("R195", True) 'IIf(IsNull(Trim((rsTagih("195K")))), "", Trim(rsTagih("195K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R195' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
320                       !dealerK195 = GetServiceDealer("R195", True)
321                       !k200 = GetService("R200", True) 'IIf(IsNull(Trim((rsTagih("200K")))), "", Trim(rsTagih("200K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R200' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
322                       !dealerK200 = GetServiceDealer("R200", True)
323                       !lastSB = GetServiceLast("fs_rc <> ''", True, False, False) 'IIf(IsNull(Trim((rsTagih("lastSB")))), "", Trim(rsTagih("lastSB"))) 'Trim(rsTagih("lastSB"))
324                       !dealerLastSB = GetServiceLast("fs_rc <> ''", True, False, True)
325                       !lastGR = GetServiceLast("fs_rc = ''", True, False, False) 'IIf(IsNull(Trim((rsTagih("lastGR")))), "", Trim(rsTagih("lastGR"))) 'Trim(rsTagih("lastGR"))
326                       !dealerLastGR = GetServiceLast("fs_rc = ''", True, False, True)
327                       !lastService = GetServiceLast("", True, False, False) 'IIf(IsNull(Trim((rsTagih("lastSVC")))), "", Trim(rsTagih("lastSVC"))) 'Trim(rsTagih("lastSVC"))
328                       !lastInfoGR = GetServiceLast("", True, True, False) 'IIf(IsNull(Trim((rsTagih("jobGR")))), "", Trim(rsTagih("jobGR"))) 'Trim(rsTagih("jobGR"))
329                       !DealerLastSVC = GetServiceLast("", True, False, True)
                          
330                       If Not rsReminder.EOF Then
331                           rsReminder.MoveFirst
332                           !lastReminder = Nul(rsReminder!lastReminder)
333                           !kecamatan = Nul(rsReminder!kecamatan)
334                           !KodePos = Nul(rsReminder!KodePos)
335                           !tipelastsb = Nul(rsReminder!tipelastsb)
336                           !umurServiceBySB = Nul(rsReminder!UMURLASTSB)
337                           !kmLast = Nul(rsReminder!KM)
338                           !saLastSV = Nul(rsReminder!FRONT)
339                           !tipeLastService = Nul(rsReminder!FS_RC)
340                           !umurLastService = Nul(rsReminder!selisihsv)
341                           !voc = Nul(rsReminder!voc)
342                           !CR7 = Nul(rsReminder!CR7)
343                           If NulSID(txtCR7.Text) <> "" Then
344                               rsReminder.filter = "CR7 Like '%" & NulSID(txtCR7.Text) & "%'"
345                               If rsReminder.EOF Then
346                                   isBoleh = False
347                               End If
348                               rsReminder.filter = ""
349                           End If
350                       End If
                          
351                       If Not rsSPK.EOF Then
352                           rsSPK.MoveFirst
353                           !retailSales = Format(Nul(rsSPK!tgl_retail), "DD/MM/YYYY")
354                           !WIRANIAGA = Nul(rsSPK!sales)
355                           !CARABAYAR = Nul(rsSPK!syarat)
356                           !NAMASPK = Nul(rsSPK!Nama)
357                           !ALAMATSPK = Nul(rsSPK!Alamat)
358                           !provinsiSPK = Nul(rsSPK!provinsi)
359                           !KOTASPK = ""  'Nul(rsSPK!Kota)
360                           !KECSPK = Nul(rsSPK!kecamatan)
361                           !KODEPOSSPK = Nul(rsSPK!KodePos)
362                           !CDMASPK = Nul(rsSPK!hp2)
363                           !HP1SPK = Nul(rsSPK!hp1)
364                           !HP2SPK = Nul(rsSPK!hp2)
365                           !namastnk = Nul(rsSPK!namafaktur)
366                           !AlamatSTNK = Nul(rsSPK!alamatfaktur)
367                           !provinsiSTNK = Nul(rsSPK!provinsifaktur)
368                           !Kotastnk = Nul(rsSPK!kotafaktur)
369                           !KECSTNK = Nul(rsSPK!kecFaktur)
370                           !KodePosSTNK = Nul(rsSPK!kodeposfaktur)
371                       End If
372                   Else
373                       rsCari.filter = ""
                          'sbi
374                       !K1 = GetService("R001")  'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R001' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK1 = GetServiceDealer("R001", True)
                          'sbe
375                       !k5 = GetService("R005") 'IIf(IsNull(Trim(rsTagih("5K"))), "", Trim(rsTagih("5K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R005' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK5 = GetServiceDealer("R005", True)
376                       !K10 = GetService("R010") 'IIf(IsNull(Trim((rsTagih("10K")))), "", Trim(rsTagih("10K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R010' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK10 = GetServiceDealer("R010", True)
377                       !k15 = GetService("R015") 'IIf(IsNull(Trim((rsTagih("15K")))), "", Trim(rsTagih("15K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R015' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK15 = GetServiceDealer("R015", True)
378                       !K20 = GetService("R020") 'IIf(IsNull(Trim((rsTagih("20K")))), "", Trim(rsTagih("20K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R020' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK20 = GetServiceDealer("R020", True)
379                       !k25 = GetService("R025") 'IIf(IsNull(Trim((rsTagih("25K")))), "", Trim(rsTagih("25K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R025' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK25 = GetServiceDealer("R025", True)
380                       !K30 = GetService("R030") 'IIf(IsNull(Trim((rsTagih("30K")))), "", Trim(rsTagih("30K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R030' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK30 = GetServiceDealer("R030", True)
381                       !k35 = GetService("R035") 'IIf(IsNull(Trim((rsTagih("35K")))), "", Trim(rsTagih("35K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R035' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK35 = GetServiceDealer("R035", True)
382                       !K40 = GetService("R040") 'IIf(IsNull(Trim((rsTagih("40K")))), "", Trim(rsTagih("40K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R040' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK40 = GetServiceDealer("R040", True)
383                       !k45 = GetService("R045") 'IIf(IsNull(Trim((rsTagih("45K")))), "", Trim(rsTagih("45K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R045' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK45 = GetServiceDealer("R045", True)
384                       !K50 = GetService("R050") 'IIf(IsNull(Trim((rsTagih("50K")))), "", Trim(rsTagih("50K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R050' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK50 = GetServiceDealer("R050", True)
385                       !k55 = GetService("R055") 'IIf(IsNull(Trim((rsTagih("55K")))), "", Trim(rsTagih("55K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R055' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK55 = GetServiceDealer("R055", True)
386                       !K60 = GetService("R060") 'IIf(IsNull(Trim((rsTagih("60K")))), "", Trim(rsTagih("60K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R060' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK60 = GetServiceDealer("R060", True)
387                       !k65 = GetService("R065") 'IIf(IsNull(Trim((rsTagih("65K")))), "", Trim(rsTagih("65K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R065' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK65 = GetServiceDealer("R065", True)
388                       !K70 = GetService("R070") 'IIf(IsNull(Trim((rsTagih("70K")))), "", Trim(rsTagih("70K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R070' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                        '  !dealerK70 = GetServiceDealer("R070", True)
389                       !k75 = GetService("R075") 'IIf(IsNull(Trim((rsTagih("75K")))), "", Trim(rsTagih("75K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R075' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK75 = GetServiceDealer("R075", True)
390                       !K80 = GetService("R080") 'IIf(IsNull(Trim((rsTagih("80K")))), "", Trim(rsTagih("80K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R080' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK80 = GetServiceDealer("R080", True)
391                       !k85 = GetService("R085") 'IIf(IsNull(Trim((rsTagih("85K")))), "", Trim(rsTagih("85K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R085' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK85 = GetServiceDealer("R085", True)
392                       !k90 = GetService("R090") 'IIf(IsNull(Trim((rsTagih("90K")))), "", Trim(rsTagih("90K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R090' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK90 = GetServiceDealer("R090", True)
393                       !k95 = GetService("R095") 'IIf(IsNull(Trim((rsTagih("95K")))), "", Trim(rsTagih("95K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R095' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK95 = GetServiceDealer("R095", True)
394                       !k100 = GetService("R100") 'IIf(IsNull(Trim((rsTagih("100K")))), "", Trim(rsTagih("100K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R100' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK100 = GetServiceDealer("R100", True)
395                       !k105 = GetService("R105") 'IIf(IsNull(Trim((rsTagih("105K")))), "", Trim(rsTagih("105K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R105' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK105 = GetServiceDealer("R105", True)
396                       !k110 = GetService("R110") 'IIf(IsNull(Trim((rsTagih("110K")))), "", Trim(rsTagih("110K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R110' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK110 = GetServiceDealer("R110", True)
397                       !k115 = GetService("R115") 'IIf(IsNull(Trim((rsTagih("115K")))), "", Trim(rsTagih("115K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R115' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK115 = GetServiceDealer("R115", True)
398                       !k120 = GetService("R120") 'IIf(IsNull(Trim((rsTagih("120K")))), "", Trim(rsTagih("120K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R120' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK120 = GetServiceDealer("R120", True)
399                       !k125 = GetService("R125") 'IIf(IsNull(Trim((rsTagih("125K")))), "", Trim(rsTagih("125K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R125' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK125 = GetServiceDealer("R125", True)
400                       !k130 = GetService("R130") 'IIf(IsNull(Trim((rsTagih("130K")))), "", Trim(rsTagih("130K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R130' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK130 = GetServiceDealer("R130", True)
401                       !k135 = GetService("R135") 'IIf(IsNull(Trim((rsTagih("135K")))), "", Trim(rsTagih("135K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R135' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK135 = GetServiceDealer("R135", True)
402                       !k140 = GetService("R140") 'IIf(IsNull(Trim((rsTagih("140K")))), "", Trim(rsTagih("140K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R140' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                          '!dealerK140 = GetServiceDealer("R140", True)
403                       !k145 = GetService("R145") 'IIf(IsNull(Trim((rsTagih("145K")))), "", Trim(rsTagih("145K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R145' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK145 = GetServiceDealer("R145", True)
404                       !k150 = GetService("R150") 'IIf(IsNull(Trim((rsTagih("150K")))), "", Trim(rsTagih("150K"))) 'ObjO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R150' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK150 = GetServiceDealer("R150", True)
405                       !k155 = GetService("R155") 'IIf(IsNull(Trim((rsTagih("155K")))), "", Trim(rsTagih("155K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R155' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK155 = GetServiceDealer("R155", True)
406                       !k160 = GetService("R160") 'IIf(IsNull(Trim((rsTagih("160K")))), "", Trim(rsTagih("160K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R160' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK160 = GetServiceDealer("R160", True)
407                       !k165 = GetService("R165") 'IIf(IsNull(Trim((rsTagih("165K")))), "", Trim(rsTagih("165K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R165' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK165 = GetServiceDealer("R165", True)
408                       !k170 = GetService("R170") 'IIf(IsNull(Trim((rsTagih("170K")))), "", Trim(rsTagih("170K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R170' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK170 = GetServiceDealer("R170", True)
409                       !k175 = GetService("R175") 'IIf(IsNull(Trim((rsTagih("175K")))), "", Trim(rsTagih("175K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R175' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK175 = GetServiceDealer("R175", True)
410                       !k180 = GetService("R180") 'IIf(IsNull(Trim((rsTagih("180K")))), "", Trim(rsTagih("180K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R180' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK180 = GetServiceDealer("R180", True)
411                       !k185 = GetService("R185") 'IIf(IsNull(Trim((rsTagih("185K")))), "", Trim(rsTagih("185K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R185' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK185 = GetServiceDealer("R185", True)
412                       !k190 = GetService("R190") 'IIf(IsNull(Trim((rsTagih("190K")))), "", Trim(rsTagih("190K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R190' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK190 = GetServiceDealer("R190", True)
413                       !k195 = GetService("R195") 'IIf(IsNull(Trim((rsTagih("195K")))), "", Trim(rsTagih("195K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R195' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                        '  !dealerK195 = GetServiceDealer("R195", True)
414                       !k200 = GetService("R200") 'IIf(IsNull(Trim((rsTagih("200K")))), "", Trim(rsTagih("200K"))) 'objO.FetchData_String("SELECT TOP 1 DATE FROM JOBDTA J WHERE FS_RC = 'R200' AND J.NO_CHASSIS = '" & Trim(rsTagih("NO_CHASSIS")) & "' ORDER BY DATE DESC")
                         ' !dealerK200 = GetServiceDealer("R200", True)
415                       !lastSB = GetServiceLast("fs_rc <> ''", False, False, False) 'IIf(IsNull(Trim((rsTagih("lastSB")))), "", Trim(rsTagih("lastSB"))) 'Trim(rsTagih("lastSB"))
416                       !dealerLastSB = GetServiceLast("fs_rc <> ''", False, False, True)
417                       !lastGR = GetServiceLast("fs_rc = ''", False, False, False) 'IIf(IsNull(Trim((rsTagih("lastGR")))), "", Trim(rsTagih("lastGR"))) 'Trim(rsTagih("lastGR"))
418                       !dealerLastGR = GetServiceLast("fs_rc = ''", False, False, True)
419                       !lastService = GetServiceLast("", False, False, False) 'IIf(IsNull(Trim((rsTagih("lastSVC")))), "", Trim(rsTagih("lastSVC"))) 'Trim(rsTagih("lastSVC"))
420                       !lastInfoGR = GetServiceLast("", False, True, False) 'IIf(IsNull(Trim((rsTagih("jobGR")))), "", Trim(rsTagih("jobGR"))) 'Trim(rsTagih("jobGR"))
421                       !DealerLastSVC = GetServiceLast("", False, False, True) 'IIf(IsNull(Trim((rsTagih("dealerlastsvc")))), "", Trim(rsTagih("dealerlastsvc")))
                          
422                       If Not rsReminder.EOF Then
423                           rsReminder.MoveFirst
424                           !lastReminder = Nul(rsReminder!lastReminder)
425                           !kecamatan = Nul(rsReminder!kecamatan)
426                           !KodePos = Nul(rsReminder!KodePos)
427                           !tipelastsb = Nul(rsReminder!tipelastsb)
428                           !umurServiceBySB = Nul(rsReminder!UMURLASTSB)
429                           !kmLast = Nul(rsReminder!KM)
430                           !saLastSV = Nul(rsReminder!FRONT)
431                           !tipeLastService = Nul(rsReminder!FS_RC)
432                           !umurLastService = Nul(rsReminder!selisihsv)
433                           !voc = Nul(rsReminder!voc)
434                           !CR7 = Nul(rsReminder!CR7)
435                           If NulSID(txtCR7.Text) <> "" Then
436                               rsReminder.filter = "CR7 Like '%" & NulSID(txtCR7.Text) & "%'"
437                               If rsReminder.EOF Then
438                                   isBoleh = False
439                               End If
440                               rsReminder.filter = ""
441                           End If
442                       End If
                          
443                       If Not rsSPK.EOF Then
444                           rsSPK.MoveFirst
445                           !retailSales = Nul(Format(Nul(rsSPK!tgl_notadebet), "DD/MM/YYYY"))
446                           !WIRANIAGA = Nul(rsSPK!namasales)
447                           !CARABAYAR = Nul(rsSPK!syarat)
448                           !NAMASPK = Nul(rsSPK!nama_pemesan)
449                           !ALAMATSPK = Nul(rsSPK!Alamat)
450                           !provinsiSPK = Nul(rsSPK!Propinsi)
451                           !KOTASPK = Nul(rsSPK!Kota)
452                           !KECSPK = Nul(rsSPK!kecamatan)
453                           !KODEPOSSPK = Nul(rsSPK!Kode_Pos)
454                           !CDMASPK = Nul(rsSPK!hp2)
455                           !HP1SPK = Nul(rsSPK!hp1)
456                           !HP2SPK = Nul(rsSPK!hp2)
457                           !namastnk = Nul(rsSPK!Nama_Faktur)
458                           !AlamatSTNK = Nul(rsSPK!alamat_faktur)
459                           !provinsiSTNK = Nul(rsSPK!propinsi_faktur)
460                           !Kotastnk = Nul(rsSPK!kota_faktur)
461                           !KECSTNK = Nul(rsSPK!kecamatan_faktur)
462                           !KodePosSTNK = Nul(rsSPK!Kode_Pos)
463                       End If
                          
464                   End If
                      
465                   !SbeNext = objO.FetchData_String("select top 1 sbenext from m_estimasi where NoRangka = '" & Trim(rsTagih("NO_CHASSIS")) & "' and nowo = '' order by (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) asc ")
466                   !tglsbenext = objO.FetchData_String("select top 1 (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) from m_estimasi where NoRangka = '" & Trim(rsTagih("NO_CHASSIS")) & "' and nowo = '' order by (CASE WHEN JT_Tgl > JT_Bln then  JT_Bln WHEN JT_Tgl < JT_Bln then  JT_Tgl WHEN JT_Tgl = JT_Bln then  Jt_Bln END) asc ")

467                   !Pilih = False
468                   !selisih = jumlahService
469                   warna = getWarnaUnit(jumlahService)
470                   !warna = warna
471                   If warna <> "" Then
472                     If UCase(warna) = UCase(LabelAktif.Caption) Then
473                        warna = "Aktif"
474                     ElseIf UCase(warna) = UCase(LabelPasif.Caption) Then
475                        warna = "Pasif1"
476                     ElseIf UCase(warna) = UCase(LabelPasif2.Caption) Then
477                        warna = "Pasif2"
478                     ElseIf UCase(warna) = UCase(LabelPasif3.Caption) Then
479                        warna = "Pasif3"
480                     Else
481                        warna = ""
482                     End If
483                   Else
484                     warna = ""
485                   End If
486                     !aktif = warna
487                   .Update
                      
488                   If isBoleh = False Then
489                       .Delete adAffectCurrent
490                   End If
491               End With
492               ProgressBar1.value = konter
493               DoEvents
494               lblJumlah.Caption = konter
495               konter = konter + 1
496               rsTagih.MoveNext
497           Loop
              
498       Else
              
499           Call createRS
500       End If
          
501       gridTagih.FetchRowStyle = True
502       lblPengambilanData.Visible = False
503       ProgressBar1.Visible = False
504       lblJumlah.Caption = rsRecord.RecordCount
505       Screen.MousePointer = vbDefault
506       Exit Sub
507       jmlTagih = 0
vb_error:
508       MsgBox "Terjadi kesalahan pengambilan data." & vbCrLf & "Error Number = " & Erl & vbCrLf & err.Description & vbCrLf & "Procedure : " & Me.Name & ".RefreshGrid"
509       lblPengambilanData.Visible = False
510       ProgressBar1.Visible = False
511       lblJumlah.Caption = rsRecord.RecordCount
512       Screen.MousePointer = vbDefault

513      On Error GoTo 0
514      Exit Sub

RefreshGridMapping_Error:

515       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure RefreshGridMapping of Form frmHOMappingCustomer")
          
End Sub

Private Sub Command9_Click()
    Frame2.Visible = False
End Sub

Private Sub Form_Load()
  '  FrameAktifRange.Visible = False
    ckAktif.value = vbChecked
    ckPasif1.value = vbChecked
    ckPasif2.value = vbChecked
    ckPasif3.value = vbChecked
    
    cbOperator.Text = cbOperator.LIST(1)
    Combo1.Text = Combo1.LIST(3)
    Combo2.Text = Combo2.LIST(3)
    Combo3.Text = Combo3.LIST(0)
    cbBulanRet.Text = Month(isnow)
    TDBNumber1.value = year(isnow)
    
    ckKolom.value = vbChecked
    Call aturKolomLevel
    
    Call refreshCombo
    
End Sub

Private Sub ckKolom_Click()
    If ckKolom = Unchecked Then
        Call aturKolomLevel
        frameKolom.Visible = True
    ElseIf ckKolom = Checked Then
        frameKolom.Visible = False
    End If
End Sub

Private Sub aturKolomLevel()
    For i = 0 To 25
        kolom(i).Enabled = True
        kolom(i).value = vbChecked
        kolom(i).Visible = True
    Next
End Sub
Private Sub createRS()
    If rsRecord.State Then rsRecord.Close
    Set rsRecord = Nothing
    With rsRecord
        'If .State <> adStateClosed Then .Close
        .Fields.Append "PILIH", adBoolean
        .Fields.Append "NO_CHASSIS", adVarChar, 100
        .Fields.Append "MDL", adVarChar, 100
        .Fields.Append "FMDL", adVarChar, 100
        .Fields.Append "NO_POLISI", adVarChar, 40
        .Fields.Append "TAHUN", adVarChar, 10
        .Fields.Append "CUSTOMER", adVarChar, 200
        .Fields.Append "CUST_ADRES", adVarChar, 300
        .Fields.Append "KOTA", adVarChar, 100
        .Fields.Append "TELP", adVarChar, 100
        .Fields.Append "USENAME", adVarChar, 100
        .Fields.Append "USEGSM", adVarChar, 100
        ''## 2016-06-28 || CDMA
        .Fields.Append "CDMA", adVarChar, 100
        .Fields.Append "GSM", adVarChar, 100
        ''## 2016-05-26 || DELIVERY DATE --- ##
        .Fields.Append "DDate", adVarChar, 100
        .Fields.Append "k1", adVarChar, 100
        .Fields.Append "k5", adVarChar, 100
        .Fields.Append "k10", adVarChar, 100
        .Fields.Append "k15", adVarChar, 100
        .Fields.Append "k20", adVarChar, 100
        .Fields.Append "k25", adVarChar, 100
        .Fields.Append "k30", adVarChar, 100
        .Fields.Append "k35", adVarChar, 100
        .Fields.Append "k40", adVarChar, 100
        .Fields.Append "k45", adVarChar, 100
        .Fields.Append "k50", adVarChar, 100
        .Fields.Append "k55", adVarChar, 100
        .Fields.Append "k60", adVarChar, 100
        .Fields.Append "k65", adVarChar, 100
        .Fields.Append "k70", adVarChar, 100
        .Fields.Append "k75", adVarChar, 100
        .Fields.Append "k80", adVarChar, 100
        .Fields.Append "k85", adVarChar, 100
        .Fields.Append "k90", adVarChar, 100
        .Fields.Append "k95", adVarChar, 100
        .Fields.Append "k100", adVarChar, 100
        .Fields.Append "k105", adVarChar, 100
        .Fields.Append "k110", adVarChar, 100
        .Fields.Append "k115", adVarChar, 100
        .Fields.Append "k120", adVarChar, 100
        .Fields.Append "k125", adVarChar, 100
        .Fields.Append "k130", adVarChar, 100
        .Fields.Append "k135", adVarChar, 100
        .Fields.Append "k140", adVarChar, 100
        .Fields.Append "k145", adVarChar, 100
        .Fields.Append "k150", adVarChar, 100
        .Fields.Append "k155", adVarChar, 100
        .Fields.Append "k160", adVarChar, 100
        .Fields.Append "k165", adVarChar, 100
        .Fields.Append "k170", adVarChar, 100
        .Fields.Append "k175", adVarChar, 100
        .Fields.Append "k180", adVarChar, 100
        .Fields.Append "k185", adVarChar, 100
        .Fields.Append "k190", adVarChar, 100
        .Fields.Append "k195", adVarChar, 100
        .Fields.Append "k200", adVarChar, 100
        
        If Check1.value = vbChecked Then
            .Fields.Append "dealerK1", adVarChar, 100
            .Fields.Append "dealerK5", adVarChar, 100
            .Fields.Append "dealerK10", adVarChar, 100
            .Fields.Append "dealerK15", adVarChar, 100
            .Fields.Append "dealerK20", adVarChar, 100
            .Fields.Append "dealerK25", adVarChar, 100
            .Fields.Append "dealerK30", adVarChar, 100
            .Fields.Append "dealerK35", adVarChar, 100
            .Fields.Append "dealerK40", adVarChar, 100
            .Fields.Append "dealerK45", adVarChar, 100
            .Fields.Append "dealerK50", adVarChar, 100
            .Fields.Append "dealerK55", adVarChar, 100
            .Fields.Append "dealerK60", adVarChar, 100
            .Fields.Append "dealerK65", adVarChar, 100
            .Fields.Append "dealerK70", adVarChar, 100
            .Fields.Append "dealerK75", adVarChar, 100
            .Fields.Append "dealerK80", adVarChar, 100
            .Fields.Append "dealerK85", adVarChar, 100
            .Fields.Append "dealerK90", adVarChar, 100
            .Fields.Append "dealerK95", adVarChar, 100
            .Fields.Append "dealerK100", adVarChar, 100
            .Fields.Append "dealerK105", adVarChar, 100
            .Fields.Append "dealerK110", adVarChar, 100
            .Fields.Append "dealerK115", adVarChar, 100
            .Fields.Append "dealerK120", adVarChar, 100
            .Fields.Append "dealerK125", adVarChar, 100
            .Fields.Append "dealerK130", adVarChar, 100
            .Fields.Append "dealerK135", adVarChar, 100
            .Fields.Append "dealerK140", adVarChar, 100
            .Fields.Append "dealerK145", adVarChar, 100
            .Fields.Append "dealerK150", adVarChar, 100
            .Fields.Append "dealerK155", adVarChar, 100
            .Fields.Append "dealerK160", adVarChar, 100
            .Fields.Append "dealerK165", adVarChar, 100
            .Fields.Append "dealerK170", adVarChar, 100
            .Fields.Append "dealerK175", adVarChar, 100
            .Fields.Append "dealerK180", adVarChar, 100
            .Fields.Append "dealerK185", adVarChar, 100
            .Fields.Append "dealerK190", adVarChar, 100
            .Fields.Append "dealerK195", adVarChar, 100
            .Fields.Append "dealerK200", adVarChar, 100
        End If
        .Fields.Append "lastSB", adVarChar, 100
        .Fields.Append "dealerLastSB", adVarChar, 100
        .Fields.Append "lastGR", adVarChar, 100
        .Fields.Append "dealerLastGR", adVarChar, 100
        .Fields.Append "lastInfoGR", adVarChar, 350
        .Fields.Append "lastService", adVarChar, 100
        .Fields.Append "DealerPenjual", adVarChar, 100
        .Fields.Append "DealerLastSVC", adVarChar, 100
        .Fields.Append "Selisih", adVarChar, 100
        .Fields.Append "Warna", adVarChar, 50
        .Fields.Append "aktif", adVarChar, 50
        
        .Fields.Append "lastReminder", adVarChar, 50
        .Fields.Append "retailSales", adVarChar, 50
        
        .Fields.Append "kelurahan", adVarChar, 50
        .Fields.Append "kecamatan", adVarChar, 50
        '.Fields.Append "kota", adVarChar, 50
        .Fields.Append "propinsi", adVarChar, 50
        .Fields.Append "kodepos", adVarChar, 50
        .Fields.Append "tipeLastSB", adVarChar, 50
        .Fields.Append "umurServiceBySB", adVarChar, 50
        .Fields.Append "kmLast", adVarChar, 50
        .Fields.Append "saLastSV", adVarChar, 50
        .Fields.Append "tipeLastService", adVarChar, 50
        .Fields.Append "umurLastService", adVarChar, 50
        .Fields.Append "wiraniaga", adVarChar, 50
        .Fields.Append "caraBayar", adVarChar, 50
        .Fields.Append "namaSPK", adVarChar, 150
        .Fields.Append "alamatSPK", adVarChar, 250
        .Fields.Append "provinsiSPK", adVarChar, 150
        .Fields.Append "kotaSPK", adVarChar, 150
        .Fields.Append "kecSPK", adVarChar, 150
        .Fields.Append "kodeposSPK", adVarChar, 50
        .Fields.Append "cdmaSPK", adVarChar, 50
        .Fields.Append "hp1SPK", adVarChar, 50
        .Fields.Append "hp2SPK", adVarChar, 50
        .Fields.Append "namaSTNK", adVarChar, 150
        .Fields.Append "alamatSTNK", adVarChar, 250
        .Fields.Append "provinsiSTNK", adVarChar, 150
        .Fields.Append "kotaSTNK", adVarChar, 150
        .Fields.Append "kecSTNK", adVarChar, 150
        .Fields.Append "kodePosSTNK", adVarChar, 150
        .Fields.Append "voc", adVarChar, 250
        .Fields.Append "cr7", adVarChar, 250
        .Fields.Append "sbenext", adVarChar, 250
        .Fields.Append "tglsbenext", adVarChar, 250
        .Open
    End With
    Set gridTagih.DataSource = rsRecord
    Call headColumns
    
End Sub

Private Sub headColumns()
1         iAlamatSPK = 11
2         iAlamatSTNK = 11
3         With gridTagih
              Dim i As Integer
              Dim kolomI As Integer
              
4              i = 0
5              Do While i <= 104
6                  .Splits(1).Columns(i).Visible = True
7                  i = i + 1
8              Loop
               
9             .Columns(0).Width = 500
10            .Columns(0).Caption = "Pilih"
11            .Columns(0).DataField = "pilih"
12            .Columns(0).Alignment = dbgCenter
      '        .Splits(0).Columns(0).Locked = False
      '        .Splits(1).Columns(0).Locked = False
              
13            .Columns(1).Width = 2000
14            .Columns(1).Caption = "Last Reminder"
15            .Columns(1).DataField = "lastReminder"
16            .Splits(1).Columns(1).Locked = False
              
17            If kolom(0).value = vbUnchecked Then
18                .Columns(1).Visible = False
19            End If
              
20            .Columns(2).Width = 2000
21            .Columns(2).Caption = "Last SVC"
22            .Columns(2).DataField = "lastService"
23            .Splits(1).Columns(2).Locked = False
              
24            If kolom(1).value = vbUnchecked Then
25                .Columns(2).Visible = False
26            End If
              
27            .Columns(3).Width = 2000
28            .Columns(3).Caption = "NAMA CUSTOMER"
29            .Columns(3).DataField = "CUSTOMER"
30            .Splits(0).Columns(3).Locked = True
              
31            If kolom(2).value = vbUnchecked Then
32                .Columns(3).Visible = False
33            End If
              
34            .Columns(4).Width = 2500
35            .Columns(4).Caption = "NAMA MODEL"
36            .Columns(4).DataField = "FMDL"
37            .Splits(0).Columns(4).Locked = True
              
38            If kolom(3).value = vbUnchecked Then
39                .Columns(4).Visible = False
40            End If
              
41            .Columns(5).Width = 1500
42            .Columns(5).Caption = "NO POLISI"
43            .Columns(5).DataField = "NO_POLISI"
44            .Splits(0).Columns(5).Locked = True
              
45            If kolom(4).value = vbUnchecked Then
46                .Columns(5).Visible = False
47            End If
              
48            .Columns(6).Width = 2000
49            .Columns(6).Caption = "No. CHASSIS"
50            .Columns(6).DataField = "NO_CHASSIS"
51            .Splits(0).Columns(6).Locked = True
              
52            If kolom(5).value = vbUnchecked Then
53                .Columns(6).Visible = False
54            End If
              
55            .Columns(7).Width = 1000
56            .Columns(7).Caption = "TAHUN"
57            .Columns(7).DataField = "TAHUN"
58            .Splits(0).Columns(7).Locked = True
              
59            If kolom(6).value = vbUnchecked Then
60                .Columns(7).Visible = False
61            End If
              
              
62            .Columns(8).Width = 2000
63            .Columns(8).Caption = "DELIVERY DATE"
64            .Columns(8).DataField = "ddate"
65            .Splits(0).Columns(8).Locked = True
              
66            If kolom(7).value = vbUnchecked Then
67                .Columns(8).Visible = False
68            End If
              
69            .Columns(9).Width = 2000
70            .Columns(9).Caption = "Retail Sales"
71            .Columns(9).DataField = "retailSales"
72            .Splits(1).Columns(9).Locked = False

73            If kolom(8).value = vbUnchecked Then
74                .Columns(9).Visible = False
75            End If
              
76            .Columns(10).Width = 1500
77            .Columns(10).Caption = "MODEL"
78            .Columns(10).DataField = "MDL"
79            .Splits(0).Columns(10).Locked = True
       
80            If kolom(9).value = vbUnchecked Then
81                .Columns(10).Visible = False
82            End If
              
83            .Columns(11).Width = 4000
84            .Columns(11).Caption = "ALAMAT CUSTOMER"
85            .Columns(11).DataField = "CUST_ADRES"
86            .Splits(0).Columns(11).Locked = True
              
              
87            .Columns(12).Width = 2000
88            .Columns(12).Caption = "Kelurahan"
89            .Columns(12).DataField = "Kelurahan"
90            .Splits(1).Columns(12).Locked = False
              
91            .Columns(13).Width = 2000
92            .Columns(13).Caption = "Kecamatan"
93            .Columns(13).DataField = "kecamatan"
94            .Splits(1).Columns(13).Locked = False
              
95            .Columns(14).Width = 2000
96            .Columns(14).Caption = "KOTA"
97            .Columns(14).DataField = "KOTA"
98            .Splits(0).Columns(14).Locked = True
              
              
99            .Columns(15).Width = 2000
100           .Columns(15).Caption = "Propinsi"
101           .Columns(15).DataField = "Propinsi"
102           .Splits(0).Columns(15).Locked = True
              
103           .Columns(16).Width = 2000
104           .Columns(16).Caption = "KodePos"
105           .Columns(16).DataField = "kodepos"
106           .Splits(1).Columns(16).Locked = False
               
107           If kolom(10).value = vbUnchecked Then
108               .Columns(11).Visible = False
109               .Columns(12).Visible = False
110               .Columns(13).Visible = False
111               .Columns(14).Visible = False
112               .Columns(15).Visible = False
113               .Columns(16).Visible = False
114           End If
              
115           .Columns(17).Width = 1500
116           .Columns(17).Caption = "TELP"
117           .Columns(17).DataField = "TELP"
118           .Splits(0).Columns(17).Locked = True
              
              
119           .Columns(18).Width = 2000
120           .Columns(18).Caption = "GSM"
121           .Columns(18).DataField = "GSM"
122           .Splits(0).Columns(18).Locked = True
              
123           .Columns(19).Width = 2000
124           .Columns(19).Caption = "CDMA"
125           .Columns(19).DataField = "CDMA"
126           .Splits(0).Columns(19).Locked = True
              
127           If kolom(11).value = vbUnchecked Then
128               .Columns(17).Visible = False
129               .Columns(18).Visible = False
130               .Columns(19).Visible = False
131           End If
               
               
132           .Columns(20).Width = 2000
133           .Columns(20).Caption = "USENAME"
134           .Columns(20).DataField = "USENAME"
135           .Splits(0).Columns(20).Locked = True
               
136           .Columns(21).Width = 2000
137           .Columns(21).Caption = "USEGSM"
138           .Columns(21).DataField = "USEGSM"
139           .Splits(0).Columns(21).Locked = True
              
140           If kolom(12).value = vbUnchecked Then
141               .Columns(20).Visible = False
142               .Columns(21).Visible = False
143           End If
              
144           .Columns(22).Width = 2000
145           .Columns(22).Caption = "SBI 1K"
146           .Columns(22).DataField = "k1"
147           .Splits(1).Columns(22).Locked = False
              
148           If Check1.value = vbChecked Then
149               .Columns(23).Width = 2000
150               .Columns(23).Caption = "SBI 1K"
151               .Columns(23).DataField = "dealerK1"
152               .Splits(1).Columns(23).Locked = False
                  
153           End If
              'sbe 5k s/d 200k
              
              
              ''### 2016-05-26 || jika centang ho maka tambah kolom dealer
              
154           i = 5
              
155           If Check1.value = vbChecked Then
156               kolomI = 24
157           Else
158               kolomI = 23
159           End If
              
160           If kolom(13).value = vbUnchecked Then
161               .Columns(kolomI - 1).Visible = False
162               .Columns(22).Visible = False
163           End If
              
164           Do While i <= 200
                  
165               .Columns(kolomI).Width = 2000
166               .Columns(kolomI).Caption = "SBE " & i & "K"
167               .Columns(kolomI).DataField = "k" & i
168               .Splits(1).Columns(kolomI).Locked = False
169               If kolom(13).value = vbUnchecked Then
170                   .Columns(kolomI).Visible = False
171               End If
172               If Check1.value = vbChecked Then
173                  kolomI = kolomI + 1
174                  .Columns(kolomI).Width = 2000
175                  .Columns(kolomI).Caption = "Dealer SBE " & i & "K"
176                  .Columns(kolomI).DataField = "dealerK" & i
177                  .Splits(1).Columns(kolomI).Locked = False
178                  If kolom(13).value = vbUnchecked Then
179                       .Columns(kolomI).Visible = False
180                   End If
181               End If
182               kolomI = kolomI + 1
183               i = i + 5
184           Loop
              
185           i = 0
              
              'If Check1.value = vbChecked Then i = 43
                

186           .Columns(kolomI).Width = 2000
187           .Columns(kolomI).Caption = "TIPE LAST SB"
188           .Columns(kolomI).DataField = "tipeLastSB"
189           .Splits(1).Columns(kolomI).Locked = False
              
190           .Columns(kolomI + 1).Width = 2000
191           .Columns(kolomI + 1).Caption = "LAST SVC SB"
192           .Columns(kolomI + 1).DataField = "lastSB"
193           .Splits(1).Columns(kolomI + 1).Locked = False
              
194           .Columns(kolomI + 2).Width = 2000
195           .Columns(kolomI + 2).Caption = "UMUR LAST SB"
196           .Columns(kolomI + 2).DataField = "umurServiceBySB"
197           .Splits(1).Columns(kolomI + 2).Locked = False
              
198           .Columns(kolomI + 3).Width = 2000
199           .Columns(kolomI + 3).Caption = "DEALER LAST SVC SB"
200           .Columns(kolomI + 3).DataField = "dealerLastSB"
201           .Splits(1).Columns(kolomI + 3).Locked = False
                      
202           .Columns(kolomI + 4).Width = 2000
203           .Columns(kolomI + 4).Caption = "LAST SVC GR"
204           .Columns(kolomI + 4).DataField = "lastGR"
205           .Splits(1).Columns(kolomI + 1).Locked = False
              
206           .Columns(kolomI + 5).Width = 2000
207           .Columns(kolomI + 5).Caption = "INFO LAST SVC GR"
208           .Columns(kolomI + 5).DataField = "lastInfoGR"
209           .Splits(1).Columns(kolomI + 5).Locked = False
              
210           .Columns(kolomI + 6).Width = 2000
211           .Columns(kolomI + 6).Caption = "KM LAST"
212           .Columns(kolomI + 6).DataField = "kmLast"
213           .Splits(1).Columns(kolomI + 6).Locked = False
              
214           .Columns(kolomI + 7).Width = 2000
215           .Columns(kolomI + 7).Caption = "SA LAST"
216           .Columns(kolomI + 7).DataField = "saLastSV"
217           .Splits(1).Columns(kolomI + 7).Locked = False
              
218           .Columns(kolomI + 8).Width = 2000
219           .Columns(kolomI + 8).Caption = "TIPE LAST SERVICE"
220           .Columns(kolomI + 8).DataField = "tipeLastService"
221           .Splits(1).Columns(kolomI + 8).Locked = False
       
222           .Columns(kolomI + 9).Width = 2000
223           .Columns(kolomI + 9).Caption = "UMUR SERVICE By LAST SVC"
224           .Columns(kolomI + 9).DataField = "umurLastService"
225           .Splits(1).Columns(kolomI + 9).Locked = False
              
226           If kolom(13).value = vbUnchecked Then
227               .Columns(kolomI).Visible = False
228               .Columns(kolomI + 1).Visible = False
229               .Columns(kolomI + 2).Visible = False
230               .Columns(kolomI + 3).Visible = False
231               .Columns(kolomI + 4).Visible = False
232               .Columns(kolomI + 5).Visible = False
233               .Columns(kolomI + 6).Visible = False
234               .Columns(kolomI + 7).Visible = False
235               .Columns(kolomI + 8).Visible = False
236               .Columns(kolomI + 9).Visible = False
237           End If
              
238           .Columns(kolomI + 10).Width = 2000
239           .Columns(kolomI + 10).Caption = "SB Next SVC"
240           .Columns(kolomI + 10).DataField = "SbeNext"
241           .Splits(1).Columns(kolomI + 10).Locked = False
              
242           .Columns(kolomI + 11).Width = 2000
243           .Columns(kolomI + 11).Caption = "Tgl Next SVC"
244           .Columns(kolomI + 11).DataField = "tglsbenext"
245           .Splits(1).Columns(kolomI + 11).Locked = False
              
246           If kolom(14).value = vbUnchecked Then
247               .Columns(kolomI + 10).Visible = False
248               .Columns(kolomI + 11).Visible = False
249           End If
                      
250           kolomI = kolomI + 2
              
251           .Columns(kolomI + 10).Width = 2000
252           .Columns(kolomI + 10).Caption = "DEALER PENJUAL"
253           .Columns(kolomI + 10).DataField = "DealerPenjual"
254           .Splits(1).Columns(kolomI + 10).Locked = False
              
255           If kolom(15).value = vbUnchecked Then
256               .Columns(kolomI + 10).Visible = False
257           End If
              
258           .Columns(kolomI + 11).Width = 2000
259           .Columns(kolomI + 11).Caption = "DEALER TERAKHIR SVC"
260           .Columns(kolomI + 11).DataField = "DealerLastSVC"
261           .Splits(1).Columns(kolomI + 11).Locked = False
              
262           If kolom(16).value = vbUnchecked Then
263               .Columns(kolomI + 11).Visible = False
264           End If
              
265           .Columns(kolomI + 12).Width = 2000
266           .Columns(kolomI + 12).Caption = "KATEGORI"
267           .Columns(kolomI + 12).DataField = "aktif"
268           .Splits(1).Columns(kolomI + 12).Locked = False
              
269           If kolom(17).value = vbUnchecked Then
270               .Columns(kolomI + 12).Visible = False
271           End If
              
272           .Columns(kolomI + 13).Width = 2000
273           .Columns(kolomI + 13).Caption = "WIRANIAGA"
274           .Columns(kolomI + 13).DataField = "WIRANIAGA"
275           .Splits(1).Columns(kolomI + 13).Locked = False
              
276           If kolom(18).value = vbUnchecked Then
277               .Columns(kolomI + 13).Visible = False
278           End If
              
279           .Columns(kolomI + 14).Width = 2000
280           .Columns(kolomI + 14).Caption = "CARA BAYAR"
281           .Columns(kolomI + 14).DataField = "caraBayar"
282           .Splits(1).Columns(kolomI + 14).Locked = False
              
283           If kolom(19).value = vbUnchecked Then
284               .Columns(kolomI + 14).Visible = False
285           End If
              
286           .Columns(kolomI + 15).Width = 2000
287           .Columns(kolomI + 15).Caption = "NAMA SPK"
288           .Columns(kolomI + 15).DataField = "namaSPK"
289           .Splits(1).Columns(kolomI + 15).Locked = False
                
290           If kolom(20).value = vbUnchecked Then
291               .Columns(kolomI + 15).Visible = False
292           End If
              
293           .Columns(kolomI + 16).Width = 2000
294           .Columns(kolomI + 16).Caption = "ALAMAT SPK"
295           .Columns(kolomI + 16).DataField = "alamatSPK"
296           .Splits(1).Columns(kolomI + 16).Locked = False
297           iAlamatSPK = kolomI + 16
              
298           .Columns(kolomI + 17).Width = 2000
299           .Columns(kolomI + 17).Caption = "PROVINSI SPK"
300           .Columns(kolomI + 17).DataField = "provinsiSPK"
301           .Splits(1).Columns(kolomI + 17).Locked = False
                
302           .Columns(kolomI + 18).Width = 2000
303           .Columns(kolomI + 18).Caption = "KOTA SPK"
304           .Columns(kolomI + 18).DataField = "kotaSPK"
305           .Splits(1).Columns(kolomI + 18).Locked = False
              
306           .Columns(kolomI + 19).Width = 2000
307           .Columns(kolomI + 19).Caption = "KEC SPK"
308           .Columns(kolomI + 19).DataField = "kecSPK"
309           .Splits(1).Columns(kolomI + 19).Locked = False
              
310           .Columns(kolomI + 20).Width = 2000
311           .Columns(kolomI + 20).Caption = "KODEPOS SPK"
312           .Columns(kolomI + 20).DataField = "kodeposSPK"
313           .Splits(1).Columns(kolomI + 20).Locked = False
              
              
314           If kolom(21).value = vbUnchecked Then
315               .Columns(kolomI + 16).Visible = False
316               .Columns(kolomI + 17).Visible = False
317               .Columns(kolomI + 18).Visible = False
318               .Columns(kolomI + 19).Visible = False
319               .Columns(kolomI + 20).Visible = False
320           End If
              
              
321           .Columns(kolomI + 21).Width = 2000
322           .Columns(kolomI + 21).Caption = "CDMA SPK"
323           .Columns(kolomI + 21).DataField = "cdmaSPK"
324           .Splits(1).Columns(kolomI + 21).Locked = False
              
325           .Columns(kolomI + 22).Width = 2000
326           .Columns(kolomI + 22).Caption = "HP 1 SPK"
327           .Columns(kolomI + 22).DataField = "hp1SPK"
328           .Splits(1).Columns(kolomI + 22).Locked = False
                      
329           .Columns(kolomI + 23).Width = 2000
330           .Columns(kolomI + 23).Caption = "HP 2 SPK"
331           .Columns(kolomI + 23).DataField = "hp2SPK"
332           .Splits(1).Columns(kolomI + 23).Locked = False
              
333           If kolom(22).value = vbUnchecked Then
334               .Columns(kolomI + 21).Visible = False
335               .Columns(kolomI + 22).Visible = False
336               .Columns(kolomI + 23).Visible = False
337           End If
              
338           .Columns(kolomI + 24).Width = 2000
339           .Columns(kolomI + 24).Caption = "NAMA STNK"
340           .Columns(kolomI + 24).DataField = "namaSTNK"
341           .Splits(1).Columns(kolomI + 24).Locked = False
              
342           If kolom(23).value = vbUnchecked Then
343               .Columns(kolomI + 24).Visible = False
344           End If
              
345           .Columns(kolomI + 25).Width = 2000
346           .Columns(kolomI + 25).Caption = "ALAMAT STNK"
347           .Columns(kolomI + 25).DataField = "alamatSTNK"
348           .Splits(1).Columns(kolomI + 25).Locked = False
349           iAlamatSTNK = kolomI + 25
              
              
350           .Columns(kolomI + 26).Width = 2000
351           .Columns(kolomI + 26).Caption = "PROVINSI STNK"
352           .Columns(kolomI + 26).DataField = "provinsiSTNK"
353           .Splits(1).Columns(kolomI + 26).Locked = False
              
354           .Columns(kolomI + 27).Width = 2000
355           .Columns(kolomI + 27).Caption = "KOTA STNK"
356           .Columns(kolomI + 27).DataField = "kotaSTNK"
357           .Splits(1).Columns(kolomI + 27).Locked = False
              
358           .Columns(kolomI + 28).Width = 2000
359           .Columns(kolomI + 28).Caption = "KECAMATAN STNK"
360           .Columns(kolomI + 28).DataField = "kecSTNK"
361           .Splits(1).Columns(kolomI + 28).Locked = False
       
362           .Columns(kolomI + 29).Width = 2000
363           .Columns(kolomI + 29).Caption = "KODEPOS STNK"
364           .Columns(kolomI + 29).DataField = "kodePosSTNK"
365           .Splits(1).Columns(kolomI + 29).Locked = False
              
366           If kolom(24).value = vbUnchecked Then
367               .Columns(kolomI + 25).Visible = False
368               .Columns(kolomI + 26).Visible = False
369               .Columns(kolomI + 27).Visible = False
370               .Columns(kolomI + 28).Visible = False
371               .Columns(kolomI + 29).Visible = False
372           End If
              
373           .Columns(kolomI + 30).Width = 2000
374           .Columns(kolomI + 30).Caption = "VOC"
375           .Columns(kolomI + 30).DataField = "voc"
376           .Splits(1).Columns(kolomI + 30).Locked = False
              
377           .Columns(kolomI + 31).Width = 2000
378           .Columns(kolomI + 31).Caption = "CR7"
379           .Columns(kolomI + 31).DataField = "cr7"
380           .Splits(1).Columns(kolomI + 31).Locked = False
       
381           If kolom(25).value = vbUnchecked Then
382               .Columns(kolomI + 30).Visible = False
383               .Columns(kolomI + 31).Visible = False
384           End If
              
              
385           If Check1.value = vbUnchecked Then
386               i = 33
387               Do While (kolomI + i) <= 137
388                   .Splits(1).Columns(kolomI + i).Visible = False
389                   i = i + 1
390               Loop
391           End If
392           i = 3
393           .Splits(1).Columns(0).Visible = False
394           .Splits(1).Columns(1).Visible = False
395           .Splits(1).Columns(2).Visible = False
396           Do While i <= 137
397               .Splits(0).Columns(i).Visible = False
398               i = i + 1
399           Loop
400       End With
End Sub
 
 

 

 

Private Sub GridDealer_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        If rsDealer.State Then
            rsDealer.MoveFirst
            Do While Not rsDealer.EOF
                rsDealer!terlibat = IIf(rsDealer!terlibat = True, False, True)
                rsDealer.MoveNext
            Loop
            rsDealer.MoveFirst
        End If
    End If
End Sub

Private Sub gridDealerNon_HeadClick(ByVal ColIndex As Integer)
    If ColIndex = 2 Then
        If rsDealerNon.State Then
            rsDealerNon.MoveFirst
            Do While Not rsDealerNon.EOF
                rsDealerNon!terlibat = IIf(rsDealerNon!terlibat = True, False, True)
                rsDealerNon.MoveNext
            Loop
            rsDealerNon.MoveFirst
        End If
    End If
End Sub

 

Private Sub gridTagih_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
    Dim rs As New ADODB.Recordset
    Set rs = rsRecord.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    RowStyle = "Putih"
    If rs("warna") <> "" Then
        If UCase(rs("Warna")) = "MERAH" Then
            RowStyle = "MERAH"
        ElseIf UCase(rs("WARNA")) = "KUNING" Then
            RowStyle = "KUNING"
        ElseIf UCase(rs("WARNA")) = "BIRU" Then
            RowStyle = "BIRU"
        ElseIf UCase(rs("WARNA")) = "HIJAU" Then
            RowStyle = "HIJAU"
        Else
            RowStyle = "PUTIH"
        End If
    End If

End Sub

Private Sub gridTagih_FilterChange()
    On Error GoTo errHandler
    Set cols = gridTagih.Columns
    Dim C As Integer
    
    C = gridTagih.col
    gridTagih.HoldFields
    rsRecord.filter = getFilter()
    If rsRecord.EOF Then
        For Each col In gridTagih.Columns
            col.FilterText = ""
        Next col
        rsRecord.filter = ""
    End If
    gridTagih.col = C
    gridTagih.EditActive = True
    lblJumlah.Caption = rsRecord.RecordCount
    
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

Private Sub gridTagih_HeadClick(ByVal ColIndex As Integer)
1     On Error GoTo err
2        Urut = IIf(Urut = "DESC", "ASC", "DESC")
3        If ColIndex <= 12 Then rsRecord.Sort = gridTagih.Columns(ColIndex).DataField & " " & Urut
4        gridTagih.SetFocus
5        Sendkeys "{Right " & Val(ColIndex) & "}"
6     Exit Sub
err:
7        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form Remainder")
End Sub

Private Sub gridTagih_AfterColUpdate(ByVal ColIndex As Integer)
    With gridTagih
        Select Case ColIndex
            Case 0
                If .Columns(ColIndex) Then
                    jmlTagih = jmlTagih + 1
                Else
                    jmlTagih = jmlTagih - 1
                End If
        End Select
 
    End With
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Sendkeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        Sendkeys ("”")
    End If
End Sub

Public Function getWarnaUnit(total As Integer) As String
    Dim hasil As String
    hasil = "Putih"
    If ckAktif.value = vbChecked Then
        If cbOperator.Text = "<" Then
            If CInt(Trim(total)) < blnAktif Then
                hasil = LabelAktif.Caption
                jumAktif = jumAktif + 1
            End If
            
        ElseIf cbOperator.Text = ">" Then
            If (total) > blnAktif Then
                hasil = LabelAktif.Caption
                jumAktif = jumAktif + 1
            End If
        ElseIf cbOperator.Text = "=" Then
            If (total) = blnAktif Then
                hasil = LabelAktif.Caption
                jumAktif = jumAktif + 1
            End If
        ElseIf cbOperator.Text = "RANGE" Then
            If (total) >= blnRangeAktifAwal And (total) <= blnRangeAktifAkhir Then
                hasil = LabelAktif.Caption
                jumAktif = jumAktif + 1
            End If
        End If
    End If
    '------------ PASIF 1 ---------------
    If ckPasif1.value = vbChecked Then
        If Combo1.Text = "<" Then
         
                If CInt((total)) < blnPasif1 Then
                    hasil = LabelPasif.Caption
                    jumPasif1 = jumPasif1 + 1
                End If
         
        ElseIf Combo1.Text = ">" Then
            If (total) > blnPasif1 Then
                hasil = LabelPasif.Caption
                jumPasif1 = jumPasif1 + 1
            End If
        ElseIf Combo1.Text = "=" Then
            If (total) = blnPasif1 Then
                hasil = LabelPasif.Caption
                jumPasif1 = jumPasif1 + 1
            End If
        ElseIf Combo1.Text = "RANGE" Then
            If total >= blnRangePasif1Awal And total <= blnRangePasif1Akhir Then
                hasil = LabelPasif.Caption
                jumPasif1 = jumPasif1 + 1
            End If
        End If
    End If
    '------------ PASIF 2 ---------------
    If ckPasif2.value = vbChecked Then
        If Combo2.Text = "<" Then
            
                If CInt((total)) < blnPasif2 Then
                    hasil = LabelPasif2.Caption
                    jumPasif2 = jumPasif2 + 1
                End If
          
        ElseIf Combo2.Text = ">" Then
            If (total) > blnPasif2 Then
                hasil = LabelPasif2.Caption
                jumPasif2 = jumPasif2 + 1
            End If
        ElseIf Combo2.Text = "=" Then
            If (total) = blnPasif2 Then
                hasil = LabelPasif2.Caption
                jumPasif2 = jumPasif2 + 1
            End If
        ElseIf Combo2.Text = "RANGE" Then
            If (total) >= blnPasif2Awal And (total) <= blnPasif2Akhir Then
                hasil = LabelPasif2.Caption
                jumPasif2 = jumPasif2 + 1
            End If
        End If
    End If
    '------------ PASIF 3 ---------------
    If ckPasif3.value = vbChecked Then
        If Combo3.Text = "<" Then
            
                If CInt((total)) < blnPasif3Bulan Then
                    hasil = LabelPasif3.Caption
                    jumPasif3 = jumPasif3 + 1
                End If
           
        ElseIf Combo3.Text = ">" Then
            If (total) > blnPasif3Bulan Then
                hasil = LabelPasif3.Caption
                 jumPasif3 = jumPasif3 + 1
            End If
        ElseIf Combo3.Text = "=" Then
            If (total) = blnPasif3Bulan Then
                hasil = LabelPasif3.Caption
                 jumPasif3 = jumPasif3 + 1
            End If
        ElseIf Combo3.Text = "RANGE" Then
            If (total) >= blnPasif3Awal And (total) <= blnPasif3Akhir Then
                hasil = LabelPasif3.Caption
                 jumPasif3 = jumPasif3 + 1
            End If
        End If
    End If
    getWarnaUnit = hasil
End Function
 
Private Sub isiDealer()
Dim rsTmp As New ADODB.Recordset

    Call HeaderCol
    ''Set rsTmp = objO.FetchData("select FS_KD_DEALER, FS_NM_DEALER from [ReplBengkel99]..T_DEALER with (nolock) where FS_LOG_DEL = '' and INTERNAL = 1 and FS_KD_DEALER <> 'NR 00' ORDER BY FS_KD_DEALER")
    Set rsTmp = objO.FetchData("select FS_KD_DEALER, FS_NM_DEALER from T_DEALER with (nolock) where FS_LOG_DEL = '' and internal = 1 and FS_KD_DEALER <> 'NR 00' ORDER BY FS_KD_DEALER")
    If Not rsTmp.EOF Then
        Call createRSDealer
        Do While Not rsTmp.EOF
            rsDealer.AddNew
            rsDealer!KodeDealer = rsTmp("FS_KD_DEALER")
            rsDealer!namaDealer = rsTmp("FS_NM_DEALER")
            rsDealer!terlibat = False
           
            rsDealer.Update
            rsTmp.MoveNext
        Loop
        
        GridDealer.DataSource = rsDealer
        GridDealer.EvenRowStyle.BackColor = vbYellow
        GridDealer.OddRowStyle.BackColor = vbWhite
        GridDealer.AlternatingRowStyle = True
    End If
End Sub

Private Sub isiDealerNon()
Dim rsTmp As New ADODB.Recordset

    Call HeaderColNon
    ''Set rsTmp = objO.FetchData("select FS_KD_DEALER, FS_NM_DEALER from [ReplBengkel99]..T_DEALER with (nolock) where FS_LOG_DEL = '' and INTERNAL = 1 and FS_KD_DEALER <> 'NR 00' ORDER BY FS_KD_DEALER")
    Set rsTmp = objO.FetchDataHO("select FS_KD_DEALER, FS_NM_DEALER from T_DEALER with (nolock) where FS_LOG_DEL = '' and internal = 0 and FS_KD_DEALER <> 'NR 00' ORDER BY FS_KD_DEALER")
    If Not rsTmp.EOF Then
        Call createRSDealerNon
        Do While Not rsTmp.EOF
            rsDealerNon.AddNew
            rsDealerNon!KodeDealer = Left(NulSID(rsTmp("FS_KD_DEALER")), 15)
            rsDealerNon!namaDealer = Left(NulSID(rsTmp("FS_NM_DEALER")), 25)
            rsDealerNon!terlibat = False
           
            rsDealerNon.Update
            rsTmp.MoveNext
        Loop
        
        gridDealerNon.DataSource = rsDealerNon
        gridDealerNon.EvenRowStyle.BackColor = vbYellow
        gridDealerNon.OddRowStyle.BackColor = vbWhite
        gridDealerNon.AlternatingRowStyle = True
    End If
End Sub

Private Sub HeaderCol()
    GridDealer.Columns(0).Width = 1000
    GridDealer.Columns(1).Width = 3500
    GridDealer.Columns(2).Width = 1000
    GridDealer.Columns(2).Alignment = dbgCenter
End Sub

Private Sub HeaderColNon()
    gridDealerNon.Columns(0).Width = 1000
    gridDealerNon.Columns(1).Width = 3500
    gridDealerNon.Columns(2).Width = 1000
    gridDealerNon.Columns(2).Alignment = dbgCenter
End Sub

Private Sub createRSDealer()
    
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

Private Sub createRSDealerNon()
    
    If rsDealerNon.State Then rsDealerNon.Close
    Set rsDealerNon = Nothing
    With rsDealerNon
        .Fields.Append "KODEDEALER", adVarChar, 15
        .Fields.Append "NAMADEALER", adVarChar, 25
        .Fields.Append "TERLIBAT", adBoolean
        .Open
    End With
    
    Set gridDealerNon.DataSource = rsDealerNon
End Sub

Private Sub isiModel()
    Set rsTemp = objO.FetchData("select distinct FS_GRP_MDL from model with (nolock) order by FS_GRP_MDL")
    If Not rsTemp.EOF Then
        Call createTbl
        Do While Not rsTemp.EOF
            rsGridMdl.AddNew
            rsGridMdl!fs_grp_mdl = Trim(rsTemp("FS_GRP_MDL"))
            rsGridMdl!ikut = False
            rsGridMdl.Update
            rsTemp.MoveNext
        Loop

        rsGridMdl.MoveFirst
        GridModel.Columns(0).Width = 4000
        GridModel.Columns(1).Alignment = dbgCenter
        GridModel.Columns(1).Width = 800
    End If
End Sub

Private Sub createTbl()
    
    If rsGridMdl.State Then rsGridMdl.Close
    Set rsGridMdl = Nothing
    With rsGridMdl
        .Fields.Append "FS_GRP_MDL", adVarChar, 50
        .Fields.Append "IKUT", adBoolean
        .Open
    End With
    
    GridModel.DataSource = rsGridMdl
    GridModel.AlternatingRowStyle = True
End Sub

