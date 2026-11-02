VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmOrderValidasi 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Validasi Order"
   ClientHeight    =   9930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   10620
   Begin VB.CommandButton cmdPsmInfo 
      BackColor       =   &H00DCFCFB&
      Caption         =   "PSM ITEM"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   261
      Top             =   9255
      Width           =   1170
   End
   Begin VB.CommandButton cmdTungguValidasiTL 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Tunggu Validasi TL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   258
      Top             =   9255
      Width           =   1890
   End
   Begin VB.Frame frameUploadKTP 
      Height          =   4815
      Left            =   4200
      TabIndex        =   253
      Top             =   2880
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton cmdTutupUploadKTP 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Tutup KTP"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   257
         Top             =   4320
         Width           =   1410
      End
      Begin VB.CommandButton cmdEditCustomer 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Edit Customer"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   6240
         Style           =   1  'Graphical
         TabIndex        =   256
         Top             =   4320
         Width           =   1410
      End
      Begin TrueOleDBGrid70.TDBGrid GridFind 
         Height          =   3735
         Left            =   120
         TabIndex        =   255
         Top             =   480
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   6588
         _LayoutType     =   4
         _RowHeight      =   13
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   4
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Pilih"
         Columns(0).DataField=   "pilih"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "KodeLgn"
         Columns(1).DataField=   "idcust"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "No KTP"
         Columns(2).DataField=   "NoKTP"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Nama  Customer"
         Columns(3).DataField=   "nama"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Alamat"
         Columns(4).DataField=   "alamat"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "No HP"
         Columns(5).DataField=   "hp"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "WO Terakhir"
         Columns(6).DataField=   "tglWO"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "KodeLgn Lama"
         Columns(7).DataField=   "Kdlgn"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   8
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=8"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=873"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=767"
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
         Splits.Count    =   1
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
         MultiSelect     =   0
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=862,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=859,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=860,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=861,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=870,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=867,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=868,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=869,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=878,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=875,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=876,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=877,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=886,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=883,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=884,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=885,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=894,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=891,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=892,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=893,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=28,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=25,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=26,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=27,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=32,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=29,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=30,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=31,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
         _StyleDefs(68)  =   "Named:id=33:Normal"
         _StyleDefs(69)  =   ":id=33,.parent=0"
         _StyleDefs(70)  =   "Named:id=34:Heading"
         _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(72)  =   ":id=34,.wraptext=-1"
         _StyleDefs(73)  =   "Named:id=35:Footing"
         _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(75)  =   "Named:id=36:Selected"
         _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H80FF&,.fgcolor=&H8000000E&"
         _StyleDefs(77)  =   "Named:id=37:Caption"
         _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(79)  =   "Named:id=38:HighlightRow"
         _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(81)  =   "Named:id=39:EvenRow"
         _StyleDefs(82)  =   ":id=39,.parent=33"
         _StyleDefs(83)  =   "Named:id=40:OddRow"
         _StyleDefs(84)  =   ":id=40,.parent=33"
         _StyleDefs(85)  =   "Named:id=41:RecordSelector"
         _StyleDefs(86)  =   ":id=41,.parent=34"
         _StyleDefs(87)  =   "Named:id=42:FilterBar"
         _StyleDefs(88)  =   ":id=42,.parent=33"
         _StyleDefs(89)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(90)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(91)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(92)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         _StyleDefs(93)  =   "Named:id=87:SudahDitagih"
         _StyleDefs(94)  =   ":id=87,.parent=132,.fgcolor=&HFF&,.locked=-1,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(95)  =   ":id=87,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(96)  =   ":id=87,.fontname=Tahoma"
         _StyleDefs(97)  =   "Named:id=514:Kuning"
         _StyleDefs(98)  =   ":id=514,.parent=87,.bgcolor=&HFFFF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
         _StyleDefs(99)  =   ":id=514,.italic=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(100) =   ":id=514,.fontname=Tahoma"
         _StyleDefs(101) =   "Named:id=515:Biru"
         _StyleDefs(102) =   ":id=515,.parent=514,.bgcolor=&HFF8080&"
         _StyleDefs(103) =   "Named:id=516:Merah"
         _StyleDefs(104) =   ":id=516,.parent=515,.bgcolor=&H88A0F7&,.fgcolor=&H0&"
         _StyleDefs(105) =   "Named:id=517:Hijau"
         _StyleDefs(106) =   ":id=517,.parent=516,.bgcolor=&HFF00&"
         _StyleDefs(107) =   "Named:id=518:Putih"
         _StyleDefs(108) =   ":id=518,.parent=517,.bgcolor=&HFFFFFF&"
      End
      Begin VB.Line Line29 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   195
         Y2              =   195
      End
      Begin VB.Line Line28 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   75
         Y2              =   75
      End
      Begin VB.Label Label56 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Customer Belum Upload KTP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4050
         TabIndex        =   254
         Top             =   0
         Width           =   3945
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   9255
      End
   End
   Begin VB.CommandButton btnUploadKTP 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Upload KTP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   252
      Top             =   0
      Width           =   1410
   End
   Begin VB.Frame frameNonCR7 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   1440
      TabIndex        =   238
      Top             =   3000
      Visible         =   0   'False
      Width           =   11655
      Begin VB.TextBox Text16 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   244
         Top             =   680
         Width           =   5655
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   243
         Top             =   3230
         Width           =   5655
      End
      Begin VB.TextBox txtCatatNonCR7 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   249
         TabIndex        =   242
         Top             =   5520
         Width           =   9210
      End
      Begin VB.TextBox Text13 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   241
         Top             =   3230
         Width           =   5655
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   240
         Top             =   680
         Width           =   5655
      End
      Begin VB.CommandButton cmdNonCR7 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOrderValidasi.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   239
         Top             =   5400
         Width           =   1245
      End
      Begin VB.Label Label54 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "CR 7 : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   251
         Top             =   2900
         Width           =   3615
      End
      Begin VB.Label Label53 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Final Inspection dan Test Drive : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5895
         TabIndex        =   250
         Top             =   375
         Width           =   5655
      End
      Begin VB.Label Label52 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Keluhan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   249
         Top             =   2900
         Width           =   3615
      End
      Begin VB.Label Label50 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   248
         Top             =   375
         Width           =   3615
      End
      Begin VB.Line Line27 
         X1              =   0
         X2              =   11640
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Label Label49 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
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
         TabIndex        =   247
         Top             =   5520
         Width           =   585
      End
      Begin VB.Line Line26 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line25 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line24 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4680
         TabIndex        =   246
         Top             =   0
         Width           =   2790
      End
      Begin VB.Line Line23 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11655
      End
      Begin VB.Line Line22 
         BorderColor     =   &H000000FF&
         X1              =   7560
         X2              =   11400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line21 
         BorderColor     =   &H000000FF&
         X1              =   7560
         X2              =   11400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line20 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   3840
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label47 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4320
         TabIndex        =   245
         Top             =   0
         Width           =   2790
      End
      Begin VB.Line Line19 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   3840
         Y1              =   120
         Y2              =   120
      End
   End
   Begin VB.Frame FrameKeluhDmsNonDSS 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   1320
      TabIndex        =   123
      Top             =   2400
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Frame frameFoto 
         Caption         =   "Foto CR 7"
         Height          =   4215
         Index           =   0
         Left            =   5640
         TabIndex        =   234
         Top             =   120
         Visible         =   0   'False
         Width           =   6015
         Begin InetCtlsObjects.Inet Inet1 
            Left            =   2400
            Top             =   1920
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
         End
         Begin VB.CommandButton cmdUpload 
            BackColor       =   &H0080FF80&
            Caption         =   "BROWSE"
            Height          =   375
            Index           =   0
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   236
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdTutupFoto 
            BackColor       =   &H000000FF&
            Caption         =   "TUTUP"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   235
            Top             =   3720
            Width           =   975
         End
         Begin VB.Image imgFotoCR7 
            Appearance      =   0  'Flat
            Height          =   3375
            Index           =   0
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameFoto 
         Caption         =   "Foto CR 7"
         Height          =   4215
         Index           =   1
         Left            =   5640
         TabIndex        =   231
         Top             =   120
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CommandButton cmdTutupFoto 
            BackColor       =   &H000000FF&
            Caption         =   "TUTUP"
            Height          =   375
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   233
            Top             =   3720
            Width           =   975
         End
         Begin VB.CommandButton cmdUpload 
            BackColor       =   &H0080FF80&
            Caption         =   "BROWSE"
            Height          =   375
            Index           =   1
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   232
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Image imgFotoCR7 
            Appearance      =   0  'Flat
            Height          =   3375
            Index           =   1
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameFoto 
         Caption         =   "Foto CR 7"
         Height          =   4215
         Index           =   2
         Left            =   5640
         TabIndex        =   228
         Top             =   120
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CommandButton cmdUpload 
            BackColor       =   &H0080FF80&
            Caption         =   "BROWSE"
            Height          =   375
            Index           =   2
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   230
            Top             =   3720
            Width           =   1215
         End
         Begin VB.CommandButton cmdTutupFoto 
            BackColor       =   &H000000FF&
            Caption         =   "TUTUP"
            Height          =   375
            Index           =   2
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   229
            Top             =   3720
            Width           =   975
         End
         Begin VB.Image imgFotoCR7 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   3375
            Index           =   2
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameFoto 
         Caption         =   "Foto CR 7"
         Height          =   4215
         Index           =   3
         Left            =   5760
         TabIndex        =   225
         Top             =   120
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CommandButton cmdTutupFoto 
            BackColor       =   &H000000FF&
            Caption         =   "TUTUP"
            Height          =   375
            Index           =   3
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   227
            Top             =   3720
            Width           =   975
         End
         Begin VB.CommandButton cmdUpload 
            BackColor       =   &H0080FF80&
            Caption         =   "BROWSE"
            Height          =   375
            Index           =   3
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   226
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Image imgFotoCR7 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   3375
            Index           =   3
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameFoto 
         Caption         =   "Foto CR 7"
         Height          =   4215
         Index           =   4
         Left            =   5760
         TabIndex        =   222
         Top             =   120
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CommandButton cmdTutupFoto 
            BackColor       =   &H000000FF&
            Caption         =   "TUTUP"
            Height          =   375
            Index           =   4
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   224
            Top             =   3720
            Width           =   975
         End
         Begin VB.CommandButton cmdUpload 
            BackColor       =   &H0080FF80&
            Caption         =   "BROWSE"
            Height          =   375
            Index           =   4
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   223
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Image imgFotoCR7 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   3375
            Index           =   4
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameFoto 
         Caption         =   "Foto CR 7"
         Height          =   4215
         Index           =   5
         Left            =   5760
         TabIndex        =   219
         Top             =   120
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CommandButton cmdTutupFoto 
            BackColor       =   &H000000FF&
            Caption         =   "TUTUP"
            Height          =   375
            Index           =   5
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   221
            Top             =   3720
            Width           =   975
         End
         Begin VB.CommandButton cmdUpload 
            BackColor       =   &H0080FF80&
            Caption         =   "BROWSE"
            Height          =   375
            Index           =   5
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   220
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Image imgFotoCR7 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   3375
            Index           =   5
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.Frame frameFoto 
         Caption         =   "Foto CR 7"
         Height          =   4215
         Index           =   6
         Left            =   5760
         TabIndex        =   216
         Top             =   120
         Visible         =   0   'False
         Width           =   6015
         Begin VB.CommandButton cmdTutupFoto 
            BackColor       =   &H000000FF&
            Caption         =   "TUTUP"
            Height          =   375
            Index           =   6
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   218
            Top             =   3720
            Width           =   975
         End
         Begin VB.CommandButton cmdUpload 
            BackColor       =   &H0080FF80&
            Caption         =   "BROWSE"
            Height          =   375
            Index           =   6
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   217
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Image imgFotoCR7 
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            Height          =   3375
            Index           =   6
            Left            =   120
            Stretch         =   -1  'True
            Top             =   240
            Width           =   5775
         End
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOrderValidasi.frx":0532
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   6360
         Width           =   1245
      End
      Begin VB.CommandButton btnPilihFoto 
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
         Index           =   6
         Left            =   11220
         Picture         =   "frmOrderValidasi.frx":0A64
         Style           =   1  'Graphical
         TabIndex        =   215
         Top             =   6030
         Width           =   350
      End
      Begin VB.CommandButton btnPilihFoto 
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
         Index           =   5
         Left            =   11220
         Picture         =   "frmOrderValidasi.frx":0F00
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   5770
         Width           =   350
      End
      Begin VB.CommandButton btnPilihFoto 
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
         Index           =   4
         Left            =   11220
         Picture         =   "frmOrderValidasi.frx":139C
         Style           =   1  'Graphical
         TabIndex        =   213
         Top             =   5520
         Width           =   350
      End
      Begin VB.CommandButton btnPilihFoto 
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
         Index           =   3
         Left            =   11220
         Picture         =   "frmOrderValidasi.frx":1838
         Style           =   1  'Graphical
         TabIndex        =   212
         Top             =   5280
         Width           =   350
      End
      Begin VB.CommandButton btnPilihFoto 
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
         Index           =   2
         Left            =   11220
         Picture         =   "frmOrderValidasi.frx":1CD4
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   5040
         Width           =   350
      End
      Begin VB.CommandButton btnPilihFoto 
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
         Index           =   1
         Left            =   11220
         Picture         =   "frmOrderValidasi.frx":2170
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   4800
         Width           =   350
      End
      Begin VB.CommandButton btnPilihFoto 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   275
         Index           =   0
         Left            =   11220
         Picture         =   "frmOrderValidasi.frx":260C
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   4560
         Width           =   350
      End
      Begin VB.TextBox txtBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   10010
         Locked          =   -1  'True
         TabIndex        =   208
         Top             =   6000
         Width           =   1200
      End
      Begin VB.TextBox txtBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   10010
         Locked          =   -1  'True
         TabIndex        =   207
         Top             =   5760
         Width           =   1200
      End
      Begin VB.TextBox txtBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   10010
         Locked          =   -1  'True
         TabIndex        =   206
         Top             =   5520
         Width           =   1200
      End
      Begin VB.TextBox txtBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   10010
         Locked          =   -1  'True
         TabIndex        =   205
         Top             =   5280
         Width           =   1200
      End
      Begin VB.TextBox txtBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   10010
         Locked          =   -1  'True
         TabIndex        =   204
         Top             =   5040
         Width           =   1200
      End
      Begin VB.TextBox txtBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   10010
         Locked          =   -1  'True
         TabIndex        =   203
         Top             =   4800
         Width           =   1200
      End
      Begin VB.TextBox txtBiaya 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   10010
         Locked          =   -1  'True
         TabIndex        =   202
         Top             =   4560
         Width           =   1200
      End
      Begin VB.TextBox txtSublet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   8960
         TabIndex        =   201
         Top             =   6000
         Width           =   1060
      End
      Begin VB.TextBox txtSublet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   8960
         TabIndex        =   200
         Top             =   5760
         Width           =   1060
      End
      Begin VB.TextBox txtSublet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   8960
         TabIndex        =   199
         Top             =   5520
         Width           =   1060
      End
      Begin VB.TextBox txtSublet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   8960
         TabIndex        =   198
         Top             =   5280
         Width           =   1060
      End
      Begin VB.TextBox txtSublet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   8960
         TabIndex        =   197
         Top             =   5040
         Width           =   1060
      End
      Begin VB.TextBox txtSublet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   8960
         TabIndex        =   196
         Top             =   4800
         Width           =   1060
      End
      Begin VB.TextBox txtSublet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   8960
         TabIndex        =   195
         Top             =   4560
         Width           =   1060
      End
      Begin VB.TextBox txtPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   7890
         TabIndex        =   194
         Top             =   6000
         Width           =   1080
      End
      Begin VB.TextBox txtPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   7890
         TabIndex        =   193
         Top             =   5760
         Width           =   1080
      End
      Begin VB.TextBox txtPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   7890
         TabIndex        =   192
         Top             =   5520
         Width           =   1080
      End
      Begin VB.TextBox txtPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   7890
         TabIndex        =   191
         Top             =   5280
         Width           =   1080
      End
      Begin VB.TextBox txtPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   7890
         TabIndex        =   190
         Top             =   5040
         Width           =   1080
      End
      Begin VB.TextBox txtPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   7890
         TabIndex        =   189
         Top             =   4800
         Width           =   1080
      End
      Begin VB.TextBox txtPart 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   7890
         TabIndex        =   188
         Top             =   4560
         Width           =   1080
      End
      Begin VB.TextBox txtJasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   187
         Top             =   6000
         Width           =   1200
      End
      Begin VB.TextBox txtJasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   186
         Top             =   5760
         Width           =   1200
      End
      Begin VB.TextBox txtJasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   185
         Top             =   5520
         Width           =   1200
      End
      Begin VB.TextBox txtJasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   184
         Top             =   5280
         Width           =   1200
      End
      Begin VB.TextBox txtJasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   183
         Top             =   5040
         Width           =   1200
      End
      Begin VB.TextBox txtJasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   182
         Top             =   4800
         Width           =   1200
      End
      Begin VB.TextBox txtJasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   6705
         Locked          =   -1  'True
         TabIndex        =   181
         Top             =   4560
         Width           =   1200
      End
      Begin VB.TextBox txtTglKembali 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   6
         Left            =   5280
         TabIndex        =   180
         Top             =   6000
         Width           =   1440
      End
      Begin VB.TextBox txtTglKembali 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   5
         Left            =   5280
         TabIndex        =   179
         Top             =   5760
         Width           =   1440
      End
      Begin VB.TextBox txtTglKembali 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   4
         Left            =   5280
         TabIndex        =   178
         Top             =   5520
         Width           =   1440
      End
      Begin VB.TextBox txtTglKembali 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   3
         Left            =   5280
         TabIndex        =   177
         Top             =   5280
         Width           =   1440
      End
      Begin VB.TextBox txtTglKembali 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   2
         Left            =   5280
         TabIndex        =   176
         Top             =   5040
         Width           =   1440
      End
      Begin VB.TextBox txtTglKembali 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   1
         Left            =   5280
         TabIndex        =   175
         Top             =   4800
         Width           =   1440
      End
      Begin VB.TextBox txtTglKembali 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   300
         Index           =   0
         Left            =   5280
         TabIndex        =   174
         Top             =   4560
         Width           =   1440
      End
      Begin VB.TextBox txtJobOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   173
         Top             =   6000
         Width           =   2535
      End
      Begin VB.TextBox txtJobOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   172
         Top             =   5760
         Width           =   2535
      End
      Begin VB.TextBox txtJobOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   171
         Top             =   5520
         Width           =   2535
      End
      Begin VB.TextBox txtJobOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   170
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox txtJobOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   169
         Top             =   5040
         Width           =   2535
      End
      Begin VB.TextBox txtJobOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   168
         Top             =   4800
         Width           =   2535
      End
      Begin VB.TextBox txtJobOrder 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   167
         Top             =   4560
         Width           =   2535
      End
      Begin VB.TextBox txtSaran 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   600
         TabIndex        =   166
         Top             =   6000
         Width           =   2175
      End
      Begin VB.TextBox txtSaran 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   600
         TabIndex        =   165
         Top             =   5760
         Width           =   2175
      End
      Begin VB.TextBox txtSaran 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   600
         TabIndex        =   164
         Top             =   5520
         Width           =   2175
      End
      Begin VB.TextBox txtSaran 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   600
         TabIndex        =   163
         Top             =   5280
         Width           =   2175
      End
      Begin VB.TextBox txtSaran 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   600
         TabIndex        =   162
         Top             =   5040
         Width           =   2175
      End
      Begin VB.TextBox txtSaran 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   600
         TabIndex        =   161
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox txtSaran 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   600
         TabIndex        =   160
         Top             =   4560
         Width           =   2175
      End
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   159
         Top             =   6000
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   158
         Top             =   5760
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   157
         Top             =   5520
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   156
         Top             =   5280
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   155
         Top             =   5040
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   154
         Top             =   4800
         Width           =   495
      End
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   153
         Top             =   4560
         Width           =   495
      End
      Begin VB.TextBox Text11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10010
         TabIndex        =   152
         Text            =   "Est Biaya"
         Top             =   4290
         Width           =   1200
      End
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8960
         TabIndex        =   151
         Text            =   "Est Sublet"
         Top             =   4290
         Width           =   1060
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7890
         TabIndex        =   150
         Text            =   "Est Part"
         Top             =   4290
         Width           =   1080
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6705
         TabIndex        =   149
         Text            =   "Est Jasa"
         Top             =   4290
         Width           =   1200
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5280
         TabIndex        =   148
         Text            =   "Est. Tgl Kembali"
         Top             =   4290
         Width           =   1440
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   147
         Text            =   "Job Order"
         Top             =   4290
         Width           =   2580
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   600
         TabIndex        =   146
         Text            =   "Temuan Saran Perbaikan"
         Top             =   4290
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   145
         Text            =   "No."
         Top             =   4290
         Width           =   495
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1500
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   128
         Top             =   680
         Width           =   5655
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1500
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   127
         Top             =   2625
         Width           =   5655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   249
         TabIndex        =   126
         Top             =   6480
         Visible         =   0   'False
         Width           =   9210
      End
      Begin VB.TextBox txtCatatanCR7 
         BackColor       =   &H00FFFFFF&
         Height          =   1500
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   125
         Top             =   2625
         Width           =   5655
      End
      Begin VB.TextBox txtFI 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   1500
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   124
         Top             =   680
         Width           =   5655
      End
      Begin MSComDlg.CommonDialog cdgopen 
         Index           =   0
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdgopen 
         Index           =   1
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdgopen 
         Index           =   2
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdgopen 
         Index           =   3
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdgopen 
         Index           =   4
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdgopen 
         Index           =   5
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog cdgopen 
         Index           =   6
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Line Line18 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   3840
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4320
         TabIndex        =   135
         Top             =   0
         Width           =   2790
      End
      Begin VB.Line Line17 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   3840
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line16 
         BorderColor     =   &H000000FF&
         X1              =   7560
         X2              =   11400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line15 
         BorderColor     =   &H000000FF&
         X1              =   7560
         X2              =   11400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11655
      End
      Begin VB.Line Line14 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label45 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4680
         TabIndex        =   134
         Top             =   0
         Width           =   2790
      End
      Begin VB.Line Line13 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
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
         TabIndex        =   133
         Top             =   6480
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Line Line10 
         X1              =   0
         X2              =   11640
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label32 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   132
         Top             =   375
         Width           =   3615
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Keluhan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   131
         Top             =   2295
         Width           =   3615
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Hasil Final Inspection dan Test Drive : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5895
         TabIndex        =   130
         Top             =   375
         Width           =   5655
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   129
         Top             =   2295
         Width           =   3615
      End
   End
   Begin VB.TextBox txtKodeLgnLama 
      Height          =   285
      Left            =   1410
      TabIndex        =   142
      Top             =   630
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.CommandButton cmdBatalRequest 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Batal Re&q Disc"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   141
      Top             =   9255
      Width           =   1890
   End
   Begin VB.Frame framePembatalanWashing 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Pembatalan Washing"
      Height          =   2295
      Left            =   7200
      TabIndex        =   137
      Top             =   2160
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command2 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Simpan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   139
         Top             =   1800
         Width           =   1725
      End
      Begin VB.TextBox txtAlasanBtlWashing 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   138
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.CommandButton tblCetakWO 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Cetak Ulang WO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   117
      Top             =   9240
      Width           =   1890
   End
   Begin VB.CommandButton TblKeluar 
      BackColor       =   &H00DCFCFB&
      Caption         =   "&Kembali ke Menu"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   9600
      Width           =   1890
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
      Height          =   300
      Left            =   13800
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   13800
      TabIndex        =   7
      Top             =   600
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
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Frame framKeluh 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   1575
      TabIndex        =   105
      Top             =   2430
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton btnSimpan 
         Appearance      =   0  'Flat
         BackColor       =   &H00DCFCFB&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10200
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmOrderValidasi.frx":2AA8
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   2790
         Width           =   1245
      End
      Begin VB.TextBox Txt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         MaxLength       =   249
         TabIndex        =   118
         Top             =   2840
         Width           =   9210
      End
      Begin VB.TextBox txtKeluhan 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   5880
         MultiLine       =   -1  'True
         TabIndex        =   107
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtPermintaan 
         BackColor       =   &H00C0E0FF&
         Height          =   2175
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   106
         Top             =   480
         Width           =   5655
      End
      Begin VB.Line Line9 
         X1              =   0
         X2              =   11640
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Catatan"
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
         TabIndex        =   119
         Top             =   2880
         Width           =   585
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   7680
         X2              =   11520
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Permintaan && Keluhan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4440
         TabIndex        =   113
         Top             =   0
         Width           =   2790
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   11655
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   12885
      Begin VB.CommandButton btnBatalWashing 
         BackColor       =   &H000000FF&
         Caption         =   "Batal Washing"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   136
         Top             =   1680
         Width           =   1890
      End
      Begin VB.CommandButton cmdKeluh 
         BackColor       =   &H0000FFFF&
         Caption         =   "&Permintaan && Keluhan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   1680
         Width           =   2130
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   930
         TabIndex        =   9
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   930
         TabIndex        =   10
         Top             =   750
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   930
         TabIndex        =   11
         Top             =   1035
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   930
         TabIndex        =   12
         Top             =   1350
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   7
         Left            =   930
         TabIndex        =   13
         Top             =   1680
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   10
         Left            =   5550
         TabIndex        =   14
         Top             =   435
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   12
         Left            =   5550
         TabIndex        =   15
         Top             =   750
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   15
         Left            =   5550
         TabIndex        =   16
         Top             =   1350
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   14
         Left            =   5550
         TabIndex        =   17
         Top             =   1050
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   13
         Left            =   7890
         TabIndex        =   18
         Top             =   750
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   960
         TabIndex        =   19
         Top             =   480
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   16
         Left            =   9510
         TabIndex        =   20
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   17
         Left            =   9510
         TabIndex        =   21
         Top             =   735
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   11
         Left            =   7470
         TabIndex        =   22
         Top             =   435
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   3480
         TabIndex        =   23
         Top             =   1680
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   9
         Left            =   6090
         TabIndex        =   24
         Top             =   1680
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   18
         Left            =   9510
         TabIndex        =   25
         Top             =   1035
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   41
         Left            =   3840
         TabIndex        =   26
         Top             =   480
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   40
         Left            =   2505
         TabIndex        =   27
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   34
         Left            =   9510
         TabIndex        =   28
         Top             =   1350
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   503
         _Version        =   393216
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
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   8640
         X2              =   12720
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   8640
         X2              =   12720
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data Customer && Kendaraan"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   4320
         TabIndex        =   112
         Top             =   45
         Width           =   3675
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   3960
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine."
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
         Left            =   8820
         TabIndex        =   43
         Top             =   1080
         Width           =   540
      End
      Begin VB.Label Label2 
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
         Left            =   5505
         TabIndex        =   42
         Top             =   1725
         Width           =   435
      End
      Begin VB.Label Label55 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DO TAM"
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
         Left            =   6720
         TabIndex        =   41
         Top             =   480
         Width           =   585
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Frame."
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
         Left            =   8850
         TabIndex        =   40
         Top             =   780
         Width           =   510
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Polisi"
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
         Left            =   8715
         TabIndex        =   39
         Top             =   495
         Width           =   645
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
         Left            =   135
         TabIndex        =   38
         Top             =   480
         Width           =   660
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Warna"
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
         Left            =   4920
         TabIndex        =   37
         Top             =   1365
         Width           =   480
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Left            =   4980
         TabIndex        =   36
         Top             =   765
         Width           =   420
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Th. Mobil"
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
         Left            =   4755
         TabIndex        =   35
         Top             =   495
         Width           =   645
      End
      Begin VB.Label Label22 
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
         Left            =   405
         TabIndex        =   34
         Top             =   1695
         Width           =   375
      End
      Begin VB.Label Label23 
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
         Left            =   450
         TabIndex        =   33
         Top             =   1395
         Width           =   330
      End
      Begin VB.Label Label24 
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
         Left            =   285
         TabIndex        =   32
         Top             =   1095
         Width           =   495
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         TabIndex        =   31
         Top             =   795
         Width           =   405
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MOU"
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
         Left            =   8955
         TabIndex        =   30
         Top             =   1395
         Width           =   345
      End
      Begin VB.Label Label3 
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
         Left            =   3060
         TabIndex        =   29
         Top             =   1725
         Width           =   315
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   12855
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   5655
      LargeChange     =   800
      Left            =   12960
      Max             =   2600
      SmallChange     =   600
      TabIndex        =   96
      Top             =   3600
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   480
      ScaleHeight     =   5595
      ScaleWidth      =   12795
      TabIndex        =   66
      Top             =   3600
      Width           =   12855
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8055
         Left            =   120
         TabIndex        =   67
         Top             =   0
         Width           =   12375
         Begin VB.TextBox txtCatatan 
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            TabIndex        =   143
            Top             =   6660
            Width           =   7725
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H0080FF80&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   735
            Left            =   1440
            TabIndex        =   121
            Top             =   7350
            Width           =   6855
            Begin VB.TextBox txtInfoNea 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2520
               MaxLength       =   49
               TabIndex        =   3
               ToolTipText     =   ".:: Keterangan NEA ::."
               Top             =   360
               Visible         =   0   'False
               Width           =   4290
            End
            Begin VB.CheckBox chPS 
               BackColor       =   &H0080FF80&
               Caption         =   "Profit Sharing"
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
               Height          =   255
               Left            =   3600
               TabIndex        =   2
               Top             =   60
               Width           =   1935
            End
            Begin VB.CheckBox chNea 
               BackColor       =   &H0080FF80&
               Caption         =   "N.E.A."
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
               Height          =   255
               Left            =   2520
               TabIndex        =   1
               Top             =   60
               Width           =   975
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Additional Order Tipe : "
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
               Height          =   195
               Index           =   7
               Left            =   120
               TabIndex        =   122
               Top             =   60
               Width           =   2265
            End
         End
         Begin VB.CommandButton TblBatal 
            BackColor       =   &H00DCFCFB&
            Caption         =   "&Batal"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   7560
            Width           =   1890
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1170
            Left            =   10320
            TabIndex        =   85
            Top             =   4800
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   26
               Left            =   915
               TabIndex        =   86
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   27
               Left            =   720
               TabIndex        =   87
               Top             =   480
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sublet"
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
               Left            =   195
               TabIndex        =   89
               Top             =   525
               Width           =   450
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               Left            =   195
               TabIndex        =   88
               Top             =   120
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   10320
            TabIndex        =   78
            Top             =   2640
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   23
               Left            =   915
               TabIndex        =   79
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   24
               Left            =   480
               TabIndex        =   80
               Top             =   480
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
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
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   25
               Left            =   480
               TabIndex        =   81
               Top             =   840
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
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
               Format          =   "#,##0.00##"
               PromptChar      =   ">"
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Oli"
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
               Left            =   240
               TabIndex        =   84
               Top             =   840
               Width           =   180
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Part"
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
               TabIndex        =   83
               Top             =   525
               Width           =   300
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               TabIndex        =   82
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin VB.CommandButton TblValid 
            BackColor       =   &H00DCFCFB&
            Caption         =   "&Validasi"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   10320
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   7560
            Width           =   1890
         End
         Begin VB.CommandButton cmdCetak 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Cetak"
            Height          =   390
            Left            =   10440
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   8520
            Width           =   1890
         End
         Begin VB.TextBox txtTotalRp 
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1425
            TabIndex        =   76
            Top             =   6990
            Width           =   10815
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1290
            Left            =   10320
            TabIndex        =   68
            Top             =   240
            Width           =   2055
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   21
               Left            =   915
               TabIndex        =   69
               Top             =   120
               Visible         =   0   'False
               Width           =   570
               _ExtentX        =   1005
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   22
               Left            =   720
               TabIndex        =   70
               Top             =   480
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   39
               Left            =   720
               TabIndex        =   71
               Top             =   840
               Width           =   1260
               _ExtentX        =   2223
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   12640511
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
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Body R"
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
               Left            =   105
               TabIndex        =   74
               Top             =   885
               Width           =   510
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Labour"
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
               TabIndex        =   73
               Top             =   525
               Width           =   495
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Discount                   %"
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
               TabIndex        =   72
               Top             =   165
               Visible         =   0   'False
               Width           =   1635
            End
         End
         Begin TrueOleDBGrid70.TDBGrid JobGrid 
            Height          =   2295
            Left            =   0
            TabIndex        =   75
            Top             =   240
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   4048
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
            Columns(8)._VlistStyle=   4
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
            Columns(10).NumberFormat=   "Standard"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   4
            Columns(11)._MaxComboItems=   5
            Columns(11).DataField=   ""
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).DataField=   ""
            Columns(12).NumberFormat=   "Standard"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).DataField=   ""
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).DataField=   ""
            Columns(14).NumberFormat=   "Standard"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).DataField=   ""
            Columns(15).NumberFormat=   "Standard"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(16)._VlistStyle=   4
            Columns(16)._MaxComboItems=   5
            Columns(16).DataField=   ""
            Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   17
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=17"
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
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6)._ColStyle=8196"
            Splits(0)._ColumnProps(29)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(30)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(31)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(32)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(33)=   "Column(7)._ColStyle=8196"
            Splits(0)._ColumnProps(34)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(35)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(36)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(37)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(38)=   "Column(8)._ColStyle=8196"
            Splits(0)._ColumnProps(39)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(40)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(41)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(42)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(43)=   "Column(9)._ColStyle=8196"
            Splits(0)._ColumnProps(44)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(45)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(46)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(48)=   "Column(10)._ColStyle=8196"
            Splits(0)._ColumnProps(49)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(50)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(51)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(53)=   "Column(11)._ColStyle=8196"
            Splits(0)._ColumnProps(54)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(55)=   "Column(12).Width=2725"
            Splits(0)._ColumnProps(56)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(57)=   "Column(12)._WidthInPix=2646"
            Splits(0)._ColumnProps(58)=   "Column(12)._ColStyle=8196"
            Splits(0)._ColumnProps(59)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(60)=   "Column(13).Width=2725"
            Splits(0)._ColumnProps(61)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(62)=   "Column(13)._WidthInPix=2646"
            Splits(0)._ColumnProps(63)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(64)=   "Column(14).Width=2725"
            Splits(0)._ColumnProps(65)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(66)=   "Column(14)._WidthInPix=2646"
            Splits(0)._ColumnProps(67)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(68)=   "Column(15).Width=2725"
            Splits(0)._ColumnProps(69)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(70)=   "Column(15)._WidthInPix=2646"
            Splits(0)._ColumnProps(71)=   "Column(15).Order=16"
            Splits(0)._ColumnProps(72)=   "Column(16).Width=2725"
            Splits(0)._ColumnProps(73)=   "Column(16).DividerColor=0"
            Splits(0)._ColumnProps(74)=   "Column(16)._WidthInPix=2646"
            Splits(0)._ColumnProps(75)=   "Column(16).Order=17"
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
            Caption         =   "JOB ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=102,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=99,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=100,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=101,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=98,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=95,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=96,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=97,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=94,.parent=13"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=91,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=92,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=93,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=90,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=87,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=88,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=89,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=86,.parent=13"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=83,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=84,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=85,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=82,.parent=13"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=79,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=80,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=81,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=70,.parent=13,.locked=-1"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=67,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=68,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=69,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13,.locked=-1"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=62,.parent=13,.locked=-1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=59,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=60,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=61,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.locked=-1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=54,.parent=13,.locked=-1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=51,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=52,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=53,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=50,.parent=13,.locked=-1"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=47,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=48,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=49,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=46,.parent=13,.locked=-1"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=43,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=44,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=45,.parent=17"
            _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=28,.parent=13"
            _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=25,.parent=14"
            _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=26,.parent=15"
            _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=27,.parent=17"
            _StyleDefs(92)  =   "Splits(0).Columns(14).Style:id=78,.parent=13"
            _StyleDefs(93)  =   "Splits(0).Columns(14).HeadingStyle:id=75,.parent=14"
            _StyleDefs(94)  =   "Splits(0).Columns(14).FooterStyle:id=76,.parent=15"
            _StyleDefs(95)  =   "Splits(0).Columns(14).EditorStyle:id=77,.parent=17"
            _StyleDefs(96)  =   "Splits(0).Columns(15).Style:id=74,.parent=13"
            _StyleDefs(97)  =   "Splits(0).Columns(15).HeadingStyle:id=71,.parent=14"
            _StyleDefs(98)  =   "Splits(0).Columns(15).FooterStyle:id=72,.parent=15"
            _StyleDefs(99)  =   "Splits(0).Columns(15).EditorStyle:id=73,.parent=17"
            _StyleDefs(100) =   "Splits(0).Columns(16).Style:id=32,.parent=13,.locked=0"
            _StyleDefs(101) =   "Splits(0).Columns(16).HeadingStyle:id=29,.parent=14"
            _StyleDefs(102) =   "Splits(0).Columns(16).FooterStyle:id=30,.parent=15"
            _StyleDefs(103) =   "Splits(0).Columns(16).EditorStyle:id=31,.parent=17"
            _StyleDefs(104) =   "Named:id=33:Normal"
            _StyleDefs(105) =   ":id=33,.parent=0"
            _StyleDefs(106) =   "Named:id=34:Heading"
            _StyleDefs(107) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(108) =   ":id=34,.wraptext=-1"
            _StyleDefs(109) =   "Named:id=35:Footing"
            _StyleDefs(110) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(111) =   "Named:id=36:Selected"
            _StyleDefs(112) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(113) =   "Named:id=37:Caption"
            _StyleDefs(114) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(115) =   "Named:id=38:HighlightRow"
            _StyleDefs(116) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(117) =   "Named:id=39:EvenRow"
            _StyleDefs(118) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(119) =   "Named:id=40:OddRow"
            _StyleDefs(120) =   ":id=40,.parent=33"
            _StyleDefs(121) =   "Named:id=41:RecordSelector"
            _StyleDefs(122) =   ":id=41,.parent=34"
            _StyleDefs(123) =   "Named:id=42:FilterBar"
            _StyleDefs(124) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid partGrid 
            Height          =   2055
            Left            =   0
            TabIndex        =   90
            Top             =   2640
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   3625
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
            Columns(3).NumberFormat=   "Standard"
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
            Columns(9)._VlistStyle=   4
            Columns(9)._MaxComboItems=   5
            Columns(9).DataField=   ""
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   4
            Columns(10)._MaxComboItems=   5
            Columns(10).DataField=   ""
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   4
            Columns(11)._MaxComboItems=   5
            Columns(11).DataField=   ""
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   12
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=12"
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
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(45)=   "Column(11).Width=2725"
            Splits(0)._ColumnProps(46)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(11)._WidthInPix=2646"
            Splits(0)._ColumnProps(48)=   "Column(11).Order=12"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "PART ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=900,.italic=0"
            _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=58,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=55,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=56,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=57,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=70,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=32,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=74,.parent=13"
            _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=71,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=72,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=73,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
            _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
            _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(86)  =   "Named:id=33:Normal"
            _StyleDefs(87)  =   ":id=33,.parent=0"
            _StyleDefs(88)  =   "Named:id=34:Heading"
            _StyleDefs(89)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(90)  =   ":id=34,.wraptext=-1"
            _StyleDefs(91)  =   "Named:id=35:Footing"
            _StyleDefs(92)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(93)  =   "Named:id=36:Selected"
            _StyleDefs(94)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(95)  =   "Named:id=37:Caption"
            _StyleDefs(96)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(97)  =   "Named:id=38:HighlightRow"
            _StyleDefs(98)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(99)  =   "Named:id=39:EvenRow"
            _StyleDefs(100) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(101) =   "Named:id=40:OddRow"
            _StyleDefs(102) =   ":id=40,.parent=33"
            _StyleDefs(103) =   "Named:id=41:RecordSelector"
            _StyleDefs(104) =   ":id=41,.parent=34"
            _StyleDefs(105) =   "Named:id=42:FilterBar"
            _StyleDefs(106) =   ":id=42,.parent=33"
         End
         Begin TrueOleDBGrid70.TDBGrid subletGrid 
            Height          =   1815
            Left            =   0
            TabIndex        =   91
            Top             =   4800
            Width           =   10335
            _ExtentX        =   18230
            _ExtentY        =   3201
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
            Columns(10)._VlistStyle=   4
            Columns(10)._MaxComboItems=   5
            Columns(10).DataField=   ""
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).RecordSelectorWidth=   688
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   14215660
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
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
            Splits(0)._ColumnProps(21)=   "Column(5).Width=2725"
            Splits(0)._ColumnProps(22)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(23)=   "Column(5)._WidthInPix=2646"
            Splits(0)._ColumnProps(24)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(25)=   "Column(6).Width=2725"
            Splits(0)._ColumnProps(26)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(27)=   "Column(6)._WidthInPix=2646"
            Splits(0)._ColumnProps(28)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(29)=   "Column(7).Width=2725"
            Splits(0)._ColumnProps(30)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(31)=   "Column(7)._WidthInPix=2646"
            Splits(0)._ColumnProps(32)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(33)=   "Column(8).Width=2725"
            Splits(0)._ColumnProps(34)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(8)._WidthInPix=2646"
            Splits(0)._ColumnProps(36)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(37)=   "Column(9).Width=2725"
            Splits(0)._ColumnProps(38)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(39)=   "Column(9)._WidthInPix=2646"
            Splits(0)._ColumnProps(40)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(41)=   "Column(10).Width=2725"
            Splits(0)._ColumnProps(42)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(43)=   "Column(10)._WidthInPix=2646"
            Splits(0)._ColumnProps(44)=   "Column(10).Order=11"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   3
            PrintInfos(0).Name=   "piInternal 0"
            PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
            PrintInfos(0).PageHeaderHeight=   0
            PrintInfos(0).PageFooterHeight=   0
            PrintInfos.Count=   1
            AllowDelete     =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1
            FootLines       =   1
            Caption         =   "SUBLET ORDER"
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12577002
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
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=208,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37,.bold=0,.fontsize=975,.italic=0"
            _StyleDefs(10)  =   ":id=4,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(11)  =   ":id=4,.fontname=Tahoma"
            _StyleDefs(12)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(13)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(14)  =   ":id=2,.fontname=Tahoma"
            _StyleDefs(15)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(16)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(17)  =   ":id=3,.fontname=Tahoma"
            _StyleDefs(18)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(19)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(20)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(21)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(22)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(23)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(24)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(25)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(26)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(27)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(28)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=70,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=67,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=68,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=69,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=66,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=63,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=64,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=65,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=62,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=59,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=60,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=61,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=58,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=54,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=51,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=52,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=53,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=50,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=46,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=43,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=44,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=45,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=32,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=29,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=30,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=31,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=28,.parent=13"
            _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=25,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=26,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=27,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=78,.parent=13"
            _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(82)  =   "Named:id=33:Normal"
            _StyleDefs(83)  =   ":id=33,.parent=0"
            _StyleDefs(84)  =   "Named:id=34:Heading"
            _StyleDefs(85)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(86)  =   ":id=34,.wraptext=-1"
            _StyleDefs(87)  =   "Named:id=35:Footing"
            _StyleDefs(88)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(89)  =   "Named:id=36:Selected"
            _StyleDefs(90)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(91)  =   "Named:id=37:Caption"
            _StyleDefs(92)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(93)  =   "Named:id=38:HighlightRow"
            _StyleDefs(94)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(95)  =   "Named:id=39:EvenRow"
            _StyleDefs(96)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(97)  =   "Named:id=40:OddRow"
            _StyleDefs(98)  =   ":id=40,.parent=33"
            _StyleDefs(99)  =   "Named:id=41:RecordSelector"
            _StyleDefs(100) =   ":id=41,.parent=34"
            _StyleDefs(101) =   "Named:id=42:FilterBar"
            _StyleDefs(102) =   ":id=42,.parent=33"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   405
            Index           =   29
            Left            =   10365
            TabIndex        =   92
            Top             =   6510
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   714
            _Version        =   393216
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox text 
            Height          =   285
            Index           =   30
            Left            =   1440
            TabIndex        =   101
            Top             =   6680
            Width           =   7515
            _ExtentX        =   13256
            _ExtentY        =   503
            _Version        =   393216
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
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Catatan"
            Height          =   195
            Index           =   0
            Left            =   720
            TabIndex        =   102
            Top             =   6700
            Width           =   555
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "__________________    +"
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
            Left            =   10320
            TabIndex        =   95
            Top             =   6240
            Width           =   2040
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Grand Total"
            Height          =   195
            Left            =   9360
            TabIndex        =   94
            Top             =   6600
            Width           =   840
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Terbilang (Rupiah)"
            Height          =   195
            Left            =   60
            TabIndex        =   93
            Top             =   7080
            Width           =   1305
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   480
      TabIndex        =   44
      Top             =   2400
      Width           =   12885
      Begin VB.ComboBox cboPelanggan 
         Height          =   315
         ItemData        =   "frmOrderValidasi.frx":2FDA
         Left            =   8880
         List            =   "frmOrderValidasi.frx":2FDC
         Style           =   2  'Dropdown List
         TabIndex        =   114
         Top             =   800
         Width           =   735
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Regular"
         Enabled         =   0   'False
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
         Left            =   840
         TabIndex        =   47
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Program"
         Enabled         =   0   'False
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
         Left            =   1800
         TabIndex        =   46
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opWO 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warranty"
         Enabled         =   0   'False
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
         Left            =   2880
         TabIndex        =   45
         Top             =   240
         Width           =   1095
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   20
         Left            =   4920
         TabIndex        =   48
         Top             =   800
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   19
         Left            =   6795
         TabIndex        =   49
         Top             =   165
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   35
         Left            =   120
         TabIndex        =   50
         Top             =   525
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   36
         Left            =   120
         TabIndex        =   51
         Top             =   810
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   37
         Left            =   4920
         TabIndex        =   52
         Top             =   170
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   38
         Left            =   6795
         TabIndex        =   53
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
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
         Mask            =   ">CCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   31
         Left            =   6795
         TabIndex        =   54
         Top             =   795
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "CCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   43
         Left            =   4920
         TabIndex        =   55
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "99-99-9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   28
         Left            =   8880
         TabIndex        =   108
         Top             =   165
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   32
         Left            =   8880
         TabIndex        =   109
         Top             =   480
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   32
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   ">CCCCCCCCCCCCCCCCCCCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   270
         Index           =   44
         Left            =   9600
         TabIndex        =   115
         Top             =   825
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   476
         _Version        =   393216
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
      Begin VB.Label JualNRM 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   12360
         TabIndex        =   140
         Top             =   600
         Width           =   60
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Karakter"
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
         Left            =   8160
         TabIndex        =   116
         Top             =   860
         Width           =   615
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Phone"
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
         Left            =   8025
         TabIndex        =   111
         Top             =   525
         Width           =   750
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pembawa"
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
         Left            =   8085
         TabIndex        =   110
         Top             =   210
         Width           =   690
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inv. No."
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
         Left            =   4200
         TabIndex        =   64
         Top             =   845
         Width           =   600
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "T SA."
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
         Left            =   6300
         TabIndex        =   63
         Top             =   210
         Width           =   390
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe WO"
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
         TabIndex        =   62
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM"
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
         Left            =   4560
         TabIndex        =   61
         Top             =   215
         Width           =   210
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mech"
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
         Left            =   6315
         TabIndex        =   60
         Top             =   525
         Width           =   375
      End
      Begin VB.Label lblPDS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PDS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   12120
         TabIndex        =   59
         Top             =   120
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Book"
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
         Left            =   6345
         TabIndex        =   58
         Top             =   840
         Width           =   345
      End
      Begin VB.Label Label19 
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
         Left            =   4200
         TabIndex        =   57
         Top             =   525
         Width           =   570
      End
      Begin VB.Label LblReturn 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RTJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   12120
         TabIndex        =   56
         Top             =   360
         Visible         =   0   'False
         Width           =   465
      End
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   2600
      TabIndex        =   97
      Top             =   120
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Enabled         =   0   'False
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   1
      Left            =   3720
      TabIndex        =   98
      Top             =   90
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">99-99-9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   375
      Index           =   33
      Left            =   2760
      TabIndex        =   144
      Top             =   120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   ">CCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.Label lblNilaiMax 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3840
      TabIndex        =   260
      Top             =   9240
      Width           =   1575
   End
   Begin VB.Label lblFOP 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   480
      TabIndex        =   259
      Top             =   9240
      Width           =   2895
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9495
      Left            =   0
      Picture         =   "frmOrderValidasi.frx":2FDE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RETURN JOB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   5280
      TabIndex        =   103
      Top             =   82
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. WO"
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
      Left            =   480
      TabIndex        =   100
      Top             =   120
      Width           =   570
   End
   Begin VB.Label Label26 
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
      Left            =   3015
      TabIndex        =   99
      Top             =   135
      Width           =   570
   End
End
Attribute VB_Name = "frmOrderValidasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsTemp As ADODB.Recordset
Dim rsKTP  As ADODB.Recordset
Dim rsPart As New ADODB.Recordset
Dim rsJob As New ADODB.Recordset
Dim rsSublet As New ADODB.Recordset
Dim objBilling As New BClsBilling
Dim newValid As Boolean
Dim gSpart As Boolean
Dim adaSublet As Boolean
Dim simpan As Boolean
Dim noInv As Boolean
Dim totRate As Double
Dim nWKT As String
Dim support As New clsFormDealer
Dim totalNilaiPart As Double
''2017-08-01 || Rev20170801A - penambahan request disc menampilkan data wo walaupun belum validasi tapi button disable
Dim isBolehBatalReqDisc As Boolean
Dim rsCetakNewWO As New ADODB.Recordset
Dim cols As TrueOleDBGrid70.Columns
Dim nDuoInv As Boolean

Private TxtAlamat       As String


Private tagihanFOP          As String
Private kodeMDL             As String
Private nilaiMaxPartNRM     As Double
Private partTerpakaiNRM     As Double

Private nilaiDiscMax As Integer
Private nilaiLockDisc As Integer
Private Declare Sub Sleep Lib "kernel32.dll" _
(ByVal dwMilliseconds As Long)
Private nPersenAR       As Double

Private Sub doCreatePart()
1            On Error GoTo doCreatePart_Error
2             If rsPart.State Then rsPart.Close
              'Set rsPart = Nothing
3             With rsPart
          ' If .State Then .Close
4                 .Fields.Append "KODEITEM", adVarChar, 50
5                 .Fields.Append "NAMAITEM", adVarChar, 200
6                 .Fields.Append "QTYOrder", adInteger
7                 .Fields.Append "QTY", adInteger
8                 .Fields.Append "QTYBO", adInteger
9                 .Fields.Append "QTYCANCEL", adInteger
10                .Fields.Append "HARGA", adDouble
11                .Fields.Append "OLI", adBoolean
12                .Fields.Append "EXTDISC", adDouble
13                .Fields.Append "TERLIBAT", adBoolean
14                .Fields.Append "FREEC", adBoolean
15                .Fields.Append "CROSS", adBoolean
16                .Fields.Append "UP", adBoolean
17                .Open
18            End With
          
19           On Error GoTo 0
20           Exit Sub
          
doCreatePart_Error:
          
21            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreatePart of Form frmOrderValidasi")
End Sub

Private Sub doHeadPart()
       On Error GoTo doHeadPart_Error
    
        With partGrid
            .Columns(0).Width = 1500
            .Columns(0).Caption = "KODE PART"
            .Columns(0).DataField = "KODEITEM"
            .Splits(0).Columns(0).Locked = True
    
            .Columns(1).Width = 2200
            .Columns(1).Caption = "NAMA PART"
            .Columns(1).DataField = "NAMAITEM"
            .Splits(0).Columns(1).Locked = True
    
            .Columns(2).Width = 1000
            .Columns(2).Caption = "QTY Order"
            .Columns(2).DataField = "QTYOrder"
            .Splits(0).Columns(2).Locked = True
    
            .Columns(3).Width = 1000
            .Columns(3).Caption = "QTY Supply"
            .Columns(3).DataField = "QTY"
            .Splits(0).Columns(3).Locked = True
    
            .Columns(4).Width = 1000
            .Columns(4).Caption = "QTY BO"
            .Columns(4).DataField = "QTYBO"
            .Splits(0).Columns(3).Locked = True
            On Error GoTo Column_Ini_Untuk_Grid_FillRate
    
    
            .Columns(5).Visible = True
            .Columns(5).Width = 1000
            .Columns(5).Caption = "HARGA"
            .Columns(5).DataField = "Harga"
            .Splits(0).Columns(5).Locked = True
    
            .Columns(6).Visible = False
            .Columns(6).Caption = "OLI"
            .Columns(6).DataField = "OLI"
            .Splits(0).Columns(6).Locked = True
    
            .Columns(7).Width = 1000
            .Columns(7).Caption = "QTY CANCEL"
            .Columns(7).DataField = "QTYCANCEL"
            .Splits(0).Columns(7).Locked = True
            
            .Columns(8).Width = 500
            .Columns(8).Caption = "Terlibat"
            .Columns(8).DataField = "TERLIBAT"
            .Columns(8).Alignment = dbgCenter
            .Splits(0).Columns(8).Locked = False
            .Columns(8).Visible = False
            
            .Columns(9).Width = 400
            .Columns(9).Caption = "Cust"
            .Columns(9).DataField = "FREEC"
            .Columns(9).Alignment = dbgCenter
            .Splits(0).Columns(9).Locked = False
            .Columns(9).Visible = False
            If tagihanFOP <> "" And nDuoInv Then
                .Columns(9).Visible = True
            End If
            
            
            .Columns(10).Width = 400
            .Columns(10).Caption = "Cross"
            .Columns(10).DataField = "Cross"
            .Columns(10).Alignment = dbgCenter
            .Splits(0).Columns(10).Locked = False
            .Columns(10).Visible = False
            If Log.DIDX = 1 Then
                .Columns(10).Visible = True
            End If
            
            
            .Columns(11).Width = 400
            .Columns(11).Caption = "Up"
            .Columns(11).DataField = "Up"
            .Columns(11).Alignment = dbgCenter
            .Splits(0).Columns(11).Locked = False
            .Columns(11).Visible = False
            If Log.DIDX = 1 Then
                .Columns(11).Visible = True
            End If
Column_Ini_Untuk_Grid_FillRate:
        End With
    
       On Error GoTo 0
       Exit Sub
    
doHeadPart_Error:
    
        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadPart of Form frmOrderValidasi")
End Sub

Private Sub doCreateJob()
1            On Error GoTo doCreateJob_Error
2             If rsJob.State Then rsJob.Close
          ' Set rsJob = Nothing
3             With rsJob
          'If .State Then .Close
4                 .Fields.Append "JOB", adVarChar, 50
5                 .Fields.Append "RATE", adDouble
6                 .Fields.Append "ACTRT", adDouble
7                 .Fields.Append "GRP", adVarChar, 6
8                 .Fields.Append "MECH", adVarChar, 10
9                 .Fields.Append "ADDJOB", adBoolean
10                .Fields.Append "HRG", adDouble
11                .Fields.Append "DISC", adDouble
12                .Fields.Append "RP", adBoolean
13                .Fields.Append "DISCRP", adDouble
14                .Fields.Append "STLDISC", adDouble
15                .Fields.Append "VALID", adVarChar, 10
16                .Fields.Append "FREEC", adBoolean           ' TANDA JIKA DIBEBANKAN KE CUSTOMER
17                If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(43), "DD/MM/YYYY")) >= 0 Then
18                    .Fields.Append "IDPROGRAM", adVarChar, 20
19                    .Fields.Append "PROGRAM", adVarChar, 50
20                    .Fields.Append "DISCMAX", adDouble
21                    .Fields.Append "ISFP", adBoolean
22                    .Fields.Append "PRICE", adDouble
23                End If
24                .Open
25            End With
          
26           On Error GoTo 0
27           Exit Sub
          
doCreateJob_Error:
          
28            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doCreateJob of Form frmOrderValidasi")
End Sub

Private Sub doHeadJob()
1         On Error GoTo doHeadJob_Error
          Dim iTambah As Integer
2         iTambah = 0
3         With JobGrid
4         If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(43), "DD/MM/YYYY")) >= 0 Then
5             isTambah = 1
6             .Columns(2).Width = 2000
7             .Columns(2).Caption = "PROGRAM"
8             .Columns(2).DataField = "PROGRAM"
9             .Columns(2).Visible = True
10            .Splits(0).Columns(2).Locked = True
              
11            .Columns(12).Width = 0
12            .Columns(12).Caption = "ACTRATE"
13            .Columns(12).DataField = "ACTRT"
14            .Columns(12).Visible = False
15            .Splits(0).Columns(12).Locked = True
              
              
16            .Columns(13).Width = 0
17            .Columns(13).Caption = "IDPROGRRAM"
18            .Columns(13).DataField = "IDPROGRRAM"
19            .Columns(13).Visible = False
20            .Splits(0).Columns(13).Locked = True
              
21            .Columns(14).Width = 0
22            .Columns(14).Caption = "DISCMAX"
23            .Columns(14).DataField = "DISCMAX"
24            .Columns(14).Visible = False
25            .Splits(0).Columns(14).Locked = True
              
26            .Columns(15).Width = 0
27            .Columns(15).Caption = "ISFP"
28            .Columns(15).DataField = "ISFP"
29            .Columns(15).Visible = False
30            .Splits(0).Columns(15).Locked = True
              
31            .Columns(16).Width = 0
32            .Columns(16).Caption = "PRICE"
33            .Columns(16).DataField = "PRICE"
34            .Columns(16).Visible = False
35            .Splits(0).Columns(16).Locked = True
36        Else
              
37            .Columns(2).Width = 0
38            .Columns(2).Caption = "ACTRATE"
39            .Columns(2).DataField = "ACTRT"
40            .Columns(2).Visible = False
41            .Splits(0).Columns(2).Locked = True
                      
42            .Columns(12).Visible = False
43            .Columns(13).Visible = False
44            .Columns(14).Visible = False
45            .Columns(15).Visible = False
46            .Columns(16).Visible = False
47        End If
          
48            .Columns(0).Width = 3000
49            .Columns(0).Caption = "JOB ORDER"
50            .Columns(0).DataField = "JOB"
51            .Splits(0).Columns(0).Locked = True
              
52            .Columns(1).Width = 450
53            .Columns(1).Caption = "RATE"
54            .Columns(1).DataField = "RATE"
55            .Splits(0).Columns(1).Locked = True
              
              
56            .Columns(3).Width = 0
57            .Columns(3).Caption = "GROUP"
58            .Columns(3).DataField = "GRP"
59            .Columns(3).Visible = False
60            .Splits(0).Columns(3).Locked = True
              
61            .Columns(4).Width = 0
62            .Columns(4).Caption = "MEKANIK"
63            .Columns(4).DataField = "MECH"
64            .Columns(4).Visible = False
65            .Splits(0).Columns(4).Locked = True
              
66            .Columns(5).Width = 0
67            .Columns(5).Caption = "Additional Job"
68            .Columns(5).DataField = "ADDJOB"
69            .Columns(5).Visible = False
70            .Splits(0).Columns(5).Locked = True
              
71            .Columns(6).Width = 1000
72            .Columns(6).Caption = "HARGA"
73            .Columns(6).DataField = "HRG"
74            .Splits(0).Columns(6).Locked = True
              
75            .Columns(7).Width = 500
76            .Columns(7).Caption = "DISC"
77            .Columns(7).DataField = "DISC"
              .Splits(0).Columns(7).Locked = False
              
78            .Columns(8).Width = 300
79            .Columns(8).Caption = "RP"
80            .Columns(8).DataField = "RP"
81            .Columns(8).Alignment = dbgCenter
82            If Log.discKacab = 1 Then
83                .Splits(0).Columns(8).Locked = False
84            Else
85                .Splits(0).Columns(8).Locked = True
86            End If
              
              
87            .Columns(9).Width = 1000
88            .Columns(9).Caption = "DISC RUPIAH"
89            .Columns(9).DataField = "DISCRP"
              
90            .Columns(10).Width = 1200
91            .Columns(10).Caption = "SUB TOTAL"
92            .Columns(10).DataField = "STLDISC"
93            .Splits(0).Columns(10).Locked = True
              
94            .Columns(11).Width = 500
95            .Columns(11).Caption = "Cust"
96            .Columns(11).DataField = "FREEC"
97            .Columns(11).Alignment = dbgCenter
98            .Splits(0).Columns(11).Locked = False
99            .Columns(11).Visible = False
100           If nDuoInv Then
101               .Columns(11).Visible = True
102           End If
103       End With
          
104       On Error GoTo 0
105       Exit Sub
          
doHeadJob_Error:
          
106       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadJob of Form frmOrderValidasi")
End Sub

Private Sub doHeadSublet()
    
       On Error GoTo doHeadSublet_Error
    
        With subletGrid
            .Columns(0).Width = 1000
            .Columns(0).Caption = "NO NOTA"
            .Columns(0).DataField = "NONOTA"
            .Splits(0).Columns(0).Locked = True
    
            .Columns(1).Width = 1000
            .Columns(1).Caption = "SUPPLIER"
            .Columns(1).DataField = "SUPPLIER"
            .Columns(1).Visible = True
            .Splits(0).Columns(1).Locked = True
    
            .Columns(2).Width = 1700
            .Columns(2).Caption = "NAMA BARANG"
            .Columns(2).DataField = "NAMABARANG"
            .Splits(0).Columns(2).Locked = True
    
            .Columns(3).Width = 800
            .Columns(3).Caption = "SATUAN"
            .Columns(3).DataField = "KODESATUAN"
            .Splits(0).Columns(3).Locked = True
    
            .Columns(4).Width = 600
            .Columns(4).Caption = "JUMLAH"
            .Columns(4).DataField = "QTY"
            .Splits(0).Columns(4).Locked = True
    
            .Columns(5).Width = 600
            .Columns(5).Caption = "VALID"
            .Columns(5).DataField = "VALID"
            .Columns(5).Visible = False
            .Splits(0).Columns(5).Locked = True
    
            .Columns(6).Width = 1200
            .Columns(6).Caption = "BELI"
            .Columns(6).DataField = "HARGA"
            .Columns(6).NumberFormat = "#,#0"
            .Splits(0).Columns(6).Locked = True
    
            .Columns(7).Width = 1200
            .Columns(7).Caption = "JUAL"
            .Columns(7).DataField = "PRICE"
            .Columns(7).NumberFormat = "#,#0"
            .Splits(0).Columns(7).Locked = True
            
            
            .Columns(8).Width = 500
            .Columns(8).Caption = "DISC"
            .Columns(8).DataField = "diskon"
            ''.Columns(8).NumberFormat = "#,#0"
            .Splits(0).Columns(8).Locked = False
            If Log.standarisasiSublet = 1 Then
                .Columns(8).Visible = True
            Else
                .Columns(8).Visible = False
            End If
            
            .Columns(9).Width = 1200
            .Columns(9).Caption = "SUB TOTAL"
            .Columns(9).DataField = "subtotal"
            .Splits(0).Columns(9).Locked = True
            .Columns(9).NumberFormat = "#,#0"
            If Log.standarisasiSublet = 1 Then
                .Columns(9).Visible = True
            Else
                .Columns(9).Visible = False
            End If
            
            .Columns(10).Width = 500
            .Columns(10).Caption = "Cust"
            .Columns(10).DataField = "FREEC"
            .Columns(10).Alignment = dbgCenter
            .Splits(0).Columns(10).Locked = False
            .Columns(10).Visible = False
            If tagihanFOP <> "" And nDuoInv Then
                .Columns(10).Visible = True
            End If
        End With
    
       On Error GoTo 0
       Exit Sub
    
doHeadSublet_Error:
    
        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure doHeadSublet of Form frmOrderValidasi")
End Sub

Private Function returnJob(labelReturn As Object) As Boolean
1            On Error GoTo returnJob_Error
    
2             returnJob = False
3             If labelReturn.Visible Then
4                 returnJob = True
5             End If
    
6            On Error GoTo 0
7            Exit Function
    
returnJob_Error:
    
8             MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure returnJob of Form frmOrderValidasi")
End Function

Private Function typeJob(statusJob As String) As String
1            On Error GoTo typeJob_Error
    
2             If statusJob = "W" Then
3                 typeJob = "WARRANTY"
4             ElseIf statusJob = "P" Then
5                 typeJob = "PROGRAM"
6             Else
7                 typeJob = "REGULAR"
8             End If
    
9            On Error GoTo 0
10           Exit Function
    
typeJob_Error:
    
11            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure typeJob of Form frmOrderValidasi")
End Function

Private Function giveDiscount(discountValue As String, discountText As Object, statusText As Boolean)
1            On Error GoTo giveDiscount_Error
    
2             discountText = discountValue
3             discountText.Enabled = statusText
4             JobGrid.Splits(0).Columns(7).Locked = Not statusText
5             JobGrid.Splits(0).Columns(9).Locked = Not statusText
    
6            On Error GoTo 0
7            Exit Function
    
giveDiscount_Error:
    
8             MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure giveDiscount of Form frmOrderValidasi")
End Function

'---------------------------------------------------------------------------------------
' Procedure : panggilWO
' DateTime  : 30/09/2019 10:56
' Author    : Rudi
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub getQuerySingleid()
    If NulSID(objO.FetchData_String("SELECT TypeCustomer FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & Nul(Text(2)) & "' ")) <> "PERORANGAN" Then
        sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & Nul(Text(2)) & "' "
        sSQL = sSQL & " and ((isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '' or isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '1900-01-01' or isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '3000-01-01') "

        sSQL = sSQL & " or ISNULL(Agama,'') = '' or ISNULL(Pekerjaan,'') = '' or ISNULL(hobi,'') = '' or ISNULL(Olahraga,'') = '' or ISNULL(Propinsi,'') = ''"
        sSQL = sSQL & " or ISNULL(kota,'') = '' or ISNULL(Kecamatan,'') = '' or ISNULL(Kelurahan,'') = '' or ISNULL(Alamat,'') = '' or ISNULL(email,'') = ''"
        sSQL = sSQL & " or ISNULL(telp+HP1+hp2,'') = '')"
    Else
        sSQL = "SELECT * FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & Nul(Text(2)) & "' "
        sSQL = sSQL & " and (isnull(replace(NoKTP,'00000000000000',''),'') = '' or (isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '' or isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '1900-01-01' or isnull(REPLACE(TglLahir,'01/01/1900',''),'') = '3000-01-01') "

        sSQL = sSQL & " or ISNULL(Agama,'') = '' or ISNULL(Pekerjaan,'') = '' or ISNULL(hobi,'') = '' or ISNULL(Olahraga,'') = '' or ISNULL(Propinsi,'') = ''"
        sSQL = sSQL & " or ISNULL(kota,'') = '' or ISNULL(Kecamatan,'') = '' or ISNULL(Kelurahan,'') = '' or ISNULL(Alamat,'') = '' or ISNULL(email,'') = ''"
        sSQL = sSQL & " or ISNULL(telp+HP1+hp2,'') = '')"
    End If
End Sub
'---------------------------------------------------------------------------------------
' Procedure : panggilWO
' DateTime  : 23/07/2026 13:23
' Author    : HO-VM-RUDI
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub panggilWO(noWO As String)
          Dim rsHlp As ADODB.Recordset
          Dim rsHlp2 As ADODB.Recordset
          Dim rsHlp3 As ADODB.Recordset
          Dim hlp As String
          Dim hlp2 As String
          Dim hlp3 As String
          Dim hlp4 As String
          Dim hlp5 As String
          Dim fOth As Boolean
          'Dim vHitung As Double
          Dim nHarga As Double
          Dim isDiscKacab     As Boolean
          
1         On Error GoTo panggilWO_Error
          
2         tagihanFOP = ""
3         kodeMDL = ""
       
4         gKdPlgDealer = ""
5         gKodePlg = ""
6         gInvCust = ""
7         gInvAlmt = ""
8         gInvKota = ""
9         gInvNPWP = ""
10        nDuoInv = False
11        nilaiDiscMax = Log.discJasa
          '' -------------------------------------------------------------------------
          '' 20190910 req mb thania penambahan nilai lock disc untuk jasa customer
          '' -------------------------------------------------------------------------
12        nilaiLockDisc = Log.discJasa
13        isDiscKacab = False
          
14        Set rsTemp = objBilling.searchOrderDetail(noWO)
15        If Not rsTemp.EOF Then
16            If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(rsTemp!Date, "DD/MM/YYYY")) >= 0 Then
17                opWO(1).Visible = False
18            Else
19                opWO(1).Visible = True
20            End If
21            Text1.Text = Trim(rsTemp!CATAT)
              ' TIPE PEMBAYARAN
22            Set rsHlp = objBilling.searchPaymentType(noWO)
23            If Not rsHlp.EOF Then
24                tipeByr = Trim(rsHlp("PEMBAYARAN"))
25                gVerifikator = Trim(rsHlp("VERIFIKATOR"))
26                cboPelanggan = Trim(rsHlp("KARAKTER"))
27                Text(44) = Trim(rsHlp("DTLKARAKTER"))
28            End If
              
29            Call cekPSMInfo(Nul(rsTemp("NO_CHASSIS")))
          
              ''20171127 CEK UNITCUSTOMER DAN NEWCUSTOMER UNTUK NORANGKA
              ''----------------------------------------------------------------
30            If cekDataSingleID("unitcustomer", "where norangka = '" & Nul(rsTemp("NO_CHASSIS")) & "'") = True Then
                  ''jika diunit customer tampilkan idcust dan kodelgn lama
                  ''cek data bridging atas kodelgn lama sudah ada belum
                  '' -----------------------------------------------------------------------
31                If cekDataSingleID("bridging", "where kodelgn = '" & Nul(rsTemp("FS_KD_PLG")) & "' and kodedealer = '" & Right(gDb, 2) & "'") = True Then
                      ''jika di bridging ada maka di ambil data idcust
32                    Set rsHlp = objO.FetchData("SELECT idcust FROM [SingleIDOri]..bridging WITH (NOLOCK) WHERE kodelgn = '" & Nul(rsTemp("FS_KD_PLG")) & "' and kodedealer = '" & Right(gDb, 2) & "'")
33                    If Not rsHlp.EOF Then
34                        Text(2) = Nul(rsHlp!idcust)
35                        txtKodeLgnLama.Text = Nul(rsTemp("FS_KD_PLG"))
                          
                          '' 20171211
                          '' PENAMBAHAN PENGECEKAN NO KTP TIDAK KOSONG DAN
                          '' ---------------------------------------------------------
36                        Set rsHlp = objO.FetchData("SELECT PinBB,isnull(kewarganegaraan,'') as kewarganegaraan,NoKTP,UPPER(TypeCustomer) AS TypeCustomer FROM [SINGLEIDORI]..NEWCUSTOMER WITH (NOLOCK) WHERE IDCUST = '" & Nul(Text(2)) & "'")
37                        If Not rsHlp.EOF Then
38                            If NulSID(rsHlp!kewarganegaraan) = "WNA" And NulSID(rsHlp!TypeCustomer) = "PERORANGAN" Then
39                                If Nul(rsHlp!noKTP) = "" Then
40                                    If Log.mandatoryKTP = 1 Then
41                                        MsgBox "Silahkan lengkapi No KTP / Passport / KITAS di master pelanggan", vbInformation
42                                        TblValid.Enabled = False
43                                    End If
44                                End If
45                            ElseIf NulSID(rsHlp!kewarganegaraan) <> "WNA" And NulSID(rsHlp!TypeCustomer) = "PERORANGAN" Then
46                                If (Nul(rsHlp!noKTP) = "" Or Val(rsHlp!noKTP) = "0") And NulSID(rsHlp!TypeCustomer) = "PERORANGAN" Then
47                                    If Log.mandatoryKTP = 1 Then
                                          
48                                        MsgBox "Silahkan lengkapi No KTP di master pelanggan", vbInformation
49                                        TblValid.Enabled = False
50                                    End If
51                                End If
52                            End If
          
53                            If NulSID(rsHlp!PinBB) = "SCORING" Then
54                                If Log.singleIDScoring = 1 And NulSID(rsHlp!TypeCustomer) = "PERORANGAN" Then
55                                    Set rsHlp2 = objO.FetchData("SELECT * FROM [singleidori]..UploadKTP WHERE IDCust = '" & Nul(Text(2)) & "'")
56                                    If rsHlp2.EOF Then
57                                        MsgBox "Silahkan lengkapi Upload Susulan No KTP di master pelanggan", vbInformation
58                                        TblValid.Enabled = False
59                                    End If
60                                End If
61                            End If
                              
62                            If Log.monitoringScoring = 1 Then
                                  
63                                Call getQuerySingleid
64                                Set rsHlp2 = objO.FetchData(sSQL)
65                                If Not rsHlp2.EOF Then
66                                    MsgBox "Silahkan lengkapi data mandatori dimaster customer terlebih dahulu untuk melanjutkan proses", vbInformation
67                                    With frmSIDMasterPelanggan
68                                        .txt(1) = NulSID(Text(2))
69                                        .isEdit (True)
70                                        .show vbModal
71                                    End With
72                                    Exit Sub
73                                End If
74                            End If
75                        Else
76                            MsgBox "ID tidak terdaftar di new customer", vbInformation
77                            Call cmdKodePlg_Click
78                        End If
          
                  '' - AKHIR PENGECEKAN NO KTP -------------------------------
          
79                    Else
80                        Text(2) = ""
81                        MsgBox "Gagal mengambil data ID Customer", vbInformation
82                        Exit Sub
83                    End If
84                Else
85                    Call cmdKodePlg_Click
86                End If
87            Else
              ''jika unit tidak ada diarahkan ke form customer single id
88                Call cmdKodePlg_Click
89            End If
              ''----------------------------------------------------------------
              
90            If rsHlp.State <> adStateClosed Then rsHlp.Close
91            Set rsHlp = objBilling.searchDtlOrder(rsTemp("NO_WO"), rsTemp("NO_CHASSIS"))
92            If Not rsHlp.EOF Then
93              nHeavy = IIf(rsHlp("HEAVYUSE") <> 0, True, False)
94            End If
95            If rsHlp.State <> adStateClosed Then rsHlp.Close
          
          
              ' TAMPIL DETAIL CUSTOMER
96            txtKodeLgnLama.Text = Nul(rsTemp("FS_KD_PLG"))
              'text(2) = rsTemp("FS_KD_PLG")
97            Text(3) = rsTemp("FS_KD_WIL")
98            Text(4) = rsTemp("CUSTOMER")
99            TxtAlamat = NulSID(rsTemp!cust_adres)
100           Text(5) = Left(rsTemp("CUST_ADRES"), 60)
101           Text(6) = rsTemp("KOTA")
102           Text(7) = rsTemp("TELPON")
103           Text(8) = rsTemp("GSM")
104           Text(9) = rsTemp("CDMA")
105           Text(10) = rsTemp("FS_TH_MOBIL")
106           Text(11).Mask = ""
107           Text(11) = Nul(IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-3000", "", IIf(Replace(rsTemp("DOTAM"), "/", "-") = "01-01-1900", "", Replace(rsTemp("DOTAM"), "/", "-"))))
108           Text(12) = rsTemp("MDL")
109           Text(13) = rsTemp("MODEL")
110           Text(14) = rsTemp("FMDL")
111           Text(15) = rsTemp("COLOR")
112           Text(16) = rsTemp("NO_POLISI")
113           Text(17) = rsTemp("NO_CHASSIS")
114           Text(18) = rsTemp("NO_ENGINE")
115           Text(19) = rsTemp("FS_LOG_ADD")
116           If rsTemp("NOINV") <> "" Then
117             Text(20) = rsTemp("NOINV")
118           End If
119           Text(31) = rsTemp("NOFAK")
              'text(33) = rsTemp("FS_RC")
120           Text(37) = rsTemp("KM")
121           Text(38) = rsTemp("MECH")
122           Text(40) = rsTemp("STATUS")
123           Text(43) = TglDDMMYYYY(rsTemp("DATE"))
124           nTHS = IIf(rsTemp("THS") = 1, True, False)
125           LblReturn.Visible = IIf(rsTemp("FN_RETURN") = "1", True, False)
126           lblPDS.Visible = IIf(rsTemp("PDS") = "1", True, False)
127           Text(28) = rsTemp("ANCUST")
128           Text(32) = rsTemp("TELPANCUST")
              
              '' 20200528 service berkala
              '' ---------------------------------------------
129           If Log.freepart = 1 And NulSID(NulSID(Text(17))) <> "" And (rsTemp("FS_RC") <> "" And rsTemp("FS_RC") <> "R001") Then
130                tagihanFOP = ""
131                kodeMDL = objO.FetchData_String("select Mdl from SingleIDOri..UnitCustomer where norangka = '" & NulSID(NulSID(Text(17))) & "'")
132                If CekUnitTerlibatFOP(NulSID(Text(17)), kodeMDL, tagihanFOP) Then
133                    lblFOP.Caption = "Free Part + Oli : " & tagihanFOP
134                Else
135                    lblFOP.Caption = ""
136                End If
137                If tagihanFOP = "NRM" Then
                      ''ambil nilai max
                      Dim kolomSBE As String
138                   If rsTemp("FS_RC") = "R010" Then
139                       kolomSBE = "sb10k"
140                   ElseIf rsTemp("FS_RC") = "R020" Then
141                       kolomSBE = "sb20k"
142                   ElseIf rsTemp("FS_RC") = "R030" Then
143                       kolomSBE = "sb30k"
144                   ElseIf rsTemp("FS_RC") = "R040" Then
145                       kolomSBE = "sb40k"
146                   ElseIf rsTemp("FS_RC") = "R050" Then
147                       kolomSBE = "sb50k"
148                   Else
149                       kolomSBE = "deposit"
150                   End If
151                   nilaiMaxPartNRM = objO.FetchData_Double("select " & kolomSBE & " from [" & dbReplikasiHO & "]..buklet_sales where NO_RANGKA = '" & NulSID(NulSID(Text(17))) & "'")
152                   partTerpakaiNRM = 0
153                   lblNilaiMax.Caption = Format(nilaiMaxPartNRM, "#,#0")
154                Else
155                   lblNilaiMax.Caption = ""
156                End If
157           End If
              ' CHECK DEFAULT DISCOUNT
158           Text(30) = ""
159           txtCatatan.Text = ""
160           Text(35).Visible = False
161           Text(36).Visible = False
162           vTipeCustomer = ""
163           myChassis = Trim(Text(17))
164           If returnJob(LblReturn) Then
165               If typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
166                   opWO(1).value = True
                  
167                   Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
168                   If Not rsHlp.EOF Then
169                       Text(35) = rsHlp("ID_PROGRAM")
170                       Text(36) = rsHlp("PROGRAM")
171                       Text(35).Visible = True
172                       Text(36).Visible = True
173                   End If
              
174                   If rsHlp.State <> adStateClosed Then rsHlp.Close
                      
175                   If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
176                       nDuoInv = True
                          'RTJ WO BERKEMBANG
                  
177                       Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
178                       If Not rsHlp.EOF Then
179                           gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
180                           gKodePlg2 = Trim(rsHlp("KDPLG"))
181                           gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
182                           gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
183                           gInvKota2 = Trim(rsHlp("FS_KOTA"))
184                           gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                  
185                           If Trim(rsHlp("INTERNAL")) = 1 Then
186                               If Trim(gNpwp) = Trim(gInvNPWP2) Then
187                                   vTipeCustomer = "AFFNONPPN"
188                               Else
189                                   vTipeCustomer = "AFFPPN"
190                               End If
191                           Else
192                               vTipeCustomer = "EXT"
193                           End If
194                       End If
195                       If rsHlp.State <> adStateClosed Then rsHlp.Close
                          ''---------------------------------------------------
                          ''20180301 validasi order defaultnya 0 (min)
                          ''---------------------------------------------------'
196                       If Log.discKacab = 1 Then
197                           Call giveDiscount("0", Text(21), False)
198                       Else
199                           Call giveDiscount("100", Text(21), False)
200                       End If
                          '
                          'text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
201                       txtCatatan.Text = "A/N " & Text(4) & "  [" & Text(2) & "]"
202                   End If
203               Else
                      ' HANYA RTJ
                      ''---------------------------------------------------
                      ''20180301 validasi order defaultnya 0 (min)
                      ''---------------------------------------------------
204                   If Log.discKacab = 1 Then
205                       Call giveDiscount("0", Text(21), False)
206                   Else
207                       Call giveDiscount("100", Text(21), False)
208                   End If
209                   vTipeCustomer = "INT"
210               End If
211           Else
                  ''bukan return job
212               If typeJob(Trim(rsTemp("TPSVC"))) = "WARRANTY" Then
213                   opWO(2).value = True
          
214                   Set rsHlp = objBilling.searchWarrantyDetail(rsTemp("KODETIPE"))
215                   If Not rsHlp.EOF Then
216                       Text(35) = rsHlp("ID_WARRANTY")
217                       Text(36) = rsHlp("WARRANTY")
218                       Text(35).Visible = True
219                       Text(36).Visible = True
220                   End If
221                   If rsHlp.State <> adStateClosed Then rsHlp.Close
222                   Call giveDiscount("0", Text(21), True)
                      'text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
223                   txtCatatan.Text = "A/N " & Text(4) & "  [" & Text(2) & "]"
224                   Set rsHlp = objO.FetchData("SELECT * FROM T_PROFIL WITH (NOLOCK)")
225                   If Not rsHlp.EOF Then
226                       gKdPlgDealer = Trim(rsHlp("FN_HOKDPLG"))
227                       gKodePlg = Trim(rsHlp("FN_HOKDPLG"))
228                       gInvCust = Trim(rsHlp("FN_HO"))
229                       gInvAlmt = Trim(rsHlp("FN_HOADRES"))
230                       gInvKota = Trim(rsHlp("FN_HOKOTA"))
231                       gInvNPWP = Trim(rsHlp("FN_HONPWP"))
232                       vTipeCustomer = "AFFPPN"
233                   End If
234                   If rsHlp.State <> adStateClosed Then rsHlp.Close
235               ElseIf typeJob(Trim(rsTemp("TPSVC"))) = "PROGRAM" Then
236                   opWO(1).value = True
          
237                   Set rsHlp = objBilling.searchProgramDetail(rsTemp("KODETIPE"))
238                   If Not rsHlp.EOF Then
239                       Text(35) = rsHlp("ID_PROGRAM")
240                       Text(36) = rsHlp("PROGRAM")
241                       Text(35).Visible = True
242                       Text(36).Visible = True
243                   End If
244                   If rsHlp.State <> adStateClosed Then rsHlp.Close
245                   If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
246                       nDuoInv = True
                          'SBI WO BERKEMBANG
247                       If rsTemp("FS_RC") = "R001" Then
          
248                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
249                           If Not rsHlp.EOF Then
250                               gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
251                               gKodePlg2 = Trim(rsHlp("KDPLG"))
252                               gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
253                               gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
254                               gInvKota2 = Trim(rsHlp("FS_KOTA"))
255                               gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                  
256                               If Trim(rsHlp("INTERNAL")) = 1 Then
257                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
258                                       vTipeCustomer = "AFFNONPPN"
259                                   Else
260                                       vTipeCustomer = "AFFPPN"
261                                   End If
262                               Else
263                                   vTipeCustomer = "EXT"
264                               End If
265                           End If
266                           If rsHlp.State <> adStateClosed Then rsHlp.Close
                              ''---------------------------------------------------
                              ''20180301 validasi order defaultnya 0 (min)
                              ''---------------------------------------------------
267                           If Log.discKacab = 1 Then
268                               nilaiDiscMax = objBilling.takeDiscountProgram(rsTemp("KODETIPE"))
269                               Call giveDiscount("0", Text(21), False)
270                           Else
271                               Call giveDiscount("100", Text(21), False)
272                           End If
                              ''text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
273                           txtCatatan.Text = "A/N " & Text(4) & "  [" & Text(2) & "]"
274                       Else
          
275                           Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
276                           If Not rsHlp.EOF Then
277                               gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
278                               gKodePlg2 = Trim(rsHlp("KDPLG"))
279                               gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
280                               gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
281                               gInvKota2 = Trim(rsHlp("FS_KOTA"))
282                               gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                  
283                               If Trim(rsHlp("INTERNAL")) = 1 Then
284                                   If Trim(gNpwp) = Trim(gInvNPWP2) Then
285                                       vTipeCustomer = "AFFNONPPN"
286                                   Else
287                                       vTipeCustomer = "AFFPPN"
288                                   End If
289                               Else
290                                   vTipeCustomer = "EXT"
291                               End If
292                           End If
293                           If rsHlp.State <> adStateClosed Then rsHlp.Close
                              ''---------------------------------------------------
                              ''20180301 validasi order defaultnya 0 (min)
                              ''---------------------------------------------------
294                           If Log.discKacab = 1 Then
295                               nilaiDiscMax = objBilling.takeDiscountProgram(rsTemp("KODETIPE"))
296                               Call giveDiscount("0", Text(21), True)
297                           Else
298                               Call giveDiscount("100", Text(21), False)
299                           End If
                              'Text(30) = "A/N " & Text(4) & "  [" & Text(2) & "]"
300                           txtCatatan.Text = "A/N " & Text(4) & "  [" & Text(2) & "]"
301                       End If
302                   Else
                          ''---------------------------------------------------
                          ''20180301 validasi order defaultnya 0 (min)
                          ''---------------------------------------------------
303                       If Log.discKacab = 1 Then
304                           nilaiDiscMax = objBilling.takeDiscountProgram(rsTemp("KODETIPE"))
305                           Call giveDiscount("0", Text(21), True)
306                       Else
307                           Call giveDiscount(objBilling.takeDiscountProgram(rsTemp("KODETIPE")), Text(21), False)
308                       End If
309                   End If
310               ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
311                   If lblPDS.Visible Then
312                       Call giveDiscount("0", Text(21), False)
313                       Set rsHlp = objBilling.takeKodeDealer(gKdCbg)
314                       If Not rsHlp.EOF Then
315                           gKdPlgDealer2 = Trim(rsHlp("KDPLG"))
316                           gKodePlg2 = Trim(rsHlp("KDPLG"))
317                           gInvCust2 = Trim(rsHlp("FS_NM_DEALER"))
318                           gInvAlmt2 = Trim(rsHlp("FS_ALAMAT"))
319                           gInvKota2 = Trim(rsHlp("FS_KOTA"))
320                           gInvNPWP2 = Trim(rsHlp("FS_NPWP"))
                  
321                           If Trim(rsHlp("INTERNAL")) = 1 Then
322                               If Trim(gNpwp) = Trim(gInvNPWP2) Then
323                                   vTipeCustomer = "AFFNONPPN"
324                               Else
325                                   vTipeCustomer = "AFFPPN"
326                               End If
327                           Else
328                               vTipeCustomer = "EXT"
329                           End If
330                       End If
331                       If rsHlp.State <> adStateClosed Then rsHlp.Close
332                   ElseIf rsTemp("FS_RC") = "R001" Then
                      ' HANYA SBI
333                       woStat = "SBI"
                  
334                       Set rsHlp = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
335                       If Not rsHlp.EOF Then
336                           gKdPlgDealer = Trim(rsHlp("KDPLG"))
337                           gKodePlg = Trim(rsHlp("KDPLG"))
338                           gInvCust = Trim(rsHlp("FS_NM_DEALER"))
339                           gInvAlmt = Trim(rsHlp("FS_ALAMAT"))
340                           gInvKota = Trim(rsHlp("FS_KOTA"))
341                           gInvNPWP = Trim(rsHlp("FS_NPWP"))
342                           If Trim(rsHlp("INTERNAL")) = 1 Then
343                               If Trim(gNpwp) = Trim(gInvNPWP) Then
344                                   Text(40) = "AFFNONPPN"
345                               Else
346                                   Text(40) = "AFFPPN"
347                               End If
348                           Else
349                               Text(40) = "EXT"
350                           End If
351                       End If
352                       If rsHlp.State <> adStateClosed Then rsHlp.Close
353                       Call giveDiscount("0", Text(21), False)
                  '                    text(30) = "A/N " & text(4) & "  [" & text(2) & "]"
354                       txtCatatan.Text = "A/N " & Text(4) & "  [" & Text(2) & "]"
355                   ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
          
356                       Set rsHlp = objBilling.searchMOUDetail(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS"))
357                       If Not rsHlp.EOF Then
358                           Text(34) = Trim(rsHlp("NO_MOU"))
359                       End If
360                       If rsHlp.State <> adStateClosed Then rsHlp.Close
                          ''---------------------------------------------------
                          ''20180301 validasi order defaultnya 0 (min)
                          ''---------------------------------------------------
361                       If Log.discKacab = 1 Then
362                           nilaiDiscMax = objBilling.takeDiscountMOU(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS"))
363                           Call giveDiscount("0", Text(21), False)
364                       Else
365                           Call giveDiscount(objBilling.takeDiscountMOU(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")), Text(21), False)
366                       End If
367                   Else
368                       Call giveDiscount("0", Text(21), True)
369                   End If
370               End If
371         End If
          
          
            ' TAMPIL JOB ORDER
            Dim nilaiRate     As Double
            Dim rsJobDetail   As ADODB.Recordset
            Dim namaProgram   As String
372         totRate = 0
373         Text(22) = 0
374         subJob = 0
375         If rsJob.State Then rsJob.Close
376         myChassis = Trim(Text(17))
            Dim nilaiDiscJasa As Integer
377         If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(43), "DD/MM/YYYY")) >= 0 Then
378           sSQL = "select * from jobdetail with (nolock) where no_wo = '" & NulSID(Text(0)) & "' and no_chassis = '" & NulSID(rsTemp("NO_CHASSIS")) & "'"
379           Set rsJobDetail = objO.FetchData(sSQL)
380           If Not rsJobDetail.EOF Then
381               rsJobDetail.MoveFirst
382               Do While Not rsJobDetail.EOF
383                   fOth = False
384                   isDiscKacab = False
385                   Set rsHlp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & NulSID(rsJobDetail!rep) & "' AND MDL = '" & rsTemp("MDL") & "'")
386                   If rsHlp.EOF Then
387                       fOth = True
388                   End If
389                   If rsHlp.State <> adStateClosed Then rsHlp.Close
                      
390                   If NulSID(rsJobDetail!id_program) = "" Then
                          ''reguler / warranty bukan program
391                       If returnJob(LblReturn) Then
392                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
393                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
394                       Else
                              '' bukan return job
395                           If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
396                               nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
397                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
398                           ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
399                               If lblPDS.Visible Then
400                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
401                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
402                               ElseIf rsTemp("FS_RC") = "R001" Then
403                                   Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
404                                   If Not rsHlp2.EOF Then
405                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
406                                       If nStatusDealer = 1 Then
407                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
408                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
409                                       Else
410                                           If nStatusDealer = 2 Then
411                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
412                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
413                                           ElseIf nStatusDealer = 3 Then
414                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
415                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
416                                           End If
417                                       End If
418                                   End If
419                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
                                      
420                               ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
421                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
422                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
423                               Else
424                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
425                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
426                               End If
427                           End If
428                       End If
429                   ElseIf NulSID(rsJobDetail!isFP) = 1 Then
                          ''jika fix price
                          
430                   Else
                          '' jika program
431                       If objBilling.isDuoInvoice(NulSID(rsJobDetail!id_program)) Then
432                           nDuoInv = True
                              
433                           If rsTemp("FS_RC") = "R001" Then
434                               If NulSID(rsJobDetail!klpkerja) = "SBI" Then
435                                   Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
436                                   If Not rsHlp2.EOF Then
437                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
438                                       If nStatusDealer = 1 Then
439                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
440                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
441                                       Else
442                                           If nStatusDealer = 2 Then
443                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
444                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
445                                           ElseIf nStatusDealer = 3 Then
446                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
447                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
448                                           End If
449                                       End If
450                                   End If
451                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
452                               Else
453                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
454                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
455                               End If
456                           Else
457                               Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
458                               If Not rsHlp2.EOF Then
          'Dim nStatusDealer As Integer
459                                   nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
460                                   If nStatusDealer = 1 Then
461                                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0  '
462                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
463                                   Else
464                                       If nStatusDealer = 2 Then
465                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
466                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
467                                       ElseIf nStatusDealer = 3 Then
468                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
469                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
470                                       End If
471                                   End If
          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
472                               End If
473                               If rsHlp2.State <> adStateClosed Then rsHlp2.Close
474                           End If
475                       Else
476                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
477                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
478                       End If
479                   End If
               
480                   nilaiRate = CDbl(NulSID(rsJobDetail!rte))
481                   nilaiDiscJasa = CInt(Text(21))
482                   If Log.discKacab = 1 Then
                          '' mengambil data yang telah di validasi dan tidak batal
483                       Set rsHlp2 = objO.FetchData("SELECT discJasaValid as jasa FROM RequestDisc WITH (NOLOCK) WHERE (userValid <> 'BTL' AND uservalid <> '') and no_wo = '" & Trim(Text(Index)) & "' ")
484                       If Not rsHlp2.EOF Then
485                           isDiscKacab = True
486                           nilaiDiscJasa = CInt(rsHlp2!jasa)
487                           nilaiDiscMax = CInt(rsHlp2!jasa)
488                       End If
489                   End If
490                   If Log.discKacab = 1 And nDuoInv Then
491                       Text(22) = 0
492                   Else
493                       Text(22) = Text(22) + (CDbl(nHarga) * nilaiRate) * ((100 - nilaiDiscJasa) / 100)
494                   End If
                      
495                   If Log.discKacab = 1 Then
                          '' mengambil data yang telah di validasi dan tidak batal
496                       Set rsHlp2 = objO.FetchData("SELECT discJasaValid as jasa FROM RequestDisc WITH (NOLOCK) WHERE (userValid <> 'BTL' AND uservalid <> '') and no_wo = '" & Trim(Text(Index)) & "' ")
497                       If Not rsHlp2.EOF And nDuoInv = False Then
498                           nilaiDiscJasa = CInt(rsHlp2!jasa)
499                           nilaiDiscMax = CInt(rsHlp2!jasa)
500                       End If
501                   End If
              
502                   With rsJob
503                   If .State = adStateClosed Then doCreateJob
504                       namaProgram = ""
505                      .AddNew
506                      !job = NulSID(rsJobDetail!rep)
507                      !Rate = nilaiRate 'CDbl(rsTemp(hlp2))
508                      !ACTRT = CDbl(NulSID(rsJobDetail!actrte))
509                      !grp = NulSID(rsJobDetail!grp)
510                      !Mech = NulSID(rsJobDetail!mekanik)
511                      If NulSID(rsJobDetail!isFP) = 1 Then
512                       !hrg = objO.FetchData_Double("select price from [" & dbReplikasiHO & "]..HO_FIXPRICE where id_FP = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
513                      Else
514                       !hrg = Round(CDbl(nHarga) * nilaiRate, 0)
515                      End If
516                      !addJob = IIf(NulSID(rsJobDetail!addJob) = 1, True, False)
517                      !disc = CDbl(nilaiDiscJasa)
518                      !DISCRP = Round((CDbl(!hrg) * (CDbl(nilaiDiscJasa) / 100)), 0)
519                      If NulSID(rsJobDetail!isFP) = 1 Then
520                       !STLDISC = !hrg
521                      Else
522                       !STLDISC = Round((CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)), 0)
523                      End If
524                      !valid = gVerifikator
525                      If NulSID(rsJobDetail!isFP) = 1 Then
526                          !FREEC = True
527                      Else
528                          !FREEC = False
529                      End If
530                      !idProgram = NulSID(rsJobDetail!id_program)
531                      If NulSID(rsJobDetail!isFP) = 1 Then
532                         !discmax = 0
533                      Else
534                           If isDiscKacab Then
535                               !discmax = nilaiDiscJasa
536                           Else
537                               If NulSID(rsJobDetail!id_program) <> "" Then
538                                   !discmax = objO.FetchData_Double("select discjasa from [" & dbReplikasiHO & "]..HO_PROGRAM where id_Program = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
539                               Else
540                                   !discmax = nilaiDiscMax 'nilaiLockDisc
541                               End If
542                           End If
543                      End If
544                      !isFP = IIf(NulSID(rsJobDetail!isFP) = 1, True, False)
545                      If NulSID(rsJobDetail!isFP) = 1 Then
546                       namaProgram = objO.FetchData_String("select namaFP from [" & dbReplikasiHO & "]..HO_FIXPRICE where id_FP = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
547                      Else
548                       namaProgram = objO.FetchData_String("select PROGRAM from [" & dbReplikasiHO & "]..HO_PROGRAM where id_Program = '" & NulSID(rsJobDetail!id_program) & "' and FS_LOG_DEL = ''")
549                      End If
550                      !program = namaProgram
551                      If NulSID(rsJobDetail!isFP) = 1 Then
552                       Text(22) = Text(22) + !hrg
553                      End If
554                      .Update
555                   End With
556                   subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
557                   totRate = totRate + CDbl(nilaiRate)
558                   rsJobDetail.MoveNext
559               Loop
560           End If
561           GoTo lanjutJobOrder
562         End If
            
563         For i = 1 To 7
564           hlp = "REP" & i
565           hlp2 = "RTE" & i
566           hlp3 = "KLPKERJA" & i
567           hlp4 = "ACTRATE" & i
568           hlp5 = "MEKANIK" & i
569           nilaiDiscJasa = 0
570           If rsTemp(hlp) <> "" Then
571               fOth = False
572               Set rsHlp = objO.FetchData("SELECT * FROM HO_FLATE WITH (NOLOCK) WHERE FS_ITEM = '" & rsTemp(hlp) & "' AND MDL = '" & rsTemp("MDL") & "'")
573               If rsHlp.EOF Then
574                   fOth = True
575               End If
576               If rsHlp.State <> adStateClosed Then rsHlp.Close
          
577               If returnJob(LblReturn) Then
578                   If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
579                       If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
580                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
581                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
582                       End If
583                   Else
584                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
585                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
586                   End If
587               Else
588                   If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
589                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
590                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
591                   ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
592                       If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
593                           If rsTemp("FS_RC") = "R001" Then
594                               If Trim(rsTemp(hlp3)) = "SBI" Then
          
595                                   Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
596                                   If Not rsHlp2.EOF Then
597                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
598                                       If nStatusDealer = 1 Then
599                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
600                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
601                                       Else
602                                           If nStatusDealer = 2 Then
603                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
604                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
605                                           ElseIf nStatusDealer = 3 Then
606                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
607                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
608                                           End If
609                                       End If
610                                   End If
611                                   If rsHlp2.State <> adStateClosed Then rsHlp2.Close
612                               Else
613                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
614                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
615                               End If
616                           Else
          
617                               Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
618                               If Not rsHlp2.EOF Then
          'Dim nStatusDealer As Integer
619                                   nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
620                                   If nStatusDealer = 1 Then
621                                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0  '
622                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
623                                   Else
624                                       If nStatusDealer = 2 Then
625                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
626                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
627                                       ElseIf nStatusDealer = 3 Then
628                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
629                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
630                                       End If
631                                   End If
          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
632                               End If
633                               If rsHlp2.State <> adStateClosed Then rsHlp2.Close
634                           End If
635                       Else
636                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
637                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
638                       End If
639                   ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
640                       If lblPDS.Visible Then
641                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
642                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
643                       ElseIf rsTemp("FS_RC") = "R001" Then
          
644                           Set rsHlp2 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
645                           If Not rsHlp2.EOF Then
646                               nStatusDealer = objBilling.checkDealerStatus(rsHlp2("FS_KD_DEALER"))
647                               If nStatusDealer = 1 Then
648                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
649                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
650                               Else
651                                   If nStatusDealer = 2 Then
652                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
653                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
654                                   ElseIf nStatusDealer = 3 Then
655                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
656                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
657                                   End If
658                               End If
          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
659                           End If
660                           If rsHlp2.State <> adStateClosed Then rsHlp2.Close
661                       ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
662                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
663                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
664                       Else
665                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
666                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
667                       End If
668                   End If
669               End If
670               nilaiRate = CDbl(rsTemp(hlp2))
671               nilaiDiscJasa = CInt(Text(21))
672               If Log.discKacab = 1 Then
                      '' mengambil data yang telah di validasi dan tidak batal
673                   Set rsHlp2 = objO.FetchData("SELECT discJasaValid as jasa FROM RequestDisc WITH (NOLOCK) WHERE (userValid <> 'BTL' AND uservalid <> '') and no_wo = '" & Trim(Text(Index)) & "' ")
674                   If Not rsHlp2.EOF And nDuoInv = False Then
675                       nilaiDiscJasa = CInt(rsHlp2!jasa)
676                       nilaiDiscMax = CInt(rsHlp2!jasa)
677                   End If
678               End If
679               If Log.discKacab = 1 And nDuoInv Then
680                   Text(22) = 0
681               Else
682                   Text(22) = Text(22) + (CDbl(nHarga) * nilaiRate) * ((100 - nilaiDiscJasa) / 100)
683               End If
                  
          
684               With rsJob
685               If .State = adStateClosed Then doCreateJob
686                  .AddNew
687                  !job = Trim(rsTemp(hlp))
688                  !Rate = nilaiRate 'CDbl(rsTemp(hlp2))
689                  !ACTRT = CDbl(rsTemp(hlp4))
690                  !grp = Trim(rsTemp(hlp3))
691                  !Mech = Trim(rsTemp(hlp5))
692                  !hrg = Round(CDbl(nHarga) * nilaiRate, 0)
693                  !addJob = False
694                  !disc = CDbl(nilaiDiscJasa)
695                  !DISCRP = Round((CDbl(!hrg) * (CDbl(nilaiDiscJasa) / 100)), 0)
696                  !STLDISC = Round((CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)), 0)
697                  !valid = gVerifikator
          '                 If typeJob(rsTemp("TPSVC")) = "REGULAR" Then
          '                    !FREEC = True
          '                 Else
          '                    !FREEC = False
          '                 End If
698                   !FREEC = False
699                  .Update
700               End With
701               subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
702               totRate = totRate + CDbl(nilaiRate)
703           End If
704         Next
          
          ' add JOB
705         Set rsHlp2 = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & noWO & "' AND NO_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
706         If Not rsHlp2.EOF Then
707       Do While Not rsHlp2.EOF
          '            For i = 1 To 10
          '                hlp = "REP" & i
          '                hlp2 = "RTE" & i
          '                hlp3 = "KLPKERJA" & i
          '                hlp4 = "ACTRATE" & i
          '                hlp5 = "MEKANIK" & i
          
708           If rsHlp2("REP") <> "" Then
709               If returnJob(LblReturn) Then
710                   If typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
711                       If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
712                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
713                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
714                       End If
715                   Else
716                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
717                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
718                   End If
719               Else
720                   If typeJob(rsTemp("TPSVC")) = "WARRANTY" Then
721                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 3)
722                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
723                   ElseIf typeJob(rsTemp("TPSVC")) = "PROGRAM" Then
724                       If objBilling.isDuoInvoice(rsTemp("KODETIPE")) Then
725                           If rsTemp("FS_RC") = "R001" Then
726                               If rsHlp2("KLPKERJA") = "SBI" Then
          
727                                   Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
728                                   If Not rsHlp3.EOF Then
729                                       nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
730                                       If nStatusDealer = 1 Then
731                                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
732                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
733                                       Else
734                                           If nStatusDealer = 2 Then
735                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
736                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
737                                           ElseIf nStatusDealer = 3 Then
738                                               nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
739                                               If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
740                                           End If
741                                       End If
          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
742                                   End If
743                                   If rsHlp3.State <> adStateClosed Then rsHlp3.Close
744                               Else
745                                   nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
746                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
747                               End If
748                           Else
          
749                               Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
750                               If Not rsHlp3.EOF Then
751                                   nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
752                                   If nStatusDealer = 1 Then
753                                       nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1) '0 '
754                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
755                                   Else
756                                       If nStatusDealer = 2 Then
757                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
758                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
759                                       ElseIf nStatusDealer = 3 Then
760                                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 2)
761                                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
762                                       End If
763                                   End If
          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
764                               End If
765                               If rsHlp3.State <> adStateClosed Then rsHlp3.Close
766                           End If
767                       Else
768                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
769                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
770                       End If
771                   ElseIf typeJob(rsTemp("TPSVC")) = "REGULAR" Then
772                       If lblPDS.Visible Then
773                           nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 4)
774                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
775                       ElseIf rsTemp("FS_RC") = "R001" Then
          
776                           Set rsHlp3 = objBilling.takeInvoiceName(rsTemp("NO_CHASSIS"))
777                           If Not rsHlp3.EOF Then
778                               nStatusDealer = objBilling.checkDealerStatus(rsHlp3("FS_KD_DEALER"))
779                               If nStatusDealer = 1 Then
780                                   nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1) '0  '
781                                   If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
782                               Else
783                                   If nStatusDealer = 2 Then
784                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 1)
785                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
786                                   ElseIf nStatusDealer = 3 Then
787                                       nHarga = objBilling.takeClassPrice(objBilling.takeCheapsClassName, 2)
788                                       If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
789                                   End If
790                               End If
          'Text(22) = Text(22) + (CDbl(nHarga) * CDbl(rsTemp(hlp2)))
791                           End If
792                           If rsHlp3.State <> adStateClosed Then rsHlp3.Close
793                       ElseIf objBilling.isMOUCustomer(gKdCbg, rsTemp("MDL"), rsTemp("NO_CHASSIS")) Then
794                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
795                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
796                       Else
797                           nHarga = objBilling.takeClassPrice(rsTemp("MODEL"), 1)
798                           If isOnCarClass = False Then MsgBox "Kelas Kendaraan Tidak Terdaftar !", vbCritical + vbOKOnly, "Car Class"
799                       End If
800                   End If
801               End If
          
          ' ADJUSTMENT RATE
802               nilaiRate = CDbl(rsHlp2("RTE"))
          '                    If Trim(rsTemp(hlp3)) <> "SBE" And Trim(rsTemp(hlp3)) <> "SBI" Then
          '                        Me.Tag = ""
          '                        gVerifikator = ""
          '                        If MsgBox("Pekerjaan : " & Trim(rsHlp2("REP")) & vbCr & "Adjustment Rate? ", vbYesNo, "PENYESUAIAN RATE") = vbYes Then
          '                            frmValidasi.Tag = "ADJ"
          '                            frmValidasi.Show vbModal
          '                        End If
          '                        If Me.Tag = "valid" Then
          '                            'nilaiRate = 0
          '                            nHarga = 0
          '                        End If
          '                    End If
803               If Log.discKacab = 1 And nDuoInv Then
804                   Text(22) = 0
805               Else
806                   Text(22) = Text(22) + (CDbl(nHarga) * CDbl(nilaiRate)) * ((100 - Text(21)) / 100)
807               End If
          
808               With rsJob
809               If .State = adStateClosed Then doCreateJob
810                   .AddNew
811                   !job = Trim(rsHlp2("REP"))
812                   !Rate = nilaiRate 'CDbl(rsHlp2("RTE"))
813                   !ACTRT = CDbl(rsHlp2("ACTRTE"))
814                   !grp = Trim(rsHlp2("KLPKERJA"))
815                   !Mech = Trim(rsHlp2("MEKANIK"))
816                   !hrg = CDbl(nHarga) * CDbl(nilaiRate)
817                   !addJob = True
818                   !disc = CDbl(Text(21))
819                   !DISCRP = CDbl(!hrg) * (CDbl(Text(21)) / 100)
820                   !STLDISC = CDbl(nHarga) * CDbl(nilaiRate) * ((100 - !disc) / 100)
821                   !valid = gVerifikator
          '                  If typeJob(rsTemp("TPSVC")) = "REGULAR" Then
          '                   !FREEC = True
          '                  Else
          '                   !FREEC = False
          '                  End If
822                    !FREEC = False
823                   .Update
824                   adaAddJob = True
825               End With
826               subJob = subJob + (CDbl(nHarga) * CDbl(nilaiRate))
827               totRate = totRate + CDbl(nilaiRate)
828           End If
          '            Next
829           rsHlp2.MoveNext
830       Loop
831         End If
            
lanjutJobOrder:
          '        subJob = CDbl(Text(22))
832         Set JobGrid.DataSource = rsJob
833         Call doHeadJob
          
          ' TAMPIL PART ORDER
834         noKFaktur = ""
835         gSpart = True
836         Text(24) = 0
837         Text(25) = 0
838         subPartGenuine = 0
839         subPartNGenuine = 0
840         subOliGenuine = 0
841         subOliNGenuine = 0
          Dim HargaDisc As Double
842         If rsPart.State Then rsPart.Close
          
          ' /// AMBIL PART YG BAWA SENDIRI //
          ' AND B.TIPE = 'O'
843         Set rsHlp = objO.FetchData("SELECT KODE_ITEM, NAMA_ITEM, PARTBAWA, QTY, isnull(fs_kd_oli,'') as OLI, TIPE, GENUINE, BAWA FROM jobpart a WITH (NOLOCK) LEFT JOIN T_oli b WITH (NOLOCK) ON LEFT(a.kode_item,1) = b.fs_kd_Oli WHERE no_WO = '" & noWO & "' " _
                            & "AND no_Chassis = '" & rsTemp("NO_CHASSIS") & "' AND GANTI = 'Y' ") 'AND BAWA <> 'Y'
844         If Not rsHlp.EOF Then
845             rsHlp.MoveFirst
846             Do While Not rsHlp.EOF
847                 With rsPart
848                 If .State = adStateClosed Then doCreatePart
849                     gSpart = True
          
850                     If rsHlp("BAWA") = "Y" Then
851                           .AddNew
852                           !KodeItem = "BAWA SENDIRI"
853                           !NAMAITEM = rsHlp("PARTBAWA")
854                           !qty = rsHlp("QTY")
855                           !QtyOrder = rsHlp("QTY")
856                           !QtyBO = 0
857                           !QtyCancel = 0
858                           !Harga = 0
859                           !oli = IIf(rsHlp("TIPE") <> "P", True, False)
860                           .Update
861                     Else
862                         If Not gFPart Then
                                '///// TIDAK CONNECT ACCOUNTING SPAREPART ////
863                           .AddNew
864                           !KodeItem = IIf(rsHlp("BAWA") = "Y", "BAWA SENDIRI", rsHlp("KODE_ITEM"))
865                           !NAMAITEM = IIf(rsHlp("BAWA") = "Y", rsHlp("PARTBAWA"), rsHlp("NAMA_ITEM"))
866                           !oli = IIf(rsHlp("TIPE") <> "P", True, False)
867                           !qty = rsHlp("QTY")
                              'Set rsHlp2 = objO.FetchDataACC("SELECT hargajualA FROM inmst WHERE kodeItem = '" & rsHlp("KODE_ITEM") & "'")
868                           Set rsHlp2 = objO.FetchDataACC("SELECT A.hargajualA as hargajualA, B.Jenis as SCC, B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.kodeItem = '" & rsHlp("KODE_ITEM") & "'")
869                           If Not rsHlp2.EOF Then
870                               HargaDisc = CDbl(rsHlp2("HARGAJUALA"))
871                               !Harga = CDbl(HargaDisc)
872                               If rsHlp("OLI") <> "" Then
873                                   If rsHlp("TIPE") = "P" Then
874                                       If rsHlp2("TGenuine") = False Then
          ' PART GENUINE
875                                           subPartNGenuine = subPartNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
876                                       ElseIf rsHlp2("TGenuine") = True Then
          ' PART NON GENUINE
877                                           subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
878                                       End If
879                                       Text(24) = Text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
880                                   ElseIf rsHlp("TIPE") = "O" Then
881                                       If rsHlp2("TGenuine") = False Then
          ' OLI GENUINE
882                                           subOliNGenuine = subOliNGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
883                                       ElseIf rsHlp2("TGenuine") = True Then
          ' OLI NON GENUINE
884                                           subOliGenuine = subOliGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
885                                       End If
886                                       Text(25) = Text(25) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
887                                   End If
888                               Else
          ' PART GENUINE
889                                   subPartGenuine = subPartGenuine + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
890                                   Text(24) = Text(24) + (CDbl(HargaDisc) * CDbl(rsHlp("QTY")))
891                               End If
892                           End If
893                           If rsHlp2.State <> adStateClosed Then rsHlp2.Close
          
894                           .Update
895                         End If
896                     End If
897                 End With
898             rsHlp.MoveNext
899             Loop
900         End If
901         If rsHlp.State <> adStateClosed Then rsHlp.Close
          
902         If gFPart Then
            '/// KONEKSI KE SPAREPART
903         If Get_ConnectionNisSP(NisSPDb, NisSPUSER, NisSPPass, NisSPKomp) Then
904           With rsPart
905               If noInv Then
                      Dim strSql As String
906                   If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
907                       strSql = "select (case When KodeItem = '' Then namaBYSA Else KodeItem end) as KodeItem, " & _
                                    "Keterangan as Namabarang, SatuanNs as Satuan, Qty as QtyOrder, (isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0)) AS Jml, " & _
                                    "(Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) as QtyBO, HargaJual as HargaJualA, NOSO as NoRack, " & _
                                    "0 as DisItem, 0 as ExtDis, Case BO When 'CANCEL' THEN (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) ELSE 0 END AS QTYCANCEL from SOORD WITH (NOLOCK) Where NOSO = '" & Trim(Text(0)) & "'"
908                   Else
909                       strSql = "Select Z.*, 0 as QtyOrder, 0 as QtyBO FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
                                    "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
                                    "(INMST.HARGAJUALA ) as HARGAJUALA, " & _
                                    "INMST.NORACK,0 as DisItem, 0 as ExtDis, 0 as QtyCANCEL FROM INMST " & _
                                    "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
                                    "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
                                    "Where Z.Norack = '" & Trim(Text(0)) & "'"
910                   End If
911                   Set rsHlp3 = objO.FetchDataACC(strSql)
          
912                   If Not rsHlp3.EOF Then
913                       Do While Not rsHlp3.EOF
914                           If .State = adStateClosed Then doCreatePart
915                             .AddNew
916                           !KodeItem = Trim(rsHlp3("KODEITEM"))
917                           !NAMAITEM = Trim(rsHlp3("NAMABARANG"))
918                           !QtyOrder = CDbl(rsHlp3("QtyOrder"))
919                           !qty = CDbl(rsHlp3("JML"))
920                            If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsTemp("date")), "DD/MM/YYYY") Then
921                               !QtyBO = IIf(UCase(isBO) = "CANCEL", 0, objO.FetchData_DoubleACC("select Qty FROM BOFILE WITH (NOLOCK) Where TutupOrderFlag = 0 and (NoSO = '" & Trim(Text(0)) & "' Or NOSO =(Select Distinct NoSPK From SOORD WITH (NOLOCK) Where NOSO ='" & Trim(Text(0)) & "' and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "')) and KodeItem = '" & Trim(rsHlp3("KODEITEM")) & "'"))
922                            Else
923                               !QtyBO = 0
924                            End If
          
925                           !QtyCancel = CDbl(rsHlp3("QtyCancel"))
926                           !Harga = CDbl(rsHlp3("HARGAJUALA"))
          '                                Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WHERE FS_KD_OLI = '" & Left(rsHlp3("KODEITEM"), 1) & "'")
          '                                If Not rsHlp.EOF Then
          '                                    If rsHlp("TIPE") = "O" Then
          '                                        !oli = True
          '                                        Text(25) = Text(25) + (CDbl(!Harga) * CDbl(!qty))
          '                                    ElseIf rsHlp("TIPE") = "P" Then
          '                                        !oli = False
          '                                        Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
          '                                    End If
          '                                Else
          '                                    !oli = False
          '                                    Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
          '                                End If
927                               Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsHlp3("KODEITEM") & "'")
928                               If Not rsHlp.EOF Then
929                                   If Left(rsHlp("SCC"), 3) = "OLI" Then
930                                       !oli = True
931                                       Text(25) = Text(25) + (CDbl(!Harga) * CDbl(!qty))
932                                   Else 'If Left(rsHlp("SCC"), 3) = "PAR" Then
933                                       !oli = False
934                                       Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
          '                                        ElseIf Left(rsHlp("SCC"), 3) = "ACC" Then
          '                                            !oli = False
          '                                            Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
935                                   End If
936                               Else
937                                   !oli = False
938                                   Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
939                               End If
940                           .Update
941                           rsHlp3.MoveNext
942                       Loop
943                   Else 'If rsHlp3.EOF And newValid Then
          Dim rsInvPart As New ADODB.Recordset
          Dim vHitung As Double
944                       Set rsInvPart = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
          & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(Text(0)) & "' and KodeLgn = '" & txtKodeLgnLama.Text & "' ")
945                       If Not rsInvPart.EOF Then
946                           Do While Not rsInvPart.EOF
947                               If .State = adStateClosed Then doCreatePart
948                               .AddNew
949                               !KodeItem = Trim(rsInvPart("KODEITEM"))
950                               !NAMAITEM = Trim(rsInvPart("KETERANGAN"))
951                               vHitung = (Koma2Titik(rsInvPart("Qty")) * (Koma2Titik(rsInvPart("HargaJual")) - (Koma2Titik(rsInvPart("ItemDisc")) / Abs(Koma2Titik(rsInvPart("Qty"))))))
952                               vHitung = vHitung - (Koma2Titik(rsInvPart("ExtDisPersen")) * vHitung)
953                               If Trim(UCase(rsInvPart("TypePpn"))) = "E" Then
954                                   vHitung = Round(vHitung + vHitung * (profilACC.nilaippn / 100), 2)
955                               End If
956                               If gJurnalPart Then
957                                   vHitung = cekIncludePPn(vHitung)
958                               End If
959                               !qty = Val(rsInvPart("QTY"))
960                               If !qty = 0 Then
961                                   !Harga = 0
962                               Else
963                                   !Harga = CDbl(vHitung / !qty)
964                               End If
965                               !QtyOrder = Val(rsInvPart("QTY"))
966                               !QtyBO = 0
967                               !QtyCancel = 0
          
          'Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WHERE FS_KD_OLI = '" & Left(rsInvPart("KODEITEM"), 1) & "'")
968                               Set rsHlp = objO.FetchDataACC("SELECT A.hargajualA, B.Jenis as SCC,B.tg as TGenuine FROM inmst A WITH (NOLOCK) inner Join ProfilACC..SCC B WITH (NOLOCK) On A.Scc_code = B.SCC_Code WHERE A.KodeItem = '" & rsInvPart("KODEITEM") & "'")
969                               If Not rsHlp.EOF Then
970                                   If Left(rsHlp("SCC"), 3) = "OLI" Then
971                                       !oli = True
972                                       Text(25) = Text(25) + (CDbl(!Harga) * CDbl(!qty))
973                                   Else 'If Left(rsHlp("SCC"), 3) = "PAR" Then
974                                       !oli = False
975                                       Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
          '                                        ElseIf Left(rsHlp("SCC"), 3) = "ACC" Then
          '                                            !oli = False
          '                                            Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
976                                   End If
977                               Else
978                                   !oli = False
979                                   Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
980                               End If
          '                                    If Not rsHlp.EOF Then
          '                                        If rsHlp("TIPE") = "O" Then
          '                                            !oli = True
          '                                            Text(25) = Text(25) + (CDbl(!Harga) * CDbl(!qty))
          '                                        ElseIf rsHlp("TIPE") = "P" Then
          '                                            !oli = False
          '                                            Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
          '                                        End If
          '                                    Else
          '                                        !oli = False
          '                                        Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
          '                                    End If
981                               .Update
982                               rsInvPart.MoveNext
983                           Loop
984                       Else
985                           gSpart = False
          '                                If TblValid.Enabled Then
          '                                    temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbYesNo, "KONFIRMASI PART")
          '                                    If temp = vbYes Then
          '                                        gSpart = False
          '                                    End If
          '                                End If
986                       End If
987                       If rsInvPart.State <> adStateClosed Then rsInvPart.Close
988                   End If
989                   If rsHlp3.State <> adStateClosed Then rsHlp3.Close
990               Else 'Not newValid Then
          
991                   Set rsHlp2 = objO.FetchData("SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND NOINV LIKE '" & Trim(Text(20)) & "%'")
992                   If Not rsHlp2.EOF Then
993                       Do While Not rsHlp2.EOF
994                           If .State = adStateClosed Then doCreatePart
995                           .AddNew
996                           !KodeItem = Trim(rsHlp2("KODEITEM"))
997                           !NAMAITEM = Trim(rsHlp2("KETERANGAN"))
998                           !qty = CDbl(rsHlp2("QTY"))
999                           !Harga = CDbl(rsHlp2("HARGA")) '* ((100 - CDbl(rsHlp2("FN_EXT_DISKON"))) / 100)
1000                          !oli = Trim(rsHlp2("OLI")) 'IIf(Trim(rsHlp2("OLI")) = 1, True, False)
1001                          If !oli Then
1002                              Text(25) = Text(25) + CDbl(!Harga)
1003                          Else
1004                              Text(24) = Text(24) + CDbl(!Harga)
1005                          End If
1006                          !EXTDISC = CDbl(rsHlp2("FN_EXT_DISKON"))
1007                          !QtyOrder = CDbl(rsHlp2("QTY"))
1008                          !QtyBO = 0
1009                          !QtyCancel = 0
1010                          .Update
1011                          rsHlp2.MoveNext
1012                      Loop
1013                  Else
1014                      gSpart = False
1015                  End If
1016                  If rsHlp2.State <> adStateClosed Then rsHlp2.Close
1017              End If
1018          End With
1019      End If
          
1020        If Not gSpart Then
1021              With rsPart
1022                  Set rsHlp3 = objO.FetchDataACC("SELECT A.NOFAKTUR, A.TglFaktur, A.TypePpn, A.ExtDisPersen, B.qty, B.ItemDisPsn, B.HargaJual, B.Keterangan, B.KodeItem, B.ItemDisc  " _
          & "FROM SIHDR A WITH (NOLOCK) INNER JOIN SISAL B WITH (NOLOCK) ON A.NOFAKTUR = B.NOFAKTUR WHERE A.NOSO = '" & Trim(Text(0)) & "' and KodeLgn = '" & txtKodeLgnLama.Text & "' ")
          
1023                  If Not rsHlp3.EOF Then
1024                      Do While Not rsHlp3.EOF
1025                          If .State = adStateClosed Then doCreatePart
1026                          .AddNew
1027                          !KodeItem = Trim(rsHlp3("KODEITEM"))
1028                          !NAMAITEM = Trim(rsHlp3("KETERANGAN"))
1029                          vHitung = (Koma2Titik(rsHlp3("Qty")) * (Koma2Titik(rsHlp3("HargaJual")) - (Koma2Titik(rsHlp3("ItemDisc")) / Abs(Koma2Titik(rsHlp3("Qty"))))))
1030                          vHitung = vHitung - (Koma2Titik(rsHlp3("ExtDisPersen")) * vHitung)
1031                          If Trim(UCase(rsHlp3("TypePpn"))) = "E" Then
1032                              vHitung = Round(vHitung + vHitung * (profilACC.nilaippn / 100), 2)
1033                          End If
1034                          If gJurnalPart Then
1035                              vHitung = cekIncludePPn(vHitung)
1036                          End If
1037                          !qty = CDbl(rsHlp3("Qty"))
1038                          !Harga = CDbl(vHitung / !qty)
1039                          Set rsHlp = objO.FetchData("SELECT * FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(rsHlp3("KODEITEM"), 1) & "'")
1040                          If Not rsHlp.EOF Then
1041                              If rsHlp("TIPE") = "O" Then
1042                                  !oli = True
1043                                  Text(25) = Text(25) + (CDbl(!Harga) * CDbl(!qty))
1044                              ElseIf rsHlp("TIPE") = "P" Then
1045                                  !oli = False
1046                                  Text(24) = Text(24) + (CDbl(!Harga) * CDbl(!qty))
1047                              End If
1048                          End If
1049                          !QtyOrder = CDbl(rsHlp3("Qty"))
1050                          !QtyBO = 0
1051                          !QtyCancel = 0
1052                          .Update
1053                          rsHlp3.MoveNext
1054                      Loop
1055                      gSpart = True
1056                  Else
1057                      If TblValid.Enabled Then
1058                          temp = MsgBox("APAKAH ADA PEMAKAIAN PART ?? ", vbYesNo, "KONFIRMASI PART")
1059                          If temp = vbYes Then
1060                              gSpart = False
1061                          Else
1062                              gSpart = True
1063                          End If
1064                      End If
1065                  End If
1066                  If rsHlp3.State <> adStateClosed Then rsHlp3.Close
1067              End With
1068          Else
1069              gSpart = True
1070          End If
1071      End If
      'End If
          
1072          subPart = CDbl(Text(24))
1073          subOli = CDbl(Text(25))
1074          If rsPart.State <> adStateClosed Then
1075            If Not rsPart.BOF Then
1076                rsPart.MoveFirst
1077                rsPart.Sort = "OLI"
1078                Set partGrid.DataSource = rsPart
1079                Call doHeadPart
1080            End If
1081          End If
              ' TAMPIL SUBLET ORDER
1082          adaSublet = False
              'cmdSublet.Visible = False
1083          Text(27) = 0
1084          If rsSublet.State Then rsSublet.Close
1085          If tagihanFOP <> "" And nDuoInv Then
1086              Call cekSubletTerlibatFOP(noWO, NulSID(rsTemp("NO_CHASSIS")))
1087          Else
1088              If Log.standarisasiSublet = 1 Then
1089                  Call cekSubletNonFOP(noWO, NulSID(rsTemp("NO_CHASSIS")))
1090              Else
1091                  Set rsSublet = objO.FetchData("SELECT NONOTA, SUPPLIER, NAMABARANG, KODESATUAN, QTY, valid, HARGA, PRICE,cast(-1 as bit) as freec,cast(0 as int) as diskon,price*qty as subtotal FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & noWO & "' AND no_CHASSIS = '" & rsTemp("NO_CHASSIS") & "'")
1092              End If
1093          End If
1094          If Not rsSublet.EOF Then
1095              rsSublet.MoveFirst
1096              Do While Not rsSublet.EOF
1097                  If rsSublet("VALID") = 0 Then
1098                      adaSublet = True
1099                      MsgBox "Sublet belum di validasi"
1100                      If rsSublet.State Then rsSublet.Close
1101                      TblValid.Enabled = False
1102                      Exit Sub
1103                  End If
1104                  If Log.standarisasiSublet = 1 Then
1105                    Text(27) = Text(27) + SpyRound(CDbl(rsSublet("PRICE")) * CDbl(rsSublet("QTY")))
1106                  Else
1107                    Text(27) = Text(27) + SpyRound(IIf(Trim(Text(20)) = "", CDbl(rsSublet("HARGA")), CDbl(rsSublet("PRICE"))) * CDbl(rsSublet("QTY")))
1108                  End If
                  
1109                  rsSublet.MoveNext
1110              Loop
              ''2016-07-25 || remark sintaks insert tempjobsublet waktu panggil order di  menu validasi order.
              '            Call objO.InsertData("INSERT INTO TEMPJOBSUBLET SELECT NOPO, no_wo, no_inv, no_chassis, nofak, kodeitem, namabarang, kodesatuan, qty, harga, nonota, keterangan, FLAG, NO_POLISI, PRICE, VALID, SUPPLIER, TAGIH, TGLVALID, TGLRETUR,ESTWAKTU FROM jobsublet WITH (NOLOCK) " _
              '                 & "WHERE NOT EXISTS (SELECT * FROM tempjobsublet WITH (NOLOCK) WHERE no_wo = '" & noWO & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "') and no_wo = '" & noWO & "' AND no_chassis = '" & rsTemp("NO_CHASSIS") & "' " _
              '                 & "ORDER BY nonota, kodeitem")
1111              subSublet = CDbl(Text(27))
1112              adaSublet = True
1113              rsSublet.MoveFirst
                  'cmdSublet.Visible = True
1114              Set subletGrid.DataSource = rsSublet
1115              Call doHeadSublet
1116              subletGrid.ReBind
1117          Else
1118            If newValid Then
1119                If TblValid.Enabled Then
1120                    If MsgBox("Ada Sublet ?", vbYesNo, "Konfirmasi Sublet") = vbNo Then
1121                        adaSublet = True
1122                    End If
1123                End If
1124            End If
1125          End If
            ''End If
          
          '---------
          
          Dim rsWashing As ADODB.Recordset
1126      Set rsWashing = objO.FetchData("select washing from jobdtl where no_wo = '" & noWO & "'")
          
1127        If rsWashing.RecordCount > 0 Then
1128          If rsWashing("washing") = True Then
              'pengecekan sudah proses cuci
1129              Set rsWashing = objO.FetchData("SELECT * FROM TEMPPROGRESSCUCI WHERE NOORDER = '" & noWO & "'")
1130              If rsWashing.RecordCount > 0 Then
1131                  btnBatalWashing.Enabled = False
1132              Else
1133                  btnBatalWashing.Enabled = True
1134              End If
1135          Else
1136              btnBatalWashing.Enabled = False
1137          End If
1138        Else
1139            btnBatalWashing.Enabled = False
1140        End If
1141        Text(29) = Round(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0) '+ CDbl(text(28))
1142        Call panggilDiscSubletValidasi
          ' TAMPIL TERBILANG
1143        txtTotalRp = Angka2Huruf(Text(29), 1)
1144      If OnTopSV = 0 Then
1145          Set rsHlp = objO.FetchDataACC("select KodeLgn, SyaratBayar from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(txtKodeLgnLama.Text) & "'")
1146      Else
          'UNTUK FEBUARI - TOP PELANGGAN
1147          Set rsHlp = objO.FetchDataACC("select KodeLgn, TOPGRP as SyaratBayar from Armst WITH (NOLOCK) where KodeLgn = '" & Trim(txtKodeLgnLama.Text) & "'")
1148      End If
1149      If Not rsHlp.BOF Then
1150          Text(41) = rsHlp("SyaratBayar")
1151      End If
1152      If rsHlp.State <> adStateClosed Then rsHlp.Close
          
1153        If Not gSpart Or Not adaSublet Then
1154      TblValid.Enabled = False
1155        End If
1156        If Log.dmsNonDss = 1 Then
          'lakukan pengecekan apakah harus menginputkan DP part
          'total > plafon = dp >= prosentase * total
          '2014-12-19 RUDI KRISTANTO
          '
          '            If Trim(Text(64)) <> "" Or Trim(Text(78)) <> "" Then
          '                totalNilaiPart = getTotalPart(IIf(Trim(Text(64)) <> "", Trim(Text(64)), Trim(Text(78))), False)
          '            Else
          '                totalNilaiPart = CDbl(Trim(Text(57))) + CDbl(Trim(Text(59)))
          '            End If
1157      totalNilaiPart = 0
1158      totalNilaiPart = CDbl(Trim(Text(24))) + CDbl(Trim(Text(25)))
          
1159          If totalNilaiPart > Log.plafonDP Then
          'lakukan pengisian DP
1160              MsgBox "Nilai part membutuhkan DP minimal Rp." & Format(((Log.persentaseDP / 100) * totalNilaiPart), "0,#"), vbInformation, "Informasi DP Part"
          'boleh = False
1161          End If
1162        End If
          
          ''2017-08-01 || Rev20170801A
1163        If Log.discKacab = 1 Then
1164          If cmdBatalRequest.Enabled = True Then
          ''tombol valid dibuat disable saat request disc belum disetujui
1165              TblValid.Enabled = False
1166          End If
          
1167          sSQL = "select REP,DISC,DISCRP,FREEC from vo_joborderlockdisc WITH (NOLOCK) WHERE FS_LOG_DEL = '' and NO_WO = '" & NulSID(Text(0)) & "' AND NO_CHASSIS = '" & NulSID(Text(17)) & "'"
              Dim rsJobOrderLD As ADODB.Recordset
              Dim isSudahVJob As Boolean
1168          isSudahVJob = False
1169          Set rsJobOrderLD = objO.FetchData(sSQL)
          ''If Not rsJobOrderLD.EOF And Not rsJob.EOF Then
1170          If Not rsJobOrderLD.EOF And rsJob.State Then
1171            rsJobOrderLD.MoveFirst
1172            isSudahVJob = True
1173            If Not rsJob.EOF Then rsJob.MoveFirst
1174            Do While Not rsJobOrderLD.EOF
1175                rsJob.Find "JOB = '" & Trim(rsJobOrderLD!rep) & "'", , adSearchForward, 1
1176                If Not rsJob.EOF Then
1177                    rsJob("DISC") = rsJobOrderLD!disc
1178                    rsJob("DISCRP") = rsJobOrderLD!DISCRP
1179                    rsJob("STLDISC") = Round(((100 - CDbl(rsJob!disc)) / 100) * CDbl(rsJob!hrg))
1180                    If rsJobOrderLD!FREEC = 1 Then
1181                        rsJob!FREEC = True
1182                    Else
1183                        rsJob!FREEC = False
1184                    End If
1185                End If
1186                rsJobOrderLD.MoveNext
1187            Loop
1188          End If
1189          If rsJob.State Then
1190              If Not rsJob.EOF Then
1191                  rsJob.MoveFirst
                      Dim nilaiJob As Double
1192                  nilaiJob = 0
1193                  Do While Not rsJob.EOF
1194                      If nDuoInv Then
1195                          If rsJob!FREEC = True Then
1196                              nilaiJob = nilaiJob + CDbl(rsJob("STLDISC"))
1197                          End If
1198                      Else
1199                          nilaiJob = nilaiJob + CDbl(rsJob("STLDISC"))
1200                      End If
                          
1201                      rsJob.MoveNext
1202                  Loop
1203                  If nDuoInv Then
1204                      If isSudahVJob Then
1205                          Text(22) = nilaiJob
1206                      Else
1207                          If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(43), "DD/MM/YYYY")) >= 0 Then
1208                              Text(22) = nilaiJob
1209                          Else
1210                              Text(22) = 0
1211                          End If
1212                      End If
1213                  Else
1214                      Text(22) = nilaiJob
1215                  End If
                      
          
1216                  Text(29) = Round(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0) '+ CDbl(text(28))
1217              End If
1218          End If
1219        End If
            
            
1220          If Log.freepart = 1 And tagihanFOP <> "" And nDuoInv Then
1221              Call cekPartFOPTerlibat
1222              Call panggilFOPbyWO(NulSID(Text(0)), NulSID(Text(17)))
1223          End If
          
1224          If Log.DIDX = 1 Then
1225              Call panggilDIDXbyWO(NulSID(Text(0)), NulSID(Text(17)))
1226          End If
          
          '' validasi jika scc blm jadi order
          '' ---------------------------------------------------
          '' 20180920 || revisi pembacaan dari historynew
          '' ---------------------------------------------------
          '' 20190614 || JIKA SUDAH DIVALIDASI TIDAK PERLU CEK SSC
          '' ----------------------------------------------------------
1227        Set rsHlp = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND USERDEL = ''")
1228        If rsHlp.EOF Then
1229      If Log.lockSCC = 1 Then
          Dim rsSCC         As ADODB.Recordset
          Dim rsOrderSCC    As ADODB.Recordset
          Dim rsTungguTL    As ADODB.Recordset
          Dim isTungguTL As Boolean
1230          isTungguTL = False
1231          If Log.RPLHistori = 1 Then
1232              sSQL = "SELECT distinct A.ID_WARRANTY,a.warranty "
1233              sSQL = sSQL & " FROM [" & dbReplikasiHO & "]..ho_warranty A WITH (NOLOCK)"
1234              sSQL = sSQL & " INNER JOIN [" & dbReplikasiHO & "]..ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = ''"
1235              sSQL = sSQL & " WHERE no_rangka1 <= '" & NulSID(Text(17)) & "' AND no_rangka2 >= '" & NulSID(Text(17)) & "'"
          
1236          Else
1237              sSQL = "SELECT distinct A.ID_WARRANTY,a.warranty "
1238              sSQL = sSQL & " FROM ho_warranty A WITH (NOLOCK)"
1239              sSQL = sSQL & " INNER JOIN ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = ''"
1240              sSQL = sSQL & " WHERE no_rangka1 <= '" & NulSID(Text(17)) & "' AND no_rangka2 >= '" & NulSID(Text(17)) & "'"
          
1241          End If
1242          Set rsSCC = objO.FetchData(sSQL)
1243          If Not rsSCC.EOF Then
1244              rsSCC.MoveFirst
1245              Do While Not rsSCC.EOF
1246                  If Log.RPLHistori = 1 Then
          ''20190315 | MB EVY | SUDAH JADI PREDIAGNOSE BAIK OK / NOK (HARUS ADA ISI NOCLAIM TOASW) BISA VALIDASI ORDER REGULER.
          ''20190408 | SERVICE | WO Reguler  tidak bisa di validasi oleh SA apabila WO Prediagnose belum di validasi oleh TL (Baik validasi OK maupun NOK)
          
1247                      If Log.newLockSSC = 1 Then
1248                        Set rsOrderSCC = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HISTORYnew WHERE KODETIPE = '" & NulSID(rsSCC!Id_warranty) & "' AND tipewo = 'W' AND NO_CHASSIS = '" & NulSID(Text(17)) & "'")
1249                        If rsOrderSCC.EOF Then
          'Set rsOrderSCC = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(text(17)) & "' AND KODETIPE = '" & NulSID(rsSCC!Id_warranty) & "' AND VALIDTWC <> 'NOK'")
                                
                                ''Set rsOrderSCC = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(Text(17)) & "' AND KODETIPE = '" & NulSID(rsSCC!Id_warranty) & "'")
                                ''20210604 || req mb thania disamakan dengan sebelumnya ''20200316
                                ''20200316 || pergantian req mb thania
1250                            Set rsOrderSCC = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(Text(17)) & "' AND KODETIPE = '" & NulSID(rsSCC!Id_warranty) & "' AND VALIDTWC <> ''")
1251                            If rsOrderSCC.EOF Then
          
1252                                MsgBox "SSC BELUM DIKERJAKAN" & vbCrLf & NulSID(rsSCC!warranty), vbInformation
1253                                Label47.Caption = "SSC BELUM DIKERJAKAN, " & NulSID(rsSCC!warranty)
1254                                TblValid.Enabled = False
1255                                tblBatal.Enabled = False
1256                            Else
          ''jika
1257                                If NulSID(rsOrderSCC!validtwc) = "OKE" And NulSID(rsOrderSCC!no_wo) = "" Then
          '' 20190711 request mb thania sampai validasi TL
          '                              MsgBox "PREDIAGNOSE " & NulSID(rsOrderSCC!no_diag) & vbCrLf & "BELUM DIJADIKAN WO WARRANTY", vbInformation
          '                              Label47.Caption = "SSC BELUM DIKERJAKAN, " & NulSID(rsSCC!warranty)
          '                              TblValid.Enabled = False
          '                              TblBatal.Enabled = False
1258                                ElseIf NulSID(rsOrderSCC!validtwc) = "" Then
1259                                    If NulSID(rsOrderSCC!flagtl) = 0 Then
1260                                      isTungguTL = True
1261                                      MsgBox "PREDIAGNOSE " & NulSID(rsOrderSCC!no_diag) & vbCrLf & "BELUM DIVALIDASI TL.", vbInformation
1262                                      Label47.Caption = "SSC BELUM DIVALIDASI, " & NulSID(rsSCC!warranty)
1263                                      TblValid.Enabled = False
1264                                      tblBatal.Enabled = False
1265                                    Else
1266                                      Set rsTungguTL = objO.FetchData("select DATEADD(day,2,tglEntryTungguTL) ,* from DIAGDTa where NO_CHASSIS = '" & NulSID(Text(17)) & "' AND KODETIPE = '" & NulSID(rsSCC!Id_warranty) & "' and flagtl = 1 and VALIDTWC = '' and GETDATE() > DATEADD(day,2,tglEntryTungguTL)")
       
1267                                          If Not rsTungguTL.EOF Then
1268                                              MsgBox "PREDIAGNOSE " & NulSID(rsOrderSCC!no_diag) & vbCrLf & " LEBIH DARI 2 HARI BELUM VALIDASI TL.", vbInformation
1269                                              Label47.Caption = "SSC BELUM DIVALIDASI, " & NulSID(rsSCC!warranty)
1270                                              TblValid.Enabled = False
1271                                              tblBatal.Enabled = False
1272                                          End If
          '                                End If
1273                                    End If
          
1274                                End If
1275                            End If
1276                        End If
1277                      Else
1278                        Set rsOrderSCC = objO.FetchData("SELECT * FROM [" & dbReplikasiHO & "]..HISTORYnew WHERE KODETIPE = '" & NulSID(rsSCC!Id_warranty) & "' AND tipewo = 'W' AND NO_CHASSIS = '" & NulSID(Text(17)) & "'")
1279                        If rsOrderSCC.EOF Then
1280                            Set rsOrderSCC = objO.FetchData("SELECT * FROM DIAGDTA WITH (NOLOCK) WHERE NO_CHASSIS = '" & NulSID(Text(17)) & "' AND KODETIPE = '" & NulSID(rsSCC!Id_warranty) & "' AND VALIDTWC <> ''")
1281                            If rsOrderSCC.EOF Then
          
1282                                MsgBox "SCC BELUM DIKERJAKAN" & vbCrLf & NulSID(rsSCC!warranty), vbInformation
1283                                Label47.Caption = "SSC BELUM DIKERJAKAN, " & NulSID(rsSCC!warranty)
1284                                TblValid.Enabled = False
1285                                tblBatal.Enabled = False
1286                            End If
1287                        End If
1288                      End If
1289                  End If
          
          
1290                  rsSCC.MoveNext
1291              Loop
          
          ''---------------------------------
          '' 2019-05-27
          ''---------------------------------
1292              If isTungguTL And Log.newLockSSC = 1 Then
1293                sSQL = "select * from [" & dbReplikasiHO & "]..oplocknormal where kodedealer = '" & Log.KodeDealer & "' and flagdel = '0' and hari = '" & IIf(Weekday(Format(isnow, "dd/mm/yyyy")) = 1, 6, Weekday(Format(isnow, "dd/mm/yyyy")) - 2) & NulSID(WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)) & "' AND (('" & Format(isnow, "HH:NN") & "' between jammulai1 AND JAMSELESAI1) OR ('" & Format(isnow, "HH:NN") & "' between jammulai2 AND JAMSELESAI2))"
1294                Set rsTemp = objO.FetchData(sSQL)
1295                If Not rsTemp.EOF Then
1296                    cmdTungguValidasiTL.Enabled = True
1297                Else
1298                    sSQL = "select * from [" & dbReplikasiHO & "]..oplockkhusus where kodedealer = '" & Log.KodeDealer & "' and flagdel = '0' and hari = '" & IIf(Weekday(Format(isnow, "dd/mm/yyyy")) = 1, 6, Weekday(Format(isnow, "dd/mm/yyyy")) - 2) & NulSID(WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)) & "' AND (('" & Format(isnow, "HH:NN") & "' between jammulai1 AND JAMSELESAI1) OR ('" & Format(isnow, "HH:NN") & "' between jammulai2 AND JAMSELESAI2))"
1299                    Set rsTemp = objO.FetchData(sSQL)
1300                    If Not rsTemp.EOF Then
1301                        cmdTungguValidasiTL.Enabled = True
1302                    End If
1303                End If
1304              End If
                  
          
1305          End If
1306        End If
1307          End If
1308      End If
1309         On Error GoTo 0
1310         Exit Sub
          
panggilWO_Error:
          
1311          MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure panggilWO of Form frmOrderValidasi")

       
End Sub

Private Sub cekPSMInfo(noKa As String)
    Dim rsHlp As ADODB.Recordset
    cmdPsmInfo.Enabled = False
        sSQL = "select b.psm_item_category,CASE b.psm_item_category WHEN '01' THEN 'Auxiliary battery (charging status)' WHEN '02' THEN 'Auxiliary battery (degraded state)'"
        sSQL = sSQL & " WHEN '03' THEN 'Engine oil degradation (Phase2)' WHEN '04' THEN 'Key battery power status' ELSE 'Unknown' END AS [deskripsi]"
        sSQL = sSQL & " from ReplBengkel99..p2_psm a with (nolock) inner join ReplBengkel99..p2_psm_item b on a.vin=b.vin"
        sSQL = sSQL & " where a.vin='" & Nul(noKa) & "' and a.vin not in ( "
        sSQL = sSQL & " select a.vin from ReplBengkel99..P2_MonitorSVCProgress a"
        sSQL = sSQL & " inner join ReplBengkel99..P2_MonitorSvcProgress_psm b on b.wo_number=a.wo_number)"
        Set rsHlp = objO.FetchData(sSQL)
        If Not rsHlp.EOF Then
            cmdPsmInfo.Enabled = True
        End If
End Sub

Private Sub hapusForm()
1            On Error GoTo hapusForm_Error
          
          'Call objO.DeleteData("DELETE TEMPJOBSUBLET WHERE NO_WO = '" & text(0) & "'")
          
2             For i = 0 To 44
3                 Select Case i
      Case 1, 11, 43
4                         Text(i).Mask = "99-99-9999"
5                     Case 10, 21, 22, 23, 24, 25, 26, 27, 29, 39
6                         Text(i) = 0
7                     Case 42, 30, 33
8                     Case Else
9                         Text(i) = ""
10                End Select
11            Next
          
12            lblFOP.Caption = ""
13            lblNilaiMax.Caption = ""
14            tagihanFOP = ""
15            kodeMDL = ""
16            nilaiMaxPartNRM = 0
17            partTerpakaiNRM = 0
          
18            VScroll1.value = 0
19            opWO(0).value = True
          
20            If rsJob.State Then rsJob.Close
21            Set rsJob = Nothing
22            Set JobGrid.DataSource = rsJob
          'JobGrid.ReBind
          
23            If rsPart.State Then rsPart.Close
24            Set rsPart = Nothing
25            Set partGrid.DataSource = rsPart
          'partGrid.ReBind
          
26            Set subletGrid.DataSource = Nothing
27            adaAddJob = False
          
          ' ISI KARAKTER PELANGGAN
28            Set rsTemp = objO.FetchData("SELECT * FROM T_Karakter WITH (NOLOCK) ORDER BY KODE")
29            If Not rsTemp.EOF Then
30                cboPelanggan.Clear
31                Do While Not rsTemp.EOF
32                    cboPelanggan.AddItem (rsTemp("KODE"))
33                    rsTemp.MoveNext
34                Loop
35            End If
36            cboPelanggan.ListIndex = 0
          
37            vTipeCustomer = ""
38            woStat = ""
39            nDuoInv = False
40            chPS.value = Unchecked
41            chNea.value = Unchecked
42            txtKodeLgnLama.Text = ""
          
43            If Log.lockSCC = 1 And Log.newLockSSC = 1 Then
44                cmdTungguValidasiTL.Enabled = False
45                If Log.discKacab = 0 Then
46                    cmdTungguValidasiTL.Left = cmdBatalRequest.Left
47                End If
          
48            Else
49                cmdTungguValidasiTL.Visible = False
50            End If
51            If Log.CR7 = 1 Then
52              txtCatatanCR7.Text = ""
53            End If
54           On Error GoTo 0
55           Exit Sub
          
hapusForm_Error:
          
56            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frmOrderValidasi")
End Sub

Private Sub btnBatalWashing_Click()
1            On Error GoTo btnBatalWashing_Click_Error
    
2             If Trim(Text(0)) <> "" Then
3                  If Log.lockWashing = 1 Then
4                     If MsgBox("Apakah Anda yakin unit tidak dicuci?" & vbCrLf & "Jika Yes maka pilihan anda akan direport ke HO!", vbYesNo, "") = vbNo Then
5                        Exit Sub
6                     End If
7                  End If
    
8                 Set rsTmp = objO.FetchData("SELECT FLAG,JAMMULAI,NOSTALL,JAMSELESAI FROM stall WITH (NOLOCK) WHERE KETERANGAN like '%|" & Trim(Text(0)) & "|%'")
9                 If Not rsTmp.EOF Then
    ' TUTUP TAMBAH PEKERJAAN
10                    If UCase(rsTmp("FLAG")) = "G" Then
11                        framePembatalanWashing.Visible = True
12                    End If
13                End If
14            End If
    
15           On Error GoTo 0
16           Exit Sub
    
btnBatalWashing_Click_Error:
    
17            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure btnBatalWashing_Click of Form frmOrderValidasi")
End Sub

Private Sub btnPilihFoto_Click(Index As Integer)
    frameFoto(Index).Visible = True
End Sub

Private Sub btnSimpan_Click()
1            On Error GoTo btnSimpan_Click_Error
    
2             Call objO.UpdateData("Update Jobdta Set Catat = '" & Trim(txt.Text) & "' Where NO_WO = '" & Trim(Text(0).Text) & "'")
3             MsgBox "Catatan telah disimpan !", vbInformation + vbOKOnly
    
4            On Error GoTo 0
5            Exit Sub
    
btnSimpan_Click_Error:
    
6             MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure btnSimpan_Click of Form frmOrderValidasi")
End Sub

Private Sub btnUploadKTP_Click()
    frameUploadKTP.Visible = True
    Call refreshCustomerNonKTP
End Sub

Private Sub chNea_Click()
    If chNea.value = 1 Then
        txtInfoNea.Visible = True
        txtInfoNea.Text = ""
    Else
        txtInfoNea.Visible = False
    End If
End Sub

Private Sub chNea_GotFocus()
    VScroll1.value = 2500
End Sub

Private Sub chPS_GotFocus()
    VScroll1.value = 2500
End Sub

Private Sub cmdBatalRequest_Click()
    If Trim(Text(0)) = "" Then Exit Sub
    If MsgBox("Apakah yakin akan membatalkan request disc?", vbYesNo) = vbYes Then
        Set rsTemp = objO.FetchData("SELECT userValid FROM RequestDisc WITH (NOLOCK) WHERE uservalid = '' and no_wo = '" & Trim(Text(0)) & "' ")
        If Not rsTemp.EOF Then
            ''Call objO.InsertData("UPDATE RequestDisc SET tglValid = getdate(),userValid = 'BTL',status = '" & gSA & "' WHERE no_wo = '" & Trim(Text(0)) & "' ")
            Call objO.InsertData("UPDATE RequestDisc SET tglValid = getdate(),status = 'BTL', userValid = '" & gSA & "' WHERE no_wo = '" & Trim(Text(0)) & "' ")
            MsgBox "Request disc berhasil dibatalkan", vbInformation
            Call panggilWO(Trim(Text(0)))
        End If
    End If
End Sub

Private Sub cmdEditCustomer_Click()
    If Not rsKTP.EOF Then
        rsKTP.MoveFirst
        rsKTP.filter = "pilih=" & True
        If Not rsKTP.EOF Then
            If Log.monitoringScoring = 0 Then
                With frmSIDMasterPelangganNonMerge
                    .txt(1) = NulSID(rsKTP!idcust)
                    .isEdit (True)
                    .show vbModal
                End With
            Else
                With frmSIDMasterPelanggan
                    .txt(1) = NulSID(rsKTP!idcust)
                    .isEdit (True)
                    .show vbModal
                End With
            End If
        Else
            MsgBox "Silahkan pilih data yang akan diubah.", vbInformation
        End If
    End If
    Call refreshCustomerNonKTP
End Sub

Private Sub cmdKeluh_Click()
          Dim rsSearch    As New ADODB.Recordset
          Dim rscr7       As ADODB.Recordset
          Dim streamx     As New ADODB.Stream
1            On Error GoTo cmdKeluh_Click_Error
          
2             If Log.dmsNonDss = 1 Then
3                 txtPermintaan = ""
4                 txtKeluhan = ""
5                 txtFI = ""
6                 txtNextJob = ""
          
7                 If Log.CR7 = 1 Then
8                     If FrameKeluhDmsNonDSS.Visible Then
9                         FrameKeluhDmsNonDSS.Visible = False
10                    Else
11                        Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(Text(0)) & "'")
12                        If Not rsSearch.EOF Then
13                            Text3.Text = Trim(rsSearch("REQ1")) & vbCrLf & Trim(rsSearch("REQ2")) & vbCrLf & Trim(rsSearch("REQ3")) & vbCrLf & Trim(rsSearch("REQ4")) & vbCrLf & Trim(rsSearch("REQ5"))
14                            Text2.Text = Trim(rsSearch("CMPL1")) & vbCrLf & Trim(rsSearch("CMPL2")) & vbCrLf & Trim(rsSearch("CMPL3")) & vbCrLf & Trim(rsSearch("CMPL4")) & vbCrLf & Trim(rsSearch("CMPL5"))
15                            txtCatatanCR7.Text = Trim(rsSearch!CATAT)
16                        End If
          ''20180515REV01'INSPECTION
17                        Set rsSearch = objO.FetchData("SELECT top 1 * FROM INSPECTION A WITH (NOLOCK) WHERE NOWO = '" & Trim(Text(0)) & "' order by tglentry desc")
18                        If Not rsSearch.EOF Then
          
19                            txtFI = Trim(rsSearch("HASILFI")) & vbCrLf & Trim(rsSearch("HasilTestDrive"))
          
          '' ==============================
          '' 20180515DEV01'INSPECTION    ==
          '' ==============================
20                            Command1.Top = 6360
21                            FrameKeluhDmsNonDSS.Height = 6855
          Dim rateJob As Double
          Dim i As Integer
22                            i = 1
23                            For i = 0 To 6
24                                txtNo(i).Visible = True
25                                txtSaran(i).Visible = True
26                                txtJobOrder(i).Visible = True
27                                txtTglKembali(i).Visible = True
28                                txtJasa(i).Visible = True
29                                txtPart(i).Visible = True
30                                txtSublet(i).Visible = True
31                                txtBiaya(i).Visible = True
32                                btnPilihFoto(i).Visible = True
          
33                                cdgopen(i).FileName = ""
34                                imgFotoCR7(i).Picture = Nothing
35                                frameFoto(i).Visible = False
36                            Next
37                            i = 0
38                            rateJob = 0
39                            If NulSID(rsSearch!snj1) <> "" Then
          
40                                If NulSID(rsSearch!AdaFoto1) = True Then
41                                    Set streamx = New ADODB.Stream
42                                    With streamx
43                                        .Type = adTypeBinary
44                                        .Open
45                                        .Write rsSearch!Pic1
46                                        .SaveToFile App.Path & "\Temp" & i & ".bmp", adSaveCreateOverWrite
47                                        imgFotoCR7(i).Picture = LoadPicture(App.Path & "\Temp" & i & ".bmp")
48                                        cdgopen(i).FileName = App.Path & "\Temp" & i & ".bmp"
49                                    End With
50                                    Set streamx = Nothing
51                                End If
52                                If InStr(1, NulSID(rsSearch!snj1), "|") > 0 Then
53                                    rateJob = Mid(NulSID(rsSearch!snj1), InStr(1, NulSID(rsSearch!snj1), "|") + 1, Len(NulSID(rsSearch!snj1)))
54                                End If
55                                txtNo(0).Text = "1."
56                                txtSaran(0).Text = NulSID(rsSearch!sp1)
57                                txtJobOrder(0).Text = NulSID(rsSearch!snj1)
58                                txtTglKembali(0).Text = Format(NulSID(rsSearch!estn1), "dd-mm-yyyy")
59                                txtJasa(0).Text = IIf(Round(NulSID(rsSearch!ests1)) = 0, hitungLabour(rateJob, NulSID(Text(13))), Round(NulSID(rsSearch!ests1)))
60                                txtPart(0).Text = Round(NulSID(rsSearch!estp1)) '"0"
61                                txtSublet(0).Text = Round(NulSID(rsSearch!ests1)) '"0"
62                                txtBiaya(0).Text = CDbl(txtJasa(0)) + CDbl(txtPart(0)) + CDbl(txtSublet(0))
63                            Else
64                                txtNo(0).Visible = False
65                                txtSaran(0).Visible = False
66                                txtJobOrder(0).Visible = False
67                                txtTglKembali(0).Visible = False
68                                txtJasa(0).Visible = False
69                                txtPart(0).Visible = False
70                                txtSublet(0).Visible = False
71                                txtBiaya(0).Visible = False
72                                btnPilihFoto(i).Visible = False
73                                Command1.Top = Command1.Top - 200
74                                FrameKeluhDmsNonDSS.Height = FrameKeluhDmsNonDSS.Height - 200
75                            End If
          
          '' 2
76                            i = 1
77                            rateJob = 0
78                            If NulSID(rsSearch!snj2) <> "" Then
79                                If NulSID(rsSearch!AdaFoto2) = True Then
80                                    Set streamx = New ADODB.Stream
81                                    With streamx
82                                        .Type = adTypeBinary
83                                        .Open
84                                        .Write (rsSearch!Pic2)
85                                        .SaveToFile App.Path & "\Temp" & i & ".bmp", adSaveCreateOverWrite
86                                        imgFotoCR7(i).Picture = LoadPicture(App.Path & "\Temp" & i & ".bmp")
87                                        cdgopen(i).FileName = App.Path & "\Temp" & i & ".bmp"
88                                    End With
89                                    Set streamx = Nothing
90                                End If
91                                If InStr(1, NulSID(rsSearch!snj2), "|") > 0 Then
92                                    rateJob = Mid(NulSID(rsSearch!snj2), InStr(1, NulSID(rsSearch!snj2), "|") + 1, Len(NulSID(rsSearch!snj2)))
93                                End If
94                                txtNo(i).Text = "2."
95                                txtSaran(i).Text = NulSID(rsSearch!sp2)
96                                txtJobOrder(i).Text = NulSID(rsSearch!snj2)
97                                txtTglKembali(i).Text = Format(NulSID(rsSearch!estn2), "dd-mm-yyyy")
98                                txtJasa(i).Text = hitungLabour(rateJob, NulSID(Text(13)))
99                                txtPart(i).Text = Round(NulSID(rsSearch!estp2)) '"0"
100                               txtSublet(i).Text = Round(NulSID(rsSearch!ests2)) '"0"
101                               txtBiaya(i).Text = CDbl(txtJasa(i)) + CDbl(txtPart(i)) + CDbl(txtSublet(i))
102                           Else
103                               txtNo(i).Visible = False
104                               txtSaran(i).Visible = False
105                               txtJobOrder(i).Visible = False
106                               txtTglKembali(i).Visible = False
107                               txtJasa(i).Visible = False
108                               txtPart(i).Visible = False
109                               txtSublet(i).Visible = False
110                               txtBiaya(i).Visible = False
111                               FrameKeluhDmsNonDSS.Height = FrameKeluhDmsNonDSS.Height - 200
112                               btnPilihFoto(i).Visible = False
113                               Command1.Top = Command1.Top - 200
114                           End If
          '' akhir 2
          
          '' 3
115                           i = 2
116                           rateJob = 0
117                           If NulSID(rsSearch!snj3) <> "" Then
118                               If NulSID(rsSearch!AdaFoto3) = True Then
119                                   Set streamx = New ADODB.Stream
120                                   With streamx
121                                       .Type = adTypeBinary
122                                       .Open
123                                       .Write (rsSearch!Pic3)
124                                       .SaveToFile App.Path & "\Temp" & i & ".bmp", adSaveCreateOverWrite
125                                       imgFotoCR7(i).Picture = LoadPicture(App.Path & "\Temp" & i & ".bmp")
126                                       cdgopen(i).FileName = App.Path & "\Temp" & i & ".bmp"
127                                   End With
128                                   Set streamx = Nothing
129                               End If
130                               If InStr(1, NulSID(rsSearch!snj3), "|") > 0 Then
131                                   rateJob = Mid(NulSID(rsSearch!snj3), InStr(1, NulSID(rsSearch!snj3), "|") + 1, Len(NulSID(rsSearch!snj3)))
132                               End If
133                               txtNo(i).Text = "3."
134                               txtSaran(i).Text = NulSID(rsSearch!sp3)
135                               txtJobOrder(i).Text = NulSID(rsSearch!snj3)
136                               txtTglKembali(i).Text = Format(NulSID(rsSearch!estn3), "dd-mm-yyyy")
137                               txtJasa(i).Text = hitungLabour(rateJob, NulSID(Text(13)))
138                               txtPart(i).Text = Round(NulSID(rsSearch!estp3)) '"0"
139                               txtSublet(i).Text = Round(NulSID(rsSearch!ests3)) '"0"
140                               txtBiaya(i).Text = CDbl(txtJasa(i)) + CDbl(txtPart(i)) + CDbl(txtSublet(i))
141                           Else
142                               txtNo(i).Visible = False
143                               txtSaran(i).Visible = False
144                               txtJobOrder(i).Visible = False
145                               txtTglKembali(i).Visible = False
146                               txtJasa(i).Visible = False
147                               txtPart(i).Visible = False
148                               txtSublet(i).Visible = False
149                               txtBiaya(i).Visible = False
150                               FrameKeluhDmsNonDSS.Height = FrameKeluhDmsNonDSS.Height - 200
151                               btnPilihFoto(i).Visible = False
152                               Command1.Top = Command1.Top - 200
153                           End If
          '' akhir 3
          
          '' 4
154                           i = 3
155                           rateJob = 0
156                           If NulSID(rsSearch!snj4) <> "" Then
157                               If NulSID(rsSearch!AdaFoto4) = True Then
158                                   Set streamx = New ADODB.Stream
159                                   With streamx
160                                       .Type = adTypeBinary
161                                       .Open
162                                       .Write (rsSearch!Pic4)
163                                       .SaveToFile App.Path & "\Temp" & i & ".bmp", adSaveCreateOverWrite
164                                       imgFotoCR7(i).Picture = LoadPicture(App.Path & "\Temp" & i & ".bmp")
165                                       cdgopen(i).FileName = App.Path & "\Temp" & i & ".bmp"
166                                   End With
167                                   Set streamx = Nothing
168                               End If
169                               If InStr(1, NulSID(rsSearch!snj4), "|") > 0 Then
170                                   rateJob = Mid(NulSID(rsSearch!snj4), InStr(1, NulSID(rsSearch!snj4), "|") + 1, Len(NulSID(rsSearch!snj4)))
171                               End If
172                               txtNo(i).Text = "4."
173                               txtSaran(i).Text = NulSID(rsSearch!sp4)
174                               txtJobOrder(i).Text = NulSID(rsSearch!snj4)
175                               txtTglKembali(i).Text = Format(NulSID(rsSearch!estn4), "dd-mm-yyyy")
176                               txtJasa(i).Text = hitungLabour(rateJob, NulSID(Text(13)))
177                               txtPart(i).Text = Round(NulSID(rsSearch!estp4)) '"0"
178                               txtSublet(i).Text = Round(NulSID(rsSearch!ests4)) '"0"
179                               txtBiaya(i).Text = CDbl(txtJasa(i)) + CDbl(txtPart(i)) + CDbl(txtSublet(i))
180                           Else
181                               txtNo(i).Visible = False
182                               txtSaran(i).Visible = False
183                               txtJobOrder(i).Visible = False
184                               txtTglKembali(i).Visible = False
185                               txtJasa(i).Visible = False
186                               txtPart(i).Visible = False
187                               txtSublet(i).Visible = False
188                               txtBiaya(i).Visible = False
189                               FrameKeluhDmsNonDSS.Height = FrameKeluhDmsNonDSS.Height - 200
190                               btnPilihFoto(i).Visible = False
191                               Command1.Top = Command1.Top - 200
192                           End If
          '' akhir 4
          
          '' 5
193                           i = 4
194                           rateJob = 0
195                           If NulSID(rsSearch!snj5) <> "" Then
196                               If NulSID(rsSearch!AdaFoto5) = True Then
197                                   Set streamx = New ADODB.Stream
198                                   With streamx
199                                       .Type = adTypeBinary
200                                       .Open
201                                       .Write (rsSearch!Pic5)
202                                       .SaveToFile App.Path & "\Temp" & i & ".bmp", adSaveCreateOverWrite
203                                       imgFotoCR7(i).Picture = LoadPicture(App.Path & "\Temp" & i & ".bmp")
204                                       cdgopen(i).FileName = App.Path & "\Temp" & i & ".bmp"
205                                   End With
206                                   Set streamx = Nothing
207                               End If
208                               If InStr(1, NulSID(rsSearch!snj5), "|") > 0 Then
209                                   rateJob = Mid(NulSID(rsSearch!snj5), InStr(1, NulSID(rsSearch!snj5), "|") + 1, Len(NulSID(rsSearch!snj5)))
210                               End If
211                               txtNo(i).Text = "5."
212                               txtSaran(i).Text = NulSID(rsSearch!SP5)
213                               txtJobOrder(i).Text = NulSID(rsSearch!snj5)
214                               txtTglKembali(i).Text = Format(NulSID(rsSearch!EstN5), "dd-mm-yyyy")
215                               txtJasa(i).Text = hitungLabour(rateJob, NulSID(Text(13)))
216                               txtPart(i).Text = Round(NulSID(rsSearch!EstP5)) '"0"
217                               txtSublet(i).Text = Round(NulSID(rsSearch!EstS5)) ' "0"
218                               txtBiaya(i).Text = CDbl(txtJasa(i)) + CDbl(txtPart(i)) + CDbl(txtSublet(i))
219                           Else
220                               txtNo(i).Visible = False
221                               txtSaran(i).Visible = False
222                               txtJobOrder(i).Visible = False
223                               txtTglKembali(i).Visible = False
224                               txtJasa(i).Visible = False
225                               txtPart(i).Visible = False
226                               txtSublet(i).Visible = False
227                               txtBiaya(i).Visible = False
228                               FrameKeluhDmsNonDSS.Height = FrameKeluhDmsNonDSS.Height - 200
229                               btnPilihFoto(i).Visible = False
230                               Command1.Top = Command1.Top - 200
231                           End If
          '' akhir 5
          
          '' 6
232                           i = 5
233                           rateJob = 0
234                           If NulSID(rsSearch!snj6) <> "" Then
235                               If NulSID(rsSearch!AdaFoto6) = True Then
236                                   Set streamx = New ADODB.Stream
237                                   With streamx
238                                       .Type = adTypeBinary
239                                       .Open
240                                       .Write (rsSearch!Pic6)
241                                       .SaveToFile App.Path & "\Temp.bmp", adSaveCreateOverWrite
242                                       imgFotoCR7(i).Picture = LoadPicture(App.Path & "\Temp.bmp")
243                                       cdgopen(i).FileName = App.Path & "\Temp" & i & ".bmp"
244                                   End With
245                                   Set streamx = Nothing
246                               End If
247                               If InStr(1, NulSID(rsSearch!snj6), "|") > 0 Then
248                                   rateJob = Mid(NulSID(rsSearch!snj6), InStr(1, NulSID(rsSearch!snj6), "|") + 1, Len(NulSID(rsSearch!snj6)))
249                               End If
250                               txtNo(i).Text = "6."
251                               txtSaran(i).Text = NulSID(rsSearch!SP6)
252                               txtJobOrder(i).Text = NulSID(rsSearch!snj6)
253                               txtTglKembali(i).Text = Format(NulSID(rsSearch!EstN6), "dd-mm-yyyy")
254                               txtJasa(i).Text = hitungLabour(rateJob, NulSID(Text(13)))
255                               txtPart(i).Text = Round(NulSID(rsSearch!EstP6)) '"0"
256                               txtSublet(i).Text = Round(NulSID(rsSearch!EstS6)) '= "0"
257                               txtBiaya(i).Text = CDbl(txtJasa(i)) + CDbl(txtPart(i)) + CDbl(txtSublet(i))
258                           Else
259                               txtNo(i).Visible = False
260                               txtSaran(i).Visible = False
261                               txtJobOrder(i).Visible = False
262                               txtTglKembali(i).Visible = False
263                               txtJasa(i).Visible = False
264                               txtPart(i).Visible = False
265                               txtSublet(i).Visible = False
266                               txtBiaya(i).Visible = False
267                               FrameKeluhDmsNonDSS.Height = FrameKeluhDmsNonDSS.Height - 200
268                               btnPilihFoto(i).Visible = False
269                               Command1.Top = Command1.Top - 200
270                           End If
          '' akhir 6
          
          '' 7
271                           i = 6
272                           rateJob = 0
273                           If NulSID(rsSearch!snj7) <> "" Then
274                               If NulSID(rsSearch!AdaFoto7) = True Then
275                                   Set streamx = New ADODB.Stream
276                                   With streamx
277                                       .Type = adTypeBinary
278                                       .Open
279                                       .Write (rsSearch!Pic7)
280                                       .SaveToFile App.Path & "\Temp" & i & ".bmp", adSaveCreateOverWrite
281                                       imgFotoCR7(i).Picture = LoadPicture(App.Path & "\Temp" & i & ".bmp")
282                                       cdgopen(i).FileName = App.Path & "\Temp" & i & ".bmp"
283                                   End With
284                                   Set streamx = Nothing
285                               End If
286                               If InStr(1, NulSID(rsSearch!snj7), "|") > 0 Then
287                                   rateJob = Mid(NulSID(rsSearch!snj7), InStr(1, NulSID(rsSearch!snj7), "|") + 1, Len(NulSID(rsSearch!snj7)))
288                               End If
289                               txtNo(i).Text = "7."
290                               txtSaran(i).Text = NulSID(rsSearch!SP7)
291                               txtJobOrder(i).Text = NulSID(rsSearch!snj7)
292                               txtTglKembali(i).Text = Format(NulSID(rsSearch!EstN7), "dd-mm-yyyy")
293                               txtJasa(i).Text = hitungLabour(rateJob, NulSID(Text(13)))
294                               txtPart(i).Text = Round(NulSID(rsSearch!EstP7)) '"0"
295                               txtSublet(i).Text = Round(NulSID(rsSearch!EstS7)) '"0"
296                               txtBiaya(i).Text = CDbl(txtJasa(i)) + CDbl(txtPart(i)) + CDbl(txtSublet(i))
297                           Else
298                               txtNo(i).Visible = False
299                               txtSaran(i).Visible = False
300                               txtJobOrder(i).Visible = False
301                               txtTglKembali(i).Visible = False
302                               txtJasa(i).Visible = False
303                               txtPart(i).Visible = False
304                               txtSublet(i).Visible = False
305                               txtBiaya(i).Visible = False
306                               FrameKeluhDmsNonDSS.Height = FrameKeluhDmsNonDSS.Height - 200
307                               btnPilihFoto(i).Visible = False
308                               Command1.Top = Command1.Top - 200
309                           End If
          '' akhir  7
          'txtNextJob = Trim(rsSearch("SNJ1")) & vbCrLf & Trim(rsSearch("SNJ2")) & vbCrLf & Trim(rsSearch("SNJ3")) & vbCrLf & Trim(rsSearch("SNJ4")) & vbCrLf & Trim(rsSearch("SNJ5")) & vbCrLf & Trim(rsSearch("SNJ6")) & vbCrLf & Trim(rsSearch("SNJ7"))
310                       End If
311                       FrameKeluhDmsNonDSS.Visible = True
312                   End If
313                   Set rscr7 = objO.FetchData("Select Catat FROM JOBDTA WITH (NOLOCK) Where NO_WO = '" & Trim(Text(0)) & "'")
314                   If rscr7.RecordCount > 0 Then
315                       Text1.Text = rscr7.Fields(0).value
316                       txtNextJob = NulSID(rscr7.Fields(0).value)
317                   Else
318                       Text1.Text = ""
319                       txtNextJob = ""
320                   End If
321               Else
          
322                   If frameNonCR7.Visible Then
323                       frameNonCR7.Visible = False
324                   Else
325                       Text12.Text = "" ' txtPermintaan = ""
326                       Text13.Text = "" ' txtKeluhan = ""
327                       Text16.Text = "" ' txtFI = ""
328                       Text15.Text = "" ' txtNextJob = ""
          
329                       Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(Text(0)) & "'")
330                       If Not rsSearch.EOF Then
331                           Text12.Text = Trim(rsSearch("REQ1")) & vbCrLf & Trim(rsSearch("REQ2")) & vbCrLf & Trim(rsSearch("REQ3")) & vbCrLf & Trim(rsSearch("REQ4")) & vbCrLf & Trim(rsSearch("REQ5"))
332                           Text13.Text = Trim(rsSearch("CMPL1")) & vbCrLf & Trim(rsSearch("CMPL2")) & vbCrLf & Trim(rsSearch("CMPL3")) & vbCrLf & Trim(rsSearch("CMPL4")) & vbCrLf & Trim(rsSearch("CMPL5"))
333                       End If
          
334                       Set rsSearch = objO.FetchData("SELECT TOP 1 * FROM INSPECTION A WITH (NOLOCK) WHERE NOWO = '" & Trim(Text(0)) & "' ORDER BY TglEntry desc")
335                       If Not rsSearch.EOF Then
336                           Text16 = Trim(rsSearch("HASILFI")) & vbCrLf & Trim(rsSearch("HasilTestDrive"))
337                           Text15 = Trim(rsSearch("SNJ1")) & vbCrLf & Trim(rsSearch("SNJ2")) & vbCrLf & Trim(rsSearch("SNJ3")) & vbCrLf & Trim(rsSearch("SNJ4")) & vbCrLf & Trim(rsSearch("SNJ5")) & vbCrLf & Trim(rsSearch("SNJ6")) & vbCrLf & Trim(rsSearch("SNJ7"))
338                       End If
          
339                       Set rscr7 = objO.FetchData("Select Catat FROM JOBDTA WITH (NOLOCK) Where NO_WO = '" & Trim(Text(0)) & "'")
340                       If rscr7.RecordCount > 0 Then
341                           txtCatatNonCR7.Text = rscr7.Fields(0).value
          'Text15 = NulSID(rscr7.Fields(0).value)
342                       Else
343                           txtCatatNonCR7.Text = ""
          ' Text15 = ""
344                       End If
345                       frameNonCR7.Visible = True
346                   End If
347               End If
348           Else
          
349               txtPermintaan = ""
350               txtKeluhan = ""
351               If framKeluh.Visible Then
352                   framKeluh.Visible = False
353               Else
354                   Set rsSearch = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_WO = '" & Trim(Text(0)) & "'")
355                   If Not rsSearch.EOF Then
356                       txtPermintaan = Trim(rsSearch("REQ1")) & vbCrLf & Trim(rsSearch("REQ2")) & vbCrLf & Trim(rsSearch("REQ3")) & vbCrLf & Trim(rsSearch("REQ4")) & vbCrLf & Trim(rsSearch("REQ5"))
357                       txtKeluhan = Trim(rsSearch("CMPL1")) & vbCrLf & Trim(rsSearch("CMPL2")) & vbCrLf & Trim(rsSearch("CMPL3")) & vbCrLf & Trim(rsSearch("CMPL4")) & vbCrLf & Trim(rsSearch("CMPL5"))
358                   End If
359                   framKeluh.Visible = True
360               End If
          
          
361               Set rscr7 = objO.FetchData("Select Catat FROM JOBDTA WITH (NOLOCK) Where NO_WO = '" & Trim(Text(0)) & "'")
362               If rscr7.RecordCount > 0 Then
363                   txt.Text = rscr7.Fields(0).value
364               Else
365                   txt.Text = ""
366               End If
          
367           End If
          
368          On Error GoTo 0
369          Exit Sub
          
cmdKeluh_Click_Error:
          
370           MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdKeluh_Click of Form frmOrderValidasi")
End Sub



Private Sub cmdNonCR7_Click()
    
1            On Error GoTo cmdNonCR7_Click_Error
    
2             Call objO.UpdateData("Update Jobdta Set Catat = '" & Trim(txtCatatNonCR7.Text) & "' Where NO_WO = '" & Trim(Text(0).Text) & "'")
3             If Log.RPLHistori = 1 Then
4                 Call objO.UpdateData("Update [" & dbReplikasiHO & "]..Jobdta Set Catat = '" & Trim(txtCatatNonCR7.Text) & "' Where NO_WO = '" & Trim(Text(0).Text) & "' and FS_KD_DEALER = '" & Trim(gKdCbg) & "'")
5             End If
6             MsgBox "Catatan telah disimpan !", vbInformation + vbOKOnly
    
    
7            On Error GoTo 0
8            Exit Sub
    
cmdNonCR7_Click_Error:
    
9             MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdNonCR7_Click of Form frmOrderValidasi"
End Sub

Private Sub cmdPsmInfo_Click()
    Dim f As New frmOrderValidasiPsmItem
    f.Text(0) = NulSID(Text(4))
    f.Text(4) = NulSID(Text(0))
    f.Text(5) = ""
    f.Text(1) = NulSID(Text(17))
    f.Text(2) = NulSID(Text(16))
    f.isLoad = False
    f.show vbModal
    
End Sub

Private Sub cmdTungguValidasiTL_Click()
    ''PROSES UPDATE PREDIAGNOSE YANG BELUM VALIDASI
    sSQL = "UPDATE DIAGDTA SET flagTL = '1',userEntryTungguTL = '" & gSA & "',tglEntryTungguTL = getdate() WHERE NO_CHASSIS = '" & NulSID(Text(17)) & "' AND VALIDTWC = ''"
    Call objO.InsertData(sSQL)
    cmdTungguValidasiTL.Enabled = False
    Call panggilWO(NulSID(Text(0)))
    
End Sub

Private Sub cmdTutupFoto_Click(Index As Integer)
    frameFoto(Index).Visible = False
End Sub

Private Sub cmdTutupUploadKTP_Click()
    frameUploadKTP.Visible = False
End Sub

Private Sub cmdUpload_Click(Index As Integer)
    Dim PictBmp As String
1            On Error GoTo cmdUpload_Click_Error
    
2             cdgopen(Index).FileName = ""
3             cdgopen(Index).filter = "Pictures (*.jpg;*.jpeg;)|*.jpg;*.jpeg;"
4             cdgopen(Index).Flags = cdlOFNHideReadOnly
5             cdgopen(Index).ShowOpen
6             PictBmp = cdgopen(Index).FileName
7             If PictBmp = "" Then Exit Sub
8             If FileLen(PictBmp) > 5000000 Then
9                 MsgBox "Ukuran foto terlalu besar, Ukuran foto maksimal 5 MB"
10                Close #1
11                Exit Sub
12            End If
13            If cdgopen(Index).FileName = "" Then Exit Sub
    
    Dim ext As String
    Dim FileName As String
14            FileName = cdgopen(Index).FileName
    
15            ext = Mid(FileName, InStrRev(FileName, ".") + 1, Len(FileName))
    '' 20180524DEV01'INSPECTION
16            If UploadFile(Log.FTPCR7Host, Log.FTPCR7User, Log.FTPCR7Pass, cdgopen(Index).FileName, Right(gDb, 2) & NulSID(Text(0)) & Index + 1 & "." & LCase(ext)) = True Then
17                imgFotoCR7(Index).Picture = LoadPicture(cdgopen(Index).FileName)
18            Else
19                MsgBox "Data gagal upload", vbInformation
    
20            End If
    
21           On Error GoTo 0
22           Exit Sub
    
cmdUpload_Click_Error:
    
23            MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cmdUpload_Click of Form frmOrderValidasi"
End Sub

Private Sub Command1_Click()
          Dim isBegin         As Boolean
          Dim beginTransCR7   As Boolean
          Dim fotoStream      As New ADODB.Stream
          Dim readStream      As ADODB.Stream
          Dim dbConCR7        As New ADODB.Connection
          Dim rscr7           As New ADODB.Recordset
          Dim ext             As String
          Dim FileName        As String
          
1            On Error GoTo Command1_Click_Error
          
          'Call cmdKeluh_Click
          '' 20180515DEV01'INSPECTION
2             beginTransCR7 = False
3             isBegin = False
          
4             Call objHT.OpenTransactionBengkel
5             If dbConCR7.State = adStateClosed Then dbConCR7.Open (SQLCONBengkel)
6             dbConCR7.BeginTrans
7             beginTransCR7 = True
8             isBegin = True
              
              ''20190930 || req mb thania | penambahan log untuk update data loginspection
              ''----------------------------------------------------------------------------------------
9             sSQL = "INSERT INTO logInspection (NoWo,TglEntry,UserEntry,kompName,keterangan,HasilFI,HasilTestDrive,SNJ1,SNJ2,SNJ3,SNJ4,SNJ5,SNJ6,SNJ7,"
10            sSQL = sSQL & " isUrgent1,isUrgent2,isUrgent3,isUrgent4,isUrgent5,isUrgent6,isUrgent7,EstN1,EstN2,EstN3,EstN4,"
11            sSQL = sSQL & " EstN5,EstN6,EstN7,SP1,SP2,SP3,SP4,SP5,SP6,SP7,EstJ1,EstJ2,EstJ3,EstJ4,EstJ5,EstJ6,"
12            sSQL = sSQL & " EstJ7,EstP1,EstP2,EstP3,EstP4,EstP5,EstP6,EstP7,EstS1,EstS2,EstS3,EstS4,EstS5,EstS6,EstS7,"
13            sSQL = sSQL & " Pic1,Pic2,Pic3,Pic4,Pic5,Pic6,Pic7,AdaFoto1,AdaFoto2,AdaFoto3,AdaFoto4,AdaFoto5,AdaFoto6,"
14            sSQL = sSQL & " AdaFoto7 , FileName1, FileName2, FileName3, FileName4, FileName5, FileName6, FileName7)"
15            sSQL = sSQL & " select NoWo,GETDATE(),UserEntry,'" & ComptName & "','Sebelum Edit',HasilFI,HasilTestDrive,SNJ1,SNJ2,SNJ3,SNJ4,SNJ5,SNJ6,SNJ7,"
16            sSQL = sSQL & " isUrgent1,isUrgent2,isUrgent3,isUrgent4,isUrgent5,isUrgent6,isUrgent7,EstN1,EstN2,EstN3,EstN4,"
17            sSQL = sSQL & " EstN5,EstN6,EstN7,SP1,SP2,SP3,SP4,SP5,SP6,SP7,EstJ1,EstJ2,EstJ3,EstJ4,EstJ5,EstJ6,"
18            sSQL = sSQL & " EstJ7,EstP1,EstP2,EstP3,EstP4,EstP5,EstP6,EstP7,EstS1,EstS2,EstS3,EstS4,EstS5,EstS6,EstS7,"
19            sSQL = sSQL & " Pic1,Pic2,Pic3,Pic4,Pic5,Pic6,Pic7,AdaFoto1,AdaFoto2,AdaFoto3,AdaFoto4,AdaFoto5,AdaFoto6,"
20            sSQL = sSQL & " AdaFoto7,FileName1,FileName2,FileName3,FileName4,FileName5,FileName6,FileName7 from inspection with (nolock) WHERE nowo = '" & Trim(Text(0).Text) & "'"
21            Call objHT.InsertData_Bengkel(sSQL)
              
              ''----------------------------------------------------------------------------------------
          
          ''Call objO.UpdateData("Update Jobdta Set Catat = '" & Trim(Text1.text) & "' Where NO_WO = '" & Trim(text(0).text) & "'")
22            Call objHT.InsertData_Bengkel("Update Jobdta Set Catat = '" & Trim(txtCatatanCR7.Text) & "' Where NO_WO = '" & Trim(Text(0).Text) & "'")
23            If Log.RPLHistori = 1 Then
24                Call objHT.InsertData_Bengkel("Update [" & dbReplikasiHO & "]..Jobdta Set Catat = '" & Trim(txtCatatanCR7.Text) & "' Where NO_WO = '" & Trim(Text(0).Text) & "' and FS_KD_DEALER = '" & Trim(gKdCbg) & "'")
25            End If
          '' tidak bisa menggunakan with (nolock) dengan error = Cursor updates are not allowed on tables opened with the NOLOCK option.
26            sSQL = "SELECT TOP 1 * FROM inspection  WHERE nowo = '" & NulSID(Text(0).Text) & "' ORDER BY tglentry DESC"
27            If rscr7.State Then rscr7.Close
28            If fotoStream.State Then fotoStream.Close
29            rscr7.Open sSQL, dbConCR7, adOpenKeyset, adLockOptimistic
          
          Dim i As Integer
30            i = 0
31            If Not rscr7.EOF Then
32                With fotoStream
33                    .Type = adTypeBinary
34                    .Open
35                    If cdgopen(i).FileName <> "" Then
36                    .LoadFromFile cdgopen(i).FileName
37                    FileName = cdgopen(Index).FileName
          
38                    ext = Mid(FileName, InStrRev(FileName, ".") + 1, Len(FileName))
39                    End If
40                    If txtNo(i).Visible Then
41                        rscr7!sp1 = NulSID(txtSaran(i).Text)
42                        rscr7!estj1 = NulSID(txtJasa(i).Text)
43                        rscr7!estp1 = NulSID(txtPart(i).Text)
44                        rscr7!ests1 = NulSID(txtSublet(i).Text)
45                        rscr7!estn1 = Format(NulSID(txtTglKembali(i).Text), "yyyy-mm-dd")
46                        rscr7!Pic1 = .Read
47                        If imgFotoCR7(i).Picture = None Then
48                            rscr7!AdaFoto1 = 0
49                        Else
50                            rscr7!AdaFoto1 = 1
51                            rscr7!FileName1 = NulSID(Right(gDb, 2) & NulSID(Text(0)) & i + 1 & "." & ext)
52                        End If
          
          
53                        rscr7.Update
54                    End If
55                End With
56                Set fotoStream = Nothing
57                i = i + 1
          
          '' 2
58                With fotoStream
59                    .Type = adTypeBinary
60                    .Open
          '.LoadFromFile cdgopen(i).FileName
61                    If cdgopen(i).FileName <> "" Then
62                    .LoadFromFile cdgopen(i).FileName
63                    FileName = cdgopen(Index).FileName
          
64                    ext = Mid(FileName, InStrRev(FileName, ".") + 1, Len(FileName))
65                    End If
66                    If txtNo(i).Visible Then
67                        rscr7!sp2 = NulSID(txtSaran(i).Text)
68                        rscr7!estj2 = NulSID(txtJasa(i).Text)
69                        rscr7!estp2 = NulSID(txtPart(i).Text)
70                        rscr7!ests2 = NulSID(txtSublet(i).Text)
71                        rscr7!estn2 = Format(NulSID(txtTglKembali(i).Text), "yyyy-mm-dd")
72                        rscr7!Pic2 = .Read
73                        If imgFotoCR7(i).Picture = None Then
74                            rscr7!AdaFoto2 = 0
75                        Else
76                            rscr7!AdaFoto2 = 1
77                            rscr7!FileName2 = NulSID(NulSID(Right(gDb, 2) & NulSID(Text(0)) & i + 1 & "." & ext))
78                        End If
79                        rscr7.Update
80                    End If
81                End With
82                Set fotoStream = Nothing
83                i = i + 1
          '' -- 2 --
          
          '' 3
84                With fotoStream
85                    .Type = adTypeBinary
86                    .Open
          '.LoadFromFile cdgopen(i).FileName
87                    If cdgopen(i).FileName <> "" Then
88                    .LoadFromFile cdgopen(i).FileName
89                    FileName = cdgopen(Index).FileName
          
90                    ext = Mid(FileName, InStrRev(FileName, ".") + 1, Len(FileName))
91                    End If
92                    If txtNo(i).Visible Then
93                        rscr7!sp3 = NulSID(txtSaran(i).Text)
94                        rscr7!estj3 = NulSID(txtJasa(i).Text)
95                        rscr7!estp3 = NulSID(txtPart(i).Text)
96                        rscr7!ests3 = NulSID(txtSublet(i).Text)
97                        rscr7!estn3 = Format(NulSID(txtTglKembali(i).Text), "yyyy-mm-dd")
98                        rscr7!Pic3 = .Read
99                        If imgFotoCR7(i).Picture = None Then
100                           rscr7!AdaFoto3 = 0
101                       Else
102                           rscr7!AdaFoto3 = 1
103                           rscr7!FileName3 = NulSID(NulSID(Right(gDb, 2) & NulSID(Text(0)) & i + 1 & "." & ext))
104                       End If
105                       rscr7.Update
106                   End If
107               End With
108               Set fotoStream = Nothing
109               i = i + 1
          '' -- 3 --
          
          '' 4
110               With fotoStream
111                   .Type = adTypeBinary
112                   .Open
          '.LoadFromFile cdgopen(i).FileName
113                   If cdgopen(i).FileName <> "" Then
114                   .LoadFromFile cdgopen(i).FileName
115                   FileName = cdgopen(Index).FileName
          
116                   ext = Mid(FileName, InStrRev(FileName, ".") + 1, Len(FileName))
117                   End If
118                   If txtNo(i).Visible Then
119                       rscr7!sp4 = NulSID(txtSaran(i).Text)
120                       rscr7!estj4 = NulSID(txtJasa(i).Text)
121                       rscr7!estp4 = NulSID(txtPart(i).Text)
122                       rscr7!ests4 = NulSID(txtSublet(i).Text)
123                       rscr7!estn4 = Format(NulSID(txtTglKembali(i).Text), "yyyy-mm-dd")
124                       rscr7!Pic4 = .Read
125                       If imgFotoCR7(i).Picture = None Then
126                           rscr7!AdaFoto4 = 0
127                       Else
128                           rscr7!AdaFoto4 = 1
129                           rscr7!FileName4 = NulSID(NulSID(Right(gDb, 2) & NulSID(Text(0)) & i + 1 & "." & ext))
130                       End If
131                       rscr7.Update
132                   End If
133               End With
134               Set fotoStream = Nothing
135               i = i + 1
          '' -- 4 --
          
          '' 5
136               With fotoStream
137                   .Type = adTypeBinary
138                   .Open
          '.LoadFromFile cdgopen(i).FileName
139                   If cdgopen(i).FileName <> "" Then
140                   .LoadFromFile cdgopen(i).FileName
141                   FileName = cdgopen(Index).FileName
          
142                   ext = Mid(FileName, InStrRev(FileName, ".") + 1, Len(FileName))
143                   End If
144                   If txtNo(i).Visible Then
145                       rscr7!SP5 = NulSID(txtSaran(i).Text)
146                       rscr7!EstJ5 = NulSID(txtJasa(i).Text)
147                       rscr7!EstP5 = NulSID(txtPart(i).Text)
148                       rscr7!EstS5 = NulSID(txtSublet(i).Text)
149                       rscr7!EstN5 = Format(NulSID(txtTglKembali(i).Text), "yyyy-mm-dd")
150                       rscr7!Pic5 = .Read
151                       If imgFotoCR7(i).Picture = None Then
152                           rscr7!AdaFoto5 = 0
153                       Else
154                           rscr7!AdaFoto5 = 1
155                           rscr7!FileName5 = NulSID(NulSID(Right(gDb, 2) & NulSID(Text(0)) & i + 1 & "." & ext))
156                       End If
157                       rscr7.Update
158                   End If
159               End With
160               Set fotoStream = Nothing
161               i = i + 1
          '' -- 5 --
          
          '' 6
162               With fotoStream
163                   .Type = adTypeBinary
164                   .Open
          '.LoadFromFile cdgopen(i).FileName
165                   If cdgopen(i).FileName <> "" Then
166                   .LoadFromFile cdgopen(i).FileName
167                   FileName = cdgopen(Index).FileName
          
168                   ext = Mid(FileName, InStrRev(FileName, ".") + 1, Len(FileName))
169                   End If
170                   If txtNo(i).Visible Then
171                       rscr7!SP6 = NulSID(txtSaran(i).Text)
172                       rscr7!EstJ6 = NulSID(txtJasa(i).Text)
173                       rscr7!EstP6 = NulSID(txtPart(i).Text)
174                       rscr7!EstS6 = NulSID(txtSublet(i).Text)
175                       rscr7!EstN6 = Format(NulSID(txtTglKembali(i).Text), "yyyy-mm-dd")
176                       rscr7!Pic6 = .Read
177                       If imgFotoCR7(i).Picture = None Then
178                           rscr7!AdaFoto6 = 0
179                       Else
180                           rscr7!AdaFoto6 = 1
181                           rscr7!FileName6 = NulSID(NulSID(Right(gDb, 2) & NulSID(Text(0)) & i + 1 & "." & ext))
182                       End If
183                       rscr7.Update
184                   End If
185               End With
186               Set fotoStream = Nothing
187               i = i + 1
          '' -- 6 --
          
          '' 7
188               With fotoStream
189                   .Type = adTypeBinary
190                   .Open
          '.LoadFromFile cdgopen(i).FileName
191                   If cdgopen(i).FileName <> "" Then
192                   .LoadFromFile cdgopen(i).FileName
193                   FileName = cdgopen(Index).FileName
          
194                   ext = Mid(FileName, InStrRev(FileName, ".") + 1, Len(FileName))
195                   End If
196                   If txtNo(i).Visible Then
197                       rscr7!SP7 = NulSID(txtSaran(i).Text)
198                       rscr7!EstJ7 = NulSID(txtJasa(i).Text)
199                       rscr7!EstP7 = NulSID(txtPart(i).Text)
200                       rscr7!EstS7 = NulSID(txtSublet(i).Text)
201                       rscr7!EstN7 = Format(NulSID(txtTglKembali(i).Text), "yyyy-mm-dd")
202                       rscr7!Pic7 = .Read
203                       If imgFotoCR7(i).Picture = None Then
204                           rscr7!AdaFoto7 = 0
205                       Else
206                           rscr7!AdaFoto7 = 1
207                           rscr7!FileName7 = NulSID(NulSID(Right(gDb, 2) & NulSID(Text(0)) & i + 1 & "." & ext))
208                       End If
209                       rscr7.Update
210                   End If
211               End With
212               Set fotoStream = Nothing
213               i = i + 1
          '' -- 7 --
                  
                  
                  ''20190930 || req mb thania | penambahan log untuk update data loginspection
                  ''----------------------------------------------------------------------------------------
214               sSQL = "INSERT INTO logInspection (NoWo,TglEntry,UserEntry,kompName,keterangan,HasilFI,HasilTestDrive,SNJ1,SNJ2,SNJ3,SNJ4,SNJ5,SNJ6,SNJ7,"
215               sSQL = sSQL & " isUrgent1,isUrgent2,isUrgent3,isUrgent4,isUrgent5,isUrgent6,isUrgent7,EstN1,EstN2,EstN3,EstN4,"
216               sSQL = sSQL & " EstN5,EstN6,EstN7,SP1,SP2,SP3,SP4,SP5,SP6,SP7,EstJ1,EstJ2,EstJ3,EstJ4,EstJ5,EstJ6,"
217               sSQL = sSQL & " EstJ7,EstP1,EstP2,EstP3,EstP4,EstP5,EstP6,EstP7,EstS1,EstS2,EstS3,EstS4,EstS5,EstS6,EstS7,"
218               sSQL = sSQL & " Pic1,Pic2,Pic3,Pic4,Pic5,Pic6,Pic7,AdaFoto1,AdaFoto2,AdaFoto3,AdaFoto4,AdaFoto5,AdaFoto6,"
219               sSQL = sSQL & " AdaFoto7 , FileName1, FileName2, FileName3, FileName4, FileName5, FileName6, FileName7)"
220               sSQL = sSQL & " select NoWo,GETDATE(),'" & gSA & "','" & ComptName & "','Sesudah Edit',HasilFI,HasilTestDrive,SNJ1,SNJ2,SNJ3,SNJ4,SNJ5,SNJ6,SNJ7,"
221               sSQL = sSQL & " isUrgent1,isUrgent2,isUrgent3,isUrgent4,isUrgent5,isUrgent6,isUrgent7,EstN1,EstN2,EstN3,EstN4,"
222               sSQL = sSQL & " EstN5,EstN6,EstN7,SP1,SP2,SP3,SP4,SP5,SP6,SP7,EstJ1,EstJ2,EstJ3,EstJ4,EstJ5,EstJ6,"
223               sSQL = sSQL & " EstJ7,EstP1,EstP2,EstP3,EstP4,EstP5,EstP6,EstP7,EstS1,EstS2,EstS3,EstS4,EstS5,EstS6,EstS7,"
224               sSQL = sSQL & " Pic1,Pic2,Pic3,Pic4,Pic5,Pic6,Pic7,AdaFoto1,AdaFoto2,AdaFoto3,AdaFoto4,AdaFoto5,AdaFoto6,"
225               sSQL = sSQL & " AdaFoto7,FileName1,FileName2,FileName3,FileName4,FileName5,FileName6,FileName7 from inspection with (nolock) WHERE nowo = '" & Trim(Text(0).Text) & "'"
226               Call objHT.InsertData_Bengkel(sSQL)
                  
227               Call objHT.CloseTransactionBengkel
228               dbConCR7.CommitTrans
229           End If
230           MsgBox "Catatan telah disimpan !", vbInformation + vbOKOnly
          
231          On Error GoTo 0
232          Exit Sub
          
Command1_Click_Error:
233           If isBegin Then
234               Call objHT.CloseTransactionBengkel_err
235           End If
236           If beginTransCR7 = True Then
237               dbConCR7.RollbackTrans
238           End If
239           MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Command1_Click of Form frmOrderValidasi")
End Sub

Private Sub Command2_Click()
    If Trim(txtAlasanBtlWashing.Text) <> "" Then
        If Len(txtAlasanBtlWashing.Text) < 5 Then
            MsgBox "Alasan Pembatalan Washing minimal 5 karakter!", vbCritical, "Informasi Pembatalan Washing"
        Else
            On Error GoTo vbError
            erTranAll = False
            objHT.OpenTransactionBengkel
            
            Call objHT.InsertData_Bengkel("INSERT INTO WASHING VALUES ('" & Trim(Text(0)) & "',GETDATE(),'" & gSA & "','" & txtAlasanBtlWashing.Text & "')")
            Call objHT.InsertData_Bengkel("UPDATE ANTRICUCI SET FLAG = 'CANCEL' WHERE NOORDER = '" & Trim(Text(0)) & "'")
            Call objHT.InsertData_Bengkel("UPDATE JOBDTL SET WASHING = 0 WHERE NO_wo = '" & Trim(Text(0)) & "'")
            Call objHT.InsertData_Bengkel("UPDATE STALL SET FLAG = 'F' WHERE KETERANGAN LIKE '%" & Trim(Text(0)) & "%'")
            
            If erTranAll = False Then
                objHT.CloseTransactionBengkel
            End If
            MsgBox "Proses Pembatalan Washing berhasil dilakukan", vbInformation, "Informasi Pembatalan Washing"
            txtAlasanBtlWashing.Text = ""
            framePembatalanWashing.Visible = False
            btnBatalWashing.Enabled = False
            Exit Sub
vbError:
            objHT.CloseTransactionBengkel_err
            MsgBox "Proses Pembatalan Washing gagal dilakukan", vbCritical
            Exit Sub
        End If
    Else
        MsgBox "Alasan Pembatalan Washing harus diisi!", vbCritical, "Informasi Pembatalan Washing"
    End If
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Form_Activate()
1            On Error GoTo Form_Activate_Error
    
2             If orderAktif Or subletAktif Or subRetAktif Then
3                 MsgBox "Tutup Menu sebelumnya terlebih Dahulu !!", vbCritical
4                 Unload Me
5             End If
    
6            On Error GoTo 0
7            Exit Sub
    
Form_Activate_Error:
    
8             MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Activate of Form frmOrderValidasi")
End Sub

Private Sub Form_Load()
1            On Error GoTo Form_Load_Error
    
2             validAktif = True
3             Text(1) = TglDDMMYYYY(isnow)
4             simpan = True
5             Frame6.Top = 0
6             If Log.dmsNonDss = 1 Then
7                 btnBatalWashing.Visible = True
8             Else
9                 btnBatalWashing.Visible = False
10            End If
    
    ''2017-08-01 || REVISI penambahan tombol batal request discount
    ''---------------------------------------------------------------------
11            If Log.discKacab = 1 Then
12                cmdBatalRequest.Visible = True
13                cmdBatalRequest.Enabled = False
14            Else
15                cmdBatalRequest.Visible = False
16            End If
    ''---------------------------------------------------------------------
    
17            If Log.lockSCC = 1 And Log.newLockSSC = 1 Then
18                If Log.discKacab = 0 Then
19                    cmdTungguValidasiTL.Left = cmdBatalRequest.Left
20                End If
21            Else
22                cmdTungguValidasiTL.Visible = False
23            End If
    
24           On Error GoTo 0
25           Exit Sub
    
Form_Load_Error:
    
26            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Load of Form frmOrderValidasi")
End Sub

Private Sub Form_Resize()
    asxPowerBanner1.Height = Me.Height - 500
End Sub

Private Sub Form_Unload(Cancel As Integer)
    validAktif = False
End Sub

Private Sub GridFind_AfterColUpdate(ByVal ColIndex As Integer)
    If ColIndex = 0 Then
        Dim k As String
        Dim filter As String
        k = NulSID(rsKTP!idcust)
        filter = rsKTP.filter
        
        If Not rsKTP.EOF Then
            rsKTP.MoveFirst
            rsKTP.filter = "Pilih=" & True
            If Not rsKTP.EOF Then
                rsKTP.MoveFirst
                Do While Not rsKTP.EOF
                    rsKTP!Pilih = False
                    rsKTP.MoveNext
                Loop
            End If
            rsKTP.filter = ""
            
            rsKTP.filter = "IDCust=" & k
            If Not rsKTP.EOF Then
                rsKTP!Pilih = True
                
            End If
            rsKTP.filter = ""
            ' rsKTP.filter = filter
        End If
        
        GridFind.Refresh
    End If
End Sub

Private Sub GridFind_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)
    If ColIndex = 0 Then
        
    Else
        'Cancel = True
    End If
End Sub

Private Sub GridFind_FilterChange()
    On Error GoTo errHandler
    Set cols = GridFind.Columns
    Dim C As Integer
    Dim filter As String
    'filter = filterRsTemp 'rsTemp.filter
    
    C = GridFind.col
    GridFind.HoldFields
    If filter = "" Then
        rsKTP.filter = getFilter()
    Else
        If getFilter() = "" Then
            rsKTP.filter = filter
        Else
            rsKTP.filter = filter & " AND " & getFilter()
        End If
        
    End If
    
    GridFind.col = C
    GridFind.EditActive = True
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
            If col.DataField = "pilih" Then
                col.FilterText = ""
            Else
                If n > 1 Then
                    tmp = tmp & " AND "
                End If
                tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
            End If
        End If
    Next col
    getFilter = tmp
End Function

Private Sub JobGrid_AfterColUpdate(ByVal ColIndex As Integer)
          Dim isMultipleProgram As Boolean
1         isMultipleProgram = False
2         If Log.multipleMarketing = 1 And DateDiff("D", Format(Log.tglmultipleMarketing, "DD/MM/YYYY"), Format(Text(43), "DD/MM/YYYY")) >= 0 Then
3             isMultipleProgram = True
4         End If
5         With JobGrid
              
6             If ColIndex = 11 Then
                  ''centangan customer
7                 rsJob("DISC") = 0
8                 rsJob("DISCRP") = 0
9                 rsJob("STLDISC") = .Columns(6).value
10                If isMultipleProgram Then
11                    If .Columns(15).value Then
12                        .Columns(ColIndex).value = True
13                        rsJob!FREEC = True
14                        GoTo refreshHarga
15                    End If
16                End If
17                If .Columns(ColIndex).value Then
18                    If woStat = "SBI" And (Trim(Text(44)) = "" Or Trim(Text(44)) = "DAN LAIN LAIN") Then
19                        .Columns(6) = hargaDealerSBI(CDbl(.Columns(1)), Trim(Text(13)), Trim(Text(17)))
20                        rsJob("HRG") = Round(hargaDealerSBI(CDbl(.Columns(1)), Trim(Text(13)), Trim(Text(17))))
21                    Else
22                        .Columns(6) = Round(objBilling.takeClassPrice(Trim(Text(13)), 1) * CDbl(.Columns(1)))
23                        rsJob("HRG") = Round(objBilling.takeClassPrice(Trim(Text(13)), 1) * CDbl(.Columns(1)))
24                        rsJob("DISC") = 0
25                        rsJob("DISCRP") = .Columns(6).value * (0 / 100)
26                        rsJob("STLDISC") = Round(.Columns(6).value - (.Columns(6).value * (0 / 100)), 0)
27                    End If
28                Else
29                    If woStat = "SBI" And (Trim(Text(44)) = "" Or Trim(Text(44)) = "DAN LAIN LAIN") Then
30                        .Columns(6) = hargaCustomerCBUNonTAM(CDbl(.Columns(1)), Trim(Text(13)))
31                        rsJob("HRG") = Round(hargaCustomerCBUNonTAM(CDbl(.Columns(1)), Trim(Text(13))))
32                    Else
33                        .Columns(6) = Round(objBilling.takeClassPrice(Trim(Text(13)), 2) * CDbl(.Columns(1)))
34                        rsJob("HRG") = Round(objBilling.takeClassPrice(Trim(Text(13)), 2) * CDbl(.Columns(1)))
35                    End If
36                    If isMultipleProgram Then
37                        rsJob("DISC") = 0
38                        rsJob("DISCRP") = 0
39                        rsJob("STLDISC") = Round(.Columns(6).value, 0)
                          
40                    Else
41                        rsJob("DISC") = nilaiDiscMax
42                        rsJob("DISCRP") = .Columns(6).value * (nilaiDiscMax / 100)
43                        rsJob("STLDISC") = Round(.Columns(6).value - (.Columns(6).value * (nilaiDiscMax / 100)), 0)
44                    End If
45                End If
46            ElseIf ColIndex = 8 Then
                  ''centangan rp
47                If .Columns(ColIndex).value Then
48                    .Columns(9).Locked = False
49                    .Columns(7).Locked = True
50                Else
51                    .Columns(9).Locked = True
52                    .Columns(7).Locked = False
53                End If
54            ElseIf ColIndex = 7 Then
                  ''diskon %
55                If isMultipleProgram Then
56                    If .Columns(15).value Then
57                        rsJob("DISC") = 0
58                        rsJob("DISCRP") = 0
59                        .Columns(ColIndex) = 0
60                        .Columns(9) = 0
61                        .ReBind
62                        GoTo refreshHarga
63                    Else
64                        If .Columns(ColIndex).value >= 0 Then
65                            If .Columns(ColIndex).value > 100 Then
66                                GoTo discPersenLebih100
67                            Else
68                                If opWO(2).value Then
69                                    GoTo discPersen
70                                ElseIf CDbl(.Columns(ColIndex).value) > CDbl(.Columns(14).value) Then
71                                    GoTo discPersenMelebihiKetentuan
72                                Else
73                                    GoTo discPersen
74                                End If
75                            End If
76                        Else
77                            GoTo discPersenKurang0
78                        End If
79                    End If
80                End If
81                If .Columns(ColIndex).value >= 0 Then
82                    If .Columns(ColIndex).value > 100 Then
discPersenLebih100:
83                        MsgBox "Diskon tidak boleh lebih dari 100.", vbInformation
84                        rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
85                        If Not rsJob.EOF Then
86                            rsJob("DISC") = 0
87                            rsJob("DISCRP") = 0
88                            .Columns(ColIndex) = 0
89                            .Columns(9) = 0
90                            rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
91                            .Columns(10) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
92                        End If
93                    Else
94                        If opWO(2).value Then
95                            GoTo discPersen
96                        End If
97                        If Log.discKacab = 1 And .Columns(11).value Then
98                            If .Columns(ColIndex).value > nilaiLockDisc Then
discPersenMelebihiKetentuan:
99                                MsgBox "Discount tidak boleh lebih besar dari ketentuan!!", vbInformation
100                               rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
101                               If Not rsJob.EOF Then
102                                   rsJob("DISC") = 0
103                                   rsJob("DISCRP") = 0
104                                   .Columns(ColIndex) = 0
105                                   .Columns(9) = 0
106                                   rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
107                                   .Columns(10) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
108                               End If
109                           Else
110                               GoTo discPersen
111                           End If
112                       ElseIf .Columns(ColIndex).value > nilaiDiscMax Then
113                           MsgBox "Discount tidak boleh lebih besar dari ketentuan!!", vbInformation
114                           rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
115                           If Not rsJob.EOF Then
116                               rsJob("DISC") = 0
117                               rsJob("DISCRP") = 0
118                               .Columns(ColIndex) = 0
119                               .Columns(9) = 0
120                               rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
121                               .Columns(10) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
122                           End If
123                       Else
discPersen:
124                           rsJob("STLDISC") = Round(.Columns(6).value * ((100 - .Columns(7).value) / 100), 0)
125                           .Columns(10) = Round(.Columns(6).value * ((100 - .Columns(7).value) / 100), 0)
                              
126                           rsJob("DISCRP") = Round(.Columns(6).value * ((.Columns(7).value) / 100), 0)
127                           .Columns(9).value = Round(.Columns(6).value * ((.Columns(7).value) / 100), 0)
128                       End If
129                   End If
130               Else
discPersenKurang0:
131                   MsgBox "Diskon tidak boleh kurang dari 0.", vbInformation
132                   rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
133                   If Not rsJob.EOF Then
134                       rsJob("DISC") = 0
135                       rsJob("DISCRP") = 0
136                       .Columns(ColIndex) = 0
137                       .Columns(9) = 0
138                       rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
139                       .Columns(10) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
140                   End If
141               End If
142           ElseIf ColIndex = 9 Then
                  ''diskon rupiah
                  Dim nilaiDiscPersen As Double
143               nilaiDiscPersen = Round((.Columns(ColIndex).value / .Columns(6).value * 100), 2)
144               If isMultipleProgram Then
145                   If .Columns(15).value Then
146                       rsJob("DISC") = 0
147                       rsJob("DISCRP") = 0
148                       .Columns(ColIndex) = 0
149                       .Columns(9) = 0
150                       .ReBind
151                       GoTo refreshHarga
152                   Else
153                       If nilaiDiscPersen >= 0 Then
154                           If nilaiDiscPersen > 100 Then
155                               GoTo discPersenLebih100
156                           Else
157                               If opWO(2).value Then
158                                   GoTo discHarga
159                               ElseIf CDbl(nilaiDiscPersen) > CDbl(.Columns(14).value) Then
160                                   GoTo discPersenMelebihiKetentuan
161                               Else
162                                   GoTo discHarga
163                               End If
164                           End If
165                       Else
166                           GoTo discPersenKurang0
167                       End If
168                   End If
169               End If
                  
170               If nilaiDiscPersen >= 0 Then
171                   If nilaiDiscPersen > 100 Then
172                       MsgBox "Diskon tidak boleh lebih dari 100.", vbInformation
173                       rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
174                       If Not rsJob.EOF Then
175                           rsJob("DISC") = 0
176                           rsJob("DISCRP") = 0
177                           .Columns(ColIndex) = 0
178                           .Columns(9) = 0
179                           rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
180                           .Columns(10) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
181                       End If
182                   Else
183                       If opWO(2).value Then
184                           GoTo discHarga
185                       End If
186                       If Log.discKacab = 1 And .Columns(11).value Then
187                           If nilaiDiscPersen > nilaiLockDisc Then
188                               MsgBox "Discount tidak boleh lebih besar dari ketentuan!!", vbInformation
189                               rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
190                               If Not rsJob.EOF Then
191                                   rsJob("DISC") = 0
192                                   rsJob("DISCRP") = 0
193                                   .Columns(ColIndex) = 0
194                                   .Columns(9) = 0
195                                   rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
196                                   .Columns(10) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
197                               End If
198                           Else
199                               GoTo discHarga
200                           End If
201                       ElseIf nilaiDiscPersen > nilaiDiscMax Then
202                           MsgBox "Discount tidak boleh lebih besar dari ketentuan!!", vbInformation
203                           rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
204                           If Not rsJob.EOF Then
205                               rsJob("DISC") = 0
206                               rsJob("DISCRP") = 0
207                               .Columns(ColIndex) = 0
208                               .Columns(9) = 0
209                               rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
210                               .Columns(10) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
211                           End If
212                       Else
discHarga:
213                           rsJob("STLDISC") = Round(.Columns(6).value - (.Columns(9).value), 0)
214                           .Columns(10) = Round(.Columns(6).value - (.Columns(9).value), 0)
                              
215                           rsJob("DISC") = Round((.Columns(9).value * 100) / .Columns(6).value, 2)
216                           .Columns(7).value = Round((.Columns(9).value * 100) / .Columns(6).value, 2)
217                       End If
218                   End If
219               Else
220                   MsgBox "Diskon tidak boleh kurang dari 0.", vbInformation
221                   rsJob.Find "JOB = '" & Trim(.Columns(0)) & "'", , adSearchForward, 1
222                   If Not rsJob.EOF Then
223                       rsJob("DISC") = 0
224                       rsJob("DISCRP") = 0
225                       .Columns(ColIndex) = 0
226                       .Columns(9) = 0
227                       rsJob("STLDISC") = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
228                       .Columns(10) = Round(((100 - CDbl(.Columns(ColIndex))) / 100) * CDbl(.Columns(6)))
229                   End If
230               End If
231           End If
232       End With
refreshHarga:
233       If Not rsJob.EOF Then
234           rsJob.MoveFirst
              Dim nilaiJob As Double
235           nilaiJob = 0
236           Do While Not rsJob.EOF
237               If nDuoInv Then
238                   If rsJob!FREEC = True Then
239                       nilaiJob = nilaiJob + CDbl(rsJob("STLDISC"))
240                   Else
                          ''nilaiJob = nilaiJob + CDbl(rsJob("STLDISC"))
241                   End If
242               Else
243                   nilaiJob = nilaiJob + CDbl(rsJob("STLDISC"))
244               End If
                  
245               rsJob.MoveNext
246           Loop
247           Text(22) = Round(nilaiJob, 0)
              
248           Text(29) = Round(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0) '+ CDbl(text(28))
249       End If
End Sub

Private Sub partGrid_AfterColUpdate(ByVal ColIndex As Integer)
1         If ColIndex = 9 Then
2             If partGrid.Columns(9).value Then
                  ''ditagihkan customer
3                 partTerpakaiNRM = partTerpakaiNRM - (rsPart!Harga * rsPart!qty)
4                 If rsPart!oli Then
5                     Text(25) = Text(25) + (rsPart!Harga * rsPart!qty)
6                 Else
                      ''part
7                     Text(24) = Text(24) + (rsPart!Harga * rsPart!qty)
8                 End If
9                 Text(29) = Round(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0)
10                rsPart!FREEC = True
11            Else
12                If rsPart!terlibat = True Then
13                    If tagihanFOP = "NRM" Then
14                        If nilaiMaxPartNRM < (partTerpakaiNRM + (rsPart!Harga * rsPart!qty)) Then
15                            MsgBox "Nilai tertagih melebihi ketentuan " & nilaiMaxPartNRM, vbInformation
16                            partGrid.Columns(9).value = True
17                            rsPart!FREEC = True
18                            Exit Sub
19                        Else
20                            partTerpakaiNRM = partTerpakaiNRM + (rsPart!Harga * rsPart!qty)
21                        End If
22                    End If
23                    If rsPart!oli Then
                          
24                        Text(25) = Text(25) - (rsPart!Harga * rsPart!qty)
25                    Else
                          ''part
26                        Text(24) = Text(24) - (rsPart!Harga * rsPart!qty)
27                    End If
28                    Text(29) = Round(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0)
29                    rsPart!FREEC = False
30                Else
31                    MsgBox "Kode part " & NulSID(rsPart!KodeItem) & " tidak terdaftar dalam program FOP", vbInformation
32                    partGrid.Columns(9).value = True
33                    rsPart!FREEC = True
34                End If
35            End If
36        End If
End Sub

Private Sub subletGrid_AfterColUpdate(ByVal ColIndex As Integer)
1          If ColIndex = 10 Then
2             If subletGrid.Columns(10).value Then
                  ''ditagihkan customer
3                 partTerpakaiNRM = partTerpakaiNRM - (rsSublet!Price * rsSublet!qty)
4                 rsSublet!diskon = 0
5                 rsSublet!Subtotal = rsSublet!Price * rsSublet!qty
6                 subletGrid.Columns(8).value = 0
7                 subletGrid.Columns(9).value = rsSublet!Price * rsSublet!qty
8                 Text(27) = Text(27) + SpyRound(rsSublet!Price * rsSublet!qty, 0)
9                 Text(29) = SpyRound(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0)
10                rsSublet!FREEC = True
11            Else
12                If rsSublet!terlibat = True Then
13                    If tagihanFOP = "NRM" Then
14                        If nilaiMaxPartNRM < (partTerpakaiNRM + (rsSublet!Price * rsSublet!qty)) Then
15                            MsgBox "Nilai tertagih melebihi ketentuan " & nilaiMaxPartNRM, vbInformation
16                            subletGrid.Columns(9).value = True
17                            rsSublet!FREEC = True
18                            Exit Sub
19                        Else
20                            partTerpakaiNRM = partTerpakaiNRM + (rsSublet!Price * rsSublet!qty)
21                        End If
22                    End If
23                    rsSublet!diskon = 0
24                    rsSublet!Subtotal = rsSublet!Price * rsSublet!qty
25                    subletGrid.Columns(8).value = 0
26                    subletGrid.Columns(9).value = rsSublet!Price * rsSublet!qty
27                    Text(27) = Text(27) - SpyRound(rsSublet!Price * rsSublet!qty, 0)
28                    Text(29) = SpyRound(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0)
29                    rsSublet!FREEC = False
30                Else
31                    MsgBox "Kode sublet " & NulSID(rsSublet!NamaBarang) & " tidak terdaftar dalam program FOP", vbInformation
32                    subletGrid.Columns(10).value = True
33                    rsSublet!FREEC = True
34                End If
35            End If
36        ElseIf ColIndex = 8 Then
37            If rsSublet!diskon < 0 Then
38                rsSublet!diskon = 0
39                subletGrid.Columns(8).value = 0
40                MsgBox "Diskon kurang dari 0 %", vbInformation
41            ElseIf rsSublet!diskon > Log.discMaxSublet Then
42                rsSublet!diskon = 0
43                subletGrid.Columns(8).value = 0
44                MsgBox "Diskon melebihi ketentuan " & Log.discMaxSublet & " %", vbInformation
45            Else
                  ''jek nilai sublet jika promo tidak boleh di diskon
46                sSQL = "select a.KDPROMO from [ReplBengkel99]..MSTSUBLETPROMO a with (nolock) INNER JOIN jobsublet b WITH (NOLOCK) ON b.kodeitem = a.kdsublet and a.kdpromo = b.kdpromo and b.no_wo = '" & NulSID(Text(0)) & "' where a.FS_LOG_DEL = '' and a.KDSUBLET = '" & NulSID(rsSublet!KodeItem) & "'"
47                If objO.FetchData_String(sSQL) <> "" Then
48                    rsSublet!diskon = 0
49                    subletGrid.Columns(8).value = 0
50                    MsgBox "Kode sublet terlibat promo" & vbCrLf & " Tidak dapat diberi diskon", vbInformation
51                Else
52                    rsSublet!Subtotal = SpyRound((rsSublet!Price * rsSublet!qty) * ((100 - rsSublet!diskon) / 100), 0)
53                    rsSublet.Update
54                    rsSublet.MoveFirst
55                    Text(27) = 0
56                    Do While Not rsSublet.EOF
57                        Text(27) = Text(27) + SpyRound((rsSublet!Price * rsSublet!qty) * ((100 - rsSublet!diskon) / 100), 0)
58                        Text(29) = SpyRound(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0)
59                        rsSublet.MoveNext
60                    Loop
61                End If
62            End If
63            subletGrid.ReBind
64        End If
End Sub
  

Private Sub tblBatal_Click()
          'If gSpart And adaSublet Then
1            On Error GoTo tblBatal_Click_Error
          
2             Set rsTemp = objO.FetchData("SELECT DISTINCT NO_WO FROM VALIDBILLING WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND DELLOG = ''")
3             If rsTemp.EOF Then
          
4                 Call objO.InsertData("UPDATE CHASSINGFORM SET TANGGALVALIDASI = NULL WHERE NOWO = '" & Trim(Text(0)) & "'")
          
5                 Call objO.UpdateData("UPDATE VALIDORDER SET USERDEL = '" & Trim(gSA) & "', USERDEL_TGL = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "' WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(Text(17)) & "' AND USERDEL = ''")
6                 Call objO.DeleteData("DELETE FROM NEA_PS Where noWO = '" & Trim(Text(0)) & "'")
7                 chNea.value = 0
8                 chPS.value = 0
9                 txtInfoNea.Text = ""
10                lblStatus.Caption = "BATAL VALIDASI"
11                tblBatal.Enabled = False
12                TblValid.Enabled = True
                  
13                sSQL = "SELECT * FROM [VO_FOPPart] WITH (NOLOCK) WHERE [NO_WO] = '" & NulSID(Text(0)) & "' AND [NO_CHASSIS] = '" & NulSID(Text(17)) & "' AND FS_LOG_DEL = ''"
14                If Log.freepart = 1 Then 'And objO.FetchData_Count(sSQL) > 0 Then
15                    If objO.FetchData_Count(sSQL) > 0 Then
16                        sSQL = "UPDATE [VO_FOPPart] SET [FS_LOG_DEL] = '" & gSA & "',FS_DATE_DEL = GETDATE() WHERE [NO_WO] = '" & NulSID(Text(0)) & "' AND [NO_CHASSIS] = '" & NulSID(Text(17)) & "' AND FS_LOG_DEL = ''"
17                        Call objO.UpdateData(sSQL)
18                    End If
19                End If
                  
20                If Log.standarisasiSublet = 1 Then
21                    sSQL = "UPDATE [vo_discsublet] SET [FS_LOG_DEL] = '" & gSA & "',FS_DATE_DEL = GETDATE() WHERE [NO_WO] = '" & NulSID(Text(0)) & "' AND [NO_CHASSIS] = '" & NulSID(Text(17)) & "' AND FS_LOG_DEL = ''"
22                    Call objO.UpdateData(sSQL)
23                End If
                  
24                If Log.DIDX = 1 Then
25                    sSQL = "UPDATE [JobPartDX] SET [FS_LOG_DEL] = '" & gSA & "',FS_DATE_DEL = GETDATE() WHERE [NO_WO] = '" & NulSID(Text(0)) & "' AND [NO_CHASSIS] = '" & NulSID(Text(17)) & "' AND FS_LOG_DEL = ''"
26                    Call objO.UpdateData(sSQL)
27                End If
                  
                  
28                    sSQL = "UPDATE ReplBengkel99..[P2_MonitorSvcProgress_psm] SET [validasi] = 1 WHERE [wo_number] = '" & Right(gDb, 2) & "." & NulSID(Text(0)) & "'"
                      
29                    Call objO.UpdateData(sSQL)
                  
30                If Log.discKacab = 1 Then
31                    sSQL = "UPDATE [VO_JOBORDERLOCKDISC] SET [FS_LOG_DEL] = '" & gSA & "',FS_DATE_DEL = GETDATE() WHERE [NO_WO] = '" & NulSID(Text(0)) & "' AND [NO_CHASSIS] = '" & NulSID(Text(17)) & "' AND FS_LOG_DEL = ''"
32                    Call objO.UpdateData(sSQL)
          
          
          ''Rev20170801A
33                    isBolehBatalReqDisc = True
34                    Set rsTemp = objO.FetchData("SELECT userValid FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(Text(Index)) & "' ")
35                    If Not rsTemp.EOF Then
          ''2017-08-01 || ENABLE / DISABLE BUTTON BATAL
36                        If rsTemp!userValid = "" Then
37                            cmdBatalRequest.Enabled = True
38                            MsgBox "Request belum divalidasi." & vbCrLf & "Proses validasi order dibatalkan.", vbInformation
          'Exit Sub
39                        Else
40                            cmdBatalRequest.Enabled = False
41                        End If
42                    End If
          
          
43                End If
          
44            Else
45                MsgBox "WO ini sudah diberi Diskon Khusus oleh HO, sehingga tidak bisa dibatal Validasi" & vbCr & "Hubungi HO untuk membatalkan diskon khusus. ", vbCritical + vbOKOnly, "Batal Diskon Khusus"
46            End If
          'End If
          
47           On Error GoTo 0
48           Exit Sub
          
tblBatal_Click_Error:
          
49            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblBatal_Click of Form frmOrderValidasi")
End Sub

Private Sub tblCetakWO_Click()
1            On Error GoTo tblCetakWO_Click_Error
    
2             If Trim(Text(0)) <> "" Then
3                 If Log.newCetak = 1 Then
4                     Call cetakSimplify
5                 Else
6                     Call cetakForm
7                 End If
8             End If
    
9            On Error GoTo 0
10           Exit Sub
    
tblCetakWO_Click_Error:
    
11            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblCetakWO_Click of Form frmOrderValidasi")
End Sub

'---------------------------------------------------------------------------------------
' Procedure : cetakForm
' DateTime  : 30/04/2015 13:58
' Author    : RUDI KRISTANTO
' Purpose   :
'---------------------------------------------------------------------------------------
'
Private Sub cetakForm()
    Dim rpt As Object
    Dim NamaSAOrder As String
    Dim rsRepo As New ADODB.Recordset
    Dim rsHlp As New ADODB.Recordset
    Dim bodyClear As Boolean
    
1                On Error GoTo cetakForm_Error
    
2                 Select Case gPrintCetak
Case "2"
3                   Set rpt = New RptOrder4            ' KGW + PEMUDA ' 2190
4               Case "1"
5                   Set rpt = New RptOrder5            ' MLATI + ALL  ' 2180
6               Case "3"
7                   Set rpt = New RptOrder6            ' SLAMET  ' LX 300+II
8               Case "4"
9                   Set rpt = New RptOrder7            ' BLANKO UTK KERTAS KOSONG
10                  rpt.txtDealer.SetText support.kopOrder(0)
11                End Select
    
12                If MsgBox("Mencetak Order ?", vbYesNo, "Konfirmasi") = vbYes Then
13              gNoOrder = Text(0) & "/" & Trim(Text(41)) 'CboStatus.text
14              gKetSA = Text(19)
15              NamaSAOrder = objO.FetchData_String("Select FS_NM_PETUGAS from T_PETUGAS WITH (NOLOCK) where FS_KD_PETUGAS ='" & Trim(gKetSA) & "'")
16              gCust = Text(4)
17              gAlmt = Text(5)
18              gKota = "" 'text(6)
19              gTelpon = Text(7)
20              gTahun = Text(10)
21              gNoPolisi = Text(16)
22              gModel = Text(14)
23              gWarna = Text(15)
24              gTglMasuk = Text(43)
25              gMDLOrder = Text(12)
26              gFRM = Text(17)
27              gEG = Text(18)
    
    
28              Set rsHlp = objO.FetchData("SELECT CATAT, DATE, TIME, B.FS_NM_DEALER, B.DDATE, REMARK, REQ1, REQ2, REQ3, REQ4, REQ5, REP1, REP2, REP3, REP4, REP5, REP6, REP7 FROM JOBDTA A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS WHERE NO_WO = '" & Text(0) & "' AND A.NO_CHASSIS = '" & Text(17) & "'")
29              If Not rsHlp.EOF Then
30                  gCatatan = Trim(rsHlp("CATAT"))
31                  gJamMasuk = Trim(rsHlp("TIME"))
32                  gDealer = Trim(rsHlp("FS_NM_DEALER"))
33                  gDD = Trim(rsHlp("DDATE"))
34                  gPKhusus = Trim(rsHlp("REMARK"))
    
35                  nWKT = hitungWaktu(CDbl(totRate), rsHlp("TIME"), rsHlp("DATE"))
36                  gPSelesai = Left(nWKT, InStr(nWKT, "|") - 1)
    
37                  gKeluhan(1) = Trim(rsHlp("REQ1"))
38                  gKeluhan(2) = Trim(rsHlp("REQ2"))
39                  gKeluhan(3) = Trim(rsHlp("REQ3"))
40                  gKeluhan(4) = Trim(rsHlp("REQ4"))
41                  gKeluhan(5) = Trim(rsHlp("REQ5"))
42                  gOrder(1) = Trim(rsHlp("REP1"))
43                  gOrder(2) = Trim(rsHlp("REP2"))
44                  gOrder(3) = Trim(rsHlp("REP3"))
45                  gOrder(4) = Trim(rsHlp("REP4"))
46                  gOrder(5) = Trim(rsHlp("REP5"))
47                  gOrder(6) = Trim(rsHlp("REP6"))
48                  gOrder(7) = Trim(rsHlp("REP7"))
49              End If
    
50              Set rsHlp = objO.FetchData("SELECT * FROM JOBSUBLET WITH (NOLOCK) WHERE NO_WO = '" & Text(0) & "' AND NO_CHASSIS = '" & Text(17) & "'")
51              If Not rsHlp.EOF Then
52                  i = 1
53                  Do While Not rsHlp.EOF
54                      If i < 10 Then
55                          gSublet(i) = rsHlp("NAMABARANG")
56                          i = i + 1
57                      End If
58                      rsHlp.MoveNext
59                  Loop
60              End If
61              If rsHlp.State <> adStateClosed Then rsHlp.Close
    
62              gKM = Format(Text(37), "#,#0")
63              gSAdvisor = NamaSAOrder
64              gKodePlg = Text(2)
    
    Dim sHlp As Object  'New ADODB.Recordset
    Dim sHlp2 As Object
    Dim sHlp3 As Object
    Dim sHlp4 As Object
    Dim rRs As ADODB.Recordset
    Dim crxSubReportHistory As CRAXDRT.Report
    Dim crxSubReportBooking As CRAXDRT.Report
    Dim crxSubReportBooking2 As CRAXDRT.Report
    Dim crxSubReportBookingFU As CRAXDRT.Report
65              If gUkuranOrder = False Then
    
66                  Set rRs = objO.FetchData("SELECT A.CUSTOMER, A.CUST_ADRES, A.KOTA, A.TELPON, A.NO_POLISI, A.NO_ENGINE, A.NO_CHASSIS, A.MDL, A.MODEL, A.FMDL, A.COLOR, " _
    & "D.FS_KD_DEALER, D.FS_NM_DEALER, D.DDATE, A.DOTAM, A.CODE, A.NPWP, A.PPN, A.DISC, A.DISCT, A.CRDTCD, A.FRONT, A.[GROUP], A.MECH, " _
    & "A.KM, A.DATE, A.TIME, A.TIME_M, A.TIME_S, A.TIME_A, A.TIME_AT, A.TIME_SA, A.PART, A.MTRL, A.OIL, A.MET, A.SBLT, A.LABOR, A.BODY, A.TOTAL, " _
    & "A.JOB, A.REQ1, A.REQ2, A.REQ3, A.REQ4, A.REQ5, A.JOBCD, A.CR, A.REP1, A.REP2, A.REP3, A.REP4, A.REP5, A.REP6, A.REP7, A.REP8, A.SLIP, " _
    & "A.RTE1, A.RTE2, A.RTE3, A.RTE4, A.RTE5, A.RTE6, A.RTE7, A.RTE8, A.INVDATE, A.NO_WO, A.TPSVC, A.NOINV, A.STATUS, A.PAID, A.NOFAK, " _
    & "A.CATAT, A.DELIV, A.SPK, A.RPR_COND, A.KD_K, A.MEKANIK1, A.MEKANIK2, A.MEKANIK3, A.MEKANIK4, A.MEKANIK5, " _
    & "A.MEKANIK6, A.MEKANIK7, A.MEKANIK8, A.GRP1, A.GRP2, A.GRP3, A.GRP4, A.GRP5, A.GRP6, A.GRP7, A.GRP8, " _
    & "A.OFF1, A.OFF2, A.OFF3, A.OFF4, A.OFF5, A.OFF6, A.OFF7, A.OFF8, " _
    & "A.KLPKERJA1, A.KLPKERJA2, A.KLPKERJA3, A.KLPKERJA4, A.KLPKERJA5, A.KLPKERJA6, A.KLPKERJA7, A.KLPKERJA8, A.ALASAN, A.FS_STATUS_INV, " _
    & "A.FS_TH_MOBIL, A.FN_RETURN, A.FS_LOG_ADD, A.FS_LOG_DEL, A.FD_TGL_DEL, A.FS_RC, A.FS_NM_HOBI, A.FS_KD_PLG, A.FS_KD_WIL, " _
    & "A.KODETIPE, A.PDS, A.THS, A.SBE, B.COND1, B.COND2, B.COND3, B.COND4, B.COND5, B.COND6, B.WAK1, B.WAK2, B.WAK3, B.WAK4, B.WAK5, " _
    & "B.WAE1, B.WAE2, B.WAE3, B.WAE4, B.WAE5, B.BBM, B.HEAVYUSE, B.GANTIPART, B.TUNGGUSVC, B.PEMBAYARAN, B.TGLFU, B.JAMFU, B.TELPFU, " _
    & "B.KETFU, B.VERIFIKATOR, B.DATESERAH, B.TIMESERAH, B.CMPL1, B.CMPL2, B.CMPL3, B.CMPL4, B.CMPL5, B.DATANG, B.DTLDATANG, " _
    & "B.KARAKTER, B.DTLKARAKTER, D.KODEPOS, D.REMARK, D.KEYNO, D.[USER], D.TGL_REP, D.FS_KM_AKHIR, D.USEJOBDSC, D.RCDATE, D.CDK, " _
    & "D.TGL_D , D.CDD, D.USEFS_NM_HOBI, D.NOTE, E.FS_KD_PETUGAS, E.FS_NM_PETUGAS, E.FS_KD_PASS, E.FS_KD_GROUP, E.tipe, E.DIVISI, A.ANCUST, D.GSM, D.CDMA " _
    & "FROM JOBDTA AS A WITH (NOLOCK) INNER JOIN JOBDTL AS B WITH (NOLOCK) ON A.NO_WO = B.NO_WO LEFT OUTER JOIN CUSTDAT AS D WITH (NOLOCK) ON A.NO_CHASSIS = D.NO_CHASSIS INNER JOIN " _
    & "T_PETUGAS AS E WITH (NOLOCK) ON A.FRONT = E.FS_KD_PETUGAS WHERE A.NO_WO = '" & Trim(Text(0)) & "'")
67                  rpt.database.SetDataSource rRs
    
68                  gReportConn = "PROVIDER=SQLOLEDB;DATA SOURCE=" & gServerDb & ";initial catalog=" & gDb & ";User ID=" & gUser & ";Password=" & gPwd & ";"
    
69                  sSQL = "SELECT * FROM historyUnitFunc('" & Trim(gFRM) & "')"
70                  Set sHlp = objO.FetchData(sSQL)
71                  Set crxSubReportHistory = rpt.Subreport2.OpenSubreport
72                  crxSubReportHistory.database.SetDataSource sHlp
    
73                  sSQL = "select DATE_BOOK, TIME_D, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(Text(31)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
74                  Set sHlp2 = objO.FetchData(sSQL)
75                  Set crxSubReportBooking = rpt.Subreport1.OpenSubreport
76                  crxSubReportBooking.database.SetDataSource sHlp2
    
77                  sSQL = "select DATE, TIME, NOFAK FROM BOOK WITH (NOLOCK) where NOFAK = '" & Trim(Text(31)) & "' AND FS_LOG_DEL = '' order by DATE DESC"
78                  Set sHlp4 = objO.FetchData(sSQL)
79                  Set crxSubReportBooking2 = rpt.Subreport4.OpenSubreport
80                  crxSubReportBooking2.database.SetDataSource sHlp4
    
81                  sSQL = "select WAKTU FROM M_FUBOOK WITH (NOLOCK) where NOFAK = '" & Trim(Text(31)) & "' "
82                  Set sHlp3 = objO.FetchData(sSQL)
83                  Set crxSubReportBookingFU = rpt.Subreport3.OpenSubreport
84                  crxSubReportBookingFU.database.SetDataSource sHlp3
    
85                  Set rsHlp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(Text(0)) & "|%'")
86                  If Not rsHlp.EOF Then
87                      rpt.txtStall.SetText Trim(rsHlp("NOSTALL"))
88                  End If
    
89                  rpt.mechPlan.SetText ""
    
    Dim n As Integer
90                  bodyClear = True
91                  Set rsHlp = objO.FetchData("SELECT * FROM JOBKOORD WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "'")
92                  If Not rsHlp.EOF Then
93                      n = 1
94                      Do While Not rsHlp.EOF
95                          Call pasangTanda(rpt, n, CDbl(rsHlp("KOORDY")), CDbl(rsHlp("KOORDX")))
96                          n = n + 1
97                          rsHlp.MoveNext
98                      Loop
99                      bodyClear = False
100                 End If
101                 rsHlp.Close
    
102                 If bodyClear Then
103                     rpt.Text93.TextColor = &HFF&
104                     rpt.Text94.TextColor = &HFF&
105                     rpt.Text95.TextColor = &HFF&
106                     rpt.Text96.TextColor = &HFF&
107                     rpt.Text97.TextColor = &HFF&
108                 End If
    
109                 rpt.Text2.TextColor = &HFF&
110                 rpt.Text3.TextColor = &HFF&
111                 rpt.Text4.TextColor = &HFF&
112                 rpt.Text5.TextColor = &HFF&
    
    
    ' ===================================================================================================================================
    ' 23 Feb 2017 | THANIA | tambahan info bill di cetak ulang WO
    ' ===================================================================================================================================
    
113                   Set rsHlp = objO.FetchData("SELECT ISNULL(INFOBILL,'') INFO FROM T_PROFIL WITH(NOLOCK)")
114                   If Not rsHlp.EOF Then
115                      rpt.txtInfo.SetText Trim(rsHlp!info)
116                   End If
    ' ===================================================================================================================================
    
117                 Set rsHlp = objO.FetchData("SELECT BBM FROM JOBDTL WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "'")
118                 If Not rsHlp.EOF Then
119                       Select Case rsHlp("BBM")
Case 0
    ' FULL
120                               rpt.Text90.TextColor = &H0&
121                           Case 1
    ' 3/4
122                               rpt.Text62.TextColor = &H0&
123                           Case 2
    ' 1/2
124                               rpt.Text64.TextColor = &H0&
125                           Case 3
    ' 1/4
126                               rpt.Text88.TextColor = &H0&
127                           Case 4
    ' EMPTY
128                               rpt.Text89.TextColor = &H0&
129                       End Select
130                   End If
    
131                   With FrmctkOrder.CRViewer1
132                       .ReportSource = rpt
133                       .ViewReport
134                       .Zoom 100
135                       .Visible = True
136                   End With
137                   FrmctkOrder.cmdTutup.Visible = True
138                   Screen.MousePointer = vbDefault
139                   FrmctkOrder.WindowState = vbMaximized
140                   FrmctkOrder.show vbModal
141               End If
142           End If
    
143          On Error GoTo 0
144          Exit Sub
    
cetakForm_Error:
    
145           MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakForm of Form frmOrderValidasi")
End Sub

Private Sub TblKeluar_Click()
1            On Error GoTo TblKeluar_Click_Error
2            Call hapusForm
    
3             Unload Me
    
4            On Error GoTo 0
5            Exit Sub
    
TblKeluar_Click_Error:
    
6             MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TblKeluar_Click of Form frmOrderValidasi")
End Sub

Private Sub TblValid_Click()
          Dim rsSearch As New ADODB.Recordset
          
1             On Error GoTo vbError
2             Set rsSearch = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND USERDEL = ''")
3             If Not rsSearch.EOF Then
4                 MsgBox "Sudah di validasi sebelumnya !!", vbInformation, "Validasi Order"
5                 lblStatus.Caption = "DATA SUDAH DIVALIDASI"
6                 TblValid.Enabled = False
7                 tblBatal.Enabled = True
          
8             Else
9                 If Log.isFill = 1 And Log.startAppRate <= Format(Trim(Text(43)), "DD/MM/YYYY") Then
          Dim rsSO As ADODB.Recordset
10                    Set rsSO = objO.FetchDataACC("select KodeItem  from SOORD WITH (NOLOCK) " & _
          "Where (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0)+ IsNull(QtyKeluar,0))) <> 0 and NOSO = '" & Trim(Text(0)) & "' and isnull(bo,'') <> 'CANCEL'")
11                    If rsSO.RecordCount > 0 Then
12                        MsgBox "Terdapat " & rsSO.RecordCount & " Line Item Part yang belum tersuplay sesuai order !" & Chr(13) & "Validasi tidak dapat diproses !", vbInformation, "Validasi Order"
13                        Exit Sub
14                    End If
15                Else
          
16                End If
                  
                          
17                If Log.LockAR = True Then
18                    If isOverdueAR(Text(2).Text) = True Then
19                        If objO.FetchIsOpenLockAR(Text(2).Text) = False Then
20                            If objO.FetchIsValid_AR_Kabeng(Text(2).Text, Text(0).Text) = False Then
21                                MsgBox "Ada AR Overdue !" & Chr(13) & Chr(13) & "WORK ORDER TIDAK DAPAT DIPROSES !" & Chr(13) & "OTORITAS KEPALA CABANG", vbInformation + vbOKOnly
22                                Exit Sub
23                            Else
24                                MsgBox "AR OVERDUE" & Chr(13) & "TELAH DIVALIDASI OLEH KACAB", vbInformation + vbOKOnly
25                            End If
26                        End If
27                    Else
28                        If isCurrentAR(Text(2).Text) = True Then
29                            If objO.FetchIsOpenLockAR(Text(2).Text) = False Then
                              
30                                If nPersenAR >= 100 Then
31                                    If objO.FetchIsValid_AR_Kabeng(Text(2).Text, Text(0).Text) = False Then
32                                        MsgBox "AR sudah >= 100% dari Batas AR !" & Chr(13) & "Work Order Tidak Dapat Di Proses !" & Chr(13) & Chr(13) & "OTORITAS KEPALA CABANG", vbInformation + vbOKOnly
33                                        Exit Sub
34                                    Else
35                                        MsgBox "AR sudah >= 100% dari Batas AR !" & Chr(13) & "TELAH DIVALIDASI OLEH KACAB", vbInformation + vbOKOnly
36                                    End If
37                                ElseIf nPersenAR >= 80 Then
38                                    If MsgBox("AR sudah >= 80% dari Batas AR !" & Chr(13) & "LANJUTKAN PROSES ?", vbQuestion + vbYesNo) = vbNo Then
39                                        Exit Sub
40                                    End If
41                                End If
42                            End If
43                        End If
44                    End If
45                End If
46                Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasi", "tblValid", 0)
47                    erTranAll = False
48                    onTR = True
49                    Call objHT.OpenTransactionBengkel
          
50                If gSpart And adaSublet And Trim(Text(17)) <> "" Then
          
51                    If Log.discKacab = 1 Then
52                        If rsJob.RecordCount > 0 Then ''If Not rsJob.EOF Then
53                            rsJob.MoveFirst
54                            Do While Not rsJob.EOF
55                                sSQL = "INSERT INTO [VO_JOBORDERLOCKDISC] ([NO_WO],[NO_CHASSIS],[DATE],[FS_LOG_DEL],[REP],[DISC],[DISCRP],[VALID],[FREEC]) VALUES "
56                                sSQL = sSQL & " ('" & NulSID(Text(0)) & "','" & NulSID(Text(17)) & "',getdate(),'','" & NulSID(rsJob!job) & "','" & k2t(NulSID(rsJob!disc)) & "','" & k2t(NulSID(rsJob!DISCRP)) & "','" & gSA & "','" & IIf(NulSID(rsJob!FREEC), "1", "0") & "')"
57                                Call objHT.InsertData_BengkelOnly(sSQL)
58                                rsJob.MoveNext
59                            Loop
          
60                        End If
61                    End If
62                    If Log.DIDX = 1 Then
63                        If rsPart.State = adStateOpen Then
64                            If rsPart.RecordCount > 0 Then
65                                rsPart.MoveFirst
66                                Do While Not rsPart.EOF
67                                    sSQL = "INSERT INTO [JobPartDX] (NO_WO,NO_CHASSIS,KODE_ITEM,nID,CrossSell,UpSell ) VALUES ( "
68                                    sSQL = sSQL & " '" & NulSID(Text(0)) & "','" & NulSID(Text(17)) & "','" & NulSID(rsPart!KodeItem) & "','0','" & IIf(rsPart!cross, "1", "0") & "','" & IIf(rsPart!up, "1", "0") & "' )"
69                                    Call objHT.InsertData_BengkelOnly(sSQL)
70                                    rsPart.MoveNext
71                                Loop
72                            End If
73                        End If
74                    End If
75                    If Log.freepart = 1 And tagihanFOP <> "" Then
76                        If rsPart.State = adStateOpen Then
77                            If rsPart.RecordCount > 0 Then
78                                rsPart.MoveFirst
79                                Do While Not rsPart.EOF
80                                    sSQL = "INSERT INTO [VO_FOPPart] (NO_WO,NO_CHASSIS,DATE,FS_LOG_DEL,kodeitem,VALID,terlibatFOP,FREEC ) VALUES ( "
81                                    sSQL = sSQL & " '" & NulSID(Text(0)) & "','" & NulSID(Text(17)) & "',GETDATE(),'','" & NulSID(rsPart!KodeItem) & "','" & gSA & "','" & IIf(rsPart!terlibat, "1", "0") & "','" & IIf(rsPart!FREEC, "1", "0") & "' )"
82                                    Call objHT.InsertData_BengkelOnly(sSQL)
83                                    rsPart.MoveNext
84                                Loop
85                            End If
86                        End If
87                        If rsSublet.State = adStateOpen Then
88                            If rsSublet.RecordCount > 0 Then
89                                rsSublet.MoveFirst
90                                Do While Not rsSublet.EOF
91                                    sSQL = "INSERT INTO [VO_FOPPart] (NO_WO,NO_CHASSIS,DATE,FS_LOG_DEL,kodeitem,VALID,terlibatFOP,FREEC ) VALUES ( "
92                                    sSQL = sSQL & " '" & NulSID(Text(0)) & "','" & NulSID(Text(17)) & "',GETDATE(),'','" & NulSID(rsSublet!KodeItem) & "','" & gSA & "','" & IIf(rsSublet!terlibat, "1", "0") & "','" & IIf(rsSublet!FREEC, "1", "0") & "' )"
93                                    Call objHT.InsertData_BengkelOnly(sSQL)
94                                    rsSublet.MoveNext
95                                Loop
96                            End If
97                        End If
98                    End If
                      
99                    If Log.standarisasiSublet = 1 Then
100                       If rsSublet.State = adStateOpen Then
101                           If rsSublet.RecordCount > 0 Then
102                               rsSublet.MoveFirst
103                               Do While Not rsSublet.EOF
104                                   sSQL = "INSERT INTO [vo_discsublet] ([NO_WO],[NO_CHASSIS],[DATE],[FS_LOG_DEL],[FS_DATE_DEL],[kodeitem],[DISC],hargajualasli) "
105                                   sSQL = sSQL & " VALUES ('" & NulSID(Text(0)) & "','" & NulSID(Text(17)) & "',GETDATE(),'','1900-01-01','" & NulSID(rsSublet!KodeItem) & "','" & NulSID(rsSublet!diskon) & "'," & k2t(rsSublet!Price) & ")"
106                                   Call objHT.InsertData_BengkelOnly(sSQL)
107                                   rsSublet.MoveNext
108                               Loop
109                           End If
110                       End If
111                   End If
                      
112                   sSQL = "UPDATE ReplBengkel99..[P2_MonitorSvcProgress_psm] SET [validasi] = 1 WHERE [wo_number] = '" & Right(gDb, 2) & "." & NulSID(Text(0)) & "'"
                      
113                   Call objHT.InsertData_BengkelOnly(sSQL)
                       
114                   sSQL = "insert into ReplBengkel99..kirimDataDidx (tipe, nobookingapps, nobooking, noSPK,Kodedealer, [status], proses, event_no, tglentry,[asal_cabang])"
115                   sSQL = sSQL & " VALUES ('MONITORINGGR', '" & Right(gDb, 2) & "." & NulSID(Text(0)) & "', '" & Right(gDb, 2) & "." & NulSID(Text(0)) & "','', '" & Right(gDb, 2) & "','PSM Update ',0,'', getdate(),'" & Right(gDb, 2) & "')"
116                   Call objHT.InsertData_Bengkel(sSQL)
                      
117                   Call objHT.InsertData_BengkelOnly("INSERT INTO VALIDORDER VALUES ('" & Trim(gSA) & "',getdate(),'" & IIf(rsJob.State = adStateClosed, "TANPA PEKERJAAN| ", "") + IIf(rsPart.State = adStateClosed, "TANPA PART| ", "") + IIf(rsSublet.State = adStateClosed, "TANPA SUBLET| ", "") & "','" & Trim(Text(0)) & "','" & Trim(Text(17)) & "','','')")
118                   Call objHT.UpdateData_BengkelOnly("UPDATE JOBDTL SET KARAKTER = '" & Trim(cboPelanggan) & "', DTLKARAKTER = '" & Trim(Text(44)) & "' WHERE NO_WO = '" & Trim(Text(0)) & "'")
119                   Call objHT.UpdateData_BengkelOnly("UPDATE CHASSINGFORM SET TANGGALVALIDASI = GETDATE() WHERE NOWO = '" & Trim(Text(0)) & "'")
120                   If chNea.value = 1 Or chPS.value = 1 Then Call objHT.InsertData_BengkelOnly("Insert Into NEA_PS values ('" & Trim(Text(0)) & "'," & chNea.value & "," & chPS.value & ",'" & Trim(txtInfoNea.Text) & "','" & Trim(gSA) & "',getdate(),'','01/01/1900','')")
121                   lblStatus.Caption = "SUDAH VALIDASI"
122                   TblValid.Enabled = False
123                   tblBatal.Enabled = True
124               Else
125                   lblStatus.Caption = "DATA BELUM LENGKAP"
126               End If
127                   If erTranAll = False Then
128                       Call objHT.CloseTransactionBengkel
129                       onTR = False
130                       Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasi", "tblValid", 1)
131                       Exit Sub
132                   Else
133                       Call objHT.CloseTransactionBengkel_err
134                       onTR = False
135                       Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasi", "tblValid", 2)
136                       MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
137                       Exit Sub
138                   End If
139           End If
140           Exit Sub
          
vbError:
141           Call objHT.CloseTransactionBengkel_err
142           onTR = False
143           Call saveLogMenu(gSA, ComptName, gVersi, "frmOrderValidasi", "tblValid", 2)
144           MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TblValid_Click of Form frmOrderValidasi"), vbInformation, "Gagal Simpan"
145           Exit Sub
End Sub

Private Sub Text_Change(Index As Integer)
1            On Error GoTo Text_Change_Error
    
2             If Index = 17 Then
3                 If Trim(Text(17)) <> "" Then
4                     myChassis = Trim(Text(17))
5                     Call CexChassis_Discon(Trim(Text(13)))
6                     If isJualNRM = True Then
7                         Me.JualNRM.Caption = nJualNRM
8                     Else
9                         Me.JualNRM.Caption = ""
10                    End If
11                End If
12            End If
    
13           On Error GoTo 0
14           Exit Sub
    
Text_Change_Error:
    
15            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Text_Change of Form frmOrderValidasi")
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
1            On Error GoTo text_GotFocus_Error
    
2             DiTab = False
3             If Text(Index).BackColor = &HC0E0FF Then
4                 DiTab = True
5             Else
6                 If gMainKeyCari <> "" Then
7                     Text(Index).Mask = ""
8                     Text(Index) = gMainKeyCari
9                     gMainKeyCari = ""
10                    DiTab = True
11                End If
12            End If
    
13            Select Case Index
Case 31
14                    VScroll1.value = 2500
15                Case 32
16                    VScroll1.value = 2700
17                Case 28
18                    VScroll1.value = 2000
19            End Select
    
20            If DiTab Then
21                Sendkeys ("{TAB}")
22            Else
23                Sendkeys ("{HOME}+{END}")
24            End If
    
25           On Error GoTo 0
26           Exit Sub
    
text_GotFocus_Error:
    
27            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_GotFocus of Form frmOrderValidasi")
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
1            On Error GoTo Text_KeyDown_Error
    
2             Call CekTombol(KeyCode)
3             Select Case KeyCode
Case vbKeyF12
4                     Select Case Index
Case 0
5                             gMainKeyCari = Text(Index)
6                             gSearchOption = "INV2"
7                             gBantuan = "AND DELIV <> 'BTL'"
8                             frmCari.show (vbModal)
9                         Case 42
10                            gMainKeyCari = Text(Index)
11                            gSearchOption = "METERAI"
12                            frmCari.show (vbModal)
13                    End Select
14                Case vbKeyEscape
15                    If simpan = False Then
16                        If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbYes Then
17                            Call hapusForm
18                            Text(0).SelLength = Len(Text(0))
19                            Text(0).SetFocus
20                        End If
21                    End If
22            End Select
    
23           On Error GoTo 0
24           Exit Sub
    
Text_KeyDown_Error:
    
25            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Text_KeyDown of Form frmOrderValidasi")
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
1            On Error GoTo text_KeyPress_Error
    
2             KeyAscii = UpCase(KeyAscii)
3             If KeyAscii <> 13 Then
4                 Select Case Index
Case 28, 42
5                         KeyAscii = CekNumerik(KeyAscii)
6                         simpan = False
7                 End Select
8             End If
    
9            On Error GoTo 0
10           Exit Sub
    
text_KeyPress_Error:
    
11            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_KeyPress of Form frmOrderValidasi")
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
          Dim rsHlp As New ADODB.Recordset
          
1         On Error GoTo text_Validate_Error
          
2         Select Case Index
          Case 0
3             If Text(Index) <> "" Then
4                 boleh = True
5                 noInv = True
6                 If simpan = False Then
7                     If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
8                         boleh = False
9                     End If
10                End If
11                If boleh Then
12                    n = Trim(Text(0))
13                    Call hapusForm
14                    Text(0) = n
15                    If rsTemp.State Then rsTemp.Close
16                    Set rsTemp = objO.FetchData("SELECT * FROM jobDta A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.no_WO = '" & Text(Index) & "' ")
17                    If Not rsTemp.EOF Then
18                        TblValid.Enabled = False
19                        newValid = False
20                        lblStatus.Visible = True
21                        lblStatus.Caption = ""
22                        Set rsHlp = objO.FetchData("SELECT * FROM VALIDORDER WITH (NOLOCK) WHERE NO_WO = '" & Trim(Text(0)) & "' AND NO_CHASSIS = '" & Trim(rsTemp("NO_CHASSIS")) & "' AND USERDEL = ''")
23                        If Not rsHlp.EOF Then
24                            lblStatus.Caption = "SUDAH DIVALIDASI"
25                            TblValid.Enabled = False
26                            tblBatal.Enabled = True
27                            adaSublet = True
                              
                              Dim rsNPS As ADODB.Recordset
28                            Set rsNPS = objO.FetchData("Select * from NEA_PS WITH (NOLOCK) Where NOWO = '" & Text(0).Text & "'")
29                            If rsNPS.RecordCount > 0 Then
30                                chNea.value = rsNPS.Fields(1).value
31                                chPS.value = rsNPS.Fields(2).value
32                                txtInfoNea.Text = Trim(rsNPS.Fields(3).value)
33                            End If
                              
34                            If rsTemp("NOINV") <> "" And rsTemp("RPR_COND") <> "K" Then
35                                MsgBox ("Telah dibuat invoice, Tanggal " & TglDDMMMMYYYY(Trim(rsTemp("INVDATE"))) & " No. Invoice = " & Trim(rsTemp("NOINV")))
36                                lblStatus.Caption = "Sudah Invoice"
37                                tblBatal.Enabled = False
38                                noInv = False
                                  'Exit Sub
39                            ElseIf rsTemp("NOINV") <> "" And rsTemp("RPR_COND") = "K" Then
40                                MsgBox ("Telah diclosing !")
41                                lblStatus.Caption = "Closing"
42                                tblBatal.Enabled = False
43                                noInv = False
                                  'Exit Sub
44                            End If
45                        Else
46                            tblBatal.Enabled = False
47                            If (rsTemp("DELIV") = "PRC" Or rsTemp("DELIV") = "OKE") And rsTemp("RPR_COND") = "S" And rsTemp("NOINV") = "" Then
48                                newValid = True
49                                Text(21).SelLength = Len(Text(21))
50                                TblValid.Enabled = True
51                                TblValid.SetFocus
52                                MsgBox ("Unit telah selesai dikerjakan !")
53                                lblStatus.Caption = "Selesai dikerjakan"
54                                GoTo prosesStall
55                            ElseIf rsTemp("DELIV") = "BTL" Then
56                                lblStatus.Caption = "WO dibatalkan"
57                            ElseIf rsTemp("RPR_COND") = "K" Then
58                                lblStatus.Caption = "WO Sudah Closing"
59                                noInv = False
60                            Else
prosesStall:
61                                Set rsHlp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & Trim(rsTemp("NO_POLISI")) & " |" & Trim(rsTemp("NO_WO")) & "%'")
62                                If Not rsHlp.EOF Then
63                                    Select Case Trim(rsHlp("FLAG"))
                                      Case ""
                                      
64                                        MsgBox "Unit belum dibuat WO masih booking waiting !", vbInformation, "Status Order"
65                                        lblStatus.Caption = "Booking"
66                                    Case "N"
67                                        MsgBox "Unit belum dikerjakan !", vbInformation, "Status Order"
68                                        lblStatus.Caption = "Menunggu Foreman di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
69                                    Case "S"
70                                        MsgBox "Unit sedang dikerjakan !", vbInformation, "Status Order"
71                                        lblStatus.Caption = "Proses di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
72                                    Case "P"
73                                        MsgBox "Unit sedang dihentikan sementara !", vbInformation, "Status Order"
74                                        lblStatus.Caption = "Pause di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
75                                    Case "H"
76                                        MsgBox "Unit menunggu proses Final Inspection", vbInformation, "Status Order"
77                                        lblStatus.Caption = "Tunggu Inspection di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
78                                    Case "Q"
79                                        MsgBox "Unit dalam proses Final Inspection", vbInformation, "Status Order"
80                                        lblStatus.Caption = "Proses Inspection di " & Trim(rsHlp("NOSTALL")) & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
81                                    Case "G"
82                                        MsgBox "Unit menunggu proses Washing", vbInformation, "Status Order"
83                                        lblStatus.Caption = "Tunggu Washing di " & getStallWashing(Trim(Text(0)), True, True) '& ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
84                                    Case "D"
85                                        MsgBox "Unit menunggu proses Drying", vbInformation, "Status Order"
86                                        lblStatus.Caption = "Tunggu Drying di " & getStallWashing(Trim(Text(0)), True, False) ' & ", tanggal " & TglDDMMMMYYYY(Trim(rsHlp("TANGGAL")))
87                                    Case "A"
88                                        MsgBox "Unit sedang proses Washing", vbInformation, "Status Order"
89                                        lblStatus.Caption = "Washing"
90                                    Case "R"
91                                        MsgBox "Unit sedang proses Drying", vbInformation, "Status Order"
92                                        lblStatus.Caption = "Drying"
93                                    Case "F"
94                                        MsgBox ("Unit telah selesai dikerjakan !")
95                                        lblStatus.Caption = "Selesai dikerjakan"
96                                    End Select
97                                End If
98                            End If
99                        End If
                          ''------------------------------------------------------------
                          ''2017-06-07 CEK JIKA ADA REQUEST DISC HARUS DIVALIDASI KACAB
                          ''------------------------------------------------------------
                          
100                       If Log.discKacab = 1 Then
                              ''Rev20170801A
101                           isBolehBatalReqDisc = True
102                           Set rsTemp = objO.FetchData("SELECT userValid FROM RequestDisc WITH (NOLOCK) WHERE no_wo = '" & Trim(Text(Index)) & "' ")
103                           If Not rsTemp.EOF Then
                                  ''2017-08-01 || ENABLE / DISABLE BUTTON BATAL
104                               If rsTemp!userValid = "" Then
105                                   cmdBatalRequest.Enabled = True
106                                   MsgBox "Request belum divalidasi." & vbCrLf & "Proses validasi order dibatalkan.", vbInformation
                                      'Exit Sub
107                               Else
108                                   cmdBatalRequest.Enabled = False
109                               End If
110                           End If
111                       End If
                          
112                       Call panggilWO(Text(Index))
113                   End If
114                   If rsTemp.State <> adStateClosed Then rsTemp.Close
115                   simpan = False
116               End If
117           End If
118       End Select
          
119       On Error GoTo 0
120       Exit Sub
          
text_Validate_Error:
          
121       MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure text_Validate of Form frmOrderValidasi")
End Sub

Public Function getStallWashing(noOrder As String, isTunggu As Boolean, isCuci As Boolean) As String
    Dim rsTemp As ADODB.Recordset
1            On Error GoTo getStallWashing_Error
    
2             If isTunggu = True Then
3                 Set rsTemp = objO.FetchData("SELECT * FROM ANTRICUCI WHERE NOORDER = '" & noOrder & "' AND TIPESTALL = '" & IIf(isCuci, "Cuci", "Pengering") & "'")
4             Else
5                 Set rsTemp = objO.FetchData("SELECT * FROM TEMPPROGRESSCUCI WHERE NOORDER = '" & noOrder & "' AND FINISH = ''")
6             End If
7             If rsTemp.RecordCount > 0 Then
8                 getStallWashing = rsTemp("NoStall")
    
9             Else
10                getStallWashing = ""
11            End If
    
12           On Error GoTo 0
13           Exit Function
    
getStallWashing_Error:
    
14            MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getStallWashing of Form frmOrderValidasi")
End Function

Private Sub txtKeluhan_GotFocus()
    Dim DiTab As Boolean
    DiTab = False
    If txtKeluhan.BackColor = &HC0E0FF Then
        DiTab = True
    End If
    
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub

Private Sub txtPart_LostFocus(Index As Integer)
    txtPart(Index).Text = Replace(Replace(txtPart(Index).Text, ".", ""), ",", "")
    txtBiaya(Index).Text = CDbl(txtJasa(Index).Text) + CDbl(txtPart(Index).Text) + CDbl(txtSublet(Index).Text)
    
End Sub

Private Sub txtPermintaan_GotFocus()
    Dim DiTab As Boolean
    DiTab = False
    If txtPermintaan.BackColor = &HC0E0FF Then
        DiTab = True
    End If
    
    If DiTab Then
        Sendkeys ("{TAB}")
    Else
        Sendkeys ("{HOME}+{END}")
    End If
End Sub


Private Sub txtSublet_LostFocus(Index As Integer)
    txtSublet(Index).Text = Replace(Replace(txtSublet(Index).Text, ".", ""), ",", "")
    txtBiaya(Index).Text = CDbl(txtJasa(Index).Text) + CDbl(txtPart(Index).Text) + CDbl(txtSublet(Index).Text)
    
End Sub

Private Sub VScroll1_Change()
    If VScroll1.value > 0 Then
        Frame6.Top = -VScroll1.value + 20
    Else
        Frame6.Top = 0
    End If
End Sub

Private Sub VScroll1_Scroll()
    If VScroll1.value > 0 Then
        Frame6.Top = -VScroll1.value + 20
    Else
        Frame6.Top = 0
    End If
End Sub

Private Sub pasangTanda(rpt As Report, Urut As Integer, nilX As Integer, nilY As Integer)
    Dim nTop As Integer
    Dim nLeft As Integer
    
1            On Error GoTo pasangTanda_Error
    
2             nTop = 6750  '6750 'repo3kosong    '7420 'repo2
3             nLeft = 4300 '4300 'repo3kosong    '4440 'repo2
    
4             Select Case Urut
Case 1
5                     rpt.imgTnd1.TextColor = &H0&
6                     rpt.imgTnd1.Top = nTop + nilX - (195 / 2)
7                     rpt.imgTnd1.Left = nLeft + nilY - (210 / 2)
8                 Case 2
9                     rpt.imgTnd2.TextColor = &H0&
10                    rpt.imgTnd2.Top = nTop + nilX - (195 / 2)
11                    rpt.imgTnd2.Left = nLeft + nilY - (210 / 2)
12                Case 3
13                    rpt.imgTnd3.TextColor = &H0&
14                    rpt.imgTnd3.Top = nTop + nilX - (195 / 2)
15                    rpt.imgTnd3.Left = nLeft + nilY - (210 / 2)
16                Case 4
17                    rpt.imgTnd4.TextColor = &H0&
18                    rpt.imgTnd4.Top = nTop + nilX - (195 / 2)
19                    rpt.imgTnd4.Left = nLeft + nilY - (210 / 2)
20                Case 5
21                    rpt.imgTnd5.TextColor = &H0&
22                    rpt.imgTnd5.Top = nTop + nilX - (195 / 2)
23                    rpt.imgTnd5.Left = nLeft + nilY - (210 / 2)
24                Case 6
25                    rpt.imgTnd6.TextColor = &H0&
26                    rpt.imgTnd6.Top = nTop + nilX - (195 / 2)
27                    rpt.imgTnd6.Left = nLeft + nilY - (210 / 2)
28                Case 7
29                    rpt.imgTnd7.TextColor = &H0&
30                    rpt.imgTnd7.Top = nTop + nilX - (195 / 2)
31                    rpt.imgTnd7.Left = nLeft + nilY - (210 / 2)
32                Case 8
33                    rpt.imgTnd8.TextColor = &H0&
34                    rpt.imgTnd8.Top = nTop + nilX - (195 / 2)
35                    rpt.imgTnd8.Left = nLeft + nilY - (210 / 2)
36                Case 9
37                    rpt.imgTnd9.TextColor = &H0&
38                    rpt.imgTnd9.Top = nTop + nilX - (195 / 2)
39                    rpt.imgTnd9.Left = nLeft + nilY - (210 / 2)
40                Case 10
41                    rpt.imgTnd10.TextColor = &H0&
42                    rpt.imgTnd10.Top = nTop + nilX - (195 / 2)
43                    rpt.imgTnd10.Left = nLeft + nilY - (210 / 2)
44                Case 11
45                    rpt.imgTnd11.TextColor = &H0&
46                    rpt.imgTnd11.Top = nTop + nilX - (195 / 2)
47                    rpt.imgTnd11.Left = nLeft + nilY - (210 / 2)
48                Case 12
49                    rpt.imgTnd12.TextColor = &H0&
50                    rpt.imgTnd12.Top = nTop + nilX - (195 / 2)
51                    rpt.imgTnd12.Left = nLeft + nilY - (210 / 2)
52                Case 13
53                    rpt.imgTnd13.TextColor = &H0&
54                    rpt.imgTnd13.Top = nTop + nilX - (195 / 2)
55                    rpt.imgTnd13.Left = nLeft + nilY - (210 / 2)
56                Case 14
57                    rpt.imgTnd14.TextColor = &H0&
58                    rpt.imgTnd14.Top = nTop + nilX - (195 / 2)
59                    rpt.imgTnd14.Left = nLeft + nilY - (210 / 2)
60                Case 15
61                    rpt.imgTnd15.TextColor = &H0&
62                    rpt.imgTnd15.Top = nTop + nilX - (195 / 2)
63                    rpt.imgTnd15.Left = nLeft + nilY - (210 / 2)
64                Case 16
65                    rpt.imgTnd16.TextColor = &H0&
66                    rpt.imgTnd16.Top = nTop + nilX - (195 / 2)
67                    rpt.imgTnd16.Left = nLeft + nilY - (210 / 2)
68                Case 17
69                    rpt.imgTnd17.TextColor = &H0&
70                    rpt.imgTnd17.Top = nTop + nilX - (195 / 2)
71                    rpt.imgTnd17.Left = nLeft + nilY - (210 / 2)
72                Case 18
73                    rpt.imgTnd18.TextColor = &H0&
74                    rpt.imgTnd18.Top = nTop + nilX - (195 / 2)
75                    rpt.imgTnd18.Left = nLeft + nilY - (210 / 2)
76                Case 19
77                    rpt.imgTnd19.TextColor = &H0&
78                    rpt.imgTnd19.Top = nTop + nilX - (195 / 2)
79                    rpt.imgTnd19.Left = nLeft + nilY - (210 / 2)
80                Case 20
81                    rpt.imgTnd20.TextColor = &H0&
82                    rpt.imgTnd20.Top = nTop + nilX - (195 / 2)
83                    rpt.imgTnd20.Left = nLeft + nilY - (210 / 2)
84                Case 21
85                    rpt.imgTnd21.TextColor = &H0&
86                    rpt.imgTnd21.Top = nTop + nilX - (195 / 2)
87                    rpt.imgTnd21.Left = nLeft + nilY - (210 / 2)
88                Case 22
89                    rpt.imgTnd22.TextColor = &H0&
90                    rpt.imgTnd22.Top = nTop + nilX - (195 / 2)
91                    rpt.imgTnd22.Left = nLeft + nilY - (210 / 2)
92                Case 23
93                    rpt.imgTnd23.TextColor = &H0&
94                    rpt.imgTnd23.Top = nTop + nilX - (195 / 2)
95                    rpt.imgTnd23.Left = nLeft + nilY - (210 / 2)
96                Case 24
97                    rpt.imgTnd24.TextColor = &H0&
98                    rpt.imgTnd24.Top = nTop + nilX - (195 / 2)
99                    rpt.imgTnd24.Left = nLeft + nilY - (210 / 2)
100               Case 25
101                   rpt.imgTnd25.TextColor = &H0&
102                   rpt.imgTnd25.Top = nTop + nilX - (195 / 2)
103                   rpt.imgTnd25.Left = nLeft + nilY - (210 / 2)
104               Case 26
105                   rpt.imgTnd26.TextColor = &H0&
106                   rpt.imgTnd26.Top = nTop + nilX - (195 / 2)
107                   rpt.imgTnd26.Left = nLeft + nilY - (210 / 2)
108               Case 27
109                   rpt.imgTnd27.TextColor = &H0&
110                   rpt.imgTnd27.Top = nTop + nilX - (195 / 2)
111                   rpt.imgTnd27.Left = nLeft + nilY - (210 / 2)
112               Case 28
113                   rpt.imgTnd28.TextColor = &H0&
114                   rpt.imgTnd28.Top = nTop + nilX - (195 / 2)
115                   rpt.imgTnd28.Left = nLeft + nilY - (210 / 2)
116               Case 29
117                   rpt.imgTnd29.TextColor = &H0&
118                   rpt.imgTnd29.Top = nTop + nilX - (195 / 2)
119                   rpt.imgTnd29.Left = nLeft + nilY - (210 / 2)
120               Case 30
121                   rpt.imgTnd30.TextColor = &H0&
122                   rpt.imgTnd30.Top = nTop + nilX - (195 / 2)
123                   rpt.imgTnd30.Left = nLeft + nilY - (210 / 2)
124               Case 31
125                   rpt.imgTnd31.TextColor = &H0&
126                   rpt.imgTnd31.Top = nTop + nilX - (195 / 2)
127                   rpt.imgTnd31.Left = nLeft + nilY - (210 / 2)
128               Case 32
129                   rpt.imgTnd32.TextColor = &H0&
130                   rpt.imgTnd32.Top = nTop + nilX - (195 / 2)
131                   rpt.imgTnd32.Left = nLeft + nilY - (210 / 2)
132               Case 33
133                   rpt.imgTnd33.TextColor = &H0&
134                   rpt.imgTnd33.Top = nTop + nilX - (195 / 2)
135                   rpt.imgTnd33.Left = nLeft + nilY - (210 / 2)
136               Case 34
137                   rpt.imgTnd34.TextColor = &H0&
138                   rpt.imgTnd34.Top = nTop + nilX - (195 / 2)
139                   rpt.imgTnd34.Left = nLeft + nilY - (210 / 2)
140               Case 35
141                   rpt.imgTnd35.TextColor = &H0&
142                   rpt.imgTnd35.Top = nTop + nilX - (195 / 2)
143                   rpt.imgTnd35.Left = nLeft + nilY - (210 / 2)
144               Case 36
145                   rpt.imgTnd36.TextColor = &H0&
146                   rpt.imgTnd36.Top = nTop + nilX - (195 / 2)
147                   rpt.imgTnd36.Left = nLeft + nilY - (210 / 2)
148               Case 37
149                   rpt.imgTnd37.TextColor = &H0&
150                   rpt.imgTnd37.Top = nTop + nilX - (195 / 2)
151                   rpt.imgTnd37.Left = nLeft + nilY - (210 / 2)
152               Case 38
153                   rpt.imgTnd38.TextColor = &H0&
154                   rpt.imgTnd38.Top = nTop + nilX - (195 / 2)
155                   rpt.imgTnd38.Left = nLeft + nilY - (210 / 2)
156               Case 39
157                   rpt.imgTnd39.TextColor = &H0&
158                   rpt.imgTnd39.Top = nTop + nilX - (195 / 2)
159                   rpt.imgTnd39.Left = nLeft + nilY - (210 / 2)
160               Case 40
161                   rpt.imgTnd40.TextColor = &H0&
162                   rpt.imgTnd40.Top = nTop + nilX - (195 / 2)
163                   rpt.imgTnd40.Left = nLeft + nilY - (210 / 2)
164               Case 41
165                   rpt.imgTnd41.TextColor = &H0&
166                   rpt.imgTnd41.Top = nTop + nilX - (195 / 2)
167                   rpt.imgTnd41.Left = nLeft + nilY - (210 / 2)
168               Case 42
169                   rpt.imgTnd42.TextColor = &H0&
170                   rpt.imgTnd42.Top = nTop + nilX - (195 / 2)
171                   rpt.imgTnd42.Left = nLeft + nilY - (210 / 2)
172               Case 43
173                   rpt.imgTnd43.TextColor = &H0&
174                   rpt.imgTnd43.Top = nTop + nilX - (195 / 2)
175                   rpt.imgTnd43.Left = nLeft + nilY - (210 / 2)
176               Case 44
177                   rpt.imgTnd44.TextColor = &H0&
178                   rpt.imgTnd44.Top = nTop + nilX - (195 / 2)
179                   rpt.imgTnd44.Left = nLeft + nilY - (210 / 2)
180               Case 45
181                   rpt.imgTnd45.TextColor = &H0&
182                   rpt.imgTnd45.Top = nTop + nilX - (195 / 2)
183                   rpt.imgTnd45.Left = nLeft + nilY - (210 / 2)
184               Case 46
185                   rpt.imgTnd46.TextColor = &H0&
186                   rpt.imgTnd46.Top = nTop + nilX - (195 / 2)
187                   rpt.imgTnd46.Left = nLeft + nilY - (210 / 2)
188               Case 47
189                   rpt.imgTnd47.TextColor = &H0&
190                   rpt.imgTnd47.Top = nTop + nilX - (195 / 2)
191                   rpt.imgTnd47.Left = nLeft + nilY - (210 / 2)
192               Case 48
193                   rpt.imgTnd48.TextColor = &H0&
194                   rpt.imgTnd48.Top = nTop + nilX - (195 / 2)
195                   rpt.imgTnd48.Left = nLeft + nilY - (210 / 2)
196               Case 49
197                   rpt.imgTnd49.TextColor = &H0&
198                   rpt.imgTnd49.Top = nTop + nilX - (195 / 2)
199                   rpt.imgTnd49.Left = nLeft + nilY - (210 / 2)
200               Case 50
201                   rpt.imgTnd50.TextColor = &H0&
202                   rpt.imgTnd50.Top = nTop + nilX - (195 / 2)
203                   rpt.imgTnd50.Left = nLeft + nilY - (210 / 2)
204               Case 51
205                   rpt.imgTnd51.TextColor = &H0&
206                   rpt.imgTnd51.Top = nTop + nilX - (195 / 2)
207                   rpt.imgTnd51.Left = nLeft + nilY - (210 / 2)
208               Case 52
209                   rpt.imgTnd52.TextColor = &H0&
210                   rpt.imgTnd52.Top = nTop + nilX - (195 / 2)
211                   rpt.imgTnd52.Left = nLeft + nilY - (210 / 2)
212               Case 53
213                   rpt.imgTnd53.TextColor = &H0&
214                   rpt.imgTnd53.Top = nTop + nilX - (195 / 2)
215                   rpt.imgTnd53.Left = nLeft + nilY - (210 / 2)
216           End Select
    
217          On Error GoTo 0
218          Exit Sub
    
pasangTanda_Error:
    
219           MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure pasangTanda of Form frmOrderValidasi")
End Sub

Private Function cekDataSingleID(table As String, where As String) As Boolean
    Dim rsHasil As ADODB.Recordset
    Set rsHasil = objO.FetchData("SELECT * FROM [SINGLEIDORI].." & table & " WITH (NOLOCK) " & where)
    If rsHasil.EOF Then
        cekDataSingleID = False
    Else
        cekDataSingleID = True
    End If
End Function

Private Sub cmdKodePlg_Click()
    On Error GoTo cmdKodePlg_Click_Error
    
    With frmSIDCustomer
        gAmbilPlg = "order"
        If Left(Text(17), 3) <> "" Then
            .Text(50) = Left(Text(17), 3)
        End If
        .Text(1) = Mid(Text(17), 4, 6)
        .Text(0) = Mid(Text(17), 10, IIf(Len(Text(17)) > 0, Len(Text(17)) - 9, 0))
        .Text(2) = Text(2)
        .Text(65) = txtKodeLgnLama.Text
        
        If Trim(Text(16)) <> "" Then
            Dim nPol As String
            nPol = Trim(Text(16))
            Dim nDpn1 As String
            Dim nTgh1 As String
            Dim nBlk1 As String
            nDpn1 = ""
            nTgh1 = ""
            nBlk1 = ""
            Call pecahNoPol(nPol, nDpn1, nTgh1, nBlk1)
            .Text(25) = nDpn1
            .Text(46) = nTgh1
            .Text(47) = Trim(nBlk1)
        End If
        .Text(26) = Text(18)
        '        gBantuan = IIf(cmdPilih.Tag = "No. Booking", "booking", "")
        '        gBantuan2 = IIf(cmdPilih.Tag = "No. Booking", text(0), "")
        frmSIDCustomer.show vbModal
    End With
    
    'On Error GoTo 0
    Exit Sub
    
cmdKodePlg_Click_Error:
    
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdKodePlg_Click of Form frmOrder")
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------------------
'   SEND FILE TO FTP
'------------------------------------------------------------------------------------------------------------------------------------------------------------
Public Function SendFTP(record1 As String, record2 As String, record3)
    Dim StErr       As Boolean
    Dim host_name   As String
    Dim v           As Integer
    
1             On Error GoTo SendFTP_Error
2             host_name = FTPAddress
3             If LCase$(Left$(host_name, 6)) <> "ftp://" Then host_name = "ftp://" & host_name
4             FrmMain.inetFTP.AccessType = icDirect
5             FrmMain.inetFTP.Protocol = icFTP
6             FrmMain.inetFTP.URL = host_name
7             FrmMain.inetFTP.UserName = FtpUser
8             FrmMain.inetFTP.Password = FtpPass
    
9             Do While FrmMain.inetFTP.StillExecuting
10                DoEvents
11            Loop
    
12            For v = 0 To 2
13                If v = 0 Then
14                    FrmMain.inetFTP.Execute , "Put " & _
    varDirx & varFilex & " " & "/DSNRMAS/" & varFilex
15                    DbConH.Execute "insert into logReportTAM values ('SEND WO',getdate(),'Finish WO :: " & record1 & " Transfered :: " & varDirx & varFilex & " ',1)"
16                    Do While FrmMain.inetFTP.StillExecuting
17                        DoEvents
18                    Loop
19                ElseIf v = 1 Then
20                    FrmMain.inetFTP.Execute , "Put " & _
    varDirx & varFilex2 & " " & "/DSNRMAS/" & varFilex2
21                    DbConH.Execute "insert into logReportTAM values ('SEND PART',getdate(),'Finish PART :: " & record2 & " Transfered :: " & varDirx & varFilex2 & " ',1)"
22                    Do While FrmMain.inetFTP.StillExecuting
23                        DoEvents
24                    Loop
25                ElseIf v = 2 Then
26                    If hariIni = UCase(Trim(Profile.TANGOHariH)) Then
27                        FrmMain.inetFTP.Execute , "Put " & _
    varDirx & varFilex3 & " " & "/DSNRMAS/" & varFilex3
28                        DbConH.Execute "insert into logReportTAM values ('SEND STOCK',getdate(),'Finish PartStock :: " & record3 & " Transfered :: " & varDirx & varFilex3 & " ',1)"
29                        Do While FrmMain.inetFTP.StillExecuting
30                            DoEvents
31                        Loop
32                    End If
33                End If
34            Next v
    
35            Do While FrmMain.inetFTP.StillExecuting
36                DoEvents
37            Loop
38            Exit Function
    
SendFTP_Error:
39            DbConH.Execute "insert into logReportTAM values ('SEND FTP',getdate(),'" & Replace(err.Description, "'", "~") & "',0)"
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------------
'   SEND FILE TO FTP
'-----------

Private Function UploadFile(ByVal sURL As String _
    , ByVal sUsername As String _
    , ByVal sPassword As String _
    , ByVal sLocalFileName As String _
    , ByVal sRemoteFileName As String) As Boolean
    'Pessimist
    Dim str As String
    On Error GoTo UploadFile_Error
    str = "Put " & sLocalFileName & " " & "/" & sRemoteFileName
    UploadFile = True
    Dim f As New frmOrderValidasiUpload
    With f
        .sUsername = sUsername
        .sPassword = sPassword
        .fileKirim = sLocalFileName
        .fileTujuan = "/" & sRemoteFileName
        .sHost = sURL
        
        Screen.MousePointer = vbHourglass
        '.Inet1.Execute , str
        .txtStr.Text = str
        .Timer1.Enabled = True
        f.show vbModal
    End With
    Screen.MousePointer = vbDefault
    
    '
    '    With Inet1
    '        .UserName = sUserName
    '        .Password = sPassword
    '        .RemotePort = 21
    '        .URL = sURL
    '        str = "PUT " & sLocalFileName & " " & "/" & sRemoteFileName
    '        Screen.MousePointer = vbHourglass
    '        DoEvents
    '        .Execute , str
    '        Sleep 1000
    '
    '        DoEvents
    '
    '        'Mayhaps, a better idea would be to implement
    '        'StateChanged event handler
    '        Do While .StillExecuting
    '            Sleep 1000
    '            DoEvents
    '        Loop
    '
    '        UploadFile = (.ResponseCode = 0)
    '        'Debug.Print .ResponseCode
    '    End With
    '    Screen.MousePointer = vbDefault
    '    DoEvents
    On Error GoTo 0
    Exit Function
    
UploadFile_Error:
    Screen.MousePointer = vbDefault
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure UploadFile of Form frmOrderValidasi"
End Function

'' ================================================================================
'' 20180906 SIMPLIFY PROSES
'' ================================================================================

Private Sub cetakSimplify()
          Dim rpt As Object
          Dim NamaSAOrder As String
          Dim i As Integer
          Dim imgCetak As String
          Dim tandaCetak As String
          Dim hlp As String
1            On Error GoTo cetakSimplify_Error
2             Set rpt = New rtpNewOrder
          
3             If MsgBox("Mencetak Order ?", vbYesNo, "Konfirmasi") = vbYes Then
4                 gKetSA = gSA
5                 isCetak = True
6                 NamaSAOrder = objO.FetchData_String("Select FS_NM_PETUGAS from T_PETUGAS WITH (NOLOCK) where FS_KD_PETUGAS ='" & Trim(gKetSA) & "'")
7                 Call createNewCetakWO
                  Dim rsC As ADODB.Recordset
8                 sSQL = "select a.CUSTOMER,a.CUST_ADRES,a.TELPON,a.ANCUST,a.NO_WO,a.NOFAK,ISNULL(c.date + c.time,'1900-01-01') as tglFollow,"
9                 sSQL = sSQL & " a.FS_KD_PLG as kodelgn,a.FRONT as sa,a.NO_POLISI,a.FMDL,a.COLOR,a.DATE,a.TIME,a.FS_NM_DEALER,a.DDATE,"
10                sSQL = sSQL & " a.MDL,a.NO_CHASSIS,a.NO_ENGINE,a.REQ1,a.REQ2,a.REQ3,a.REQ4,a.REQ5,b.CMPL1,b.CMPL2,b.CMPL3,b.CMPL4,b.CMPL5,"
11                sSQL = sSQL & " a.REP1 , a.REP2, a.REP3, a.REP4, a.REP5, a.ACTRATE1, a.ACTRATE2, a.ACTRATE3, a.ACTRATE4, a.ACTRATE5, a.Off1, a.Off2, a.Off3, a.Off4, a.Off5, a.LABOR"
12                sSQL = sSQL & " ,a.[code] as EM,a.KD_K as SBNP,a.FN_RETURN,a.PDS,a.KM,b.GANTIPART,a.TPSVC,a.FS_RC,b.TUNGGUSVC,b.BBM,b.PEMBAYARAN,a.MECH,a.GRP1,ISNULL(c.DATE_BOOK,'1900-01-01') DATE_BOOK,isnull(m.[GROUP],'') as FO,b.timeserah"
13                sSQL = sSQL & " FROM JOBDTA AS A WITH (NOLOCK)"
14                sSQL = sSQL & " INNER JOIN JOBDTL  AS B WITH (NOLOCK) ON A.NO_WO = B.NO_WO"
15                sSQL = sSQL & " LEFT JOIN MEKANIK m WITH (NOLOCK) ON a.MECH = m.MECH"
16                sSQL = sSQL & " LEFT JOIN BOOK c WITH (NOLOCK) ON ISNULL(A.NOFAK,'') = C.NOFAK and c.NOFAK <> ''"
17                sSQL = sSQL & " WHERE a.no_wo = '" & NulSID(Text(0)) & "'"
                  
18                sSQL = "select a.CUSTOMER,a.CUST_ADRES,a.TELPON,a.ANCUST,a.NO_WO,a.NOFAK,ISNULL((SELECT TOP 1 WAKTU FROM M_FUBOOK with (nolock) where nofak = c.nofak order by waktu desc),'1900-01-01') as tglFollow,"
19                sSQL = sSQL & " a.FS_KD_PLG as kodelgn,a.FRONT as sa,a.NO_POLISI,a.FMDL,a.COLOR,a.DATE,a.TIME,d.FS_NM_DEALER,a.DDATE,"
20                sSQL = sSQL & " a.MDL,a.NO_CHASSIS,a.NO_ENGINE,a.REQ1,a.REQ2,a.REQ3,a.REQ4,a.REQ5,b.CMPL1,b.CMPL2,b.CMPL3,b.CMPL4,b.CMPL5,"
21                sSQL = sSQL & " a.REP1 , a.REP2, a.REP3, a.REP4, a.REP5, a.rte1 ACTRATE1, a.rte2 ACTRATE2, a.rte3 ACTRATE3, a.rte4 ACTRATE4, a.rte5 ACTRATE5, a.Off1, a.Off2, a.Off3, a.Off4, a.Off5, a.LABOR"
                  ''sSQL = sSQL & " ,a.[code] as EM,a.KD_K as SBNP,a.FN_RETURN,a.PDS,a.KM,b.GANTIPART,a.TPSVC,a.FS_RC,b.TUNGGUSVC,b.BBM,b.PEMBAYARAN,a.MECH,a.GRP1,ISNULL(c.DATE_BOOK + c.TIME_D,'1900-01-01') DATE_BOOK,isnull(m.[GROUP],'') as FO,b.timeserah,a.TELPANCUST,d.gsm,d.cdma"
                  ''2020-03-18 | REQ TANGGAL BOOKING KEMBALI KE DATE + TIME SEBELUMNYA DATE_BOOK + TIME_D
                  ''--------------------------------------------------------------------------------------------------
22                sSQL = sSQL & " ,a.[code] as EM,a.KD_K as SBNP,a.FN_RETURN,a.PDS,a.KM,b.GANTIPART,a.TPSVC,a.FS_RC,b.TUNGGUSVC,b.BBM,b.PEMBAYARAN,a.MECH,a.GRP1,ISNULL(c.date + c.time,'1900-01-01') DATE_BOOK,isnull(m.[GROUP],'') as FO,b.timeserah,a.TELPANCUST,d.gsm,d.cdma"
23                sSQL = sSQL & " FROM JOBDTA AS A WITH (NOLOCK)"
24                sSQL = sSQL & " INNER JOIN JOBDTL  AS B WITH (NOLOCK) ON A.NO_WO = B.NO_WO"
25                sSQL = sSQL & " LEFT JOIN MEKANIK m WITH (NOLOCK) ON a.MECH = m.MECH"
26                sSQL = sSQL & " LEFT JOIN CUSTDAT d WITH (NOLOCK) ON d.no_chassis = A.no_chassis"
27                sSQL = sSQL & " LEFT JOIN BOOK c WITH (NOLOCK) ON ISNULL(A.NOFAK,'') = C.NOFAK and c.NOFAK <> ''"
28                sSQL = sSQL & " WHERE a.no_wo = '" & NulSID(Text(0)) & "'"
          'sSQL = sSQL & " WHERE a.no_wo = '1022127'"
          
29                Set rsC = objO.FetchData(sSQL)
30                If Not rsC.EOF Then
          Dim total As Double
31                    total = 0
32                    With rsCetakNewWO
33                        .AddNew
34                        !customer = Left(NulSID(rsC!customer), 40)
35                        !alamat1 = Left(NulSID(rsC!cust_adres), 40)
36                        !alamat1 = Mid(NulSID(rsC!cust_adres), 41, 40)
37                        !Telpon = NulSID(rsC!Telpon) & "," & NulSID(rsC!GSM) & "," & NulSID(rsC!CDMA) ' gsm,cdma
38                        !pembawa = NulSID(rsC!ANCUST)
39                        !noOrder = NulSID(rsC!no_wo)
40                        !NoBooking = IIf(NulSID(rsC!NoFak) <> "", Format(NulSID(rsC!date_book), "dd/mm/yyyy hh:mm") & " [" & NulSID(rsC!NoFak) & "]", "")
41                        !followBooking = NulSID(rsC!TglFollow)
42                        !kodeLgn = NulSID(rsC!kodeLgn)
43                        !sa = NulSID(objO.FetchData_String("SELECT FS_NM_PETUGAS FROM T_PETUGAS WHERE FS_KD_PETUGAS = '" & NulSID(rsC!sa) & "'")) ''NulSID(NamaSAOrder)
44                        !NoPolisi = NulSID(rsC!NO_POLISI)
45                        !model = Left(NulSID(rsC!Fmdl), 40)
46                        !warna = NulSID(rsC!Color)
47                        !tglMasuk = NulSID(rsC!Date)
48                        !jamMasuk = NulSID(rsC!Time)
49                        !dealer = NulSID(rsC!fs_nm_dealer)
50                        !dDate = tglNotNull(NulSID(rsC!dDate))
51                        !MDL = NulSID(rsC!MDL)
52                        !noRangka = NulSID(rsC!NO_CHASSIS)
53                        !noMesin = NulSID(rsC!No_Engine)
54                        !REQ1 = Left(NulSID(rsC!REQ1), 60)
55                        !REQ2 = Left(NulSID(rsC!REQ2), 60)
56                        !REQ3 = Left(NulSID(rsC!REQ3), 60)
57                        !REQ4 = Left(NulSID(rsC!REQ4), 60)
58                        !REQ5 = Left(NulSID(rsC!REQ5), 60)
59                        !comp1 = Left(NulSID(rsC!CMPL1), 60)
60                        !comp2 = Left(NulSID(rsC!CMPL2), 60)
61                        !comp3 = Left(NulSID(rsC!CMPL3), 60)
62                        !comp4 = Left(NulSID(rsC!CMPL4), 60)
63                        !comp5 = Left(NulSID(rsC!CMPL5), 60)
64                        !jasa1 = NulSID(rsC!rep1)
65                        !jasa2 = NulSID(rsC!REP2)
66                        !jasa3 = NulSID(rsC!REP3)
67                        !jasa4 = NulSID(rsC!REP4)
68                        !jasa5 = NulSID(rsC!REP5)
69                        !rate1 = k2t(NulSID(rsC!ACTRATE1))
70                        !rate2 = k2t(NulSID(rsC!ACTRATE2))
71                        !rate3 = k2t(NulSID(rsC!ACTRATE3))
72                        !rate4 = k2t(NulSID(rsC!ACTRATE4))
73                        !rate5 = k2t(NulSID(rsC!ACTRATE5))
74                        !totalJasa1 = NulSID(rsC!Off1)
75                        !totalJasa2 = NulSID(rsC!Off2)
76                        !totalJasa3 = NulSID(rsC!Off3)
77                        !totalJasa4 = NulSID(rsC!Off4)
78                        !totalJasa5 = NulSID(rsC!Off5)
79                        !jasa = getTotalLaborWO(NulSID(rsC!no_wo)) ''NulSID(rsC!labor)
          'Dim hlp  As String
          Dim hlp2 As String
          Dim hlp3 As String
80                        total = 0
          Dim strSql As String
81                        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(rsC!Date), "DD/MM/YYYY") Then
82                            strSql = "select (case When KodeItem = '' Then namaBYSA Else KodeItem end) as KodeItem, " & _
          "Keterangan as Namabarang, SatuanNs as Satuan, Qty as QtyOrder, (isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0)) AS Jml, " & _
          "(Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) as QtyBO, HargaJual as HargaJualA, NOSO as NoRack, " & _
          "0 as DisItem, 0 as ExtDis, Case BO When 'CANCEL' THEN (Qty-(isNull(QtyKirimBI,0) + IsNull(QtyKirimBL,0) + IsNull(QtyKeluar,0))) ELSE 0 END AS QTYCANCEL from SOORD WITH (NOLOCK) Where NOSO = '" & NulSID(rsC!no_wo) & "'"
83                        Else
84                            strSql = "Select Z.*, 0 as QtyOrder, 0 as QtyBO FROM (SELECT INMST.NOITEM AS KODEITEM, INMST.NAMABARANG as NAMABARANG, INMST.KodeSatuan as Satuan, " & _
          "abs((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = INMST.kodeitem)) AS JML, " & _
          "(INMST.HARGAJUALA ) as HARGAJUALA, " & _
          "INMST.NORACK,0 as DisItem, 0 as ExtDis, 0 as QtyCANCEL FROM INMST " & _
          "WHERE ((select Isnull(Sum(INTRN.qty),0) FROM INTRN where INTRN.kodeitem = inmst.kodeitem)) > 0 " & _
          "And INMST.KodeGdg = '" & Trim(gGudang) & "') Z " & _
          "Where Z.Norack = '" & NulSID(rsC!no_wo) & "'"
85                        End If
          Dim rsHlp3 As ADODB.Recordset
86                        Set rsHlp3 = objO.FetchDataACC(strSql)
87                        If Not rsHlp3.EOF Then
88                            rsHlp3.MoveFirst
89                            i = 1
90                            Do While Not rsHlp3.EOF
91                                If i < 12 Then
92                                    hlp = "part" & i
93                                    hlp2 = "qtyPart" & i
94                                    hlp3 = "hargaPart" & i
          
95                                    .Fields.item(hlp).value = NulSID(rsHlp3!NamaBarang)
96                                    .Fields.item(hlp2).value = NulSID(rsHlp3!QtyOrder)
                                      ''.Fields.Item(hlp3).value = NulSID(rsHlp3!HargaJualA)
97                                    .Fields.item(hlp3).value = NulSID(rsHlp3!QtyOrder) * NulSID(rsHlp3!HargaJualA)
                                      ''total = total + NulSID(rsHlp3!HargaJualA)
98                                    total = total + (NulSID(rsHlp3!QtyOrder) * NulSID(rsHlp3!HargaJualA))
99                                Else
100                                   total = total + (NulSID(rsHlp3!QtyOrder) * NulSID(rsHlp3!HargaJualA))
101                               End If
102                               i = i + 1
103                               rsHlp3.MoveNext
104                           Loop
105                       Else
106                           strSql = "SELECT * FROM T_PART WITH (NOLOCK) WHERE NO_WO = '" & NulSID(rsC!no_wo) & "' and fs_log_del = ''"
107                           Set rsHlp3 = objO.FetchData(strSql)
108                           If Not rsHlp3.EOF Then
109                               rsHlp3.MoveFirst
110                               i = 1
111                               Do While Not rsHlp3.EOF
112                                   If i < 12 Then
113                                       hlp = "part" & i
114                                       hlp2 = "qtyPart" & i
115                                       hlp3 = "hargaPart" & i
          
116                                       .Fields.item(hlp).value = NulSID(rsHlp3!keterangan)
117                                       .Fields.item(hlp2).value = NulSID(rsHlp3!qty)
                                          ''.Fields.Item(hlp3).value = NulSID(rsHlp3!Harga)
118                                       .Fields.item(hlp3).value = NulSID(rsHlp3!qty) * NulSID(rsHlp3!Harga)
                                          ''total = total + NulSID(rsHlp3!Harga)
119                                       total = total + (NulSID(rsHlp3!qty) * NulSID(rsHlp3!Harga))
120                                   Else
121                                       total = total + (NulSID(rsHlp3!qty) * NulSID(rsHlp3!Harga))
122                                   End If
123                                   i = i + 1
124                                   rsHlp3.MoveNext
125                               Loop
126                           End If
127                       End If
128                       !part = total
129                       total = 0
130                       sSQL = "SELECT NAMABARANG, QTY, PRICE FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & NulSID(rsC!no_wo) & "' AND no_CHASSIS = '" & NulSID(rsC!NO_CHASSIS) & "'"
131                       Set rsHlp3 = objO.FetchData(sSQL)
132                       If Not rsHlp3.EOF Then
133                           rsHlp3.MoveFirst
134                           i = 1
135                           Do While Not rsHlp3.EOF
136                               If i < 5 Then
137                                   hlp = "sublet" & i
138                                   hlp2 = "qtySublet" & i
139                                   hlp3 = "hargaSublet" & i
          
140                                   .Fields.item(hlp).value = NulSID(rsHlp3!NamaBarang)
141                                   .Fields.item(hlp2).value = NulSID(rsHlp3!qty)
      '                                .Fields.Item(hlp3).value = NulSID(rsHlp3!Price)
      '                                total = total + NulSID(rsHlp3!Price)
                                      ''20200429 | req mb thania qty x harga
142                                   .Fields.item(hlp3).value = NulSID(rsHlp3!qty) * NulSID(rsHlp3!Price)
143                                   total = total + (NulSID(rsHlp3!Price) * NulSID(rsHlp3!qty))
144                               Else
145                                   total = total + (NulSID(rsHlp3!Price) * NulSID(rsHlp3!qty))
146                               End If
147                               i = i + 1
148                               rsHlp3.MoveNext
149                           Loop
150                       End If
151                       !SUBLET = total
152                       !total = !jasa + !part + !SUBLET
153                       If Log.RPLHistori = 1 Then
154                           sSQL = "select top 3 d.INVDATE tanggal,d.KM,d.FRONT,d.FS_NM_DEALER,case when d.FS_RC = '' then 'GRP' else 'SBE' end tipe, REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder from [" & dbReplikasiHO & "]..T_INVOICEnew D with (nolock) where NO_CHASSIS ='" & NulSID(rsC!NO_CHASSIS) & "' AND D.FS_LOG_DEL = '' and d.NOINV not like '%B' order by d.INVDATE desc"
155                       Else
156                           sSQL = "select top 3 d.INVDATE tanggal,d.KM,d.FRONT,d.FS_NM_DEALER,case when d.FS_RC = '' then 'GRP' else 'SBE' end tipe, REPLACE(REPLACE(REPLACE(D.REP1 + '; ' + D.REP2 + '; ' + D.REP3 + '; ' + D.REP4 + '; ' + D.REP5 + '; ' + D.REP6 + '; ' + D.REP7,'; ;',''),'; ;',''),'; ;','') as JobOrder from T_INVOICE D with (nolock) where NO_CHASSIS ='" & NulSID(rsC!NO_CHASSIS) & "' AND D.FS_LOG_DEL = '' and d.NOINV not like '%B' order by d.INVDATE desc"
157                       End If
158                       Set rsHlp3 = objO.FetchData(sSQL)
159                       If Not rsHlp3.EOF Then
160                           rsHlp3.MoveFirst
161                           i = 1
162                           Do While Not rsHlp3.EOF
163                               If i < 4 Then
164                                   hlp = "histTgl" & i
165                                   hlp2 = "histKM" & i
166                                   hlp3 = "histSA" & i
          
167                                   .Fields.item(hlp).value = Format(NulSID(rsHlp3!Tanggal), "dd-mm-yyyy")
168                                   .Fields.item(hlp2).value = IIf(NulSID(rsHlp3!km) = "", 0, NulSID(rsHlp3!km)) ''NulSID(rsHlp3!KM)
169                                   .Fields.item(hlp3).value = NulSID(rsHlp3!FRONT)
          
170                                   hlp = "histDealer" & i
171                                   hlp2 = "histTipe" & i
172                                   hlp3 = "histPekerjaan" & i
          
173                                   .Fields.item(hlp).value = NulSID(rsHlp3!fs_nm_dealer)
174                                   .Fields.item(hlp2).value = NulSID(rsHlp3!tipe)
175                                   .Fields.item(hlp3).value = Left(NulSID(rsHlp3!jobOrder), 120)
176                               End If
177                               i = i + 1
178                               rsHlp3.MoveNext
179                           Loop
180                       End If
181                       .Update
182                   End With
          Dim rsLayanan As ADODB.Recordset
183                   Set rsLayanan = objO.FetchData("SELECT jamlayanan1,jamlayanan2 FROM t_profil WITH (NOLOCK)")
184                   If Not rsLayanan.EOF Then
185                       rsLayanan.MoveFirst
186                       rpt.txtLayanan1.SetText Trim(rsLayanan!jamlayanan1)
187                       rpt.txtLayanan2.SetText Trim(rsLayanan!jamlayanan2)
188                   End If
          
          
189                   rpt.txtKM.SetText NulSID(rsC!km)
          
190                   If NulSID(rsC!GANTIPART) = "1" Then
191                       rpt.txtKonfirmasi.SetText "N. KONF"
192                   Else
193                       rpt.txtKonfirmasi.SetText "KONF"
194                   End If
          
195                   If NulSID(rsC!TPSVC) = "W" Then
196                       rpt.txtJob.SetText "TWC"
197                   ElseIf NulSID(rsC!FN_RETURN) = 1 Then
198                       rpt.txtJob.SetText "RTJ"
199                   ElseIf NulSID(rsC!pds) = 1 Then
200                       rpt.txtJob.SetText "PDS"
201                   ElseIf NulSID(rsC!FS_RC) <> "" Then
202                       If NulSID(rsC!FS_RC) = "R001" Then
203                           rpt.txtJob.SetText "SBI"
204                       Else
205                           rpt.txtJob.SetText "SBE"
206                       End If
207                   ElseIf NulSID(rsC!FS_RC) = "" Then
208                       rpt.txtJob.SetText "GRP"
209                   End If
          
210                   If NulSID(rsC!TUNGGUSVC) = "1" Then
211                       rpt.txtStatus.SetText "Tinggal"
212                   Else
213                       rpt.txtStatus.SetText "Tunggu"
214                   End If
          
215                   If NulSID(rsC!PEMBAYARAN) = "CS" Then
216                       rpt.txtBayar.SetText "Tunai"
217                   ElseIf NulSID(rsC!PEMBAYARAN) = "CR" Then
218                       rpt.txtBayar.SetText "Card"
219                   ElseIf NulSID(rsC!PEMBAYARAN) = "T" Then
220                       rpt.txtBayar.SetText "Tagihan"
221                   End If
          
222                   rpt.txtFO.SetText NulSID(rsC!FO)
223                   rpt.txtMech.SetText NulSID(rsC!Mech)
224                   If Format(NulSID(rsC!date_book), "yyyy-mm-dd") <> "1900-01-01" Then
                          '' 20200305 || TAMBAHAN JAM BOOKING
225                       rpt.txtBooking.SetText Format(NulSID(rsC!date_book), "dd-mm-yyyy hh:mm")
226                   Else
227                       rpt.txtBooking.SetText ""
228                   End If
229                   rpt.txtSA.SetText NulSID(objO.FetchData_String("SELECT FS_NM_PETUGAS FROM T_PETUGAS WHERE FS_KD_PETUGAS = '" & NulSID(rsC!sa) & "'"))
230                   rpt.txtCustomer.SetText NulSID(rsC!ANCUST)
231                   rpt.Text115.SetText "( " & NulSID(rsC!FO) & " )"
232                   rpt.Text116.SetText "( " & NulSID(objO.FetchData_String("SELECT FS_NM_PETUGAS FROM T_PETUGAS WHERE FS_KD_PETUGAS = '" & NulSID(rsC!sa) & "'")) & " )"
          
233                   If NulSID(rsC!em) = 1 Then
234                       rpt.txtEM.SetText "EM"
235                   ElseIf NulSID(rsC!SBNP) = 1 Then
236                       rpt.txtEM.SetText "SBNP"
237                   Else
238                       rpt.txtEM.SetText ""
239                   End If
          
240                   If Log.FIR = 1 Then
241                       Set rsLayanan = objO.FetchData("SELECT TOP 1 * FROM DETAILFIR WITH (NOLOCK) WHERE nowo = '" & NulSID(rsC!no_wo) & "' ")
242                       If Not rsLayanan.EOF Then
243                           rpt.txtFIR.SetText "FIR"
244                       Else
245                           rpt.txtFIR.SetText ""
246                       End If
247                   Else
248                       rpt.txtFIR.SetText ""
249                   End If
          
250                   If NulSID(rsC!bbm) = 0 Then
251                       rpt.txtFuel.SetText "            /"
252                   ElseIf NulSID(rsC!bbm) = 1 Then
253                       rpt.txtFuel.SetText "         /"
254                   ElseIf NulSID(rsC!bbm) = 2 Then
255                       rpt.txtFuel.SetText "      |"
256                   ElseIf NulSID(rsC!bbm) = 3 Then
257                       rpt.txtFuel.SetText "   \"
258                   Else
259                       rpt.txtFuel.SetText "\"
260                   End If
                      
261                   If lblFOP.Caption = "" Then
262                       rpt.Box18.Suppress = True
263                       rpt.Text134.Suppress = True
264                   Else
265                       rpt.Box18.Suppress = False
266                       rpt.Text134.Suppress = False
267                   End If
                      
          
          '' ==============================================
          '' jobkoord
                      Dim koor1, koor2, koor3, koor4, koor5, koor6, koor7, koor8, koor9, koor10, koor11, koor12, koor13, koor14, koor15, koor16, koor17, koor18, koor19, koor20 As Boolean
                      Dim koor21, koor22, koor23, koor24, koor25, koor26, koor27, koor28, koor29, koor30, koor31, koor32, koor33, koor34, koor35, koor36, koor37, koor38, koor39, koor40 As Boolean
                      Dim koor41, koor42, koor43, koor44, koor45, koor46, koor47, koor48, koor49, koor50, koor51, koor52, koor53, koor54  As Boolean
                      
268                   koor1 = False: koor2 = False: koor3 = False: koor4 = False: koor5 = False: koor6 = False: koor7 = False: koor8 = False: koor9 = False: koor10 = False: koor11 = False: koor12 = False: koor13 = False: koor14 = False: koor15 = False: koor16 = False: koor17 = False: koor18 = False: koor19 = False: koor20 = False
269                   koor21 = False: koor22 = False: koor23 = False: koor24 = False: koor25 = False: koor26 = False: koor27 = False: koor28 = False: koor29 = False: koor30 = False: koor31 = False: koor32 = False: koor33 = False: koor34 = False: koor35 = False: koor36 = False: koor37 = False: koor38 = False: koor39 = False: koor40 = False
270                   koor41 = False: koor42 = False: koor43 = False: koor44 = False: koor45 = False: koor46 = False: koor47 = False: koor48 = False: koor49 = False: koor50 = False: koor51 = False: koor52 = False: koor53 = False: koor54 = False
                      
271                   Set rsLayanan = objO.FetchData("SELECT * FROM JOBKOORD WITH (NOLOCK) WHERE NO_WO = '" & NulSID(Text(0)) & "'")
272                   If Not rsLayanan.EOF Then
273                       rpt.txtClear.Suppress = True
274                       rsLayanan.MoveFirst
275                       Do While Not rsLayanan.EOF
276                           If pasangTandaCetakanSimplify(0, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor1 = False Then
277                               rpt.txtKoord1.SetText pasangTandaCetakanSimplify(0, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
278                               koor1 = True
279                               GoTo lanjutloop
280                           End If
                              
281                           If pasangTandaCetakanSimplify(0, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor2 = False Then
282                               rpt.txtKoord2.SetText pasangTandaCetakanSimplify(0, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
283                               koor2 = True
284                               GoTo lanjutloop
285                           End If
                              
286                           If pasangTandaCetakanSimplify(0, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor3 = False Then
287                               rpt.txtKoord3.SetText pasangTandaCetakanSimplify(0, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
288                               koor3 = True
289                               GoTo lanjutloop
290                           End If
                              
291                           If pasangTandaCetakanSimplify(0, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor4 = False Then
292                               rpt.txtKoord4.SetText pasangTandaCetakanSimplify(0, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
293                               koor4 = True
294                               GoTo lanjutloop
295                           End If
                              
296                           If pasangTandaCetakanSimplify(0, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor5 = False Then
297                               rpt.txtKoord5.SetText pasangTandaCetakanSimplify(0, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
298                               koor5 = True
299                               GoTo lanjutloop
300                           End If
                              
301                           If pasangTandaCetakanSimplify(0, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor6 = False Then
302                               rpt.txtKoord6.SetText pasangTandaCetakanSimplify(0, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
303                               koor6 = True
304                               GoTo lanjutloop
305                           End If
                              
306                           If pasangTandaCetakanSimplify(1, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor7 = False Then
307                               rpt.txtKoord7.SetText pasangTandaCetakanSimplify(1, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
308                               koor7 = True
309                               GoTo lanjutloop
310                           End If
                              
311                           If pasangTandaCetakanSimplify(1, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor8 = False Then
312                               rpt.txtKoord8.SetText pasangTandaCetakanSimplify(1, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
313                               koor8 = True
314                               GoTo lanjutloop
315                           End If
                              ''rpt.txtKoord8.SetText pasangTandaCetakanSimplify(1, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
                              
316                           If pasangTandaCetakanSimplify(1, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor9 = False Then
317                               rpt.txtKoord9.SetText pasangTandaCetakanSimplify(1, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
318                               koor9 = True
319                               GoTo lanjutloop
320                           End If
                              
321                           If pasangTandaCetakanSimplify(1, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor10 = False Then
322                               rpt.txtKoord10.SetText pasangTandaCetakanSimplify(1, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
323                               koor10 = True
324                               GoTo lanjutloop
325                           End If
                              
326                           If pasangTandaCetakanSimplify(1, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor11 = False Then
327                               rpt.txtKoord11.SetText pasangTandaCetakanSimplify(1, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
328                               koor11 = True
329                               GoTo lanjutloop
330                           End If
                              
331                           If pasangTandaCetakanSimplify(1, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor12 = False Then
332                               rpt.txtKoord12.SetText pasangTandaCetakanSimplify(1, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
333                               koor12 = True
334                               GoTo lanjutloop
335                           End If
                              
336                           If pasangTandaCetakanSimplify(2, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor13 = False Then
337                               rpt.txtKoord13.SetText pasangTandaCetakanSimplify(2, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
338                               koor13 = True
339                               GoTo lanjutloop
340                           End If
                              
341                           If pasangTandaCetakanSimplify(2, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor14 = False Then
342                               rpt.txtKoord14.SetText pasangTandaCetakanSimplify(2, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
343                               koor14 = True
344                               GoTo lanjutloop
345                           End If
                              
346                           If pasangTandaCetakanSimplify(2, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor15 = False Then
347                               rpt.txtKoord15.SetText pasangTandaCetakanSimplify(2, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
348                               koor15 = True
349                               GoTo lanjutloop
350                           End If
                              
351                           If pasangTandaCetakanSimplify(2, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor16 = False Then
352                               rpt.txtKoord16.SetText pasangTandaCetakanSimplify(2, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
353                               koor16 = True
354                               GoTo lanjutloop
355                           End If
                              
356                           If pasangTandaCetakanSimplify(2, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor17 = False Then
357                               rpt.txtKoord17.SetText pasangTandaCetakanSimplify(2, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
358                               koor17 = True
359                               GoTo lanjutloop
360                           End If
                              
361                           If pasangTandaCetakanSimplify(2, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor18 = False Then
362                               rpt.txtKoord18.SetText pasangTandaCetakanSimplify(2, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
363                               koor18 = True
364                               GoTo lanjutloop
365                           End If
                              
366                           If pasangTandaCetakanSimplify(3, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor19 = False Then
367                               rpt.txtKoord19.SetText pasangTandaCetakanSimplify(3, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
368                               koor19 = True
369                               GoTo lanjutloop
370                           End If
          
371                           If pasangTandaCetakanSimplify(3, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor20 = False Then
372                               rpt.txtKoord20.SetText pasangTandaCetakanSimplify(3, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
373                               koor20 = True
374                               GoTo lanjutloop
375                           End If
          
376                           If pasangTandaCetakanSimplify(3, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor21 = False Then
377                               rpt.txtKoord21.SetText pasangTandaCetakanSimplify(3, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
378                               koor21 = True
379                               GoTo lanjutloop
380                           End If
                              
381                           If pasangTandaCetakanSimplify(3, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor22 = False Then
382                               rpt.txtKoord22.SetText pasangTandaCetakanSimplify(3, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
383                               koor22 = True
384                               GoTo lanjutloop
385                           End If
                              
386                           If pasangTandaCetakanSimplify(3, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor23 = False Then
387                               rpt.txtKoord23.SetText pasangTandaCetakanSimplify(3, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
388                               koor23 = True
389                               GoTo lanjutloop
390                           End If
                              
391                           If pasangTandaCetakanSimplify(3, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor24 = False Then
392                               rpt.txtKoord24.SetText pasangTandaCetakanSimplify(3, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
393                               koor24 = True
394                               GoTo lanjutloop
395                           End If
                              
396                           If pasangTandaCetakanSimplify(4, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor25 = False Then
397                               rpt.txtKoord25.SetText pasangTandaCetakanSimplify(4, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
398                               koor25 = True
399                               GoTo lanjutloop
400                           End If
                              
401                           If pasangTandaCetakanSimplify(4, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor26 = False Then
402                               rpt.txtKoord26.SetText pasangTandaCetakanSimplify(4, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
403                               koor26 = True
404                               GoTo lanjutloop
405                           End If
                              
406                           If pasangTandaCetakanSimplify(4, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor27 = False Then
407                               rpt.txtKoord27.SetText pasangTandaCetakanSimplify(4, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
408                               koor27 = True
409                               GoTo lanjutloop
410                           End If
                              
411                           If pasangTandaCetakanSimplify(4, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor28 = False Then
412                               rpt.txtKoord28.SetText pasangTandaCetakanSimplify(4, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
413                               koor28 = True
414                               GoTo lanjutloop
415                           End If
                              
416                           If pasangTandaCetakanSimplify(4, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor29 = False Then
417                               rpt.txtKoord29.SetText pasangTandaCetakanSimplify(4, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
418                               koor29 = True
419                               GoTo lanjutloop
420                           End If
                              
421                           If pasangTandaCetakanSimplify(4, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor30 = False Then
422                               rpt.txtKoord30.SetText pasangTandaCetakanSimplify(4, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
423                               koor30 = True
424                               GoTo lanjutloop
425                           End If
                         
426                           If pasangTandaCetakanSimplify(5, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor31 = False Then
427                               rpt.txtKoord31.SetText pasangTandaCetakanSimplify(5, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
428                               koor31 = True
429                               GoTo lanjutloop
430                           End If
                              
431                           If pasangTandaCetakanSimplify(5, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor32 = False Then
432                               rpt.txtKoord32.SetText pasangTandaCetakanSimplify(5, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
433                               koor32 = True
434                               GoTo lanjutloop
435                           End If
                              
436                           If pasangTandaCetakanSimplify(5, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor33 = False Then
437                               rpt.txtKoord33.SetText pasangTandaCetakanSimplify(5, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
438                               koor33 = True
439                               GoTo lanjutloop
440                           End If
                              
441                           If pasangTandaCetakanSimplify(5, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor34 = False Then
442                               rpt.txtKoord34.SetText pasangTandaCetakanSimplify(5, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
443                               koor34 = True
444                               GoTo lanjutloop
445                           End If
                              
446                           If pasangTandaCetakanSimplify(5, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor35 = False Then
447                               rpt.txtKoord35.SetText pasangTandaCetakanSimplify(5, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
448                               koor35 = True
449                               GoTo lanjutloop
450                           End If
                              
451                           If pasangTandaCetakanSimplify(5, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor36 = False Then
452                               rpt.txtKoord36.SetText pasangTandaCetakanSimplify(5, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
453                               koor36 = True
454                               GoTo lanjutloop
455                           End If
                              
456                           If pasangTandaCetakanSimplify(6, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor37 = False Then
457                               rpt.txtKoord37.SetText pasangTandaCetakanSimplify(6, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
458                               koor37 = True
459                               GoTo lanjutloop
460                           End If
                              
461                           If pasangTandaCetakanSimplify(6, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor38 = False Then
462                               rpt.txtKoord38.SetText pasangTandaCetakanSimplify(6, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
463                               koor38 = True
464                               GoTo lanjutloop
465                           End If
                              
466                           If pasangTandaCetakanSimplify(6, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor39 = False Then
467                               rpt.txtKoord39.SetText pasangTandaCetakanSimplify(6, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
468                               koor39 = True
469                               GoTo lanjutloop
470                           End If
                              
471                           If pasangTandaCetakanSimplify(6, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor40 = False Then
472                               rpt.txtKoord40.SetText pasangTandaCetakanSimplify(6, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
473                               koor40 = True
474                               GoTo lanjutloop
475                           End If
                              
476                           If pasangTandaCetakanSimplify(6, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor41 = False Then
477                               rpt.txtKoord41.SetText pasangTandaCetakanSimplify(6, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
478                               koor41 = True
479                               GoTo lanjutloop
480                           End If
                              
481                           If pasangTandaCetakanSimplify(6, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor42 = False Then
482                               rpt.txtKoord42.SetText pasangTandaCetakanSimplify(6, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
483                               koor42 = True
484                               GoTo lanjutloop
485                           End If
                              
486                           If pasangTandaCetakanSimplify(7, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor43 = False Then
487                               rpt.txtKoord43.SetText pasangTandaCetakanSimplify(7, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
488                               koor43 = True
489                               GoTo lanjutloop
490                           End If
                              
491                           If pasangTandaCetakanSimplify(7, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor44 = False Then
492                               rpt.txtKoord44.SetText pasangTandaCetakanSimplify(7, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
493                               koor44 = True
494                               GoTo lanjutloop
495                           End If
                              
496                           If pasangTandaCetakanSimplify(7, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor45 = False Then
497                               rpt.txtKoord45.SetText pasangTandaCetakanSimplify(7, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
498                               koor45 = True
499                               GoTo lanjutloop
500                           End If
                              
501                           If pasangTandaCetakanSimplify(7, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor46 = False Then
502                               rpt.txtKoord46.SetText pasangTandaCetakanSimplify(7, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
503                               koor46 = True
504                               GoTo lanjutloop
505                           End If
                              
506                           If pasangTandaCetakanSimplify(7, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor47 = False Then
507                               rpt.txtKoord47.SetText pasangTandaCetakanSimplify(7, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
508                               koor47 = True
509                               GoTo lanjutloop
510                           End If
                              
511                           If pasangTandaCetakanSimplify(7, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor48 = False Then
512                               rpt.txtKoord48.SetText pasangTandaCetakanSimplify(7, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
513                               koor48 = True
514                               GoTo lanjutloop
515                           End If
                              
                              
516                           If pasangTandaCetakanSimplify(8, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor49 = False Then
517                               rpt.txtKoord49.SetText pasangTandaCetakanSimplify(8, 0, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
518                               koor49 = True
519                               GoTo lanjutloop
520                           End If
                              
521                           If pasangTandaCetakanSimplify(8, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor50 = False Then
522                               rpt.txtKoord50.SetText pasangTandaCetakanSimplify(8, 1, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
523                               koor50 = True
524                               GoTo lanjutloop
525                           End If
                              
526                           If pasangTandaCetakanSimplify(8, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor51 = False Then
527                               rpt.txtKoord51.SetText pasangTandaCetakanSimplify(8, 2, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
528                               koor51 = True
529                               GoTo lanjutloop
530                           End If
                              
531                           If pasangTandaCetakanSimplify(8, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor52 = False Then
532                               rpt.txtKoord52.SetText pasangTandaCetakanSimplify(8, 3, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
533                               koor52 = True
534                               GoTo lanjutloop
535                           End If
                              
536                           If pasangTandaCetakanSimplify(8, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor53 = False Then
537                               rpt.txtKoord53.SetText pasangTandaCetakanSimplify(8, 4, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
538                               koor53 = True
539                               GoTo lanjutloop
540                           End If
                              
541                           If pasangTandaCetakanSimplify(8, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY)) <> "" And koor54 = False Then
542                               rpt.txtKoord54.SetText pasangTandaCetakanSimplify(8, 5, NulSID(rsLayanan!koordX), NulSID(rsLayanan!koordY))
543                               koor54 = True
544                               GoTo lanjutloop
545                           End If
                              
lanjutloop:
546                           rsLayanan.MoveNext
547                       Loop
548                   Else
                          
549                       rpt.txtClear.Suppress = False
550                   End If
          '' ==============================================
          
          
551                   rpt.database.SetDataSource rsCetakNewWO
552                   rpt.txtJanjiPenyerahan.SetText NulSID(rsC!timeserah)
553                   rpt.txtTelpCust.SetText NulSID(rsC!TELPANCUST)
554                   rpt.txtNotelpSA.SetText NulSID(objO.FetchData_String("SELECT nohp FROM T_PETUGAS WHERE FS_KD_PETUGAS = '" & NulSID(rsC!sa) & "'"))
                      
555                   rpt.txtBarangDlmKendaraan.SetText Left(NulSID(objO.FetchData_String("select replace(WAE1+','+WAE2+','+WAE3+','+WAE4+','+WAE5,',,','') as kondisi from jobdtl with (nolock) where NO_WO = '" & NulSID(Text(0)) & "'")), 130)
                      ''20200128 || Permintaan tambahan Service Terkait Simplify
556                   rpt.Text96.SetText NulSID(objO.FetchData_String("SELECT NOSTALL FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%|" & Trim(Text(0)) & "|%'"))
                  
557                   With FrmctkOrder.CRViewer1
558                       .ReportSource = rpt
559                       .ViewReport
560                       .Zoom 100
561                       .Visible = True
562                   End With
563                   FrmctkOrder.cmdTutup.Visible = True
564                   Screen.MousePointer = vbDefault
565                   FrmctkOrder.WindowState = vbMaximized
566                   FrmctkOrder.show vbModal
567               End If
568           End If
          
569          On Error GoTo 0
570          Exit Sub
cetakSimplify_Error:
571           MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure cetakSimplify of Form frmOrder"
End Sub

Private Sub createNewCetakWO()
    
    On Error GoTo createNewCetakWO_Error
    If rsCetakNewWO.State Then rsCetakNewWO.Close
    Set rsCetakNewWO = Nothing
    With rsCetakNewWO
        .Fields.Append "customer", adVarChar, 40
        .Fields.Append "alamat1", adVarChar, 40
        .Fields.Append "alamat2", adVarChar, 40
        .Fields.Append "telpon", adVarChar, 40
        .Fields.Append "pembawa", adVarChar, 40
        .Fields.Append "noorder", adVarChar, 40
        .Fields.Append "nobooking", adVarChar, 40
        .Fields.Append "followBooking", adDate
        .Fields.Append "kodelgn", adVarChar, 40
        .Fields.Append "sa", adVarChar, 40
        .Fields.Append "nopolisi", adVarChar, 40
        .Fields.Append "model", adVarChar, 40
        .Fields.Append "warna", adVarChar, 40
        .Fields.Append "tglMasuk", adDate
        .Fields.Append "jamMasuk", adVarChar, 40
        .Fields.Append "dealer", adVarChar, 40
        .Fields.Append "ddate", adDate
        .Fields.Append "mdl", adVarChar, 40
        .Fields.Append "norangka", adVarChar, 40
        .Fields.Append "nomesin", adVarChar, 40
        .Fields.Append "req1", adVarChar, 60
        .Fields.Append "req2", adVarChar, 60
        .Fields.Append "req3", adVarChar, 60
        .Fields.Append "req4", adVarChar, 60
        .Fields.Append "req5", adVarChar, 60
        .Fields.Append "comp1", adVarChar, 60
        .Fields.Append "comp2", adVarChar, 60
        .Fields.Append "comp3", adVarChar, 60
        .Fields.Append "comp4", adVarChar, 60
        .Fields.Append "comp5", adVarChar, 60
        .Fields.Append "jasa1", adVarChar, 60
        .Fields.Append "jasa2", adVarChar, 60
        .Fields.Append "jasa3", adVarChar, 60
        .Fields.Append "jasa4", adVarChar, 60
        .Fields.Append "jasa5", adVarChar, 60
        .Fields.Append "rate1", adSingle
        .Fields.Append "rate2", adSingle
        .Fields.Append "rate3", adSingle
        .Fields.Append "rate4", adSingle
        .Fields.Append "rate5", adSingle
        .Fields.Append "totalJasa1", adDouble
        .Fields.Append "totalJasa2", adDouble
        .Fields.Append "totalJasa3", adDouble
        .Fields.Append "totalJasa4", adDouble
        .Fields.Append "totalJasa5", adDouble
        .Fields.Append "part1", adVarChar, 60
        .Fields.Append "part2", adVarChar, 60
        .Fields.Append "part3", adVarChar, 60
        .Fields.Append "part4", adVarChar, 60
        .Fields.Append "part5", adVarChar, 60
        .Fields.Append "part6", adVarChar, 60
        .Fields.Append "part7", adVarChar, 60
        .Fields.Append "part8", adVarChar, 60
        .Fields.Append "part9", adVarChar, 60
        .Fields.Append "part10", adVarChar, 60
        .Fields.Append "part11", adVarChar, 60
        .Fields.Append "qtyPart1", adInteger
        .Fields.Append "qtyPart2", adInteger
        .Fields.Append "qtyPart3", adInteger
        .Fields.Append "qtyPart4", adInteger
        .Fields.Append "qtyPart5", adInteger
        .Fields.Append "qtyPart6", adInteger
        .Fields.Append "qtyPart7", adInteger
        .Fields.Append "qtyPart8", adInteger
        .Fields.Append "qtyPart9", adInteger
        .Fields.Append "qtyPart10", adInteger
        .Fields.Append "qtyPart11", adInteger
        .Fields.Append "hargaPart1", adDouble
        .Fields.Append "hargaPart2", adDouble
        .Fields.Append "hargaPart3", adDouble
        .Fields.Append "hargaPart4", adDouble
        .Fields.Append "hargaPart5", adDouble
        .Fields.Append "hargaPart6", adDouble
        .Fields.Append "hargaPart7", adDouble
        .Fields.Append "hargaPart8", adDouble
        .Fields.Append "hargaPart9", adDouble
        .Fields.Append "hargaPart10", adDouble
        .Fields.Append "hargaPart11", adDouble
        .Fields.Append "sublet1", adVarChar, 60
        .Fields.Append "sublet2", adVarChar, 60
        .Fields.Append "sublet3", adVarChar, 60
        .Fields.Append "sublet4", adVarChar, 60
        .Fields.Append "qtySublet1", adInteger
        .Fields.Append "qtySublet2", adInteger
        .Fields.Append "qtySublet3", adInteger
        .Fields.Append "qtySublet4", adInteger
        .Fields.Append "hargaSublet1", adDouble
        .Fields.Append "hargaSublet2", adDouble
        .Fields.Append "hargaSublet3", adDouble
        .Fields.Append "hargaSublet4", adDouble
        .Fields.Append "histTgl1", adDate
        .Fields.Append "histKM1", adInteger
        .Fields.Append "histSA1", adVarChar, 10
        .Fields.Append "histDealer1", adVarChar, 40
        .Fields.Append "histTipe1", adVarChar, 10
        .Fields.Append "histPekerjaan1", adVarChar, 120
        .Fields.Append "histTgl2", adDate
        .Fields.Append "histKM2", adInteger
        .Fields.Append "histSA2", adVarChar, 10
        .Fields.Append "histDealer2", adVarChar, 40
        .Fields.Append "histTipe2", adVarChar, 10
        .Fields.Append "histPekerjaan2", adVarChar, 120
        .Fields.Append "histTgl3", adDate
        .Fields.Append "histKM3", adInteger
        .Fields.Append "histSA3", adVarChar, 10
        .Fields.Append "histDealer3", adVarChar, 40
        .Fields.Append "histTipe3", adVarChar, 10
        .Fields.Append "histPekerjaan3", adVarChar, 120
        .Fields.Append "jasa", adDouble
        .Fields.Append "part", adDouble
        .Fields.Append "sublet", adDouble
        .Fields.Append "total", adDouble
        
        .Open
    End With
    On Error GoTo 0
    Exit Sub
createNewCetakWO_Error:
    MsgBox "Error " & err.Number & " (" & err.Description & ") in procedure createNewCetakWO of Form frmOrder"
End Sub


Public Function pasangTandaCetakanSimplify(matY As Integer, matX As Integer, koordX As Integer, koordY As Integer) As String
          Dim hasil As String
1         hasil = ""
          
          ''465
          ''705
          ''945
          ''1185
          ''1425
          ''1665
          
          ''695
          ''935
          ''1175
          ''1415
          ''1655
          ''1895
          ''2135
          ''2375
          ''2615
          
          
2         If matX = 0 And koordX <= 465 Then
3             If matY = 0 And koordY <= 700 Then
4                 hasil = "X"
5             ElseIf matY = 1 And (koordY > 700 And koordY <= 945) Then
6                 hasil = "X"
7             ElseIf matY = 2 And (koordY > 945 And koordY <= 1195) Then
8                 hasil = "X"
9             ElseIf matY = 3 And (koordY > 1195 And koordY <= 1445) Then
10                hasil = "X"
11            ElseIf matY = 4 And (koordY > 1445 And koordY <= 1695) Then
12                hasil = "X"
13            ElseIf matY = 5 And (koordY > 1695 And koordY <= 1945) Then
14                hasil = "X"
15            ElseIf matY = 6 And (koordY > 1945 And koordY <= 2195) Then
16                hasil = "X"
17            ElseIf matY = 7 And (koordY > 2195 And koordY <= 2445) Then
18                hasil = "X"
19            ElseIf matY = 8 And (koordY > 2445 And koordY <= 2685) Then
20                hasil = "X"
21            End If
22        ElseIf matX = 1 And (koordX <= 715 And koordX > 465) Then
23            If matY = 0 And koordY <= 700 Then
24                hasil = "X"
25            ElseIf matY = 1 And (koordY > 700 And koordY <= 945) Then
26                hasil = "X"
27            ElseIf matY = 2 And (koordY > 945 And koordY <= 1195) Then
28                hasil = "X"
29            ElseIf matY = 3 And (koordY > 1195 And koordY <= 1445) Then
30                hasil = "X"
31            ElseIf matY = 4 And (koordY > 1445 And koordY <= 1695) Then
32                hasil = "X"
33            ElseIf matY = 5 And (koordY > 1695 And koordY <= 1945) Then
34                hasil = "X"
35            ElseIf matY = 6 And (koordY > 1945 And koordY <= 2195) Then
36                hasil = "X"
37            ElseIf matY = 7 And (koordY > 2195 And koordY <= 2445) Then
38                hasil = "X"
39            ElseIf matY = 8 And (koordY > 2445 And koordY <= 2685) Then
40                hasil = "X"
41            End If
42        ElseIf matX = 2 And (koordX <= 965 And koordX > 715) Then
43            If matY = 0 And koordY <= 700 Then
44                hasil = "X"
45            ElseIf matY = 1 And (koordY > 700 And koordY <= 945) Then
46                hasil = "X"
47            ElseIf matY = 2 And (koordY > 945 And koordY <= 1195) Then
48                hasil = "X"
49            ElseIf matY = 3 And (koordY > 1195 And koordY <= 1445) Then
50                hasil = "X"
51            ElseIf matY = 4 And (koordY > 1445 And koordY <= 1695) Then
52                hasil = "X"
53            ElseIf matY = 5 And (koordY > 1695 And koordY <= 1945) Then
54                hasil = "X"
55            ElseIf matY = 6 And (koordY > 1945 And koordY <= 2195) Then
56                hasil = "X"
57            ElseIf matY = 7 And (koordY > 2195 And koordY <= 2445) Then
58                hasil = "X"
59            ElseIf matY = 8 And (koordY > 2445 And koordY <= 2685) Then
60                hasil = "X"
61            End If
62        ElseIf matX = 3 And (koordX <= 1215 And koordX > 965) Then
63            If matY = 0 And koordY <= 700 Then
64                hasil = "X"
65            ElseIf matY = 1 And (koordY > 700 And koordY <= 945) Then
66                hasil = "X"
67            ElseIf matY = 2 And (koordY > 945 And koordY <= 1195) Then
68                hasil = "X"
69            ElseIf matY = 3 And (koordY > 1195 And koordY <= 1445) Then
70                hasil = "X"
71            ElseIf matY = 4 And (koordY > 1445 And koordY <= 1695) Then
72                hasil = "X"
73            ElseIf matY = 5 And (koordY > 1695 And koordY <= 1945) Then
74                hasil = "X"
75            ElseIf matY = 6 And (koordY > 1945 And koordY <= 2195) Then
76                hasil = "X"
77            ElseIf matY = 7 And (koordY > 2195 And koordY <= 2445) Then
78                hasil = "X"
79            ElseIf matY = 8 And (koordY > 2445 And koordY <= 2685) Then
80                hasil = "X"
81            End If
82        ElseIf matX = 4 And (koordX <= 1425 And koordX > 1215) Then
83            If matY = 0 And koordY <= 700 Then
84                hasil = "X"
85            ElseIf matY = 1 And (koordY > 700 And koordY <= 945) Then
86                hasil = "X"
87            ElseIf matY = 2 And (koordY > 945 And koordY <= 1195) Then
88                hasil = "X"
89            ElseIf matY = 3 And (koordY > 1195 And koordY <= 1445) Then
90                hasil = "X"
91            ElseIf matY = 4 And (koordY > 1445 And koordY <= 1695) Then
92                hasil = "X"
93            ElseIf matY = 5 And (koordY > 1695 And koordY <= 1945) Then
94                hasil = "X"
95            ElseIf matY = 6 And (koordY > 1945 And koordY <= 2195) Then
96                hasil = "X"
97            ElseIf matY = 7 And (koordY > 2195 And koordY <= 2445) Then
98                hasil = "X"
99            ElseIf matY = 8 And (koordY > 2445 And koordY <= 2685) Then
100               hasil = "X"
101           End If
102       ElseIf matX = 5 And (koordX <= 1730 And koordX > 1425) Then
103           If matY = 0 And koordY <= 700 Then
104               hasil = "X"
105           ElseIf matY = 1 And (koordY > 700 And koordY <= 945) Then
106               hasil = "X"
107           ElseIf matY = 2 And (koordY > 945 And koordY <= 1195) Then
108               hasil = "X"
109           ElseIf matY = 3 And (koordY > 1195 And koordY <= 1445) Then
110               hasil = "X"
111           ElseIf matY = 4 And (koordY > 1445 And koordY <= 1695) Then
112               hasil = "X"
113           ElseIf matY = 5 And (koordY > 1695 And koordY <= 1945) Then
114               hasil = "X"
115           ElseIf matY = 6 And (koordY > 1945 And koordY <= 2195) Then
116               hasil = "X"
117           ElseIf matY = 7 And (koordY > 2195 And koordY <= 2445) Then
118               hasil = "X"
119           ElseIf matY = 8 And (koordY > 2445 And koordY <= 2685) Then
120               hasil = "X"
121           End If
122       End If
123       pasangTandaCetakanSimplify = hasil
End Function

Private Sub createHeaderKTP()
    Set rsKTP = New ADODB.Recordset
    If rsKTP.State Then rsKTP.Close
    With rsKTP
        .Fields.Append "Pilih", adBoolean
        .Fields.Append "IDCust", adVarChar, 20
        .Fields.Append "NoKTP", adVarChar, 50
        .Fields.Append "Nama", adVarChar, 210
        .Fields.Append "Alamat", adVarChar, 350
        .Fields.Append "HP", adVarChar, 30
        .Fields.Append "tglWO", adDate
        .Fields.Append "KdLgn", adVarChar, 20
        .Open
        
    End With
End Sub

Private Sub refreshCustomerNonKTP()
    '    sSQL = "select a.IDCust,a.TypeCustomer,a.NoKTP,a.NamaCust1 + ' ' + a.NamaCust2 + ' ' + a.NamaCust3 as nama,a.Alamat + ' ' + a.Kelurahan + ' ' +  a.Kecamatan + ' ' +  a.Kota + ' ' +  a.Propinsi as alamat,a.HP1"
    '    sSQL = sSQL & " from [singleidori]..newcustomer a with (nolock)"
    '    sSQL = sSQL & " where PinBB = 'scoring' and TypeCustomer = 'PERORANGAN' and a.IDCust not in (select distinct IDCust from [singleidori]..UploadKTP)"
    
    '' --------------------------------------
    '' 20190212 | irwan | data upload ktp default sorting by wo terakhir
    '' --------------------------------------
1             sSQL = "select a.IDCust,a.TypeCustomer,a.NoKTP,a.NamaCust1 + ' ' + a.NamaCust2 + ' ' + a.NamaCust3 as nama,a.Alamat + ' ' + a.Kelurahan + ' ' +  a.Kecamatan + ' ' +  a.Kota + ' ' +  a.Propinsi as alamat,a.HP1,b.KodeLgn"
2             sSQL = sSQL & " ,ISNULL((select Top 1 [DATE] from jobdta where fs_kd_plg = b.KodeLgn order by [DATE] desc),'1900-01-01') as tglwo "
3             sSQL = sSQL & " from [singleidori]..newcustomer a with (nolock)"
4             sSQL = sSQL & " inner join [SingleIDOri]..Bridging b with (nolock) on a.IDCust = b.IDCust and b.KodeDealer = '" & Right(gDb, 2) & "'"
5             sSQL = sSQL & " where PinBB = 'scoring' and TypeCustomer = 'PERORANGAN' and a.IDCust not in (select distinct IDCust from [singleidori]..UploadKTP)"
6             sSQL = sSQL & " order by tglwo desc"
7             Set rsTemp = objO.FetchData(sSQL)
8             Call createHeaderKTP
    '    For Each col In cols
    '        col.FilterText = ""
    '    Next
9             If Not rsTemp.EOF Then
10                rsTemp.MoveFirst
11                Do While Not rsTemp.EOF
12                    With rsKTP
13                        .AddNew
14                        !Pilih = False
15                        !idcust = NulSID(rsTemp!idcust)
16                        !noKTP = NulSID(rsTemp!noKTP)
17                        !Nama = NulSID(rsTemp!Nama)
18                        !Alamat = NulSID(rsTemp!Alamat)
19                        !HP = NulSID(rsTemp!hp1)
20                        !tglWO = NulSID(rsTemp!tglWO)
21                        !kdLgn = NulSID(rsTemp!kodeLgn)
22                        .Update
23                    End With
24                    rsTemp.MoveNext
25                Loop
26            End If
27            Set GridFind.DataSource = rsKTP
    
End Sub

Private Function getTotalLaborWO(noWO As String) As Double
    sSQL = "select SUM(total) as total from ("
    sSQL = sSQL & " select off1+off2+off3+off4+off5+off6+off7+off8 as total from JOBDTA where NO_WO = '" & noWO & "' and FS_LOG_DEL = ''"
    sSQL = sSQL & " Union"
    sSQL = sSQL & " select SUM(offn) as total from JOBADD2 where NO_WO = '" & noWO & "'"
    sSQL = sSQL & " ) a"
    getTotalLaborWO = objO.FetchData_Double(sSQL)
End Function


Private Sub cekPartFOPTerlibat()
          Dim kodeItemPart As String
          Dim terlibat     As Boolean
1         If rsPart.State = adStateOpen Then
2             If rsPart.RecordCount > 0 Then
3                 rsPart.MoveFirst
4                 Do While Not rsPart.EOF
                      ''lakukan pengecekan untuk kodeitem tersebut apakah masuk ke ho_fop_program / ho_fop_program_sub
                      '' ----------------------------------------------------------------------------------------------
5                     terlibat = False
                      ''menghilangkan kode frc pada kodeitem
                      
                      ''kodeItemPart = Mid(Replace(NulSID(rsPart!KodeItem), "-", ""), 2, Len(NulSID(rsPart!KodeItem)))
                      ''kodeItemPart = Mid(NulSID(rsPart!KodeItem), 3, Len(NulSID(rsPart!KodeItem)))
6                     kodeItemPart = NulSID(rsPart!KodeItem)
                      
7                     If objO.FetchData_String("select ITEM_PART from [" & dbReplikasiHO & "]..HO_FOP_PROGRAM with (nolock) where MDL = '" & kodeMDL & "' and ITEM_PART = '" & kodeItemPart & "' and fs_log_del = ''") <> "" Then
8                         terlibat = True
                      ''ElseIf  objO.FetchData_String("select * from [" & dbReplikasiHO & "]..HO_FOP_PROGRAM_SUB with (nolock) where MDL = '" & kodeMDL & "' and ITEM_PARTsub = '" & KodeItem & "'") <> "" Then
9                     ElseIf tagihanFOP = "TAM" And objO.FetchData_String("select ITEM_PARTsub from [" & dbReplikasiHO & "]..HO_FOP_PROGRAM_SUB with (nolock) where MDL = '" & kodeMDL & "' and ITEM_PARTsub = '" & kodeItemPart & "' and fs_log_del = ''") <> "" Then
10                        terlibat = True
11                    End If
12                    rsPart!terlibat = terlibat
13                    rsPart!FREEC = True
14                    rsPart.MoveNext
15                Loop
16                Call doHeadPart
17            End If
              
18        End If
End Sub

Private Sub cekSubletNonFOP(noWO As String, noKa As String)
          Dim rsHlp           As ADODB.Recordset
          Dim terlibat        As Boolean
          Dim isSudahDiskon   As Boolean
          ''jika sudah didiskon price diambil dr jobsublet = nilai setelah diskon , jika tidak diskon price hargajual
1         isSudahDiskon = False
2         sSQL = "select * from vo_discsublet with (nolock) where FS_LOG_DEL = '' and NO_WO = '" & noWO & "' AND no_CHASSIS = '" & noKa & "'"
3         If objO.FetchData_Count(sSQL) > 0 Then
4             isSudahDiskon = True
5         End If
6         Call doCreateSublet
7         Set rsHlp = objO.FetchData("SELECT KDPROMO,KODEITEM,NONOTA, SUPPLIER, NAMABARANG, KODESATUAN, QTY, valid, HARGA, PRICE,cast(-1 as bit) as freec FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & noWO & "' AND no_CHASSIS = '" & noKa & "'")
8         If Not rsHlp.EOF Then
9             rsHlp.MoveFirst
10            Do While Not rsHlp.EOF
11                With rsSublet
12                    .AddNew
13                    !nonota = NulSID(rsHlp!nonota)
14                    !supplier = NulSID(rsHlp!supplier)
15                    !NamaBarang = NulSID(rsHlp!NamaBarang)
16                    !KODESATUAN = NulSID(rsHlp!KODESATUAN)
17                    !qty = NulSID(rsHlp!qty)
18                    !valid = NulSID(rsHlp!valid)
19                    !Harga = NulSID(SpyRound(rsHlp!Harga, 0))
20                    If isSudahDiskon = True Then
21                        !Price = NulSID(rsHlp!Price)
22                    Else
23                        If NulSID(rsHlp!KDPROMO) <> "" Then
24                            sSQL = "select HARGAJUAL * (" & k2t(profilACC.nilaitotal / profilACC.nilaidpp) & ") from [" & dbReplikasiHO & "]..MSTSUBLETPROMO with (NOLOCK) where KDPROMO = '" & NulSID(rsHlp!KDPROMO) & "' and FS_LOG_DEL = ''"
25                        Else
26                            sSQL = "SELECT HARGAJUAL * " & k2t(profilACC.nilaitotal / profilACC.nilaidpp) & " from [" & dbReplikasiHO & "]..MSTSUBLET WHERE KDSUBLET = '" & NulSID(rsHlp!KodeItem) & "' AND FS_LOG_DEL = ''"
27                        End If
28                        !Price = SpyRound(objO.FetchData_Double(sSQL), 0)
29                    End If
30                    !FREEC = True
31                    !KodeItem = NulSID(rsHlp!KodeItem)
32                    !Subtotal = NulSID(rsHlp!Price) * NulSID(rsHlp!qty)
33                    terlibat = False
34                    !diskon = 0
35                    .Update
36                End With
37                rsHlp.MoveNext
38            Loop
39        End If
End Sub

Private Sub cekSubletTerlibatFOP(noWO As String, noKa As String)
          Dim rsHlp           As ADODB.Recordset
          Dim terlibat        As Boolean
1         Call doCreateSublet
2         Set rsHlp = objO.FetchData("SELECT KODEITEM,NONOTA, SUPPLIER, NAMABARANG, KODESATUAN, QTY, valid, HARGA, PRICE,cast(-1 as bit) as freec FROM jobSublet WITH (NOLOCK) WHERE no_WO = '" & noWO & "' AND no_CHASSIS = '" & noKa & "'")
3         If Not rsHlp.EOF Then
4             rsHlp.MoveFirst
5             Do While Not rsHlp.EOF
6                 With rsSublet
7                     .AddNew
8                     !nonota = NulSID(rsHlp!nonota)
9                     !supplier = NulSID(rsHlp!supplier)
10                    !NamaBarang = NulSID(rsHlp!NamaBarang)
11                    !KODESATUAN = NulSID(rsHlp!KODESATUAN)
12                    !qty = NulSID(rsHlp!qty)
13                    !valid = NulSID(rsHlp!valid)
14                    !Harga = NulSID(rsHlp!Harga)
15                    !Price = NulSID(rsHlp!Price)
16                    !FREEC = True
17                    !KodeItem = NulSID(rsHlp!KodeItem)
18                    terlibat = False
19                    !diskon = 0
20                    If objO.FetchData_String("select ITEM_PART from [" & dbReplikasiHO & "]..HO_FOP_PROGRAM with (nolock) where MDL = '" & kodeMDL & "' and ITEM_PART = '" & NulSID(rsHlp!KodeItem) & "'") <> "" Then
21                        terlibat = True
22                    End If
23                    !terlibat = terlibat
24                    .Update
25                End With
26                rsHlp.MoveNext
27            Loop
28        End If
End Sub

Private Sub doCreateSublet()

   On Error GoTo doCreateSublet_Error
    If rsSublet.State = adStateOpen Then rsSublet.Close
    With rsSublet
        .Fields.Append "NONOTA", adVarChar, 50
        .Fields.Append "KODEITEM", adVarChar, 50
        .Fields.Append "SUPPLIER", adVarChar, 200
        .Fields.Append "NAMABARANG", adVarChar, 200
        .Fields.Append "KODESATUAN", adVarChar, 50
        .Fields.Append "qty", adDouble
        .Fields.Append "diskon", adDouble
        .Fields.Append "VALID", adInteger
        .Fields.Append "Harga", adDouble
        .Fields.Append "subtotal", adDouble
        .Fields.Append "Price", adDouble
        .Fields.Append "FREEC", adBoolean
        .Fields.Append "TERLIBAT", adBoolean
        .Open
    End With
    

   On Error GoTo 0
   Exit Sub

doCreateSublet_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, err.Description) & " in procedure doCreateSublet of Form frmOrderValidasi"


End Sub


Private Sub panggilDIDXbyWO(noWO As String, noRangka As String)
    Dim rsFOPTerpilih As ADODB.Recordset
    Set rsFOPTerpilih = objO.FetchData("select kode_item,CrossSell,UpSell from JobPartDX where NO_WO = '" & noWO & "' and NO_CHASSIS = '" & noRangka & "' and FS_LOG_DEL = ''")
    If rsFOPTerpilih.RecordCount > 0 Then
        rsFOPTerpilih.MoveFirst
        If rsPart.State = adStateOpen Then
            If rsPart.RecordCount > 0 Then
                rsPart.MoveFirst
                Do While Not rsPart.EOF
                    rsFOPTerpilih.MoveFirst
                    Do While Not rsFOPTerpilih.EOF
                        If rsPart!KodeItem = rsFOPTerpilih!KODE_ITEM Then
                            rsPart!cross = IIf(rsFOPTerpilih!CrossSell = 1, True, False)
                            rsPart!up = IIf(rsFOPTerpilih!UpSell = 1, True, False)
                        End If
                        rsFOPTerpilih.MoveNext
                    Loop
                    rsPart.MoveNext
                Loop
                rsPart.MoveFirst
            End If
        End If
    End If
End Sub

Private Sub panggilFOPbyWO(noWO As String, noRangka As String)
    Dim rsFOPTerpilih As ADODB.Recordset
    Set rsFOPTerpilih = objO.FetchData("select kodeitem,FREEC,terlibatFOP from VO_FOPPart where NO_WO = '" & noWO & "' and NO_CHASSIS = '" & noRangka & "' and FS_LOG_DEL = ''")
    If rsFOPTerpilih.RecordCount > 0 Then
        rsFOPTerpilih.MoveFirst
        If rsPart.State = adStateOpen Then
            If rsPart.RecordCount > 0 Then
                rsPart.MoveFirst
                Do While Not rsPart.EOF
                    rsFOPTerpilih.MoveFirst
                    Do While Not rsFOPTerpilih.EOF
                        If rsPart!KodeItem = rsFOPTerpilih!KodeItem Then
                            rsPart!FREEC = IIf(rsFOPTerpilih!FREEC = 1, True, False)
                        End If
                        rsFOPTerpilih.MoveNext
                    Loop
                    rsPart.MoveNext
                Loop
                rsPart.MoveFirst
                Do While Not rsPart.EOF
                    If rsPart!FREEC = False Then
                        If rsPart!oli Then
                            Text(25) = Text(25) - (rsPart!Harga * rsPart!qty)
                        Else
                            ''part
                            Text(24) = Text(24) - (rsPart!Harga * rsPart!qty)
                        End If
                        Text(29) = Round(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0)
                    End If
                    rsPart.MoveNext
                Loop
            End If
        End If
        rsFOPTerpilih.MoveFirst
        If rsSublet.State = adStateOpen Then
            If rsSublet.RecordCount > 0 Then
                rsSublet.MoveFirst
                Do While Not rsSublet.EOF
                    rsFOPTerpilih.MoveFirst
                    Do While Not rsFOPTerpilih.EOF
                        If rsSublet!KodeItem = rsFOPTerpilih!KodeItem Then
                            rsSublet!FREEC = IIf(rsFOPTerpilih!FREEC = 1, True, False)
                        End If
                        rsFOPTerpilih.MoveNext
                    Loop
                    rsSublet.MoveNext
                Loop
                rsSublet.MoveFirst
                Do While Not rsSublet.EOF
                    If rsSublet!FREEC = False Then
                        Text(27) = Text(27) - (rsSublet!Price * rsSublet!qty)
                        Text(29) = Round(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0)
                    End If
                    rsSublet.MoveNext
                Loop
            End If
        End If
        
    End If
End Sub


Private Sub panggilDiscSubletValidasi()
          Dim diskon As Double
          Dim hargaJualAsli As Double
1         If Log.standarisasiSublet = 1 Then
2             If rsSublet.State = adStateOpen Then
3                 If rsSublet.RecordCount > 0 Then
4                     rsSublet.MoveFirst
5                     Do While Not rsSublet.EOF
6                         sSQL = "SELECT [DISC],hargajualasli FROM [vo_discsublet] WITH (NOLOCK) WHERE [NO_WO] = '" & NulSID(Text(0)) & "' and [NO_CHASSIS] = '" & NulSID(Text(17)) & "' and [kodeitem] = '" & NulSID(rsSublet!KodeItem) & "' and [FS_LOG_DEL] = ''"
7                         diskon = objO.FetchData_Double(sSQL)
8                         rsSublet!diskon = diskon
                          
                          ''20210303 jika sudah di invoice tidak perlu hitung ulang sub total
                          ''--------------------------------------------------------------------------
9                         If NulSID(Text(20)) = "" Then
                              ''20210305 jika diskon ada mengambil price dari vo_discsublet hargajualasli
10                            If diskon > 0 Then
11                              rsSublet!Price = objO.FetchData_Double("SELECT hargajualasli FROM [vo_discsublet] WITH (NOLOCK) WHERE [NO_WO] = '" & NulSID(Text(0)) & "' and [NO_CHASSIS] = '" & NulSID(Text(17)) & "' and [kodeitem] = '" & NulSID(rsSublet!KodeItem) & "' and [FS_LOG_DEL] = ''")
12                            End If
13                            rsSublet!Subtotal = ((rsSublet!Price * rsSublet!qty) * ((100 - rsSublet!diskon) / 100))
14                        Else
                              
15                            sSQL = "SELECT hargajual*" & k2t(profilACC.nilaitotal / profilACC.nilaidpp) & " FROM [" & dbReplikasiHO & "]..MSTSUBLET WITH (NOLOCK) WHERE KDSUBLET = '" & NulSID(rsSublet!KodeItem) & "' AND fs_log_del = '' "
                              hargaJualAsli = objO.FetchData_Double(sSQL)
                              If hargaJualAsli = 0 Then
                                rsSublet!diskon = 0
                              Else
                                rsSublet!diskon = Round(((hargaJualAsli - rsSublet!Price) / hargaJualAsli) * 100, 2)
                              End If
16                            rsSublet!Price = hargaJualAsli ''objO.FetchData_Double(sSQL) ''SpyRound((100 * rsSublet!Price) / (100 - diskon), 0)
17                        End If
18                        rsSublet.MoveNext
19                    Loop
20                    rsSublet.MoveFirst
21                    If NulSID(Text(20)) = "" Then
22                        Text(27) = 0
23                        Do While Not rsSublet.EOF
24                            Text(27) = Text(27) + ((rsSublet!Price * rsSublet!qty) * ((100 - rsSublet!diskon) / 100))
25                            Text(29) = Round(CDbl(Text(22)) + CDbl(Text(24)) + CDbl(Text(25)) + CDbl(Text(27)), 0)
26                            rsSublet.MoveNext
27                        Loop
28                    End If
29                End If
30            End If
31        End If
End Sub


Private Function isCurrentAR(kodeLgn As String) As Boolean
   On Error GoTo isCurrentAR_Error

    isCurrentAR = False
    Dim aBatasKr As Double
    Dim aCurrent As Double
    aBatasKr = objO.Fetch_BatasKredit_Ar(Trim(Replace(kodeLgn, ">", "")))
    isCurrentAR = False
    aCurrent = objO.Fetch_Current_Ar(Trim(kodeLgn))
    If aCurrent <> 0 And aBatasKr <> 0 Then
        nPersenAR = Round((aCurrent / aBatasKr) * 100, 0)
    Else
        nPersenAR = 0
    End If
    If aBatasKr = 0 Then GoTo TanpaBatasAR
    If nPersenAR >= 100 Then
        isCurrentAR = True
    ElseIf nPersenAR >= 80 Then
        isCurrentAR = True
    ElseIf nPersenAR < 80 Then
        isCurrentAR = False
TanpaBatasAR:
        isCurrentAR = False
    End If

   'On Error GoTo 0
   Exit Function

isCurrentAR_Error:
        isCurrentAR = False

End Function

Private Function isOverdueAR(kodeLgn As String) As Boolean
    Dim nOAR As ADODB.Recordset
    On Error GoTo isOverdueAR_Error

    Set nOAR = objO.FetchOverdueAR(Trim(Replace(kodeLgn, ">", "")))
    If nOAR.RecordCount > 0 Then
        Set nOAR = objO.FetchOverdueAR_Data(Trim(Replace(kodeLgn, ">", "")))
        If nOAR.RecordCount > 0 Then
            isOverdueAR = True
        Else
            isOverdueAR = False
        End If
    Else
        isOverdueAR = False
    End If

       'On Error GoTo 0
       Exit Function

isOverdueAR_Error:
End Function
