VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm3DFU 
   BackColor       =   &H00FFFFFF&
   Caption         =   "3 Day's Follow Up"
   ClientHeight    =   10680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   10680
   ScaleWidth      =   13260
   Begin VB.Frame FrmCekLockFU 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame10"
      Height          =   2775
      Left            =   5040
      TabIndex        =   118
      Top             =   840
      Visible         =   0   'False
      Width           =   5775
      Begin TrueOleDBGrid70.TDBGrid GridLockFU 
         Height          =   2535
         Left            =   120
         TabIndex        =   119
         Top             =   120
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   4471
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
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
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
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
         _StyleDefs(69)  =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(70)  =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(71)  =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(72)  =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         _StyleDefs(73)  =   "Named:id=91:responCL"
         _StyleDefs(74)  =   ":id=91,.parent=131,.fgcolor=&HFF00&,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(75)  =   ":id=91,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(76)  =   ":id=91,.fontname=MS Sans Serif"
      End
   End
   Begin VB.Frame frmeTidakDiterima 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   1560
      TabIndex        =   70
      Top             =   3360
      Visible         =   0   'False
      Width           =   7095
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   60
         Picture         =   "frm3DFU.frx":0000
         ScaleHeight     =   2535
         ScaleWidth      =   6975
         TabIndex        =   90
         Top             =   0
         Width           =   6975
         Begin VB.TextBox txtAlasanLain 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   96
            Top             =   1580
            Width           =   5295
         End
         Begin VB.OptionButton OptAlasan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "4. Lainnya"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   97
            Top             =   1560
            Width           =   6615
         End
         Begin VB.OptionButton OptAlasan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "3. Tidak bertemu dengan yang bersangkutan (untuk Fix Phone)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   95
            Top             =   1200
            Width           =   6615
         End
         Begin VB.OptionButton OptAlasan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "2. No Telp Tidak Aktif / Sibuk"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   94
            Top             =   840
            Width           =   6615
         End
         Begin VB.OptionButton OptAlasan 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "1. No Telp Tidak Valid"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   93
            Top             =   480
            Value           =   -1  'True
            Width           =   6615
         End
         Begin VB.CommandButton simpanAlasan 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   2400
            Picture         =   "frm3DFU.frx":25E8
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   2040
            Width           =   1215
         End
         Begin VB.CommandButton cmdBatalAlsn 
            BackColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   3600
            Picture         =   "frm3DFU.frx":3CCE
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   2040
            Width           =   1215
         End
      End
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
      Left            =   15240
      Style           =   1  'Graphical
      TabIndex        =   61
      Top             =   600
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   15120
      TabIndex        =   62
      Top             =   120
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
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   9615
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   14415
      Begin VB.CommandButton cmdCekFU 
         BackColor       =   &H00DCFCFB&
         Caption         =   "Cek WO lock FU"
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
         TabIndex        =   117
         Top             =   100
         Width           =   1890
      End
      Begin VB.Frame frmeDial 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   8640
         TabIndex        =   52
         Top             =   5400
         Visible         =   0   'False
         Width           =   4575
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1695
            Left            =   120
            Picture         =   "frm3DFU.frx":53B4
            ScaleHeight     =   1695
            ScaleWidth      =   4455
            TabIndex        =   72
            Top             =   0
            Width           =   4455
            Begin VB.CommandButton cmdTerima 
               BackColor       =   &H00DCFCFB&
               Caption         =   "Diterima"
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
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   75
               Top             =   960
               Width           =   1890
            End
            Begin VB.CommandButton cmdTolak 
               BackColor       =   &H00DCFCFB&
               Caption         =   "Tidak Ada Respon"
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
               Left            =   2400
               Style           =   1  'Graphical
               TabIndex        =   74
               Top             =   960
               Width           =   1890
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "No. WO"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   0
               Left            =   240
               TabIndex        =   73
               Top             =   480
               Width           =   690
            End
         End
      End
      Begin VB.Frame framLS 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "                       Cetak LS                             "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2030
         Left            =   9600
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   4555
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   60
            Picture         =   "frm3DFU.frx":6637
            ScaleHeight     =   1935
            ScaleWidth      =   4455
            TabIndex        =   76
            Top             =   60
            Width           =   4455
            Begin VB.CommandButton cetakRepoLS 
               BackColor       =   &H00DCFCFB&
               Caption         =   "&Cetak"
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
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   79
               Top             =   1320
               Width           =   1890
            End
            Begin VB.OptionButton opTgl 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Tgl Follow"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   960
               TabIndex        =   78
               Top             =   480
               Width           =   1215
            End
            Begin VB.OptionButton opTgl 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "Tgl Unit Entry"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   2160
               TabIndex        =   77
               Top             =   480
               Value           =   -1  'True
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker tanggal1 
               Height          =   315
               Left            =   960
               TabIndex        =   80
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Format          =   433848321
               CurrentDate     =   39859
            End
            Begin MSComCtl2.DTPicker tanggal2 
               Height          =   315
               Left            =   2760
               TabIndex        =   81
               Top             =   840
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   556
               _Version        =   393216
               Format          =   433848321
               CurrentDate     =   39859
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Range By"
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
               TabIndex        =   84
               Top             =   480
               Width           =   690
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Range Tgl"
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
               Left            =   90
               TabIndex        =   83
               Top             =   840
               Width           =   720
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   15.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2580
               TabIndex        =   82
               Top             =   760
               Width           =   135
            End
         End
      End
      Begin VB.Frame framKeluhan 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   $"frm3DFU.frx":7CF4
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   840
         TabIndex        =   57
         Top             =   6000
         Width           =   8175
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1935
            Left            =   120
            Picture         =   "frm3DFU.frx":7D92
            ScaleHeight     =   1935
            ScaleWidth      =   8055
            TabIndex        =   85
            Top             =   0
            Width           =   8055
            Begin VB.TextBox Text1 
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
               Height          =   285
               Index           =   2
               Left            =   600
               MaxLength       =   200
               TabIndex        =   89
               Top             =   1010
               Width           =   7335
            End
            Begin VB.TextBox Text1 
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
               Height          =   285
               Index           =   1
               Left            =   600
               MaxLength       =   200
               TabIndex        =   88
               Top             =   740
               Width           =   7335
            End
            Begin VB.TextBox Text1 
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
               Height          =   285
               Index           =   0
               Left            =   600
               MaxLength       =   200
               TabIndex        =   87
               Top             =   480
               Width           =   7335
            End
            Begin VB.CommandButton cmdSaveKeluhan 
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
               Left            =   5280
               Picture         =   "frm3DFU.frx":9F3E
               Style           =   1  'Graphical
               TabIndex        =   86
               Top             =   1420
               Width           =   1290
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Q3 :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   3
               Left            =   120
               TabIndex        =   122
               Top             =   1032
               Width           =   390
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Q2 :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   2
               Left            =   120
               TabIndex        =   121
               Top             =   762
               Width           =   390
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Q1 :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Index           =   1
               Left            =   120
               TabIndex        =   120
               Top             =   480
               Width           =   390
            End
            Begin VB.Image btnBatal 
               Height          =   360
               Left            =   6600
               Picture         =   "frm3DFU.frx":B624
               Stretch         =   -1  'True
               Top             =   1440
               Width           =   1335
            End
         End
      End
      Begin VB.CommandButton cmdLS 
         Caption         =   "Cetak &LS"
         Height          =   390
         Left            =   12240
         TabIndex        =   59
         Top             =   120
         Width           =   1890
      End
      Begin VB.Frame FU 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   120
         TabIndex        =   20
         Top             =   4320
         Width           =   14055
         Begin VB.Frame frmeFU 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4215
            Left            =   8400
            TabIndex        =   26
            Top             =   120
            Visible         =   0   'False
            Width           =   5500
            Begin VB.Frame frameStatus 
               BackColor       =   &H00FFFFFF&
               Height          =   1575
               Left            =   885
               TabIndex        =   103
               Top             =   2385
               Visible         =   0   'False
               Width           =   4095
               Begin VB.CommandButton cmdStatusOk 
                  BackColor       =   &H0000FF00&
                  Caption         =   "OK"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   1680
                  Style           =   1  'Graphical
                  TabIndex        =   109
                  Top             =   1200
                  Width           =   735
               End
               Begin VB.OptionButton optStatus 
                  BackColor       =   &H000000FF&
                  Height          =   375
                  Index           =   2
                  Left            =   2280
                  TabIndex        =   107
                  Top             =   720
                  Width           =   255
               End
               Begin VB.OptionButton optStatus 
                  BackColor       =   &H0000FFFF&
                  Height          =   375
                  Index           =   1
                  Left            =   1920
                  TabIndex        =   106
                  Top             =   720
                  Width           =   255
               End
               Begin VB.OptionButton optStatus 
                  BackColor       =   &H0000C000&
                  Height          =   375
                  Index           =   0
                  Left            =   1560
                  TabIndex        =   105
                  Top             =   720
                  Width           =   255
               End
               Begin VB.Label Label19 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Pilih Status Customer"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   120
                  TabIndex        =   108
                  Top             =   430
                  Width           =   3855
               End
               Begin VB.Label Label18 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "Status Customer"
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
                  Left            =   0
                  TabIndex        =   104
                  Top             =   120
                  Width           =   4095
               End
            End
            Begin VB.CommandButton cmdBatal 
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
               Left            =   3000
               Picture         =   "frm3DFU.frx":CD0A
               Style           =   1  'Graphical
               TabIndex        =   67
               Top             =   3720
               Width           =   1290
            End
            Begin VB.CommandButton tblSimpan 
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
               Left            =   1680
               Picture         =   "frm3DFU.frx":D1D2
               Style           =   1  'Graphical
               TabIndex        =   51
               Top             =   3720
               Width           =   1290
            End
            Begin VB.Frame Frame8 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2040
               TabIndex        =   45
               Top             =   480
               Width           =   3135
               Begin VB.OptionButton opTepat 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Tepat"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   48
                  Top             =   0
                  Width           =   735
               End
               Begin VB.OptionButton opTepat 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Cukup"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   47
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.OptionButton opTepat 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Kurang"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   46
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   855
               End
            End
            Begin VB.Frame Frame7 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2040
               TabIndex        =   41
               Top             =   840
               Width           =   3135
               Begin VB.OptionButton opLama 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Cepat"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   44
                  Top             =   0
                  Width           =   735
               End
               Begin VB.OptionButton opLama 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Cukup"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   43
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.OptionButton opLama 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Lama"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   42
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   855
               End
            End
            Begin VB.Frame Frame6 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Left            =   960
               TabIndex        =   35
               Top             =   3075
               Width           =   2535
               Begin VB.OptionButton opRespon 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Cukup Baik"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   1
                  Left            =   840
                  TabIndex        =   13
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   1215
               End
               Begin VB.OptionButton opRespon 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Baik"
                  ForeColor       =   &H80000008&
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   12
                  Top             =   120
                  Width           =   975
               End
            End
            Begin VB.Frame Frame5 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2040
               TabIndex        =   30
               Top             =   1560
               Width           =   3135
               Begin VB.OptionButton opRamah 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Tidak"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   9
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   855
               End
               Begin VB.OptionButton opRamah 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Cukup"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   8
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.OptionButton opRamah 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Ya"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   7
                  Top             =   0
                  Width           =   735
               End
            End
            Begin VB.Frame Frame4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2040
               TabIndex        =   29
               Top             =   1200
               Width           =   3135
               Begin VB.OptionButton opJelas 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Kurang"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   6
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton opJelas 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Cukup"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   5
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.OptionButton opJelas 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Jelas"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   4
                  Top             =   0
                  Width           =   735
               End
            End
            Begin VB.Frame Frame3 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   2040
               TabIndex        =   28
               Top             =   120
               Width           =   3135
               Begin VB.OptionButton opKualitas 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Kurang"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   2
                  Left            =   1080
                  TabIndex        =   2
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton opKualitas 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Cukup"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   1
                  Left            =   2160
                  TabIndex        =   3
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   975
               End
               Begin VB.OptionButton opKualitas 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  Caption         =   "Puas"
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Index           =   0
                  Left            =   0
                  TabIndex        =   1
                  Top             =   0
                  Width           =   735
               End
            End
            Begin MSMask.MaskEdBox Text 
               Height          =   270
               Index           =   5
               Left            =   960
               TabIndex        =   10
               Top             =   2400
               Width           =   4500
               _ExtentX        =   7938
               _ExtentY        =   476
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
               Height          =   270
               Index           =   6
               Left            =   960
               TabIndex        =   11
               Top             =   2700
               Width           =   4500
               _ExtentX        =   7938
               _ExtentY        =   476
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
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Ketepatan Penyerahan"
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
               TabIndex        =   49
               Top             =   480
               Width           =   1665
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Lamanya Servis"
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
               TabIndex        =   40
               Top             =   840
               Width           =   1125
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Respon"
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
               TabIndex        =   34
               Top             =   3160
               Width           =   540
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Komentar"
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
               TabIndex        =   33
               Top             =   2445
               Width           =   690
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Kemungkinan Kembali"
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
               TabIndex        =   32
               Top             =   1560
               Width           =   1530
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Penjelasan Bengkel"
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
               Left            =   420
               TabIndex        =   31
               Top             =   1200
               Width           =   1380
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackColor       =   &H00E0E0E0&
               BackStyle       =   0  'Transparent
               Caption         =   "Kualitas Pekerjaan"
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
               TabIndex        =   27
               Top             =   120
               Width           =   1320
            End
         End
         Begin MSChart20Lib.MSChart MSChart1 
            Height          =   2655
            Left            =   9960
            OleObjectBlob   =   "frm3DFU.frx":D704
            TabIndex        =   116
            Top             =   960
            Visible         =   0   'False
            Width           =   3975
         End
         Begin VB.CommandButton cmdKirimCSL 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Kirim CSL  (Belum Follow)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   8520
            Style           =   1  'Graphical
            TabIndex        =   115
            Top             =   1200
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            Caption         =   "KURANG"
            Height          =   615
            Left            =   12960
            TabIndex        =   112
            Top             =   120
            Visible         =   0   'False
            Width           =   975
            Begin VB.Label lblKurang 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "999"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   114
               Top             =   240
               Width           =   615
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Unit CPUS"
            Height          =   615
            Left            =   11880
            TabIndex        =   111
            Top             =   120
            Visible         =   0   'False
            Width           =   975
            Begin VB.Label lblUnitEntry 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "9999"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   75
               TabIndex        =   113
               Top             =   240
               Width           =   750
            End
         End
         Begin VB.CommandButton cmdRecoveryList 
            BackColor       =   &H0080FF80&
            Caption         =   "Recovery List"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   9720
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CommandButton cmdHistory 
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
            Height          =   390
            Left            =   8400
            Picture         =   "frm3DFU.frx":F57D
            Style           =   1  'Graphical
            TabIndex        =   100
            ToolTipText     =   "On Line HO"
            Top             =   4320
            Width           =   1980
         End
         Begin VB.CommandButton tblKeluar 
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
            Height          =   390
            Left            =   12000
            Picture         =   "frm3DFU.frx":FD3A
            Style           =   1  'Graphical
            TabIndex        =   99
            Top             =   4680
            Width           =   1980
         End
         Begin VB.CommandButton Command1 
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
            Height          =   390
            Left            =   8400
            Picture         =   "frm3DFU.frx":10409
            Style           =   1  'Graphical
            TabIndex        =   98
            ToolTipText     =   "On Line HO"
            Top             =   4680
            Visible         =   0   'False
            Width           =   1980
         End
         Begin MSFlexGridLib.MSFlexGrid gridKeluhan 
            Height          =   1575
            Left            =   40
            TabIndex        =   58
            Top             =   1120
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   2778
            _Version        =   393216
            ScrollBars      =   1
            Appearance      =   0
         End
         Begin VB.CommandButton cmdCall 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Call"
            Height          =   855
            Left            =   8520
            Picture         =   "frm3DFU.frx":10BC6
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton tblBooking 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Booking"
            Height          =   270
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   5280
            Width           =   1050
         End
         Begin MSFlexGridLib.MSFlexGrid JobGrid 
            Height          =   2415
            Left            =   40
            TabIndex        =   25
            Top             =   2720
            Width           =   8295
            _ExtentX        =   14631
            _ExtentY        =   4260
            _Version        =   393216
            Appearance      =   0
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   1
            Left            =   3645
            TabIndex        =   21
            Top             =   810
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   2
            Left            =   1125
            TabIndex        =   23
            Top             =   3315
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   3
            Left            =   885
            TabIndex        =   36
            Top             =   525
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            PromptInclude   =   0   'False
            AutoTab         =   -1  'True
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   ">CCCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   4
            Left            =   885
            TabIndex        =   38
            Top             =   810
            Width           =   2100
            _ExtentX        =   3704
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   7
            Left            =   885
            TabIndex        =   53
            Top             =   240
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   476
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
            Height          =   270
            Index           =   8
            Left            =   3645
            TabIndex        =   54
            Top             =   525
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   476
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
            Mask            =   ">CCCC-CCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   9
            Left            =   6120
            TabIndex        =   63
            Top             =   240
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
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
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   10
            Left            =   6120
            TabIndex        =   65
            Top             =   525
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
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
            Mask            =   ">CCCC-CCCCCCCCCC"
            PromptChar      =   ">"
         End
         Begin MSMask.MaskEdBox Text 
            Height          =   270
            Index           =   11
            Left            =   6120
            TabIndex        =   68
            Top             =   840
            Width           =   2220
            _ExtentX        =   3916
            _ExtentY        =   476
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
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
         Begin VB.Label Label59 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "On Line HO"
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
            Height          =   240
            Left            =   10440
            TabIndex        =   102
            Top             =   4440
            Width           =   1050
         End
         Begin VB.Label Label58 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "On LOCAL"
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
            Height          =   240
            Left            =   10440
            TabIndex        =   101
            Top             =   4800
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Ket"
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
            Left            =   5850
            TabIndex        =   69
            Top             =   885
            Width           =   240
         End
         Begin VB.Label Label26 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telp."
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
            Left            =   5730
            TabIndex        =   66
            Top             =   570
            Width           =   360
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   5400
            TabIndex        =   64
            Top             =   285
            Width           =   690
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            TabIndex        =   56
            Top             =   285
            Width           =   690
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Telp."
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
            Left            =   3240
            TabIndex        =   55
            Top             =   570
            Width           =   360
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Rangka"
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
            TabIndex        =   39
            Top             =   855
            Width           =   540
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   180
            TabIndex        =   37
            Top             =   570
            Width           =   645
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "No. Invoice"
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
            TabIndex        =   24
            Top             =   3360
            Width           =   825
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
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
            Left            =   3030
            TabIndex        =   22
            Top             =   855
            Width           =   570
         End
      End
      Begin MSComCtl2.DTPicker Tgl 
         Height          =   300
         Left            =   840
         TabIndex        =   15
         Top             =   165
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   529
         _Version        =   393216
         Format          =   102301697
         CurrentDate     =   39859
      End
      Begin MSMask.MaskEdBox Text 
         Height          =   300
         Index           =   0
         Left            =   3000
         TabIndex        =   18
         Top             =   165
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         BackColor       =   12640511
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
         Mask            =   ">&&&&&"
         PromptChar      =   ">"
      End
      Begin TrueOleDBGrid70.TDBGrid GridFU 
         Height          =   3735
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   6588
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   ""
         Columns(0).NumberFormat=   "Standard"
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
         Columns.Count   =   14
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).AlternatingRowStyle=   -1  'True
         Splits(0).DividerColor=   14933984
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=14"
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
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         Appearance      =   0
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=82,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=79,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=80,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=81,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=78,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=75,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=76,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=77,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=74,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=71,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=72,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=73,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
         _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=66,.parent=13"
         _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=63,.parent=14"
         _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=64,.parent=15"
         _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=65,.parent=17"
         _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13"
         _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
         _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
         _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
         _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=90,.parent=13"
         _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=87,.parent=14"
         _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=88,.parent=15"
         _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=89,.parent=17"
         _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=86,.parent=13"
         _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=83,.parent=14"
         _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=84,.parent=15"
         _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=85,.parent=17"
         _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=58,.parent=13"
         _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=55,.parent=14"
         _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=56,.parent=15"
         _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=57,.parent=17"
         _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13"
         _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
         _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
         _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
         _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=50,.parent=13"
         _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=47,.parent=14"
         _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=48,.parent=15"
         _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=49,.parent=17"
         _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=46,.parent=13"
         _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=43,.parent=14"
         _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=44,.parent=15"
         _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=45,.parent=17"
         _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=28,.parent=13"
         _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=25,.parent=14"
         _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=26,.parent=15"
         _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=27,.parent=17"
         _StyleDefs(88)  =   "Splits(0).Columns(13).Style:id=32,.parent=13"
         _StyleDefs(89)  =   "Splits(0).Columns(13).HeadingStyle:id=29,.parent=14"
         _StyleDefs(90)  =   "Splits(0).Columns(13).FooterStyle:id=30,.parent=15"
         _StyleDefs(91)  =   "Splits(0).Columns(13).EditorStyle:id=31,.parent=17"
         _StyleDefs(92)  =   "Named:id=33:Normal"
         _StyleDefs(93)  =   ":id=33,.parent=0"
         _StyleDefs(94)  =   "Named:id=34:Heading"
         _StyleDefs(95)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(96)  =   ":id=34,.wraptext=-1"
         _StyleDefs(97)  =   "Named:id=35:Footing"
         _StyleDefs(98)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(99)  =   "Named:id=36:Selected"
         _StyleDefs(100) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(101) =   "Named:id=37:Caption"
         _StyleDefs(102) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(103) =   "Named:id=38:HighlightRow"
         _StyleDefs(104) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(105) =   "Named:id=39:EvenRow"
         _StyleDefs(106) =   ":id=39,.parent=33,.bgcolor=&H80FFFF&"
         _StyleDefs(107) =   "Named:id=40:OddRow"
         _StyleDefs(108) =   ":id=40,.parent=33"
         _StyleDefs(109) =   "Named:id=41:RecordSelector"
         _StyleDefs(110) =   ":id=41,.parent=34"
         _StyleDefs(111) =   "Named:id=42:FilterBar"
         _StyleDefs(112) =   ":id=42,.parent=33"
         _StyleDefs(113) =   "Named:id=131:TipeTanpaTipe"
         _StyleDefs(114) =   ":id=131,.parent=33,.fgcolor=&HFF&"
         _StyleDefs(115) =   "Named:id=132:TigaDigitKosong"
         _StyleDefs(116) =   ":id=132,.parent=33,.fgcolor=&HFF00&"
         _StyleDefs(117) =   "Named:id=91:responCL"
         _StyleDefs(118) =   ":id=91,.parent=131,.fgcolor=&HFF00&,.bold=-1,.fontsize=825,.italic=0"
         _StyleDefs(119) =   ":id=91,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(120) =   ":id=91,.fontname=MS Sans Serif"
      End
      Begin VB.Label jmlBlmFollow 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Unit Belum Follow Up : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4560
         TabIndex        =   71
         Top             =   165
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "SA"
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
         Left            =   2655
         TabIndex        =   17
         Top             =   210
         Width           =   195
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         TabIndex        =   16
         Top             =   210
         Width           =   570
      End
   End
   Begin VB.Image Image1 
      Height          =   10095
      Left            =   0
      Picture         =   "frm3DFU.frx":10EE6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "frm3DFU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rsTemp      As New ADODB.Recordset
Dim rsTemp2     As New ADODB.Recordset
Dim rsTemp3     As New ADODB.Recordset

Dim rsGrid      As New ADODB.Recordset
Dim nCount      As Integer
Dim cols        As TrueOleDBGrid70.Columns
Dim tglAwal     As String
Dim tglAkhir    As String
Dim totalUE     As Double
Dim totalFollow As Double
Dim sSQL        As String


Private Sub hapusForm()
1        On Error GoTo hapusForm_Error

2         For i = 1 To 10
3             Text(i) = ""
4         Next
5         opKualitas(2).value = True
6         opJelas(2).value = True
7         opRamah(2).value = True
8         opLama(2).value = True
9         opTepat(2).value = True
10        Frame8.Enabled = False
11        Frame7.Enabled = False
12        framKeluhan.Visible = False
13        frameStatus.Visible = False
14       On Error GoTo 0
15       Exit Sub

hapusForm_Error:

16        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure hapusForm of Form frm3DFU")
End Sub

Private Sub aturJob(nRows As Integer)
1        On Error GoTo aturJob_Error
2         JobGrid.Clear
3         JobGrid.cols = 5
4         JobGrid.Rows = nRows
5         JobGrid.FormatString = " | Job | Job Order              | Rate | Mekanik "
6         JobGrid.ColWidth(0) = 0
7         JobGrid.ColWidth(1) = 800
8         JobGrid.ColWidth(2) = 4000
9         JobGrid.ColWidth(3) = 500
10        JobGrid.ColWidth(4) = 1000
11        JobGrid.row = 0

12       On Error GoTo 0
13       Exit Sub

aturJob_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure aturJob of Form frm3DFU")
End Sub

Private Sub aturKeluhan()
1        On Error GoTo aturKeluhan_Error

2         With gridKeluhan
3             .Clear
4             .cols = 1
5             .Rows = 6
6             .FormatString = " | Keluhan"
7             .ColWidth(0) = 0
8             .ColWidth(1) = 8000
9             .row = 0
10        End With

11       On Error GoTo 0
12       Exit Sub

aturKeluhan_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure aturKeluhan of Form frm3DFU")
End Sub

Private Sub headColumns()
      '[DATE], CUSTOMER, CUST_ADRES, TELPFU, JAMFU, KETFU, A.MDL + ' | ' + A.FMDL, A.COLOR, A.NO_CHASSIS, A.NO_WO, A.NO_POLISI, KM

1        On Error GoTo headColumns_Error

2         With GridFU
3             .Columns(0).Caption = "TANGGAL SERVICE"
4             .Columns(0).DataField = "DATE"
5             .Columns(0).Width = 1200
6             .Columns(1).Caption = "CUSTOMER"
7             .Columns(1).DataField = "USENAME"
8             .Columns(1).Width = 2500
9             .Columns(2).Caption = "ALAMAT"
10            .Columns(2).DataField = "USEADDR"
11            .Columns(2).Width = 5500
12            .Columns(3).Caption = "TELEPON"
13            .Columns(3).DataField = "TELPFU"
14            .Columns(3).Width = 1500
15            .Columns(4).Caption = "JAM"
16            .Columns(4).DataField = "JAMFU"
17            .Columns(4).Width = 800
18            .Columns(5).Caption = "KETERANGAN"
19            .Columns(5).DataField = "KETFU"
20            .Columns(5).Width = 1500
21            .Columns(6).Caption = "PEMBAWA"
22            .Columns(6).DataField = "ANCUST"
23            .Columns(6).Width = 2000
24            .Columns(7).Caption = "TELP. PEMBAWA"
25            .Columns(7).DataField = "TELPANCUST"
26            .Columns(7).Width = 1500
              
27            .Columns(8).Caption = "MODEL"
28            .Columns(8).DataField = "MODEL"
29            .Columns(8).Width = 5500
30            .Columns(9).Caption = "WARNA"
31            .Columns(9).DataField = "COLOR"
32            .Columns(9).Width = 2000
33            .Columns(10).Caption = "NO RANGKA"
34            .Columns(10).DataField = "NO_CHASSIS"
35            .Columns(10).Width = 2000
36            .Columns(11).Caption = "NO WO"
37            .Columns(11).DataField = "NO_WO"
38            .Columns(11).Width = 1500
39            .Columns(12).Caption = "NO POLISI"
40            .Columns(12).DataField = "NO_POLISI"
41            .Columns(12).Width = 1200
42            .Columns(13).Caption = "KM"
43            .Columns(13).DataField = "KM"
44            .Columns(13).Width = 1200
              
45            .Columns(14).Caption = "TELEPON"
46            .Columns(14).DataField = "TELPON"
47            .Columns(14).Width = 2500
              
48        End With

49       On Error GoTo 0
50       Exit Sub

headColumns_Error:

51        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumns of Form frm3DFU")
End Sub

Private Sub refreshGrid()
1        On Error GoTo refreshGrid_Error
         
          

2         GridFU.ClearFields
3         If rsGrid.State Then rsGrid.Close
      '    Set rsGrid = objO.FetchData("SELECT [DATE], (select USENAME from custdat WITH(NOLOCK) where no_Chassis = a.no_chassis and no_polisi = a.no_polisi) as usename, (select USEADDR from custdat WITH(NOLOCK) where no_Chassis = a.no_chassis and no_polisi = a.no_polisi) as useaddr, TELPFU, JAMFU, KETFU, ANCUST, TELPANCUST, A.MDL + ' | ' + A.FMDL AS MODEL, A.COLOR, A.NO_CHASSIS, A.NO_WO, A.NO_POLISI, KM, (select TELPON+' , '+GSM+' , '+CDMA+' , '+USEPSTN+' , '+USEGSM+' , '+USECDMA from custdat WITH(NOLOCK) where no_Chassis = a.no_chassis and no_polisi = a.no_polisi) AS TELPON FROM JOBDTA A WITH(NOLOCK) " _
      '                             & "INNER JOIN JOBDTL B WITH(NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO " _
      '                             & "WHERE TGLFU <= '" & Format(Tgl, "MM/DD/YYYY") & "' AND RPR_COND <> '' AND DELIV <> 'BTL' AND FRONT = '" & gSA & "' AND YEAR(TGLFU) > 1900 " _
      '                             & "AND A.NO_WO NOT IN (SELECT DISTINCT NO_WO FROM DTLFU WITH (NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS AND NO_WO = A.NO_WO AND (SAMBUNG = 'Y' OR COUNTER = 3) ) AND A.FS_LOG_DEL = '' ORDER BY TGLFU DESC, JAMFU ASC")      'AND NOINV = A.NOINV
4         Set rsGrid = objO.FetchData("SELECT [DATE], C.usename, C.useaddr, TELPFU, JAMFU, KETFU, ANCUST, TELPANCUST, A.MDL + ' | ' + A.FMDL AS MODEL, A.COLOR, A.NO_CHASSIS, A.NO_WO, A.NO_POLISI, KM, C.TELPON+' , '+C.GSM+' , '+C.CDMA+' , '+C.USEPSTN+' , '+C.USEGSM+' , '+C.USECDMA AS TELPON, ISNULL(D.RESPON,'') RESPON FROM JOBDTA A WITH (NOLOCK) " _
                                    & "INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS AND A.NO_WO = B.NO_WO " _
                                    & "LEFT JOIN CUSTDAT C WITH (NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS AND A.NO_POLISI = C.NO_POLISI " _
                                    & "LEFT JOIN DTLFU D WITH (NOLOCK) ON A.NO_CHASSIS = D.NO_CHASSIS AND A.NO_WO = D.NO_WO " _
                                    & "WHERE TGLFU <= '" & Format(Tgl, "MM/DD/YYYY") & "' AND RPR_COND <> '' AND DELIV <> 'BTL' AND FRONT = '" & gSA & "' AND YEAR(TGLFU) > 1900 " _
                                    & "AND A.NO_WO NOT IN (SELECT DISTINCT NO_WO FROM DTLFU WITH (NOLOCK) WHERE NO_CHASSIS = A.NO_CHASSIS AND NO_WO = A.NO_WO AND (SAMBUNG = 'Y' OR COUNTER = 3) ) AND A.FS_LOG_DEL = '' ORDER BY TGLFU DESC, JAMFU ASC")

5         If Not rsGrid.EOF Then
6             jmlBlmFollow.Caption = "Unit Belum Follow Up : " & rsGrid.RecordCount
7             GridFU.DataSource = rsGrid
8             GridFU.FetchRowStyle = True
9             Call headColumns
10        Else
11            tblKeluar.Enabled = True
12        End If
          
13        If Log.followUp = 1 Then
              
               '' 20170109
               '' khusus hari sabtu masih display periode sebelumnya
14             If Weekday(Format(isnow, "dd/mm/yyyy")) < 7 Then
15                Set rsTemp = objO.FetchData("SELECT minggu,tglawal,tglakhir FROM periode WITH(NOLOCK) WHERE '" & Format(isnow, "mm/dd/yyyy") & "' BETWEEN tglawal AND tglakhir")
16             Else
17                Set rsTemp = objO.FetchData("SELECT minggu,tglawal,tglakhir FROM periode WITH(NOLOCK) WHERE '" & Format(DateAdd("d", -1, Format(isnow, "dd/mm/yyyy")), "mm/dd/yyyy") & "' BETWEEN tglawal AND tglakhir")
18             End If
19             lblUnitEntry.Caption = "0" ''lblPersen.Caption = "0"
20             lblKurang.Caption = "0"
21             If Not rsTemp.EOF Then
22                tglAwal = Format(Trim(rsTemp!tglAwal), "mm/dd/yyyy")
23                tglAkhir = Format(Trim(rsTemp!tglAkhir), "mm/dd/yyyy")
                  ''2016-08-23 || 1. Perhitungan 85% tidak dihitung dari total unit entry, tapi unit CPUS yg difollow / Total unit entry CPUS ' CPUS = GR+SBE
                  ''Nilai kurang dihitung pula dari CPUS. Unit entry ' judul diganti "Unit CPUS", nilainya juga total unit CPUS.
                  ''totalUE = objO.FetchData_Double("SELECT COUNT(NO_WO) AS TOTAL FROM JOBDTA WITH (NOLOCK) WHERE ([DATE] BETWEEN '" & TglAwal & "' AND '" & TglAkhir & "') AND TPSVC <> 'W' AND PDS = 0 AND FRONT = '" & gSA & "' AND FS_LOG_DEL = '' ")
                  
24                totalUE = objO.FetchData_Double("SELECT COUNT(NO_WO) AS TOTAL FROM JOBDTA WITH (NOLOCK) WHERE ([DATE] BETWEEN '" & tglAwal & "' AND '" & tglAkhir & "') AND TPSVC <> 'W'  AND PDS = 0 AND fs_rc <> 'R001' AND FRONT = '" & gSA & "' AND FS_LOG_DEL = '' ")
                  
                  ''totalFollow = objO.FetchData_Double("select COUNT(no_wo) as TOTAL from DTLFU where respon = 'k' AND sa = '" & gSA & "' and tanggal between '" & TglAwal & "' and '" & TglAkhir & "'")
                  ''totalFollow = objO.FetchData_Double("select COUNT(no_wo) as TOTAL from DTLFU where (respon = 'k' OR respon = 'LS') AND sa = '" & gSA & "' and NO_WO IN (SELECT DISTINCT (NO_WO) AS TOTAL FROM JOBDTA WITH (NOLOCK) WHERE ([DATE] BETWEEN '" & TglAwal & "' AND '" & TglAkhir & "') AND TPSVC <> 'W' AND PDS = 0 AND FRONT = '" & gSA & "' AND FS_LOG_DEL = '' )")
                  
25                totalFollow = objO.FetchData_Double("select COUNT(no_wo) as TOTAL from DTLFU where (respon = 'k' OR respon = 'LS') AND sa = '" & gSA & "' and NO_WO IN (SELECT DISTINCT (NO_WO) AS TOTAL FROM JOBDTA WITH (NOLOCK) WHERE ([DATE] BETWEEN '" & tglAwal & "' AND '" & tglAkhir & "') AND MDL not like 'zoth%' AND TPSVC <> 'W' AND fs_rc <> 'R001' AND PDS = 0 AND FRONT = '" & gSA & "' AND FS_LOG_DEL = '' )")
                  
                  ''2018-05-02 mengambil nilai minimal LS untuk di bandingkan dengan total unit entry
26                If Log.RPLHistori = 1 Then
27                    Log.prosentaseFollow = objO.FetchData_Integer("SELECT MINLS FROM [" & dbReplikasiHO & "]..master_setting WITH (NOLOCK) WHERE [week] = '" & NulSID(rsTemp!minggu) & "'")
28                Else
29                    Log.prosentaseFollow = objO.FetchData_Integer("SELECT MINLS FROM master_setting WITH (NOLOCK) WHERE [week] = '" & NulSID(rsTemp!minggu) & "'")
30                End If
                  
31                If totalUE > 0 Then
32                   lblUnitEntry.Caption = totalUE ''lblPersen.Caption = Round(((totalFollow / totalUE) * 100), 0)
                     ''2016-08-18 || round diganti round down
33                   If Int((Log.prosentaseFollow / 100) * totalUE) >= totalFollow Then
34                      lblKurang.Caption = Int((Log.prosentaseFollow / 100) * totalUE) - totalFollow  ''lblKurang.Caption = Round((Log.prosentaseFollow / 100) * totalUE, 0) - totalFollow
35                   Else
36                      lblKurang.Caption = 0
37                   End If
38                Else
                     ''jika unit entry kosong
                       
39                   lblUnitEntry.Caption = 0 ''lblPersen.Caption = 0
40                   lblKurang.Caption = 0
41                End If
                  Dim rsChart As New ADODB.Recordset
42                If rsChart.State Then rsChart.Close
43                Set rsChart = Nothing
44                With rsChart
45                    .Fields.Append "JUMLAH", adInteger
46                    .Open
47                End With
                  '   If ckAktif.value = vbChecked Then
48                    With rsChart
49                        .AddNew
50                         If totalUE > 0 Then
51                            !JUMLAH = IIf(totalUE > 0, Round(((totalFollow / totalUE) * 100), 0), 0)
52                         Else
53                            !JUMLAH = 0
54                         End If
55                        .Update
56                    End With
                  ' End If
                  ' If ckPasif1.value = vbChecked Then
57                    With rsChart
58                        .AddNew
59                         If totalUE > 0 Then
60                            !JUMLAH = IIf(totalUE > 0, (100 - Round(((totalFollow / totalUE) * 100), 0)), 0)
61                         Else
62                            !JUMLAH = 0
63                         End If
            
64                        .Update
65                    End With
                   
66                rsChart.MoveFirst
                  
67                MSChart1.ShowWhatsThis
68                MSChart1.SeriesType = VtChSeriesType2dPie
69                With MSChart1
70                    For i = 1 To .Plot.SeriesCollection.Count
71                        With .Plot.SeriesCollection(i).DataPoints(-1).DataPointLabel
72                            .LocationType = VtChLabelLocationTypeInside
                              
73                            .Component = VtChLabelComponent.VtChLabelComponentValue
74                            .PercentFormat = "0%"
75                            .VtFont.Size = 10
                  '                .LocationType = VtChLabelLocationTypeInside
                  '                .Component = VtChLabelComponentValue '& " " & VtChLabelComponentPercent
                              
76                        End With
77                        If i = 1 Then
                              ''2016-09-22 || ganti label sent
78                            .Plot.SeriesCollection(i).LegendText = "Label Sent" '' "Sudah Follow Up"
79                        ElseIf i = 2 Then
80                            .Plot.SeriesCollection(i).LegendText = "Kekurangan Label Sent" ''"Belum Follow Up"
            '              ElseIf i = 3 Then
            '                  .Plot.SeriesCollection(i).LegendText = "Kategori Pasif2"
            '              ElseIf i = 4 Then
            '                  .Plot.SeriesCollection(i).LegendText = "Kategori Pasif3"
81                        End If
                          
82                    Next i
83                End With
84                Set MSChart1.DataSource = rsChart
85             End If
86        End If

         
87       On Error GoTo 0
88       Exit Sub

refreshGrid_Error:

89        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure refreshGrid of Form frm3DFU")
End Sub

Private Function BOOKRTN(noRgk As String) As Boolean
      Dim rsHlp As ADODB.Recordset
          
1        On Error GoTo BOOKRTN_Error

2         BOOKRTN = False
3         Set rsHlp = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NO_CHASSIS = '" & noRgk & "' AND DATE_BOOK = '" & Format(isnow, "MM/DD/YYYY") & "'")
4         If Not rsHlp.EOF Then
5             BOOKRTN = True
6         End If

7        On Error GoTo 0
8        Exit Function

BOOKRTN_Error:

9         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure BOOKRTN of Form frm3DFU")
End Function

Private Sub btnBatal_Click()
    framKeluhan.Visible = False
End Sub

Private Sub cetakRepoLS_Click()
    Call cetakLSRepo
    framLS.Visible = False
End Sub

Private Sub cetakLSRepo()
      Dim Report As New rptMapCust 'RptRepoLS
1        On Error GoTo cetakLSRepo_Error

2         If opTgl(1).value Then
3         sSQL = "SELECT DISTINCT DATE, B.FS_TH_MOBIL, CASE WHEN B.CUSTOMER = B.USENAME THEN '' ELSE B.CUSTOMER END AS PERUSAHAAN, B.USENAME AS PEMILIK, " _
               & "USEGSM+CASE WHEN USEGSM = '' THEN '' ELSE (CASE WHEN USECDMA = '' THEN '' ELSE ',' END) END+USECDMA AS SELULAR, B.USEPSTN AS TELPRUMAH, B.TELPON AS TELPKANTOR, REPLACE(A.NO_POLISI,'-','') AS NO_POLISI, " _
               & "CASE WHEN FS_RC = 'R001' THEN '1000' WHEN FS_RC <> '' THEN 'Servis Berkala' ELSE (CASE TPSVC WHEN 'W' THEN 'Perbaikan Umum (Warranty)' ELSE 'Perbaikan Umum (Non Warranty)' END) END AS TIPESVC, " _
               & "FRONT , a.NO_CHASSIS, C.MDL, C.FMDL, B.DOTAM, B.ddate, ANCUST, TELPANCUST, USEADDR, USECITY  " _
               & "FROM DTLFU A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN JOBDTA C WITH (NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS AND A.NO_WO = C.NO_WO " _
               & "WHERE C.DATE BETWEEN '" & Format(tanggal1, "MM/DD/YYYY") & "' AND '" & Format(tanggal2, "MM/DD/YYYY") & "' AND SA = '" & Trim(Text(0)) & "' AND RESPON = 'K'"
4         Else
5         sSQL = "SELECT DISTINCT DATE, B.FS_TH_MOBIL, CASE WHEN B.CUSTOMER = B.USENAME THEN '' ELSE B.CUSTOMER END AS PERUSAHAAN, B.USENAME AS PEMILIK, " _
               & "USEGSM+CASE WHEN USEGSM = '' THEN '' ELSE (CASE WHEN USECDMA = '' THEN '' ELSE ',' END) END+USECDMA AS SELULAR, B.USEPSTN AS TELPRUMAH, B.TELPON AS TELPKANTOR, REPLACE(A.NO_POLISI,'-','') AS NO_POLISI, " _
               & "CASE WHEN FS_RC = 'R001' THEN '1000' WHEN FS_RC <> '' THEN 'Servis Berkala' ELSE (CASE TPSVC WHEN 'W' THEN 'Perbaikan Umum (Warranty)' ELSE 'Perbaikan Umum (Non Warranty)' END) END AS TIPESVC, " _
               & "FRONT , a.NO_CHASSIS, C.MDL, C.FMDL, B.DOTAM, B.ddate, ANCUST, TELPANCUST, USEADDR, USECITY  " _
               & "FROM DTLFU A WITH (NOLOCK) INNER JOIN CUSTDAT B WITH (NOLOCK) ON A.NO_CHASSIS = B.NO_CHASSIS INNER JOIN JOBDTA C WITH (NOLOCK) ON A.NO_CHASSIS = C.NO_CHASSIS AND A.NO_WO = C.NO_WO " _
               & "WHERE TANGGAL BETWEEN '" & Format(tanggal1, "MM/DD/YYYY") & "' AND '" & Format(tanggal2, "MM/DD/YYYY") & "' AND SA = '" & Trim(Text(0)) & "' AND RESPON = 'K'"
6         End If
          
7         Set oRs = objO.FetchData(sSQL)
8         Report.database.SetDataSource oRs
9         Report.txtPerusahaan.SetText gPerush
10        Report.TxtAlamat.SetText gAlamat
11        Report.txtTanggal.SetText "TANGGAL : " & Format(tanggal1, "DD-MM-YYYY") & " s/d " & Format(tanggal2, "DD-MM-YYYY")
          
12        With CRViewer1
13            .ReportSource = Report
14            .ViewReport
15            .Zoom 100
              
16            .Left = 0
17            .Top = Frame1.Top
18            .Width = Me.Width
19            .Height = Frame1.Height
20            cmdTutup.Visible = True
21            cmdTutup.Left = .Left
22            cmdTutup.Top = .Top + 520
              
23            .Visible = True
24        End With
25        cmdTutup.Visible = True
26        Screen.MousePointer = vbDefault

27       On Error GoTo 0
28       Exit Sub

cetakLSRepo_Error:

29        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cetakLSRepo of Form frm3DFU")
End Sub

Private Sub cmdBatal_Click()
    If Not framKeluhan.Visible Then
        Call updateFU(Trim(Text(1)), Trim(Text(4)), Trim(Text(2)), 0)   'batal sambung
        frmeFU.Visible = False
        tblKeluar.Enabled = True
    End If
End Sub

Private Sub cmdBatalAlsn_Click()
    frmeTidakDiterima.Visible = False
End Sub

Private Sub cmdCall_Click()
     
    If Trim(Text(7)) = "" Then Exit Sub
    Label10(0).Caption = "Mengubungi " & Trim(Text(7)) & vbCr & "Di Nomor " & Trim(Text(8))
    nCount = 0
    Timer1.Enabled = True
    frmeDial.Visible = True
End Sub

Private Sub cmdCekFU_Click()
' ============================================================================================
' 01 Mar 2017 | Thania | tambahan menu utk cek WO Belum proses sudah waktu FU
' ============================================================================================
    
On Error GoTo error_cekFU
    
    Dim rsGridLockWO As New ADODB.Recordset
    If FrmCekLockFU.Visible = True Then
        FrmCekLockFU.Visible = False
        Exit Sub
    End If
    
    GridLockFU.ClearFields
    Set rsGridLockWO = objO.FetchData("select a.NO_WO, A.DATE, L.TGLfu from JOBDTL l with(nolock) " & _
                                      "inner join jobdta a with(nolock) on a.no_wo = l.no_wo and a.no_chassis = l.no_chassis and a.fs_log_del = '' " & _
                                      "left join DTLFU d with(nolock) on a.NO_WO = d.NO_WO and a.NO_CHASSIS = d.NO_CHASSIS " & _
                                      "where DATEDIFF(day,tglfu,getdate()) = 1 AND a.FRONT = '" & Trim(gSA) & "' " & _
                                      "and ISNULL(d.no_wo,'') = ''")
    If Not rsGridLockWO.EOF Then
        GridLockFU.DataSource = rsGridLockWO
        Call headColumns2
        FrmCekLockFU.Top = 840
        FrmCekLockFU.Left = 5040
        FrmCekLockFU.Visible = True
    Else
        MsgBox "Tidak ada data WO belum Follow Up", vbOKOnly, "Data Wo Belum Follow Up"
    End If
    Exit Sub
    
error_cekFU:
    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdCekFU_Click of Form frm3DFU")
End Sub

Private Sub headColumns2()

1        On Error GoTo headColumns2_Error

2         With GridLockFU
3             .Columns(0).Caption = "No WO"
4             .Columns(0).DataField = "no_wo"
5             .Columns(0).Width = 1200
6             .Columns(1).Caption = "Tanggal WO"
7             .Columns(1).DataField = "date"
8             .Columns(1).Width = 1500
9             .Columns(2).Caption = "Tanggal FU"
10            .Columns(2).DataField = "tglFU"
11            .Columns(2).Width = 1500
12        End With

13       On Error GoTo 0
14       Exit Sub

headColumns2_Error:

15        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure headColumns2 of Form frm3DFU")
End Sub

Private Sub cmdHistory_Click()
'    If Trim(Text(4)) = "" Then Exit Sub
'    gTag = "3dfu"
'    frmHistory.Text(0) = Text(4)
'    frmHistory.Text(1) = cariNamaCustomer(Text(4))
'    frmHistory.Text(2) = Replace(Replace(Text(3), ">", ""), " ", "")
'    frmHistory.Show (vbModal)
    If Text(4) = "" Then Exit Sub
    If Log.LogHOConn = 0 Then
        '' 2015-09-08 | Konek HO ''
        If MsgBox("Gagal Koneksi ke Server HO !" & vbCrLf & "Apakah ingin menyambungkan koneksi kembali?", vbYesNo) = vbYes Then
            If Get_ConnectionProfil_Bengkel99 Then
                Log.LogHOConn = 1
            Else
                Exit Sub
            End If
        Else
            Exit Sub
        End If
'        MsgBox "Gagal Koneksi ke Server HO !", vbInformation + vbOKOnly, "History On Line HO"
'        Exit Sub
    End If
    gTag = "ONLINE"
    dRangka = Text(4)
    nData(0) = Text(4)
    nData(1) = cariNamaCustomer(Text(4))
    nData(2) = Replace(Replace(Text(3), ">", ""), " ", "")
    Dim fH As New FrmHistory_new
    fH.show (vbModal)
End Sub

Private Sub cmdKirimCSL_Click()
1        If Trim(Text(7)) = "" Then Exit Sub
2        If MsgBox("Apakah ingin mengirim data CSL?", vbYesNo) = vbYes Then
3           Set rsTemp = objO.FetchData("SELECT * FROM dtlfu WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(Text(4)) & "' AND NO_WO = '" & Trim(Text(1)) & "'")
4           If rsTemp.EOF Then
5              objO.InsertData ("INSERT INTO DTLFU (NO_CHASSIS, NO_WO, NOINV,RESPON,COUNTER,SA) VALUES ('" & Trim(Text(4)) & "', '" & Trim(Text(1)) & "', '" & Trim(Text(2)) & "','LS',0,'" & gSA & "')")
6              MsgBox "Data berhasil disimpan", vbInformation
7              Call refreshGrid
8           End If
9        End If
End Sub

Private Sub cmdRecoveryList_Click()
   ''2016-07-12 || cek hari minggu
   If WeekdayName(IIf(Weekday(Format(isnow, "dd/mm/yyyy")) = 1, 8, Weekday(Format(isnow, "dd/mm/yyyy"))) - 1) = "Minggu" Then
      MsgBox "Hari minggu, tidak dapat follow up ulang", vbInformation
      Exit Sub
   Else
   frm3DFURecovery.show (vbModal)
   End If
   
   
End Sub

Private Sub cmdStatusOk_Click()
1        If optStatus(0).value = False And optStatus(1).value = False And optStatus(2).value = False Then
2           MsgBox "Silahkan pilih status customer terlebih dahulu", vbInformation
3           Exit Sub
4        Else
            Dim Status As String
5           If optStatus(0).value = True Then
6              Status = "0"
7           ElseIf optStatus(1).value = True Then
8              Status = "1"
9           ElseIf optStatus(2).value = True Then
10             Status = "2"
11          End If
12          Call objO.InsertData("UPDATE dtlfu SET stCust = '" & Status & "' WHERE no_wo = '" & Trim(Text(1)) & "'")
13          If opKualitas(2).value Or opJelas(2).value Or opLama(2).value Or opRamah(2).value Or opTepat(2).value Then
14              If MsgBox("Akan Melakukan booking untuk Return Job ?", vbYesNo, "BOOKING RTJ") = vbYes Then
15                  tblBooking_Click
16              End If
17          End If
18          Call hapusForm
19          Call refreshGrid
20          frmeFU.Visible = False
21          tblKeluar.Enabled = True
22       End If
End Sub

Private Sub Command1_Click()
    If Text(4) = "" Then Exit Sub
    gTag = "LOCAL"
    dRangka = Text(4)
    nData(0) = Text(4)
    nData(1) = cariNamaCustomer(Text(4))
    nData(2) = Replace(Replace(Text(3), ">", ""), " ", "")
    If Log.dmsNonDss = "1" Then
        Dim fHDND As New FrmHistory_newDMSNonDSS
        fHDND.isOrder = False
        fHDND.noBaris = 1 'baris
        fHDND.modelKendaraan = "" 'Trim(Text(10))
        fHDND.show (vbModal)
    Else
        Dim fH As New FrmHistory_new
        fH.show (vbModal)
    End If
'
'    Dim fH As New FrmHistory_new
'    fH.Show (vbModal)
End Sub

Private Function cariNamaCustomer(noRangka As String) As String
1        On Error GoTo cariNamaCustomer_Error

2         cariNamaCustomer = objO.FetchData_String("SELECT CUSTOMER FROM CUSTDAT WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(noRangka) & "'")

3        On Error GoTo 0
4        Exit Function

cariNamaCustomer_Error:

5         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cariNamaCustomer of Form frm3DFU")
End Function

Private Sub cmdLS_Click()
1        On Error GoTo cmdLS_Click_Error

2         If framLS.Visible = True Then
3             framLS.Visible = False
4         Else
5             framLS.Visible = True
6             tanggal1 = Format(isnow, "dd/MM/yyyy")
7             tanggal2 = GetTglAkhir(Month(isnow), year(isnow), 0)
8         End If

9        On Error GoTo 0
10       Exit Sub

cmdLS_Click_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdLS_Click of Form frm3DFU")
End Sub

Private Sub cmdSaveKeluhan_Click()
1        On Error GoTo cmdSaveKeluhan_Click_Error

2         If (Trim(Text1(0)) <> "" And Len(Trim(Text1(0))) > 10) Or (Trim(Text1(1)) <> "" And Len(Trim(Text1(1))) > 10) Or (Trim(Text1(2)) <> "" And Len(Trim(Text1(2))) > 10) Then
3             Call objO.UpdateData("UPDATE DTLFU SET KELUHAN1 = '" & Trim(Text1(0)) & "', KELUHAN2 = '" & Trim(Text1(1)) & "', KELUHAN3 = '" & Trim(Text1(2)) & "' WHERE NO_CHASSIS = '" & Trim(Text(4)) & "' AND NO_WO = '" & Trim(Text(1)) & "'")
4             framKeluhan.Visible = False
5             Call Simpan_OK
6         ElseIf Trim(Text1(0)) = "" And Trim(Text1(1)) = "" And Trim(Text1(2)) = "" Then
7             MsgBox "Isi alasan customer memberikan jawaban tidak puas !", vbCritical, "Keluhan"
8         ElseIf Len(Trim(Text1(0))) <= 10 Or Len(Trim(Text1(1))) <= 10 Or Len(Trim(Text1(2))) <= 10 Then
9             MsgBox "Jawaban terlalu pendek, minimal 10 karakter !", vbCritical, "Keluhan"
10        End If

11       On Error GoTo 0
12       Exit Sub

cmdSaveKeluhan_Click_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdSaveKeluhan_Click of Form frm3DFU")
End Sub

Private Sub cmdTerima_Click()

1        On Error GoTo cmdTerima_Click_Error

2         Timer1.Enabled = False
3         frmeDial.Visible = False
4         frmeFU.Visible = True
5         frmeFU.Top = 120
6         frmeFU.Left = 8520
7         Call updateFU(Trim(Text(1)), Trim(Text(4)), Trim(Text(2)), 1)  'sambung
8         tblKeluar.Enabled = False

9        On Error GoTo 0
10       Exit Sub

cmdTerima_Click_Error:

11        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdTerima_Click of Form frm3DFU")
End Sub

Private Sub cmdTolak_Click()
1        On Error GoTo cmdTolak_Click_Error

2         Timer1.Enabled = False
3         frmeDial.Visible = False
          
4         frmeTidakDiterima.Visible = True

5        On Error GoTo 0
6        Exit Sub

cmdTolak_Click_Error:

7         MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cmdTolak_Click of Form frm3DFU")
End Sub

Private Sub updateFU(nWO As String, nChassis As String, nINV As String, kdTipe As Integer, Optional reason As String)
Dim rsHlp As ADODB.Recordset
Dim cCall As Integer
    
   On Error GoTo updateFU_Error

    Call saveLogMenu(gSA, ComptName, gVersi, "frm3DFU", "updateFU", 0)

    cCall = objO.FetchData_Double("SELECT COUNTER FROM DTLFU WITH (NOLOCK) WHERE NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "'") ' AND SAMBUNG <> 'Y'")
    
    sSQL = ""
    sSQL1 = ""
    Select Case kdTipe
        Case 0
            Select Case cCall
                Case 0
                Case 1
                    sSQL1 = "DELETE DTLFU WHERE NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "'"
                Case 2
                    sSQL1 = "UPDATE DTLFU SET CALL2 = '', SAMBUNG = 'N', COUNTER = COUNTER -1 WHERE NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "'"
                Case 3
                    sSQL1 = "UPDATE DTLFU SET CALL3 = '', SAMBUNG = 'N', COUNTER = COUNTER -1 WHERE NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "'"
            End Select
        Case 1  ' sambung
            Select Case cCall
                Case 0
                    ''hapus data respon csl
                    sSQL1 = "DELETE DTLFU WHERE respon = 'LS' AND NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "'"
                    sSQL = "INSERT INTO DTLFU (NO_CHASSIS, NO_WO, NOINV, CALL1, SAMBUNG, COUNTER, SA) VALUES ('" & nChassis & "', '" & nWO & "', '" & nINV & "','" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', 'Y','1','" & Trim(Text(0)) & "')"
                Case 1
                    sSQL1 = "UPDATE DTLFU SET CALL2 = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', SAMBUNG = 'Y', COUNTER = '2' WHERE NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "' AND SAMBUNG <> 'Y'"
                Case 2
                    sSQL1 = "UPDATE DTLFU SET CALL3 = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', SAMBUNG = 'Y', COUNTER = '3' WHERE NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "' AND SAMBUNG <> 'Y'"
            End Select
        Case 2  ' tidak sambung
            Select Case cCall
                Case 0
                    sSQL1 = "DELETE DTLFU WHERE respon = 'LS' AND NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "'"
                    sSQL = "INSERT INTO DTLFU (NO_CHASSIS, NO_WO, NOINV, CALL1, REASON1, SAMBUNG, COUNTER, SA) VALUES ('" & nChassis & "', '" & nWO & "', '" & nINV & "', '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "','" & Trim(reason) & "', 'N', '1','" & Trim(Text(0)) & "')"
                Case 1
                    sSQL1 = "UPDATE DTLFU SET CALL2 = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', REASON2 = '" & Trim(reason) & "', SAMBUNG = 'N', COUNTER = '2' WHERE NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "' AND SAMBUNG <> 'Y'"
                Case 2
                    sSQL1 = "UPDATE DTLFU SET CALL3 = '" & Format(isnow, "MM/DD/YYYY HH:NN:SS") & "', REASON3 = '" & Trim(reason) & "', SAMBUNG = 'N', COUNTER = '3' WHERE NO_CHASSIS = '" & nChassis & "' AND NO_WO = '" & nWO & "' AND SAMBUNG <> 'Y'"
            End Select
    End Select

    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM DTLFU WITH (NOLOCK) WHERE NO_CHASSIS = '" & Trim(nChassis) & "' AND NO_WO = '" & Trim(nWO) & "'")

    Call objHT.OpenTransactionBengkel
    erTranAll = False
    onTR = True
    
    If rsHlp.EOF Then
        If Log.followUp = 1 And sSQL1 <> "" Then
            Call objHT.UpdateData_BengkelOnly(sSQL1)
        End If
        Call objHT.UpdateData_BengkelOnly(sSQL)
    Else
        Call objHT.UpdateData_BengkelOnly(sSQL1)
    End If

        If erTranAll = False Then
            Call objHT.CloseTransactionBengkel
            onTR = False
            Call saveLogMenu(gSA, ComptName, gVersi, "frm3DFU", "updateFU", 1)
        Else
            Call objHT.CloseTransactionBengkel_err
            onTR = False
            Call saveLogMenu(gSA, ComptName, gVersi, "frm3DFU", "updateFU", 2)
            MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
        End If
    Exit Sub
    
vbError:
    Call objHT.CloseTransactionBengkel_err
    onTR = False
    Call saveLogMenu(gSA, ComptName, gVersi, "frm3DFU", "updateFU", 2)
    MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
    Exit Sub

   On Error GoTo 0
   Exit Sub

updateFU_Error:

    MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure updateFU of Form frm3DFU")
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
End Sub

Private Sub Form_Load()
1        On Error GoTo Form_Load_Error
2           If Log.followUp = 1 Then
3              cmdRecoveryList.Visible = True
4              MSChart1.Visible = True
5              cmdKirimCSL.Visible = True
6              Frame2.Visible = True
7              Frame9.Visible = True
               'frameStatus.Visible = True
8              Label12.Visible = False
9              opRespon(0).Visible = False
10             opRespon(1).Visible = False
               
               ''2016-07-13 || cek periode untuk minggu ini sudah ada / belum
11             Set rsTemp = objO.FetchData("select ID from periode with (nolock) where GETDATE() between tglAwal and tglAkhir + '23:59:59'")
12             If rsTemp.EOF Then
                  ''jika belum ada data pada periode
                  Dim minggu As Integer
                  Dim hariIni As Integer
                  Dim tanggalAwal As String '' date
                  Dim tanggalAkhir As String '' date
                  
                  ''mengambil data hari ini minggu keberapa
13                minggu = objO.FetchData_Integer("select datediff(week, dateadd(week, datediff(week, 0, dateadd(month, datediff(month, 0, GETDATE()), 0)), 0), getdate() - 1) + 1")
                  
                  ''mencari tanggal awal = sabtu dan tanggal akhir = jumat
14                hariIni = Weekday(Format(isnow, "dd/mm/yyyy"))
15                If hariIni < 7 Then
                     ''jika hari minggu - jumat
                     ''mencari hari sabtu minggu sebelumnya
16                   tanggalAwal = Format((DateAdd("d", Abs(hariIni) * -1, Format(isnow, "dd/mm/YYYY"))), "mm/dd/YYYY")
                  
                     ''mencari hari jumat
17                   If hariIni = 6 Then
18                      tanggalAkhir = Format(isnow, "mm/dd/yyyy")
19                   Else
20                      tanggalAkhir = Format((DateAdd("d", 6 - hariIni, Format(isnow, "dd/mm/yyyy"))), "mm/dd/yyyy")
21                   End If
22                Else
                      ''jika hari sabtu
23                    tanggalAwal = Format(isnow, "mm/dd/yyyy")
24                    tanggalAkhir = Format((DateAdd("d", 6, Format(isnow, "dd/mm/yyyy"))), "mm/dd/yyyy")
                      ''jika awal hari sabtu minggu ditambah 1
25                    minggu = minggu + 1
26                End If
                  
                  ''insert periode
27                Call objO.InsertData("INSERT INTO PERIODE([tahun],[bulan],[minggu],[tglAwal],[tglAkhir]) VALUES (" & year(isnow) & "," & Month(isnow) & "," & minggu & ",'" & tanggalAwal & "','" & tanggalAkhir & "')")
                  
28             End If
29          End If
30        Call hapusForm
31        Tgl = Format(isnow, "dd-mm-yyyy")
32        Text(0) = gSA
33        Call refreshGrid
          'Call transferDataPenjualanNew
          ''-------------------------------------------------------------
          '' 2017-09-19 | cek data replikasi
          ''-------------------------------------------------------------
      '    sSQL = "SELECT rplSalesUnit FROM profilbengkel..profil"
      '    Set rsTemp = objO.FetchData(sSQL)
      '    If Not rsTemp.EOF Then
      '        If rsTemp!rplSalesUnit = 0 Then
      '            Call transferDataPenjualanNew
      '        End If
      '    End If
34        If Log.RPLHistori = 0 Then
35            Call transferDataPenjualanNew
36        End If
37       On Error GoTo 0
38       Exit Sub

Form_Load_Error:

39        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Form_Load of Form frm3DFU")
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        Image1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub GridFU_FetchRowStyle(ByVal Split As Integer, Bookmark As Variant, ByVal RowStyle As TrueOleDBGrid70.StyleDisp)
   Dim rs As New ADODB.Recordset
    Set rs = rsGrid.Clone
    If Bookmark > 1 Then
        rs.Bookmark = Bookmark
    End If
    'RowStyle = "putih"
    If rs("RESPON") <> "" Then
        If Trim(UCase(rs("RESPON"))) = "LS" Then
            RowStyle = "responCL"
 
        End If
    End If
End Sub

Private Sub GridFU_FilterChange()
1        On Error GoTo GridFU_FilterChange_Error

2         On Error GoTo errHandler
3         Set cols = GridFU.Columns
          Dim C As Integer
          
4         C = GridFU.col
5         GridFU.HoldFields
6         rsGrid.filter = getFilter()
7         GridFU.col = C
8         GridFU.EditActive = True
9         Exit Sub
          
errHandler:
10        MsgBox err.Source & ":" & vbCrLf & err.Description

11       On Error GoTo 0
12       Exit Sub

GridFU_FilterChange_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GridFU_FilterChange of Form frm3DFU")
End Sub

Private Function getFilter() As String
      Dim tmp As String
      Dim n As Integer
1        On Error GoTo getFilter_Error

2         For Each col In cols
3             If Trim(col.FilterText) <> "" Then
4                 n = n + 1
5                 If n > 1 Then
6                     tmp = tmp & " AND "
7                 End If
8                 tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
              
9             End If
10        Next col
11        getFilter = tmp

12       On Error GoTo 0
13       Exit Function

getFilter_Error:

14        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure getFilter of Form frm3DFU")
End Function

Private Sub GridFU_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
      Dim hlp, hlp2, hlp3, hlp4 As String
      Dim rsHlp As New ADODB.Recordset
      Dim nADD As Integer
      Dim nCounter As Integer

1        On Error GoTo GridFU_RowColChange_Error

2         If rsGrid.State <> adStateClosed Then
3         If rsTemp.State Then rsTemp.Close
4         Set rsTemp = objO.FetchData("SELECT * FROM JOBDTA A WITH (NOLOCK) INNER JOIN JOBDTL B WITH (NOLOCK) ON A.NO_WO = B.NO_WO AND A.NO_CHASSIS = B.NO_CHASSIS WHERE A.NO_CHASSIS = '" & GridFU.Columns(10) & "' AND A.NO_WO = '" & GridFU.Columns(11) & "'")
5         If Not rsTemp.EOF Then
6             Text(1) = rsTemp("NO_WO")
7             Text(2) = rsTemp("NOINV")
8             Text(3) = rsTemp("NO_POLISI")
9             Text(4) = rsTemp("NO_CHASSIS")
10            Text(7) = rsTemp("CUSTOMER")
11            Text(8) = rsTemp("TELPFU")
12            Text(9) = rsTemp("ANCUST")
13            Text(10) = rsTemp("TELPANCUST")
14            Text(11) = rsTemp("KETFU")
              
15            nADD = objO.FetchData_Count("SELECT NO_WO FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsTemp("NO_WO")) & "'")
              
16            nADD = nADD + objO.FetchData_Count("SELECT DISTINCT A.NAMABARANG FROM HIST_SUBLET A  WITH (NOLOCK) inner Join T_invoice B WITH (NOLOCK) On A.NoInv = B.NoInv WHERE B.Fs_log_Del = '' and A.NO_WO = '" & Trim(rsTemp("NO_WO")) & "'")
              
17            Call aturJob(CDbl(9 + nADD))
18            nCounter = 1
19            For i = 1 To 8
20                hlp = "KLPKERJA" & i
21                hlp2 = "REP" & i
22                hlp3 = "RTE" & i
23                hlp4 = "MECH"
                  
24                If rsTemp(hlp2) <> "" Then
25                    With JobGrid
26                        .TextMatrix(nCounter, 1) = Trim(rsTemp(hlp))
27                        .TextMatrix(nCounter, 2) = Trim(rsTemp(hlp2))
28                        .TextMatrix(nCounter, 3) = Trim(rsTemp(hlp3))
29                        .TextMatrix(nCounter, 4) = Trim(rsTemp(hlp4))
30                    End With
31                    nCounter = nCounter + 1
32                End If
33            Next
              
              ' TAMPIL JOBADD
               
34            Set rsHlp = objO.FetchData("SELECT * FROM JOBADD2 WITH (NOLOCK) WHERE NO_WO = '" & Trim(rsTemp("NO_WO")) & "'")
35            If Not rsHlp.EOF Then
36                Do While Not rsHlp.EOF
37                    With JobGrid
38                        .TextMatrix(nCounter, 1) = Trim(rsHlp("KLPKERJA"))
39                        .TextMatrix(nCounter, 2) = Trim(rsHlp("REP"))
40                        .TextMatrix(nCounter, 3) = Trim(rsHlp("RTE"))
41                        .TextMatrix(nCounter, 4) = Trim(rsTemp("MECH")) 'Trim(rsHlp("MEKANIK"))
42                    End With
43                    nCounter = nCounter + 1
44                    rsHlp.MoveNext
45                Loop
46            End If
              
              ' TAMPIL SUBLET
               
47            Set rsHlp = objO.FetchData("SELECT A.* FROM HIST_SUBLET A WITH (NOLOCK) inner Join T_invoice B WITH (NOLOCK) On A.NoInv = B.NoInv WHERE B.Fs_log_Del = '' and A.NO_WO = '" & Trim(rsTemp("NO_WO")) & "'")
48            If Not rsHlp.EOF Then
49                Do While Not rsHlp.EOF
50                    With JobGrid
51                        .TextMatrix(nCounter, 1) = "SUB"
52                        .TextMatrix(nCounter, 2) = Trim(rsHlp("NAMABARANG"))
53                        .TextMatrix(nCounter, 3) = Trim(rsHlp("QTY")) + " " + Trim(rsHlp("KODESATUAN"))
54                        .TextMatrix(nCounter, 4) = ""
55                    End With
56                    nCounter = nCounter + 1
57                    rsHlp.MoveNext
58                Loop
59            End If
              
60            FU.Enabled = True
              
61            Call aturKeluhan
62            For i = 1 To 5
63                hlp = "CMPL" & i
                  
64                If rsTemp(hlp) <> "" Then
65                    With gridKeluhan
66                        .TextMatrix(i, 1) = " " & Trim(rsTemp(hlp))
67                    End With
68                End If
69            Next
70        End If
71        End If

72       On Error GoTo 0
73       Exit Sub

GridFU_RowColChange_Error:

74        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure GridFU_RowColChange of Form frm3DFU")
End Sub

Private Sub opKualitas_Click(Index As Integer)
    Select Case Index
        Case 0
            Frame8.Enabled = True
            If opTepat(2).value Then
                Frame7.Enabled = False
            Else
                Frame7.Enabled = True
            End If
        Case 2
            Frame8.Enabled = False
            Frame7.Enabled = False
            opTepat(2).value = True
            opLama(2).value = True
    End Select
End Sub

Private Sub opRespon_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub opTepat_Click(Index As Integer)
    Select Case Index
        Case 0
            Frame7.Enabled = True
        Case 2
            Frame7.Enabled = False
            opLama(2).value = True
    End Select
End Sub

Private Sub simpanAlasan_Click()
      Dim boleh As Boolean
      Dim pesan As String
      Dim txtKetTelp As String

1        On Error GoTo simpanAlasan_Click_Error

2         boleh = True
3         If OptAlasan(3).value Then
4             If Trim(txtAlasanLain) = "" Then
5                 boleh = False
6                 pesan = "Isi alasannya !!!"
7             End If
8         End If
          
9         If boleh Then
10            If OptAlasan(0).value Then
11                txtKetTelp = OptAlasan(0).Caption & " : " & Trim(txtAlasanLain)
12            Else
13                If OptAlasan(1).value Then
14                    txtKetTelp = OptAlasan(1).Caption & " : " & Trim(txtAlasanLain)
15                Else
16                    If OptAlasan(2).value Then
17                        txtKetTelp = OptAlasan(2).Caption & " : " & Trim(txtAlasanLain)
18                    Else
19                        txtKetTelp = "4. " & Trim(txtAlasanLain)
20                    End If
21                End If
22            End If
23            Call updateFU(Trim(Text(1)), Trim(Text(4)), Trim(Text(2)), 2, Left(txtKetTelp, 50)) 'tidak sambung
24            Call refreshGrid
25            frmeTidakDiterima.Visible = False
26        Else
27            MsgBox pesan
28        End If

29       On Error GoTo 0
30       Exit Sub

simpanAlasan_Click_Error:

31        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure simpanAlasan_Click of Form frm3DFU")
End Sub

Private Sub tblBooking_Click()
1        On Error GoTo tblBooking_Click_Error

2         With frmBooking
              ' RETURN JOB
3             .Tag = "FU"
4             .cmdNo.Caption = "No. &Polisi"
5             .Text(0).Mask = ""
              
              Dim nDpn1 As String
              Dim nTgh1 As String
              Dim nBlk1 As String
6             nDpn1 = ""
7             nTgh1 = ""
8             nBlk1 = ""
9             Call pecahNoPol(Trim(Text(3)), nDpn1, nTgh1, nBlk1)
10            .Text(0) = Trim(nDpn1)
11            .Text(56) = Trim(nTgh1)
12            .Text(57) = Trim(nBlk1)
13            .chkRtn = Checked
14            .Text(41) = Trim(Text(4))
15            TOChild frm3DFU.hWnd, frmBooking.hWnd
16            frmBooking.show
17        End With

18       On Error GoTo 0
19       Exit Sub

tblBooking_Click_Error:

20        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure tblBooking_Click of Form frm3DFU")
End Sub

Private Sub TblKeluar_Click()
    Unload Me
End Sub

Private Function cekInputKeluhan() As Boolean
          
1        On Error GoTo cekInputKeluhan_Error

2         cekInputKeluhan = False
3         If opKualitas(2) Or opTepat(2) Or opLama(2) Then
4             For i = 0 To 2
5                 Text1(i) = ""
6             Next
7             framKeluhan.Visible = True
8             cekInputKeluhan = True
9         End If

10       On Error GoTo 0
11       Exit Function

cekInputKeluhan_Error:

12        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure cekInputKeluhan of Form frm3DFU")
End Function

Private Sub Simpan_OK()
1        On Error GoTo Simpan_OK_Error
          Dim statusCustomer As Boolean
2         statusCustomer = False
3         If Log.followUp = 1 Then
4           If opKualitas(0).value And opTepat(0).value And opLama(0).value And opJelas(0).value And opRamah(0).value Then
5             opRespon(0).value = True
6             statusCustomer = False
7           Else
8             opRespon(1).value = True
9             statusCustomer = True
10          End If
11        End If
          
12        ClearsSql
13        gsSql1(0) = "NO_CHASSIS"
14        gsSql2(0) = Text(4)
15        gsSql1(1) = "NO_WO"
16        gsSql2(1) = Text(1)
17        gsSql1(2) = "NOINV"
18        gsSql2(2) = Text(2)
19        gsSql1(3) = "NO_POLISI"
20        gsSql2(3) = Text(3)
21        gsSql1(4) = "KUALITAS"
22        gsSql2(4) = IIf(opKualitas(0).value, "P", IIf(opKualitas(1).value, "C", "K"))
23        gsSql1(5) = "PENJELASAN"
24        gsSql2(5) = IIf(opJelas(0).value, "P", IIf(opJelas(1).value, "C", "K"))
25        gsSql1(6) = "RAMAH"
26        gsSql2(6) = IIf(opRamah(0).value, "P", IIf(opRamah(1).value, "C", "K"))
27        gsSql1(7) = "LAMA"
28        gsSql2(7) = IIf(opLama(0).value, "P", IIf(opLama(1).value, "C", "K"))
29        gsSql1(8) = "TEPAT"
30        gsSql2(8) = IIf(opTepat(0).value, "P", IIf(opTepat(1).value, "C", "K"))
31        gsSql1(9) = "KOMENTAR1"
32        gsSql2(9) = Text(5)
33        gsSql1(10) = "KOMENTAR2"
34        gsSql2(10) = Text(6)
35        gsSql1(11) = "TANGGAL"
36        gsSql2(11) = Format(isnow, "MM/DD/YYYY")
37        gsSql1(12) = "BOOK"
38        gsSql2(12) = IIf(BOOKRTN(Text(4)), 1, 0)
39        gsSql1(13) = "SA"
40        gsSql2(13) = Text(0)
41        gsSql1(14) = "RESPON"
42        gsSql2(14) = IIf(opRespon(0).value, "K", "TK")
          
43        If rsTemp.State Then rsTemp.Close
44        Set rsTemp = objO.FetchData("SELECT * FROM DTLFU WITH (NOLOCK) WHERE NO_CHASSIS = '" & Text(4) & "' AND NO_WO = '" & Text(1) & "'")
45        If Not rsTemp.EOF Then
46            sSQL = GensSql2("DTLFU", True) + "WHERE NO_CHASSIS = '" & Text(4) & "' AND NO_WO = '" & Text(1) & "'"
47        Else
48            sSQL = GensSql2("DTLFU", False)
49        End If
50        Call objO.UpdateData(sSQL)
          
51        If statusCustomer = True Then
            ''munculkan status customer
52          optStatus(0).value = False
53          optStatus(1).value = False
54          optStatus(2).value = False
55          frameStatus.Visible = True
            
56        Else
            
57          If opKualitas(2).value Or opJelas(2).value Or opLama(2).value Or opRamah(2).value Or opTepat(2).value Then
58              If MsgBox("Akan Melakukan booking untuk Return Job ?", vbYesNo, "BOOKING RTJ") = vbYes Then
59                  tblBooking_Click
60              End If
61          End If
62          Call hapusForm
63          Call refreshGrid
64          frmeFU.Visible = False
65          tblKeluar.Enabled = True

66        End If
          
          
67       On Error GoTo 0
68       Exit Sub

Simpan_OK_Error:

69        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure Simpan_OK of Form frm3DFU")
End Sub

Private Sub tblSimpan_Click()
        
    If Not cekInputKeluhan Then
        Call Simpan_OK
    End If
End Sub

Private Sub Text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    If Text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            Text(Index).Mask = ""
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
End Sub

Private Sub Tgl_CloseUp()
    Call refreshGrid
End Sub

Private Sub transferDataPenjualanNew()
1     On Error GoTo transferDataPenjualanNew_error
          Dim rsTgl As New ADODB.Recordset
2         sSQL = "select top 1 tglentry as tgl,getdate() as skrg,dateadd(d,1,tglentry) tgl1,dateadd(d,-1,getdate()) skrg1 from salesunit order by tglentry desc"
3         Set rsTgl = objO.FetchData(sSQL)
4         If Not rsTgl.EOF Then
5             If DateDiff("d", Format(Trim(rsTgl("TGL")), "dd/mm/yyyy"), Format(Trim(rsTgl("skrg")), "dd/mm/yyyy")) > 1 Then
6                 erTranAll = False
7                 sSQL = "select NORANGKA, TANGGAL AS TANGGAL, B.kodeNRService, B.NAMADEALER from SUMMARY A LEFT JOIN PROFILCR..DEALER b on a.KODEDEALER = b.KODEDEALER  " & _
                          "WHERE ISNULL(HAPUS,0) = 0 and TIPETR = 'ND' AND TANGGAL between '" & TglNul2(Trim(rsTgl("TGL1"))) & "' and '" & TglNul2(Trim(rsTgl("skrg1"))) & "' "
                  
8                 Set rsTemp = objO.FetchDataSPKHO(sSQL)
                  
9                 Call objHT.OpenTransactionBengkel
10                If Not rsTemp.EOF Then
11                    rsTemp.MoveFirst
12                    Do While Not rsTemp.EOF
13                        Call objO.InsertData("insert into SalesUnit(NORANGKA,TGLENTRY,FS_KD_DEALER,FS_NM_DEALER) VALUES ('" & Trim(rsTemp!noRangka) & "','" & Format(Trim(rsTemp!Tanggal), "mm/dd/yyyy") & "','" & Trim(rsTemp!kodeNRService) & "','" & Trim(rsTemp!namaDealer) & "')")
14                        rsTemp.MoveNext
15                    Loop
16                End If
17                If erTranAll = False Then Call objHT.CloseTransactionBengkel
18            End If
19        End If
20        On Error GoTo 0
21        Exit Sub
transferDataPenjualanNew_error:
22        erTranAll = True
23        Call objHT.CloseTransactionBengkel_err
24        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure transferDataPenjualanNew of Form frm3DFU")
End Sub

Private Sub TransferDataPenjualan()
1        On Error GoTo TransferDataPenjualan_Error

      '   Call objO.InsertData_Penjualan_Vehicle
         ''2016-09-07
         ''status dihilangkan agar mendapatkan data penjualan terakhir jika terjadi retur penjualan disales
         ''
2        sSQL = "select NORANGKA, TANGGAL AS TANGGAL, fs_kd_dealer, fs_nm_dealer from OPENROWSET('SQLOLEDB','DRIVER={SQL Server};SERVER=" & Log.LogHOKompSPK & ";Uid=" & Log.LogHOUserSPK & ";PWD=" & Log.LogHOPassSPK & "',[" & Log.LogHODbSPK & "].dbo.SUMMARY) A LEFT JOIN T_DEALER b on a.KODEDEALER = b.KODECABANG  " & _
                           "WHERE ISNULL(HAPUS,0) = 0 and TIPETR = 'ND' AND TANGGAL between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "' " & _
                           "AND NORANGKA NOT IN (SELECT NORANGKA FROM SALESUNIT Where TglEntry between '" & TglNul2(DateAdd("d", -7, isnow)) & "' and '" & TglNul2(DateAdd("d", -1, isnow)) & "')"


3        Set rsTemp = objO.FetchData(sSQL)
4        If Not rsTemp.EOF Then
5           rsTemp.MoveFirst
6           Do While Not rsTemp.EOF
7              Call objO.InsertData("insert into SalesUnit(NORANGKA,TGLENTRY,FS_KD_DEALER,FS_NM_DEALER) VALUES ('" & Trim(rsTemp!noRangka) & "','" & Format(Trim(rsTemp!Tanggal), "mm/dd/yyyy") & "','" & Trim(rsTemp!FS_KD_DEALER) & "','" & Trim(rsTemp!fs_nm_dealer) & "')")
               
8              rsTemp.MoveNext
9           Loop
10       End If
11       On Error GoTo 0
12       Exit Sub

TransferDataPenjualan_Error:

13        MsgBox "Terjadi Kesalahan " & vbCrLf & ErrMessage(Erl, err.Number, " (" & err.Description & ") in procedure TransferDataPenjualan of Form frm3DFU")
End Sub

