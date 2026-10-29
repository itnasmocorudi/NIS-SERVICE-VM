VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Object = "{49CBFCC0-1337-11D2-9BBF-00A024695830}#1.0#0"; "tinumb6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{531653E6-69A2-40A4-8734-91C78B33DD08}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmBookingDMSNonDSS 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Booking"
   ClientHeight    =   11010
   ClientLeft      =   1965
   ClientTop       =   1815
   ClientWidth     =   11400
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
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
      Height          =   375
      Left            =   10860
      Picture         =   "frmBookingDMSNonDSS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   238
      Top             =   2760
      Width           =   1935
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
      Height          =   375
      Left            =   9000
      Picture         =   "frmBookingDMSNonDSS.frx":07BD
      Style           =   1  'Graphical
      TabIndex        =   237
      Top             =   2760
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   10695
      Left            =   0
      Picture         =   "frmBookingDMSNonDSS.frx":0F7A
      ScaleHeight     =   10695
      ScaleWidth      =   495
      TabIndex        =   236
      Top             =   0
      Width           =   495
   End
   Begin VB.Frame frameDPPart 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PENGISIAN DP PART"
      Height          =   1815
      Left            =   5160
      TabIndex        =   228
      Top             =   3840
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CommandButton Command3 
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
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   229
         Top             =   1320
         Width           =   1725
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   121
         Left            =   1680
         TabIndex        =   230
         Top             =   240
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   122
         Left            =   1680
         TabIndex        =   231
         Top             =   600
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
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
         Index           =   123
         Left            =   1680
         TabIndex        =   232
         Top             =   960
         Width           =   1740
         _ExtentX        =   3069
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
      Begin VB.Label Label65 
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
         Left            =   1200
         TabIndex        =   235
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label64 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP Part Minimal"
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
         TabIndex        =   234
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label63 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DP Part"
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
         Left            =   990
         TabIndex        =   233
         Top             =   1020
         Width           =   540
      End
   End
   Begin VB.Frame frmEstimasiSublet 
      BackColor       =   &H0080FF80&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   5400
      TabIndex        =   224
      Top             =   5760
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton cmdStalOk 
         BackColor       =   &H00DCFCFB&
         Caption         =   "&OK"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   960
         Width           =   1815
      End
      Begin MSMask.MaskEdBox text 
         Height          =   270
         Index           =   91
         Left            =   3000
         TabIndex        =   45
         Top             =   550
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
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
         Mask            =   "99:99"
         PromptChar      =   ">"
      End
      Begin VB.Label Label74 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "ESTIMASI PENGERJAAN SUBLET"
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
         Left            =   480
         TabIndex        =   227
         Top             =   0
         Width           =   3915
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   4815
      End
      Begin VB.Label Label62 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Konfirmasi Stall"
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
         Left            =   1320
         TabIndex        =   226
         Top             =   30
         Width           =   2040
      End
      Begin VB.Line Line12 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1200
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line11 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1200
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line10 
         BorderColor     =   &H000000FF&
         X1              =   3600
         X2              =   4680
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line9 
         BorderColor     =   &H000000FF&
         X1              =   3600
         X2              =   4680
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label Label61 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "ESTIMASI PENGERJAAN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   600
         TabIndex        =   225
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdNo 
      BackColor       =   &H00FFFF80&
      Caption         =   "No &Polisi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   223
      Top             =   50
      Width           =   1170
   End
   Begin VB.CommandButton cmdBersih 
      BackColor       =   &H0000FF00&
      Caption         =   "Bersih"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   181
      Top             =   75
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdSettingPrinter 
      BackColor       =   &H0000FFFF&
      Caption         =   "&Setting Printer"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   14520
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   176
      Top             =   435
      Visible         =   0   'False
      Width           =   1395
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
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   165
      Top             =   435
      Visible         =   0   'False
      Width           =   1050
   End
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   8055
      Left            =   13320
      TabIndex        =   166
      Top             =   0
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
   Begin VB.Frame frmeWaktu 
      BackColor       =   &H00C0FFFF&
      Height          =   3375
      Left            =   12960
      TabIndex        =   108
      Top             =   3360
      Visible         =   0   'False
      Width           =   6135
      Begin TrueOleDBGrid70.TDBGrid GridWaktu1 
         Height          =   3855
         Left            =   8640
         TabIndex        =   109
         Top             =   1440
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   6800
         _LayoutType     =   0
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
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowAddNew     =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
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
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=32,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
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
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin MSFlexGridLib.MSFlexGrid GridWaktu 
         Height          =   2625
         Left            =   120
         TabIndex        =   110
         Top             =   240
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   4630
         _Version        =   393216
         Cols            =   6
         HighLight       =   2
         SelectionMode   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hari H : Tanggal Customer Walk In"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   120
         TabIndex        =   115
         Top             =   2880
         Width           =   4425
      End
   End
   Begin MSMask.MaskEdBox text 
      Height          =   405
      Index           =   1
      Left            =   7515
      TabIndex        =   88
      Top             =   30
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   714
      _Version        =   393216
      BackColor       =   12640511
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   270
      Left            =   3120
      TabIndex        =   90
      Top             =   60
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
      Left            =   1680
      TabIndex        =   0
      Top             =   60
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   503
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      MaxLength       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "??"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   18
      Left            =   10200
      TabIndex        =   102
      Top             =   690
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
      Height          =   285
      Index           =   19
      Left            =   10200
      TabIndex        =   103
      Top             =   1020
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   503
      _Version        =   393216
      BackColor       =   12640511
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
      Mask            =   ">99:99"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   56
      Left            =   2040
      TabIndex        =   1
      Top             =   60
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   503
      _Version        =   393216
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
      Mask            =   "9999"
      PromptChar      =   ">"
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   57
      Left            =   2640
      TabIndex        =   2
      Top             =   60
      Width           =   420
      _ExtentX        =   741
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
      Mask            =   "CCC"
      PromptChar      =   ">"
   End
   Begin VB.CommandButton cmdWaktu 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Waktu Terpakai"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   107
      Top             =   360
      Width           =   1890
   End
   Begin VB.CommandButton cmdHO 
      BackColor       =   &H00DCFCFB&
      Caption         =   "Check &HO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   106
      Top             =   120
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Data Pelanggan"
      Height          =   2895
      Left            =   360
      TabIndex        =   75
      Top             =   480
      Width           =   8565
      Begin VB.CommandButton cmdPlgn 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Pelanggan"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3405
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   465
         Width           =   1050
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   3
         Left            =   915
         TabIndex        =   4
         Top             =   735
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   4
         Left            =   915
         TabIndex        =   5
         Top             =   1035
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   5
         Left            =   915
         TabIndex        =   6
         Top             =   1335
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   6
         Left            =   915
         TabIndex        =   8
         Top             =   1935
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   16
         Left            =   5430
         TabIndex        =   71
         Top             =   2520
         Width           =   1185
         _ExtentX        =   2090
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
         Height          =   285
         Index           =   17
         Left            =   7245
         TabIndex        =   72
         Top             =   2535
         Width           =   1245
         _ExtentX        =   2196
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
         Height          =   285
         Index           =   9
         Left            =   5430
         TabIndex        =   10
         Top             =   435
         Width           =   570
         _ExtentX        =   1005
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
         Mask            =   "9999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   13
         Left            =   5430
         TabIndex        =   12
         Top             =   1320
         Width           =   2925
         _ExtentX        =   5159
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   12
         Left            =   5430
         TabIndex        =   74
         Top             =   1035
         Width           =   3060
         _ExtentX        =   5398
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   11
         Left            =   7695
         TabIndex        =   73
         Top             =   735
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
         Mask            =   ">CCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox MaskEdBox2 
         Height          =   270
         Left            =   7560
         TabIndex        =   85
         Top             =   750
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
         Index           =   15
         Left            =   5430
         TabIndex        =   86
         Top             =   2235
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   2
         Left            =   915
         TabIndex        =   3
         Top             =   480
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   16777215
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   14
         Left            =   5430
         TabIndex        =   70
         Top             =   1935
         Width           =   3060
         _ExtentX        =   5398
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
         Mask            =   ">CCCCCCCCCCCCCCC"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   7
         Left            =   915
         TabIndex        =   9
         Top             =   2235
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   8
         Left            =   915
         TabIndex        =   7
         Top             =   1635
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         _Version        =   393216
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
         Mask            =   ">9999-9999999999"
         PromptChar      =   ">"
      End
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   41
         Left            =   5430
         TabIndex        =   111
         Top             =   1635
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   43
         Left            =   915
         TabIndex        =   69
         Top             =   2535
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   44
         Left            =   2640
         TabIndex        =   114
         Top             =   450
         Width           =   630
         _ExtentX        =   1111
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
         Index           =   10
         Left            =   5430
         TabIndex        =   11
         Top             =   735
         Width           =   2070
         _ExtentX        =   3651
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   71
         Left            =   2865
         TabIndex        =   144
         Top             =   2520
         Visible         =   0   'False
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox MaskEdBox13 
         Height          =   270
         Left            =   8400
         TabIndex        =   164
         Top             =   1320
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox14 
         Height          =   270
         Left            =   2500
         TabIndex        =   180
         Top             =   457
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   476
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
         PromptChar      =   "_"
      End
      Begin VB.Line Line8 
         BorderColor     =   &H000000FF&
         X1              =   6240
         X2              =   8400
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line7 
         BorderColor     =   &H000000FF&
         X1              =   6240
         X2              =   8400
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line6 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line5 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   1920
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label55 
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
         Left            =   2280
         TabIndex        =   175
         Top             =   45
         Width           =   3675
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   8535
      End
      Begin VB.Label Label33 
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
         Left            =   450
         TabIndex        =   113
         Top             =   2580
         Width           =   345
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No. Rangka"
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
         Left            =   4440
         TabIndex        =   112
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   4695
         TabIndex        =   93
         Top             =   1950
         Width           =   585
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "PSTN No."
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
         Left            =   -60
         TabIndex        =   92
         Top             =   1650
         Width           =   945
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "CDMA No."
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
         Left            =   -60
         TabIndex        =   91
         Top             =   2250
         Width           =   945
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -165
         TabIndex        =   87
         Top             =   495
         Width           =   1065
      End
      Begin VB.Label Label19 
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
         Left            =   4800
         TabIndex        =   84
         Top             =   1380
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
         Left            =   4860
         TabIndex        =   83
         Top             =   780
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
         Left            =   4635
         TabIndex        =   82
         Top             =   480
         Width           =   645
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DDTE."
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
         TabIndex        =   81
         Top             =   2580
         Width           =   450
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOTAM"
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
         Left            =   4740
         TabIndex        =   80
         Top             =   2550
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "GSM No."
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
         Left            =   -60
         TabIndex        =   79
         Top             =   1950
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -60
         TabIndex        =   78
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -60
         TabIndex        =   77
         Top             =   1050
         Width           =   945
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -60
         TabIndex        =   76
         Top             =   750
         Width           =   945
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "History Service"
      Height          =   1935
      Left            =   9000
      TabIndex        =   94
      Top             =   1440
      Width           =   3735
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   22
         Left            =   1080
         TabIndex        =   97
         Top             =   990
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   21
         Left            =   1080
         TabIndex        =   99
         Top             =   705
         Width           =   2340
         _ExtentX        =   4128
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   12640511
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
      Begin MSMask.MaskEdBox text 
         Height          =   285
         Index           =   20
         Left            =   1080
         TabIndex        =   95
         Top             =   435
         Width           =   1245
         _ExtentX        =   2196
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
         Mask            =   "99-99-9999"
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
         Left            =   480
         TabIndex        =   220
         Top             =   1680
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
         Left            =   2400
         TabIndex        =   219
         Top             =   1680
         Width           =   930
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   3120
         X2              =   3600
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   3120
         X2              =   3600
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   600
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         X1              =   120
         X2              =   600
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Label Label54 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Last Service"
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
         Left            =   1200
         TabIndex        =   174
         Top             =   45
         Width           =   1620
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   0
         Width           =   3735
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer"
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
         Left            =   360
         TabIndex        =   100
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "KM Akhir"
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
         Left            =   360
         TabIndex        =   98
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label10 
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
         Left            =   360
         TabIndex        =   96
         Top             =   480
         Width           =   570
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6240
      LargeChange     =   350
      Left            =   12480
      Max             =   5250
      SmallChange     =   350
      TabIndex        =   116
      Top             =   3360
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00808080&
      Height          =   6495
      Left            =   480
      ScaleHeight     =   6435
      ScaleWidth      =   12435
      TabIndex        =   117
      Top             =   3360
      Width           =   12495
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Input Booking"
         Height          =   11175
         Left            =   0
         TabIndex        =   118
         Top             =   -4600
         Width           =   11955
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame12"
            Height          =   255
            Left            =   5760
            TabIndex        =   216
            Top             =   120
            Width           =   6015
            Begin VB.Label JualNRM 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   75
               TabIndex        =   217
               Top             =   0
               Width           =   5865
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Caption         =   "Frame11"
            Height          =   495
            Left            =   7800
            TabIndex        =   215
            Top             =   10440
            Width           =   4095
            Begin ucXPButton.XPButton cmdcetak 
               Height          =   495
               Left            =   1200
               TabIndex        =   239
               Top             =   0
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
               MICON           =   "frmBookingDMSNonDSS.frx":2FC3
               PICN            =   "frmBookingDMSNonDSS.frx":2FDF
               PICH            =   "frmBookingDMSNonDSS.frx":3436
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ucXPButton.XPButton tblSimpan 
               Height          =   495
               Left            =   0
               TabIndex        =   240
               Top             =   0
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
               MICON           =   "frmBookingDMSNonDSS.frx":3920
               PICN            =   "frmBookingDMSNonDSS.frx":393C
               PICH            =   "frmBookingDMSNonDSS.frx":3E4F
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin ucXPButton.XPButton TblKeluar 
               Height          =   495
               Left            =   2880
               TabIndex        =   241
               Top             =   0
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
               MICON           =   "frmBookingDMSNonDSS.frx":43F1
               PICN            =   "frmBookingDMSNonDSS.frx":440D
               PICH            =   "frmBookingDMSNonDSS.frx":48E1
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Left            =   2160
            ScaleHeight     =   855
            ScaleWidth      =   3090
            TabIndex        =   209
            Top             =   10440
            Width           =   3090
            Begin TDBNumber6Ctl.TDBNumber txtRp 
               Height          =   315
               Index           =   0
               Left            =   1200
               TabIndex        =   210
               Top             =   30
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   556
               Calculator      =   "frmBookingDMSNonDSS.frx":4E2C
               Caption         =   "frmBookingDMSNonDSS.frx":4E4C
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmBookingDMSNonDSS.frx":4EB8
               Keys            =   "frmBookingDMSNonDSS.frx":4ED6
               Spin            =   "frmBookingDMSNonDSS.frx":4F20
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   12640511
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   ","
               DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###,##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   -99999999999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   "."
               ShowContextMenu =   -1
               ValueVT         =   1245189
               Value           =   0
               MaxValueVT      =   1397882885
               MinValueVT      =   1112342533
            End
            Begin TDBNumber6Ctl.TDBNumber txtRp 
               Height          =   315
               Index           =   1
               Left            =   1200
               TabIndex        =   211
               Top             =   360
               Width           =   1845
               _Version        =   65536
               _ExtentX        =   3254
               _ExtentY        =   556
               Calculator      =   "frmBookingDMSNonDSS.frx":4F48
               Caption         =   "frmBookingDMSNonDSS.frx":4F68
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               DropDown        =   "frmBookingDMSNonDSS.frx":4FD4
               Keys            =   "frmBookingDMSNonDSS.frx":4FF2
               Spin            =   "frmBookingDMSNonDSS.frx":503C
               AlignHorizontal =   1
               AlignVertical   =   0
               Appearance      =   0
               BackColor       =   12640511
               BorderStyle     =   1
               BtnPositioning  =   0
               ClipMode        =   0
               ClearAction     =   0
               DecimalPoint    =   ","
               DisplayFormat   =   "###,###,###,##0.00;(###,###,###,##0.00)"
               EditMode        =   0
               Enabled         =   -1
               ErrorBeep       =   0
               ForeColor       =   -2147483640
               Format          =   "###,###,###,##0.00"
               HighlightText   =   0
               MarginBottom    =   1
               MarginLeft      =   1
               MarginRight     =   1
               MarginTop       =   1
               MaxValue        =   99999999999
               MinValue        =   -99999999999
               MousePointer    =   0
               MoveOnLRKey     =   0
               NegativeColor   =   255
               OLEDragMode     =   0
               OLEDropMode     =   0
               ReadOnly        =   0
               Separator       =   "."
               ShowContextMenu =   -1
               ValueVT         =   1245189
               Value           =   0
               MaxValueVT      =   1397882885
               MinValueVT      =   1112342533
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Current AR"
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
               Left            =   0
               TabIndex        =   213
               Top             =   45
               Width           =   1050
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Platform AR"
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
               Left            =   0
               TabIndex        =   212
               Top             =   390
               Width           =   1155
            End
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   4215
            Left            =   120
            TabIndex        =   184
            Top             =   6120
            Width           =   11775
            _ExtentX        =   20770
            _ExtentY        =   7435
            _Version        =   393216
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   617
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Job Order && Sublet"
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame5"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "cmdStall"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).ControlCount=   2
            TabCaption(1)   =   "Part Order"
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame10"
            Tab(1).ControlCount=   1
            Begin VB.Frame Frame10 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   3735
               Left            =   -74970
               TabIndex        =   196
               Top             =   360
               Width           =   11655
               Begin MSFlexGridLib.MSFlexGrid PartGrid 
                  Height          =   3105
                  Left            =   120
                  TabIndex        =   51
                  Top             =   600
                  Width           =   7665
                  _ExtentX        =   13520
                  _ExtentY        =   5477
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  SelectionMode   =   1
               End
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   35
                  Left            =   120
                  TabIndex        =   49
                  Top             =   240
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   36
                  Left            =   3165
                  TabIndex        =   50
                  Top             =   240
                  Width           =   450
                  _ExtentX        =   794
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "99"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox MaskEdBox5 
                  Height          =   270
                  Left            =   3045
                  TabIndex        =   197
                  Top             =   240
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
                  Index           =   70
                  Left            =   3240
                  TabIndex        =   198
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   690
                  _ExtentX        =   1217
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
                  PromptChar      =   ">"
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Order Part"
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
                  TabIndex        =   199
                  Top             =   0
                  Width           =   765
               End
            End
            Begin VB.CommandButton cmdStall 
               BackColor       =   &H00DCFCFB&
               Caption         =   "S&tall"
               Height          =   615
               Left            =   10680
               Style           =   1  'Graphical
               TabIndex        =   48
               Top             =   360
               Width           =   930
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   3735
               Left            =   30
               TabIndex        =   185
               Top             =   360
               Width           =   11655
               Begin VB.CheckBox chSublet 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Part"
                  Height          =   255
                  Left            =   11040
                  TabIndex        =   43
                  Top             =   720
                  Width           =   615
               End
               Begin MSFlexGridLib.MSFlexGrid DGrid 
                  Height          =   2625
                  Left            =   120
                  TabIndex        =   38
                  Top             =   1080
                  Width           =   5505
                  _ExtentX        =   9710
                  _ExtentY        =   4630
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  ScrollBars      =   1
                  SelectionMode   =   1
               End
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   34
                  Left            =   900
                  TabIndex        =   36
                  Top             =   720
                  Width           =   3135
                  _ExtentX        =   5530
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox3 
                  Height          =   270
                  Left            =   4050
                  TabIndex        =   186
                  Top             =   720
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
                  Index           =   32
                  Left            =   120
                  TabIndex        =   35
                  Top             =   720
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox6 
                  Height          =   270
                  Left            =   750
                  TabIndex        =   187
                  Top             =   720
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
                  Index           =   42
                  Left            =   3240
                  TabIndex        =   188
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   690
                  _ExtentX        =   1217
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
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
                  Mask            =   "99,99"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   64
                  Left            =   4200
                  TabIndex        =   37
                  Top             =   720
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox12 
                  Height          =   270
                  Left            =   4860
                  TabIndex        =   189
                  Top             =   720
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
                  Index           =   33
                  Left            =   1920
                  TabIndex        =   193
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSFlexGridLib.MSFlexGrid subletGrid 
                  Height          =   2625
                  Left            =   5640
                  TabIndex        =   44
                  Top             =   1080
                  Width           =   5985
                  _ExtentX        =   10557
                  _ExtentY        =   4630
                  _Version        =   393216
                  FixedCols       =   0
                  HighLight       =   2
                  ScrollBars      =   2
                  SelectionMode   =   1
               End
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   79
                  Left            =   5640
                  TabIndex        =   39
                  Top             =   720
                  Width           =   2895
                  _ExtentX        =   5106
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   80
                  Left            =   8640
                  TabIndex        =   40
                  Top             =   720
                  Width           =   330
                  _ExtentX        =   582
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
                  PromptInclude   =   0   'False
                  AutoTab         =   -1  'True
                  MaxLength       =   2
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   "99"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox MaskEdBox15 
                  Height          =   270
                  Left            =   8550
                  TabIndex        =   194
                  Top             =   720
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
                  Index           =   81
                  Left            =   9000
                  TabIndex        =   41
                  Top             =   720
                  Width           =   810
                  _ExtentX        =   1429
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
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
                  Mask            =   ">?????"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   82
                  Left            =   9840
                  TabIndex        =   42
                  Top             =   720
                  Width           =   1170
                  _ExtentX        =   2064
                  _ExtentY        =   503
                  _Version        =   393216
                  BackColor       =   16777215
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
                  Format          =   "#,##0"
                  PromptChar      =   ">"
               End
               Begin MSMask.MaskEdBox text 
                  Height          =   315
                  Index           =   85
                  Left            =   3720
                  TabIndex        =   47
                  Top             =   120
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   556
                  _Version        =   393216
                  BackColor       =   16777215
                  AutoTab         =   -1  'True
                  MaxLength       =   5
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Courier New"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Mask            =   ">99:99"
                  PromptChar      =   ">"
               End
               Begin VB.Label Label60 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Estimasi Adjust"
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
                  Left            =   2160
                  TabIndex        =   222
                  Top             =   200
                  Width           =   1440
               End
               Begin VB.Label Label56 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Sublet Order"
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
                  Left            =   5640
                  TabIndex        =   195
                  Top             =   480
                  Width           =   915
               End
               Begin VB.Label Label52 
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
                  Left            =   4200
                  TabIndex        =   192
                  Top             =   480
                  Width           =   375
               End
               Begin VB.Label Label31 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Job Order"
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
                  Left            =   900
                  TabIndex        =   191
                  Top             =   510
                  Width           =   720
               End
               Begin VB.Label Label27 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tipe Job"
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
                  TabIndex        =   190
                  Top             =   510
                  Width           =   600
               End
            End
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   5880
            Left            =   5760
            TabIndex        =   132
            Top             =   240
            Width           =   6015
            Begin VB.ComboBox cboPelanggan 
               Height          =   315
               Left            =   840
               Style           =   2  'Dropdown List
               TabIndex        =   66
               Top             =   3360
               Width           =   1215
            End
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   39
               Left            =   3495
               TabIndex        =   133
               Top             =   3900
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   40
               Left            =   3495
               TabIndex        =   134
               Top             =   4200
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   52
               Left            =   3495
               TabIndex        =   68
               Top             =   4860
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   53
               Left            =   3495
               TabIndex        =   135
               Top             =   5460
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   55
               Left            =   3495
               TabIndex        =   136
               Top             =   4500
               Width           =   2250
               _ExtentX        =   3969
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   ">"
            End
            Begin VB.Frame Frame9 
               BackColor       =   &H00FFFFFF&
               Caption         =   "                                                                                                                                "
               Height          =   2535
               Left            =   0
               TabIndex        =   145
               Top             =   120
               Width           =   6015
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   46
                  Left            =   3765
                  TabIndex        =   52
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox4 
                  Height          =   270
                  Left            =   3645
                  TabIndex        =   146
                  Top             =   360
                  Visible         =   0   'False
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
                  Index           =   47
                  Left            =   3765
                  TabIndex        =   54
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox7 
                  Height          =   270
                  Left            =   3645
                  TabIndex        =   147
                  Top             =   720
                  Visible         =   0   'False
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
                  Index           =   48
                  Left            =   3765
                  TabIndex        =   56
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox8 
                  Height          =   270
                  Left            =   3645
                  TabIndex        =   148
                  Top             =   1080
                  Visible         =   0   'False
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
                  Index           =   49
                  Left            =   3765
                  TabIndex        =   58
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox9 
                  Height          =   270
                  Left            =   3720
                  TabIndex        =   149
                  Top             =   1440
                  Visible         =   0   'False
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
                  Index           =   50
                  Left            =   3765
                  TabIndex        =   60
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox10 
                  Height          =   270
                  Left            =   3645
                  TabIndex        =   150
                  Top             =   1800
                  Visible         =   0   'False
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
                  Index           =   51
                  Left            =   3765
                  TabIndex        =   62
                  Top             =   2040
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox MaskEdBox11 
                  Height          =   270
                  Left            =   3645
                  TabIndex        =   151
                  Top             =   2160
                  Visible         =   0   'False
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
                  Index           =   54
                  Left            =   3840
                  TabIndex        =   53
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   65
                  Left            =   3840
                  TabIndex        =   55
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   66
                  Left            =   3840
                  TabIndex        =   57
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   67
                  Left            =   3840
                  TabIndex        =   59
                  Top             =   1320
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   68
                  Left            =   3840
                  TabIndex        =   61
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin MSMask.MaskEdBox text 
                  Height          =   285
                  Index           =   69
                  Left            =   3840
                  TabIndex        =   63
                  Top             =   2040
                  Visible         =   0   'False
                  Width           =   2370
                  _ExtentX        =   4180
                  _ExtentY        =   503
                  _Version        =   393216
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
               Begin VB.Label Label57 
                  BackColor       =   &H00C0FFFF&
                  Caption         =   "-------------------------------------------------Kondisi ketika terjadi Keluhan ---------------------------------"
                  Height          =   255
                  Left            =   0
                  TabIndex        =   218
                  Top             =   120
                  Width           =   5895
               End
               Begin VB.Label Label73 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "(Berhenti, Mesin dingin, Panas)"
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
                  Left            =   1440
                  TabIndex        =   173
                  Top             =   2205
                  Width           =   2235
               End
               Begin VB.Label Label72 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "(Starter, Pemanasan, Konstan)"
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
                  Left            =   1440
                  TabIndex        =   172
                  Top             =   1920
                  Width           =   2250
               End
               Begin VB.Label Label71 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   " (Lampu menyala, Kedip sekali)"
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
                  Left            =   1440
                  TabIndex        =   171
                  Top             =   1485
                  Width           =   2205
               End
               Begin VB.Label Label70 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "(Jalan datar, Jalan Tol, Tanjakan)"
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
                  Left            =   1440
                  TabIndex        =   170
                  Top             =   1125
                  Width           =   2430
               End
               Begin VB.Label Label69 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   "(Selalu, Kadang-kadang)"
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
                  Left            =   1440
                  TabIndex        =   169
                  Top             =   765
                  Width           =   1770
               End
               Begin VB.Label Label68 
                  AutoSize        =   -1  'True
                  BackColor       =   &H00FFFFFF&
                  BackStyle       =   0  'Transparent
                  Caption         =   " (Akhir-akhir ini, Seminggu lalu)"
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
                  Left            =   1440
                  TabIndex        =   168
                  Top             =   405
                  Width           =   2205
               End
               Begin VB.Label Label39 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lainnya"
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
                  Left            =   720
                  TabIndex        =   156
                  Top             =   1920
                  Width           =   555
               End
               Begin VB.Label Label41 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lampu Peringatan"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   435
                  Left            =   -500
                  TabIndex        =   155
                  Top             =   1485
                  Width           =   1890
               End
               Begin VB.Label Label42 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Lokasi"
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
                  Left            =   840
                  TabIndex        =   154
                  Top             =   1125
                  Width           =   435
               End
               Begin VB.Label Label43 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Frekuensi"
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
                  Left            =   600
                  TabIndex        =   153
                  Top             =   765
                  Width           =   690
               End
               Begin VB.Label Label44 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Waktu"
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
                  Left            =   840
                  TabIndex        =   152
                  Top             =   405
                  Width           =   465
               End
            End
            Begin MSMask.MaskEdBox text 
               Height          =   270
               Index           =   75
               Left            =   2040
               TabIndex        =   67
               Top             =   3345
               Width           =   3885
               _ExtentX        =   6853
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   37
               Left            =   840
               TabIndex        =   64
               Top             =   2760
               Width           =   5100
               _ExtentX        =   8996
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   38
               Left            =   840
               TabIndex        =   65
               Top             =   3030
               Width           =   5100
               _ExtentX        =   8996
               _ExtentY        =   503
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   45
               Left            =   1080
               TabIndex        =   200
               Top             =   4320
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   59
               Left            =   1080
               TabIndex        =   201
               Top             =   3960
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
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
               Height          =   315
               Index           =   83
               Left            =   1080
               TabIndex        =   206
               Top             =   5160
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   84
               Left            =   1080
               TabIndex        =   207
               Top             =   4800
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
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
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Janji Penyerahan"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   0
               TabIndex        =   208
               Top             =   4920
               Width           =   960
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Estimasi Selesai  "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   240
               TabIndex        =   202
               Top             =   4080
               Width           =   720
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pesan Khusus"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               TabIndex        =   161
               Top             =   2640
               Width           =   615
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Saran"
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
               Left            =   330
               TabIndex        =   160
               Top             =   3075
               Width           =   420
            End
            Begin VB.Label Label37 
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
               Left            =   135
               TabIndex        =   159
               Top             =   3405
               Width           =   615
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL"
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
               Left            =   2850
               TabIndex        =   142
               Top             =   5520
               Width           =   480
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SUBLET"
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
               Left            =   2790
               TabIndex        =   141
               Top             =   4920
               Width           =   540
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PART"
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
               Left            =   2940
               TabIndex        =   140
               Top             =   4260
               Width           =   390
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "________________________  +"
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
               Left            =   3525
               TabIndex        =   139
               Top             =   5175
               Width           =   2370
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "JASA"
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
               Left            =   2955
               TabIndex        =   138
               Top             =   3960
               Width           =   375
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "OLI"
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
               Left            =   3075
               TabIndex        =   137
               Top             =   4560
               Width           =   255
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   4440
            Left            =   120
            TabIndex        =   128
            Top             =   1560
            Width           =   5640
            Begin VB.Frame Frame7 
               BackColor       =   &H00FFFFFF&
               Caption         =   "       Service       "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   1080
               TabIndex        =   203
               Top             =   3960
               Width           =   2175
               Begin VB.OptionButton opSvc 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Ditunggu"
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
                  Index           =   0
                  Left            =   120
                  TabIndex        =   205
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   975
               End
               Begin VB.OptionButton opSvc 
                  BackColor       =   &H00FFFFFF&
                  Caption         =   "Ditinggal"
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
                  Index           =   1
                  Left            =   1080
                  TabIndex        =   204
                  Top             =   240
                  Width           =   975
               End
            End
            Begin VB.TextBox Text1 
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
               Index           =   4
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   25
               Top             =   2280
               Width           =   4455
            End
            Begin VB.TextBox Text1 
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
               Index           =   3
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   24
               Top             =   2040
               Width           =   4455
            End
            Begin VB.TextBox Text1 
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
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   23
               Top             =   1800
               Width           =   4455
            End
            Begin VB.TextBox Text1 
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
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   22
               Top             =   1560
               Width           =   4455
            End
            Begin VB.TextBox Text1 
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
               Left            =   1080
               MaxLength       =   500
               TabIndex        =   21
               Top             =   1320
               Width           =   4455
            End
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   27
               Left            =   1095
               TabIndex        =   16
               Top             =   0
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   28
               Left            =   1095
               TabIndex        =   17
               Top             =   240
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   29
               Left            =   1095
               TabIndex        =   18
               Top             =   480
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   30
               Left            =   1095
               TabIndex        =   19
               Top             =   720
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   31
               Left            =   1095
               TabIndex        =   20
               Top             =   975
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   315
               Index           =   58
               Left            =   120
               TabIndex        =   143
               Top             =   840
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   12640511
               AutoTab         =   -1  'True
               MaxLength       =   5
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   72
               Left            =   5295
               TabIndex        =   26
               Top             =   1320
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   73
               Left            =   5295
               TabIndex        =   27
               Top             =   1560
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   74
               Left            =   5295
               TabIndex        =   28
               Top             =   1785
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   76
               Left            =   5295
               TabIndex        =   29
               Top             =   2010
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   240
               Index           =   77
               Left            =   5295
               TabIndex        =   30
               Top             =   2235
               Visible         =   0   'False
               Width           =   4440
               _ExtentX        =   7832
               _ExtentY        =   423
               _Version        =   393216
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   60
               Left            =   1095
               TabIndex        =   31
               Top             =   2760
               Width           =   4440
               _ExtentX        =   7832
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   61
               Left            =   1095
               TabIndex        =   32
               Top             =   3045
               Width           =   4440
               _ExtentX        =   7832
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   62
               Left            =   1095
               TabIndex        =   33
               Top             =   3360
               Width           =   4440
               _ExtentX        =   7832
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   63
               Left            =   1095
               TabIndex        =   34
               Top             =   3645
               Width           =   4440
               _ExtentX        =   7832
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
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   26
               Left            =   4440
               TabIndex        =   162
               Top             =   4080
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
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
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "User"
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
               Left            =   3990
               TabIndex        =   163
               Top             =   4125
               Width           =   330
            End
            Begin VB.Label Label50 
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
               Left            =   360
               TabIndex        =   158
               Top             =   2805
               Width           =   585
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Keluhan Customer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   157
               Top             =   1320
               Width           =   945
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Permintaan Customer"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   495
               Left            =   0
               TabIndex        =   129
               Top             =   120
               Width           =   945
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   120
            TabIndex        =   121
            Top             =   240
            Width           =   5625
            Begin VB.CheckBox chkKirim 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Kirim"
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
               Left            =   4680
               TabIndex        =   179
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox chSBNP 
               BackColor       =   &H00FFFFFF&
               Caption         =   "SB NP"
               Height          =   255
               Left            =   2520
               TabIndex        =   178
               Top             =   940
               Width           =   855
            End
            Begin VB.CheckBox chEM 
               BackColor       =   &H00FFFFFF&
               Caption         =   "EM"
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
               Left            =   1800
               TabIndex        =   177
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox chkAmbil 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Ambil"
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
               Left            =   3840
               TabIndex        =   167
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox chkThs 
               BackColor       =   &H00FFFFFF&
               Caption         =   "THS"
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
               Left            =   840
               TabIndex        =   131
               Top             =   940
               Width           =   735
            End
            Begin VB.CheckBox chkRtn 
               BackColor       =   &H00FFFFFF&
               Caption         =   "RTJ"
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
               Left            =   120
               TabIndex        =   130
               Top             =   940
               Width           =   735
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   23
               Left            =   1260
               TabIndex        =   13
               Top             =   0
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               Height          =   285
               Index           =   24
               Left            =   1260
               TabIndex        =   14
               Top             =   300
               Width           =   675
               _ExtentX        =   1191
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
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
               Mask            =   ">99:99"
               PromptChar      =   ">"
            End
            Begin MSMask.MaskEdBox text 
               Height          =   285
               Index           =   25
               Left            =   1260
               TabIndex        =   15
               Top             =   600
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               PromptInclude   =   0   'False
               AutoTab         =   -1  'True
               MaxLength       =   2
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Courier New"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Mask            =   ">CC"
               PromptChar      =   ">"
            End
            Begin VB.Shape Shape3 
               Height          =   315
               Left            =   0
               Top             =   930
               Width           =   5535
            End
            Begin VB.Label Label20 
               BackStyle       =   0  'Transparent
               Caption         =   "(HH:MM)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2040
               TabIndex        =   127
               Top             =   322
               Width           =   1140
            End
            Begin VB.Label Label30 
               BackStyle       =   0  'Transparent
               Caption         =   "(DD-MM-YYYY)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2535
               TabIndex        =   126
               Top             =   0
               Width           =   1140
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "D [ Datang ] / T [Telepon] / S [Sms]"
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
               Left            =   1710
               TabIndex        =   125
               Top             =   645
               Width           =   2550
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Via"
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
               Left            =   1005
               TabIndex        =   124
               Top             =   645
               Width           =   210
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jam Booking"
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
               Left            =   345
               TabIndex        =   123
               Top             =   345
               Width           =   885
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tanggal Booking"
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
               Left            =   60
               TabIndex        =   122
               Top             =   0
               Width           =   1170
            End
         End
         Begin VB.CommandButton cmdWarranty 
            BackColor       =   &H00C0C0C0&
            Caption         =   "&Warranty"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   10440
            Width           =   1890
         End
         Begin VB.CommandButton TblBatal 
            Caption         =   "&Batal"
            Height          =   390
            Left            =   9240
            TabIndex        =   119
            Top             =   11640
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.Label nAR2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0 % "
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   5280
            TabIndex        =   214
            Top             =   10800
            Width           =   840
         End
      End
   End
   Begin MSMask.MaskEdBox text 
      Height          =   285
      Index           =   78
      Left            =   11280
      TabIndex        =   183
      Top             =   240
      Visible         =   0   'False
      Width           =   1620
      _ExtentX        =   2858
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
      Mask            =   ">CCCCCCCCCCCCCCC"
      PromptChar      =   ">"
   End
   Begin VB.Label lblEdit 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   221
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image asxPowerBanner1 
      Height          =   9855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   375
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "(DD-MM-YYYY)"
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
      Left            =   11520
      TabIndex        =   105
      Top             =   720
      Width           =   1050
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "(HH:MM)"
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
      Left            =   10965
      TabIndex        =   104
      Top             =   1065
      Width           =   630
   End
   Begin VB.Label Label24 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Booking Order"
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
      Left            =   9030
      TabIndex        =   101
      Top             =   735
      Width           =   1020
   End
   Begin VB.Label Label40 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Booking"
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
      Left            =   6600
      TabIndex        =   89
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmBookingDMSNonDSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public baris As Integer
Dim barispart As Integer
Dim barisSublet As Integer
'Dim Enter As Integer
Dim stkPart As Integer
Dim namaPart As String
Dim HrgPart As Double
Dim satPart As String
Dim oliPart As Boolean
Public hrgLabour As Double
Dim Jrate As Double
Dim fOther As Boolean
Dim noPol As String
Public nJRate As Double
Dim nWKT As String
Dim nMech As String
Dim bKondisi(6) As String
Private nX As Boolean
Dim rpt As Object

Const MaxBrs = 8
Const MaxBrsPart = 30
Dim simpan As Boolean
Dim gNoRangka As String
Dim rsTemp As ADODB.Recordset

Dim namabySA As String
Dim NewId As Double
Dim estimasiSublet As String
     
'BOOKING ONLINE
Public isBookingOnline As Boolean
Public noPolBookingOnline As String
Public IDBookingOnline As Integer
Public noTelpBookingOnline As String
Public namaCustomerBookingOnline As String
Public kotaBookingOnline As String
Public tanggalBookingOnline As String
Public jamBookingOnline As String

Private isBookDiag As Boolean
Private no_approval As String


Public nilaiDPPart As Double
Public Property Let passingISBookingOnline(VALUES As Boolean)
    isBookingOnline = VALUES
End Property

Public Property Let passingNoPolBookingOnline(VALUES As String)
    noPolBookingOnline = VALUES
End Property

Public Property Let passingIDBookingOnline(VALUES As Integer)
    IDBookingOnline = VALUES
End Property

Public Property Let passingNoTelpBookingOnline(VALUES As String)
    noTelpBookingOnline = VALUES
End Property

Public Property Let passingNamaCustomerBookingOnline(VALUES As String)
    namaCustomerBookingOnline = VALUES
End Property

Public Property Let passingKotaBookingOnline(VALUES As String)
    kotaBookingOnline = VALUES
End Property

Public Property Let passingTanggalBookingOnline(VALUES As String)
    tanggalBookingOnline = VALUES
End Property

Public Property Let passingJamBookingOnline(VALUES As String)
    jamBookingOnline = VALUES
End Property




Public Property Let passingX(VALUES As Boolean)
    nX = VALUES
End Property

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

Private Sub hapusForm()
' *** FUNGSI UNTUK MEMBERSIHKAN TEXTBOX PADA FORM ***
    Frame2.Top = -100
    For i = 0 To 85
        Select Case i
            Case 0
                text(i) = ""
                text(i).Mask = "??"
            Case 56
                text(i) = ""
                text(i).Mask = "9999"
            Case 57
                text(i) = ""
                text(i).Mask = "CCC"
            Case 6, 7, 8
                text(i) = ""
                text(i).Mask = "9999-9999999999"
            Case 9
                text(i) = "0"
                text(i).Mask = "9999"
            Case 16, 17, 18, 20, 23, 59, 84
                text(i) = ">>->>->>>>"
                text(i).Mask = "99-99-9999"
            Case 19, 24, 45, 58, 83, 85
                text(i) = ">>:>>"
                text(i).Mask = "99:99"
            Case 36, 39, 40, 55, 52, 53, 70, 82
                text(i) = 0
            Case 42
                text(i) = ">>,>>"
                text(i).Mask = "99,99"
                text(i).Enabled = False
            Case Else
                text(i) = ""
        End Select
    Next
    
    For i = 0 To 4
        Text1(i) = ""
    Next
    
    For i = 1 To 7
        DGrid.TextMatrix(i, 1) = ""
        DGrid.TextMatrix(i, 2) = ""
        DGrid.TextMatrix(i, 3) = ""
        DGrid.TextMatrix(i, 4) = ""
        PartGrid.TextMatrix(i, 1) = ""
        PartGrid.TextMatrix(i, 2) = ""
        PartGrid.TextMatrix(i, 3) = ""
        PartGrid.TextMatrix(i, 4) = ""
        PartGrid.TextMatrix(i, 5) = ""
        PartGrid.TextMatrix(i, 6) = ""
        PartGrid.TextMatrix(i, 7) = ""
        PartGrid.TextMatrix(i, 8) = ""
    Next
    
    For i = 1 To 30
        subletGrid.TextMatrix(i, 0) = ""
        subletGrid.TextMatrix(i, 1) = ""
        subletGrid.TextMatrix(i, 2) = ""
        subletGrid.TextMatrix(i, 3) = ""
        subletGrid.TextMatrix(i, 4) = ""
        subletGrid.TextMatrix(i, 5) = ""
        subletGrid.TextMatrix(i, 6) = ""
    Next
    
    baris = 1
    barispart = 1
    barisSublet = 1
    
    
    For i = 0 To 5
        bKondisi(i) = ""
    Next
    
    nJRate = 0
    cmdStall.Tag = "belum"
    cmdWarranty.Enabled = False
    cmdWarranty.BackColor = &HC0C0C0
    cmdWaktu.Enabled = False
    cmdHistory.Enabled = False
    Command1.Enabled = False
    chkRtn = Unchecked
    chkThs = Unchecked
    VScroll1.value = 0
    isJualNRM = False
    text(54).Visible = False
    For i = 65 To 69
        text(i).Visible = False
    Next
    chkAmbil = Unchecked
    chkKirim = Unchecked
    chkRtn = Unchecked
    chkThs = Unchecked
    
    text(23).Enabled = True
    text(24).Enabled = True
    
    ' DEFAULT DOTAM & DDATE
    text(16) = "01-01-3000"
    text(17) = "01-01-3000"
    gNoRangka = ""
    text(18) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
    text(19) = Format(isnow, "HH:NN")
    If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then Call objO.UpdateDataACC("Update SOORD Set Status = '' Where Status = '" & gSA & "'")
    'Call isiNoBook
End Sub

Private Sub isiNoBook()
' *** FUNGSI UNTUK MENGISI NOMOR URUT BOOKING, 6 DIGIT ***
    '- proses mengisi no booking berdasarkan t_profil
    'Set oRs2 = objO.FetchData("SELECT FN_NO_BOOK from t_profil WITH (NOLOCK)")
    Set oRs2 = objO.FetchData("SELECT FN_NO_BOOK from t_profil")
    Dim no_urut As String
    
    '- pengisian no booking lama
    'Set oRs2 = objO.FetchData("SELECT nofak FROM Book WHERE nofak <> 'BOOKED' AND left(nofak,2) = '" & Right(Year(isnow), 2) & "' AND right(left(nofak,4),2) = '" & Format(Month(isnow), "0#") & "' ORDER BY nofak DESC")
'    If Not oRs2.EOF Then
'        Text(1) = Right(Year(isnow), 2) & Format(Month(isnow), "0#") & Format(CDbl(Right(oRs2!NoFak, 4)) + 1, "0###")
'    Else
'        Text(1) = Right(Year(isnow), 2) & Format(Month(isnow), "0#") & Format(1, "0###")
'    End If
'    oRs2.Close

    no_urut = ""
    If Not oRs2.EOF Then
        If IsNull(oRs2.Fields(0)) Then
            'jika ada yang null
            text(1) = Right(Year(isnow), 2) & Format(Month(isnow), "0#") & Format(1, "0###")
            no_urut = Format(1, "0###")
        Else
            text(1) = Right(Year(isnow), 2) & Format(Month(isnow), "0#") & Format(CDbl(Right(oRs2!FN_NO_BOOK, 4)) + 1, "0###")
            no_urut = Format(CDbl(Right(oRs2!FN_NO_BOOK, 4)) + 1, "0###")
        End If
    Else
        text(1) = Right(Year(isnow), 2) & Format(Month(isnow), "0#") & Format(1, "0###")
        no_urut = Format(1, "0###")
    End If
    
    objO.InsertData ("UPDATE t_profil SET FN_NO_BOOK = '" & no_urut & "' ")
    oRs2.Close
End Sub

Private Sub saveSublet()
    Call objHT.DeleteData_Bengkel("DELETE BOOK_SUBLET WHERE NOFAK = '" & Trim(text(1)) & "' AND NO_POLISI = '" & noPol & "' AND NO_CHASSIS = '" & Trim(text(41)) & "'")
    For i = 1 To barisSublet - 1
        With subletGrid
            If .TextMatrix(i, 1) <> "" Then
                Call objHT.DeleteData_Bengkel("DELETE FROM BOOK_SUBLET WHERE NOFAK = '" & Trim(text(1)) & "' AND NO_POLISI = '" & noPol & "' AND NO_CHASSIS = '" & Trim(text(41)) & "' AND NAMABARANG = '" & .TextMatrix(i, 1) & "'")
                Call objHT.InsertData_Bengkel("INSERT INTO BOOK_SUBLET VALUES ('" & Trim(text(1)) & "','" & Trim(text(41)) & "','" & Trim(noPol) & "','', '" & .TextMatrix(i, 1) & "', '" & .TextMatrix(i, 3) & "', '" & CDbl(.TextMatrix(i, 2)) & "', " & CDbl(.TextMatrix(i, 4)) & ",'" & IIf(.TextMatrix(i, 5) = "J", "J", "P") & "','" & .TextMatrix(i, 6) & "')")
            End If
        End With
    Next
End Sub

Private Sub savePart(noBok As String)
    Dim rsS As ADODB.Recordset
    Dim Flg As String
    Dim iNo As Double
    
    Dim rsSO As ADODB.Recordset
    If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then
        Set rsSO = objO.FetchDataACC("Select KodeItem + NamaBySA + convert(varchar(10),Nomor) as Kode from SOORD WHERE NOSPK = '" & noBok & "' Order By Nomor DESC")
    End If
    
    Dim rsSoH As ADODB.Recordset
    
    Set rsSoH = objO.FetchDataACC("Select NoSPK from SOHDR WHERE NOSPK = '" & noBok & "'")
    
    Dim nSyaratBayar As String
    Dim rsSb As ADODB.Recordset
    'Set rsSb = objO.FetchDataACC("SELECT TOPGRP as SyaratBayar FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & Trim(Text(2)) & "'")
    Set rsSb = objO.FetchDataACC("SELECT TOPGRP as SyaratBayar FROM ARMST WHERE KODELGN = '" & Trim(text(2)) & "'")
    If rsSb.RecordCount > 0 Then
        nSyaratBayar = rsSb.Fields(0).value
        Set rsSb = Nothing
    Else
        nSyaratBayar = "D0"
    End If
    
    Dim nDelIndex As String
    nDelIndex = GetDelIndex("SV")
    
    Call objHT.DeleteData_Bengkel("DELETE BOOK_PART WHERE NOFAK = '" & noBok & "'")
    
    Dim nSat As String
    nSat = ""

     Dim xNomor As String
     xNomor = ""
     Dim rsSOno As ADODB.Recordset
     'Set rsSOno = objO.FetchDataACC("Select top 1 Nomor from SOORD WITH (NOLOCK) WHERE NOSO = '" & Trim(Text(64)) & "' Order By Nomor DESC")
     Set rsSOno = objO.FetchDataACC("Select top 1 Nomor from SOORD WHERE NOSO = '" & Trim(text(64)) & "' Order By Nomor DESC")
     
     If rsSOno.RecordCount > 0 Then
         iNo = CDbl(rsSOno.Fields(0).value + 1)
         Set rsSOno = Nothing
     Else
         iNo = 1
     End If
        For i = 1 To barispart - 1
            If Trim(PartGrid.TextMatrix(i, 1)) <> "" Or Trim(PartGrid.TextMatrix(i, 7)) <> "" Then
                Set rsS = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & PartGrid.TextMatrix(i, 1) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan ORDER BY A.kodeitem")
                If Not rsS.EOF Then
                    If rsS("QTYSTOCK") < CDbl(PartGrid.TextMatrix(i, 4)) Then
                        Flg = "KURANG"
                    Else
                        Flg = "LENGKAP"
                    End If
                Else
                    Flg = "KURANG"
                End If
                
                'Call objHT.InsertData_Bengkel("INSERT INTO book_part VALUES ('" & noBok & "','" & Trim(partGrid.TextMatrix(i, 1)) & "','" & Trim(partGrid.TextMatrix(i, 2)) & "','" & IIf(rsS.RecordCount > 0, rsS("Kodesatuan"), "ITEM") & "','" & CDbl(partGrid.TextMatrix(i, 4)) & "','" & Flg & "','" & Format(Trim(Text(23)), "mm/dd/yyyy") & "','', '" & partGrid.TextMatrix(i, 7) & "'," & Val(partGrid.TextMatrix(i, 8)) & ")")
                Call objHT.InsertData_Bengkel("INSERT INTO book_part VALUES ('" & noBok & "','" & Trim(PartGrid.TextMatrix(i, 1)) & "','" & Trim(PartGrid.TextMatrix(i, 2)) & "','" & IIf(rsS.RecordCount > 0, rsS("Kodesatuan"), "ITEM") & "','" & CDbl(PartGrid.TextMatrix(i, 4)) & "','" & Flg & "','" & Format(Trim(text(23)), "mm/dd/yyyy") & "','', '" & PartGrid.TextMatrix(i, 7) & "'," & iNo & ")")
'                xNomor = xNomor & "," & iNo 'Trim(Val(partGrid.TextMatrix(i, 8)))
                If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then
                    If rsSO.RecordCount > 0 Then
                        rsSO.MoveFirst
                        rsSO.Find "Kode = '" & Trim(PartGrid.TextMatrix(i, 1)) + Trim(PartGrid.TextMatrix(i, 7)) + Trim(PartGrid.TextMatrix(i, 8)) & "'", , adSearchForward, 1
                        'rsSO.Find "Kode = '" & Trim(partGrid.TextMatrix(i, 1)) + Trim(partGrid.TextMatrix(i, 7)) + iNo & "'", , adSearchForward, 1
                        
                        'rsSO.Find "Kode = '" & Trim(PartGrid.TextMatrix(i, 8)) & "'", , adSearchForward, 1
                        If rsSO.EOF = True Then
                            xNomor = xNomor & "," & Trim(Val(PartGrid.TextMatrix(i, 8)))
                            Call objHT.InsertData_ACC("Insert Into SOORD (NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) Values ('" & noBok & "', '" & PartGrid.TextMatrix(i, 1) & "', '" & PartGrid.TextMatrix(i, 2) & "', '" & PartGrid.TextMatrix(i, 4) & "','1', '" & PartGrid.TextMatrix(i, 3) & "', '" & noBok & "', '" & gSA & "', Getdate(),'" & Val(PartGrid.TextMatrix(i, 8)) & "','" & PartGrid.TextMatrix(i, 7) & "', '0', '" & TglNul2(Trim(text(23))) + " " + Trim(text(24)) & "','" & gSA & "','" & TglNul2(Trim(text(23))) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), Trim(text(24)))), "HH:MM:SS") & "')")
                        Else
                            xNomor = xNomor & "," & Trim(Val(PartGrid.TextMatrix(i, 8)))
                        End If
                    Else
                        xNomor = xNomor & "," & Trim(Val(PartGrid.TextMatrix(i, 8)))
                        Call objHT.InsertData_ACC("Insert Into SOORD " & _
                            "(NOSO, KodeITEM, Keterangan,Qty, TypeSatuan, SatuanNs, NOSPK, KodeUser, TglEntry, Nomor, NAMABYSA, ADDJOB, TglPerlu, Status, TglRange ) " & _
                            "Values ('" & noBok & "', '" & PartGrid.TextMatrix(i, 1) & "', '" & PartGrid.TextMatrix(i, 2) & "', '" & PartGrid.TextMatrix(i, 4) & "','1', '" & PartGrid.TextMatrix(i, 3) & "', '" & noBok & "', '" & gSA & "', Getdate(),'" & Val(PartGrid.TextMatrix(i, 8)) & "','" & PartGrid.TextMatrix(i, 7) & "', '0', '" & TglNul2(Trim(text(23))) + " " + Trim(text(24)) & "','" & gSA & "','" & TglNul2(Trim(text(23))) + " " + Format((DateAdd(Left(Log.AppRate, 1), Right(Log.AppRate, 1), Trim(text(24)))), "HH:MM:SS") & "')")
                    End If
                End If
            End If
        Next
    'End If
    
        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then
          '  Set rsSoH = objO.FetchDataACC("Select NoSPK from SOHDR WHERE noso = '" & noBok & "'")
            If barispart > 1 Then
            
                If rsSoH.RecordCount > 0 Then
                    Call objHT.UpdateData_ACC("Update SOHDR Set TglPerlu = '" & TglNul2(Trim(text(23))) + " " + Trim(text(24)) & "' Where NoSPK = '" & noBok & "'")
                Else
                    If barispart > 1 Then
                        Call objHT.InsertData_ACC("insert into SOHDR " & _
                        "(NOSO,TglSO, TypeTrn, TglPerlu, KodeLgn, NamaKirim, AlamatKirim1, AlamatKirim2,AlamatKirim3,SyaratBayar, KODESUMBER, DELINDEX, TglEntry, KodeUser, NoSPK, TUTUPORDER ) " & _
                        "values ('" & noBok & "', Getdate(), 'A','" & TglNul2(Trim(text(23))) + " " + Trim(text(24)) & "', '" & Trim(text(2)) & "', '" & Trim(text(3)) & "', '" & Trim(text(4)) & "', '" & Trim(text(5)) & "', '', '" & nSyaratBayar & "', 'SV', '" & nDelIndex & "', getdate(), '" & gSA & "', '" & noBok & "',0 )")
                    End If
                End If
            End If
            If xNomor = "" Then
            
            Else
                Call objHT.UpdateData_ACC("Delete SOORD Where NOSPK = '" & noBok & "' AND IsNull(qtyreserve,0) = 0 and Nomor not in (" & Right(xNomor, Len(xNomor) - 1) & ")")
            End If
            
        End If
End Sub

Private Sub TampilWaktu(Tggl As Date)
Dim rsWaktu As ADODB.Recordset
Dim nStall As String
Dim i As Integer

    GridWaktu.TextMatrix(0, 0) = "..."
    GridWaktu.TextMatrix(0, 1) = "STALL"
    
    'FUNGSI MENGHASILKAN WAKTU TERSEDIA BERDASARKAN TGL TTN (HARI H)
    If DateDiff("d", isnow, Tggl) >= 2 Then
        Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl, "mm/dd/yyyy") & "'")
        GridWaktu.TextMatrix(0, 2) = "H - 2"
        GridWaktu.TextMatrix(0, 3) = "H - 1"
        GridWaktu.TextMatrix(0, 4) = "  H  "
        GridWaktu.TextMatrix(0, 5) = "H + 1"
    ElseIf DateDiff("d", isnow, Tggl) = 1 Then
        Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 1, "mm/dd/yyyy") & "'")
        GridWaktu.TextMatrix(0, 2) = "H - 1"
        GridWaktu.TextMatrix(0, 3) = "  H  "
        GridWaktu.TextMatrix(0, 4) = "H + 1"
        GridWaktu.TextMatrix(0, 5) = "H + 2"
    ElseIf DateDiff("d", isnow, Tggl) <= 0 Then
        Call objO.UpdateData("EXEC sp_lihatStall '" & Format(Tggl + 2, "mm/dd/yyyy") & "'")
        GridWaktu.TextMatrix(0, 2) = "  H  "
        GridWaktu.TextMatrix(0, 3) = "H + 1"
        GridWaktu.TextMatrix(0, 4) = "H + 2"
        GridWaktu.TextMatrix(0, 5) = "H + 3"
    End If
    
    'FUNSGI UNTUK MENAMPILKAN WAKTU TERSEDIA H-2 SAMPAI H+1
    Set rsWaktu = objO.FetchData("SELECT * FROM STALLWKF WITH (NOLOCK) ORDER BY LEFT(NOSTALL,1),CASE LEFT(NOSTALL,1) WHEN 'S' THEN CAST(SUBSTRING(NOSTALL, 7, LEN(NOSTALL)-6) AS INTEGER) ELSE CAST(SUBSTRING(NOSTALL, 5, LEN(NOSTALL)-4) AS INTEGER) END, [H-2], [H-1], H, [H+1]")
    If Not rsWaktu.EOF Then
        i = 1
        GridWaktu.Rows = rsWaktu.RecordCount + 1
        
        GridWaktu.ColWidth(0) = 200
        GridWaktu.ColWidth(1) = 700
        GridWaktu.ColWidth(2) = 1200
        GridWaktu.ColWidth(3) = 1200
        GridWaktu.ColWidth(4) = 1200
        GridWaktu.ColWidth(5) = 1200
        nStall = ""
        
        Do While Not rsWaktu.EOF
            If Trim(rsWaktu("NOSTALL")) = nStall Then
                GridWaktu.TextMatrix(i, 1) = ""
                GridWaktu.TextMatrix(i, 2) = Nul(Trim(rsWaktu("H-2")), "")
                GridWaktu.TextMatrix(i, 3) = Nul(Trim(rsWaktu("H-1")), "")
                GridWaktu.TextMatrix(i, 4) = Nul(Trim(rsWaktu("H")), "")
                GridWaktu.TextMatrix(i, 5) = Nul(Trim(rsWaktu("H+1")))
            Else
                GridWaktu.TextMatrix(i, 0) = "*"
                GridWaktu.TextMatrix(i, 1) = Trim(rsWaktu("NOSTALL"))
                GridWaktu.TextMatrix(i, 2) = Nul(Trim(rsWaktu("H-2")), "")
                GridWaktu.TextMatrix(i, 3) = Nul(Trim(rsWaktu("H-1")), "")
                GridWaktu.TextMatrix(i, 4) = Nul(Trim(rsWaktu("H")), "")
                GridWaktu.TextMatrix(i, 5) = Nul(Trim(rsWaktu("H+1")))
                nStall = Trim(rsWaktu("NOSTALL"))
            End If
            i = i + 1
            rsWaktu.MoveNext
        Loop
    End If
End Sub

Public Sub panggilWarrantyBok(noRangka As String, modL As String)
Dim rsRec As ADODB.Recordset
Dim terlibat As Boolean

    Set rsTemp = objO.FetchData("SELECT * FROM ho_warranty A WITH (NOLOCK) INNER JOIN ho_wardtl B WITH (NOLOCK) ON A.id_warranty = B.id_warranty and A.FS_LOG_DEL = '' " _
               & "WHERE MDL = '" & modL & "' AND no_rangka1 <= '" & noRangka & "' AND no_Rangka2 >= '" & noRangka & "' " _
               & "ORDER BY A.id_warranty,mdl,no_Rangka1")
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            Set rsRec = objO.FetchData("SELECT no_chassis FROM History WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
            If Not rsRec.EOF Then
                terlibat = False
                Exit Do
            Else
                terlibat = True
'                Exit Do
            End If
            rsRec.Close
            
            Set rsRec = objO.FetchData("SELECT NO_RANGKA FROM TEMPWAR WITH (NOLOCK) WHERE NO_RANGKA = '" & noRangka & "' AND ID_WARRANTY = '" & rsTemp("ID_WARRANTY") & "'")
            If Not rsRec.EOF Then
                terlibat = False
                Exit Do
            Else
                terlibat = True
'                Exit Do
            End If
            rsRec.Close
    
            If Log.LogHOConn = 1 Then
                If ConHO = True Then
                    Set rsRec = objO.FetchDataHO("SELECT no_chassis FROM Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
                    If Not rsRec.EOF Then
                        terlibat = False
                        Exit Do
                    Else
                        terlibat = True
    '                    Exit Do
                    End If
                End If
            End If
            rsTemp.MoveNext
        Loop
        
        If terlibat Then
            MsgBox "UNIT TERLIBAT WARRANTY !!!", vbCritical, "UNIT SSC"
            cmdWarranty.BackColor = &HFF&
            cmdWarranty.Enabled = True
        End If
    Else
        cmdWarranty.BackColor = &HC0C0C0
        cmdWarranty.Enabled = False
    End If
End Sub

Private Sub panggilCustHO(nPol As String, rs As ADODB.Recordset)
    
    If Not rs.EOF Then
        text(2) = rs("FS_KD_PLG")
        text(44) = rs("FS_KD_WIL")
        text(3) = rs("CUSTOMER")
        text(4) = rs("CUST_ADRES")
        text(5) = rs("KOTA")
        text(6) = rs("GSM")
        text(7) = rs("CDMA")
        text(8) = rs("TELPON")
        text(9) = rs("FS_TH_MOBIL")
        text(10) = rs("MDL")
        text(11) = rs("MODEL")
                    Call CexChassis_Discon(text(11))
                    If isOnCarClass = False Then
                        MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
                    End If
        text(12) = rs("FMDL")
        text(13) = rs("COLOR")
        text(14) = rs("FS_KD_DEALER")
        text(15) = rs("FS_NM_DEALER")
        text(16) = TglDDMMYYYY(IIf(rs("DOTAM") = "", "01-01-3000", rs("DOTAM")))    'TglDDMMYYYY(rs("DOTAM"))
        text(17) = TglDDMMYYYY(IIf(rs("DDATE") = "", "01-01-3000", rs("DDATE")))
        gNoRangka = rs("NO_CHASSIS")
        text(41) = gNoRangka
        text(26) = gSA
        cmdHistory.Enabled = True
        Command1.Enabled = True
    End If
End Sub

Private Sub cboPelanggan_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub chEM_Click()
    Call hitungEstimasiSelesai
End Sub

Private Sub chSBNP_Click()
    Call hitungEstimasiSelesai
End Sub

Private Sub chSublet_KeyPress(KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Public Sub PindahStallEM()
    If Log.isEmStall = 1 Then
        If chEM = Checked Then
            Dim rsAStl As ADODB.Recordset
            Dim keterangan As String
            Dim noPolEM As String
            Dim jamJM As Integer
            Dim menitJM As Integer
            Dim jamSelesai As String
            
            noPolEM = Trim(text(0)) & "-" & Trim(text(56)) & "-" & Trim(text(57))
            
            If noPolEM <> "" Then
                If Trim(text(1)) <> "" Then
                    keterangan = noPolEM & " |" & Trim(text(1))
                Else
                    keterangan = noPolEM & " |"
                End If
                Set rsAStl = objO.FetchData("SELECT * FROM stall WITH(NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "%' ORDER BY noStall")
                If Not rsAStl.EOF Then
                    If rsAStl("ADJUST") <> "00:30" Then
                        'lakukan pergantian nostall tersebut dengan adjust EM
                        jamJM = CInt(Left(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") - 1))
                        menitJM = Mid(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") + 1, Len(Trim(rsAStl("jammulai"))) - InStr(Trim(rsAStl("jammulai")), ":"))
                                                
                        menitJM = menitJM + 30
                        
                        If menitJM >= 60 Then
                           jamJM = jamJM + Val(menitJM / 60)
                           menitJM = menitJM Mod 60
                        End If
                        jamSelesai = Format(jamJM, "0#") & ":" & Format(menitJM, "0#")
                        objO.InsertData ("UPDATE STALL SET JAMSELESAI = '" & jamSelesai & "',ADJUST = '00:30' WHERE tanggal = '" & Format(Trim(text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "%'")
                    End If
                Else
                    Set rsAStl = objO.FetchData("SELECT * FROM tempstall WITH(NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "' ORDER BY noStall")
                    If Not rsAStl.EOF Then
                        If rsAStl("ADJUST") <> "00:30" Then
                            'lakukan pergantian nostall tersebut dengan adjust EM
                            jamJM = CInt(Left(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") - 1))
                            menitJM = Mid(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") + 1, Len(Trim(rsAStl("jammulai"))) - InStr(Trim(rsAStl("jammulai")), ":"))
                                                    
                            menitJM = menitJM + 30
                            
                            If menitJM >= 60 Then
                               jamJM = jamJM + Val(menitJM / 60)
                               menitJM = menitJM Mod 60
                            End If
                            jamSelesai = Format(jamJM, "0#") & ":" & Format(menitJM, "0#")
                            objO.InsertData ("UPDATE TEMPSTALL SET JAMSELESAI = '" & jamSelesai & "',ADJUST = '00:30' WHERE tanggal = '" & Format(Trim(text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "'")
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Sub PindahStallSBNP()
    If Log.isEmStall = 1 Then
        If chEM = Checked Then
            Dim rsAStl As ADODB.Recordset
            Dim keterangan As String
            Dim noPolEM As String
            Dim jamJM As Integer
            Dim menitJM As Integer
            Dim jamSelesai As String
            
            noPolEM = Trim(text(0)) & "-" & Trim(text(56)) & "-" & Trim(text(57))
            
            If noPolEM <> "" Then
                If Trim(text(1)) <> "" Then
                    keterangan = noPolEM & " |" & Trim(text(1))
                Else
                    keterangan = noPolEM & " |"
                End If
                Set rsAStl = objO.FetchData("SELECT * FROM stall WITH(NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "%' ORDER BY noStall")
                If Not rsAStl.EOF Then
                    If rsAStl("ADJUST") <> "00:30" Then
                        'lakukan pergantian nostall tersebut dengan adjust EM
                        jamJM = CInt(Left(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") - 1))
                        menitJM = Mid(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") + 1, Len(Trim(rsAStl("jammulai"))) - InStr(Trim(rsAStl("jammulai")), ":"))
                                                
                        menitJM = menitJM + 60
                        
                        If menitJM >= 60 Then
                           jamJM = jamJM + Val(menitJM / 60)
                           menitJM = menitJM Mod 60
                        End If
                        jamSelesai = Format(jamJM, "0#") & ":" & Format(menitJM, "0#")
                        objO.InsertData ("UPDATE STALL SET JAMSELESAI = '" & jamSelesai & "',ADJUST = '00:30' WHERE tanggal = '" & Format(Trim(text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "%'")
                    End If
                Else
                    Set rsAStl = objO.FetchData("SELECT * FROM tempstall WITH(NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "' ORDER BY noStall")
                    If Not rsAStl.EOF Then
                        If rsAStl("ADJUST") <> "00:30" Then
                            'lakukan pergantian nostall tersebut dengan adjust EM
                            jamJM = CInt(Left(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") - 1))
                            menitJM = Mid(Trim(rsAStl("jammulai")), InStr(Trim(rsAStl("jammulai")), ":") + 1, Len(Trim(rsAStl("jammulai"))) - InStr(Trim(rsAStl("jammulai")), ":"))
                                                    
                            menitJM = menitJM + 60
                            
                            If menitJM >= 60 Then
                               jamJM = jamJM + Val(menitJM / 60)
                               menitJM = menitJM Mod 60
                            End If
                            jamSelesai = Format(jamJM, "0#") & ":" & Format(menitJM, "0#")
                            objO.InsertData ("UPDATE TEMPSTALL SET JAMSELESAI = '" & jamSelesai & "',ADJUST = '00:30' WHERE tanggal = '" & Format(Trim(text(23)), "MM/DD/YYYY") & "' AND FLAG <> 'X' AND keterangan LIKE '" & keterangan & "'")
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub


Private Sub cmdBersih_Click()
Dim nPol1 As String
Dim nPol2 As String
Dim nPol3 As String
    
    nPol1 = Trim(text(0))
    nPol2 = Trim(text(56))
    nPol3 = Trim(text(57))
    Call hapusForm
    text(0) = nPol1
    text(56) = nPol2
    text(57) = nPol3
    
    text(3).SelLength = Len(text(3))
    text(3).SetFocus
End Sub

Private Sub cmdCetak_Click()
Dim rsSearch As ADODB.Recordset

    Set rsSearch = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & Trim(text(1)) & "' AND DELIV <> 'BTL'")
    If Not rsSearch.EOF Then
        'Call CetakForm
        Call CetakForm1
    Else
        MsgBox "Booking belum disimpan", vbCritical, "SIMPAN BOOKING"
        Exit Sub
    End If
End Sub

Private Sub cmdHistory_Click()
'    If gNoRangka <> "" Then
'        gTag = "BOOKING"
'        frmHistory.Text(0) = gNoRangka
'        frmHistory.Text(1) = Text(3)
'        frmHistory.Text(2) = noPol
'        frmHistory.Show (vbModal)
'    End If
    If gNoRangka = "" Then Exit Sub
    If Log.LogHOConn = 0 Then
        MsgBox "Gagal Koneksi ke Server HO !", vbInformation + vbOKOnly, "History On Line HO"
        Exit Sub
    End If
    gTag = "ONLINE"
    dRangka = gNoRangka
    nData(0) = gNoRangka
    nData(1) = text(3)
    nData(2) = noPol
    Dim fH As New FrmHistory_new
    fH.Show (vbModal)
End Sub

Private Sub cmdStalOk_Click()
    
    estimasiSublet = ""
    Dim jJ As Integer
    Dim jM As Integer
    Dim total As Integer
    
    jJ = CInt(Left(Trim(text(91)), InStr(Trim(text(91)), ":") - 1))
    jM = Mid(Trim(text(91)), InStr(Trim(text(91)), ":") + 1, Len(Trim(text(91))) - InStr(Trim(text(91)), ":"))
    
    total = (jJ * 60) + jM
    If total = 0 Then
        MsgBox "Estimasi pengerjaan sublet harus diisi", vbCritical
    ElseIf total > 240 Then
        MsgBox "Estimasi pengerjaan sublet maksimal 4 Jam", vbCritical
    Else
        estimasiSublet = Trim(text(91))
        frmEstimasiSublet.Visible = False
        
        If MaxBrs >= barisSublet Then
            With subletGrid
                .TextMatrix(barisSublet, 1) = Trim(text(79))                                ' NAMA SUBLET
                .TextMatrix(barisSublet, 2) = CDbl(text(80))                                ' QTY
                .TextMatrix(barisSublet, 3) = Trim(text(81))                                ' SATUAN
                .TextMatrix(barisSublet, 4) = Format(CDbl(text(82)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
                .TextMatrix(barisSublet, 5) = IIf(chSublet = Checked, 1, 0)
                .TextMatrix(barisSublet, 6) = estimasiSublet
            End With
            barisSublet = barisSublet + 1
            text(52) = Format(CDbl(text(52)) + (CDbl(text(80)) * CDbl(text(82))), "#0,#")
            text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#0,#")
            
            text(79) = ""
            text(80) = 0
            text(81) = ""
            text(82) = 0
            chSublet = Unchecked
            Call hitungEstimasiSelesai
            text(79).SetFocus
        End If
    End If
    
End Sub

Private Sub Command1_Click()
    If Replace(Replace(text(15), ">", ""), " ", "") = "" Then Exit Sub
    gTag = "LOCAL"
    dRangka = gNoRangka
    nData(0) = gNoRangka
    nData(1) = text(3)
    nData(2) = noPol
    
   ' fH.Show (vbModal)
    If Log.DmsNonDss = "1" Then
        Dim fHDND As New FrmHistory_newDMSNonDSS
        fHDND.isOrder = False
        fHDND.noBaris = baris
        fHDND.modelKendaraan = Trim(text(10))
        fHDND.Show (vbModal)
    Else
        Dim fH As New FrmHistory_new
        fH.Show (vbModal)
    End If
    
End Sub

Private Sub cmdHO_Click()
Dim rsHlp As ADODB.Recordset
    If noPol <> "" Then
        If Log.LogHOConn = 1 Then
            If ConHO = True Then
                Set rsHlp = objO.FetchDataHO("SELECT * FROM CUSTDAT WITH (NOLOCK) WHERE NO_POLISI = '" & noPol & "'")
                If Not rsHlp.EOF Then
                    Call panggilCustHO(noPol, rsHlp)
                    Call panggilHistoryHO(rsHlp("NO_CHASSIS"))
                End If
            End If
        End If
    End If
End Sub

Private Sub cmdNo_Click()
    If cmdNo.Caption = "No &Polisi" Then
        Call hapusForm
        cmdNo.Caption = "No &Rekom"
    ElseIf cmdNo.Caption = "No &Rekom" Then
        Call hapusForm
        cmdNo.Caption = "No &Polisi"
    End If
End Sub

Private Sub cmdNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        text(0).SetFocus
    End If
End Sub

Private Sub cmdPlgn_Click()
    With frmtPlgn
        .text(57) = text(2)
        .text(1) = text(3)
        .text(2) = Left(text(4), 40)
        .text(3) = text(5)
        .text(4) = text(8) + "," + Trim(text(6)) + "," + Trim(text(7))
        gBantuan2 = "kosong"
        frmtPlgn.Show vbModal, Me
    End With
End Sub

Private Sub cmdSettingPrinter_Click()
    rpt.PrinterSetup Me.hWnd
    CRViewer1.Refresh
End Sub

Private Sub cmdStall_Click()
    Dim nRate As Double
    Dim pesan As String
    Dim boleh As Boolean
    
    noPol = Trim(text(0)) & IIf(Trim(text(0)) <> "", "-", "") & Trim(text(56)) & IIf(Trim(text(57)) <> "", "-", "") & Trim(text(57))
    If nMech = "" Then
        nMech = Trim(text(64))
    End If
    boleh = True
    nRate = 0
    For i = 1 To baris - 1
        If Trim(DGrid.TextMatrix(i, 1)) <> "" Then
            Set rsTemp = objO.FetchData("SELECT RATE FROM actRAte WITH (NOLOCK) WHERE MDL = '" & text(10) & "' AND FS_ITEM = '" & Trim(DGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
            If Not rsTemp.EOF Then
                nRate = nRate + CDbl(rsTemp("RATE")) 'DGrid.TextMatrix(i, 2)
            Else
                boleh = False
                pesan = pesan + vbCr & "Actual Rate untuk " & Trim(DGrid.TextMatrix(i, 1)) & " belum ada "
                Exit For
            End If
        End If
    Next
    
    ' HITUNG NJRATE
    If nJRate = 0 Then
        For i = 1 To baris - 1
            If Trim(DGrid.TextMatrix(i, 1)) <> "" Then
                nJRate = nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(DGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
            End If
        Next
    End If
    
    If baris = 1 Then
        pesan = pesan + "Job Order. "
        boleh = False
    Else
        If nRate = 0 Then
            pesan = pesan + "Nilai Actual Rate 0. "
            boleh = False
        End If
    End If
    If Trim(noPol) = "" Then
        noPol = Trim(text(0)) & IIf(Trim(text(0)) <> "", "-", "") & Trim(text(56)) & IIf(Trim(text(57)) <> "", "-", "") & Trim(text(57))
        If Trim(noPol) = "" Then
            pesan = pesan + "No Polisi. "
            boleh = False
        End If
    End If
    If Trim(text(3)) = "" Then
        pesan = pesan + "Nama Customer. "
        boleh = False
    End If
    If Trim(text(23)) = ">>->>->>>>" Then
        pesan = pesan + "Tanggal Booking. "
        boleh = False
    End If
    
    If Log.isEmStall = 1 Then
        If Trim(text(85)) = ">>:>>" Then
            pesan = pesan + "Estimasi Pengerjaan service harus diisi. "
            boleh = False
        End If
    End If
    
    If boleh Then
        If chEM = Checked Then
            Call PindahStallEM
            text(85) = "00:30"
        End If
        If chSBNP = Checked Then
            Call PindahStallSBNP
            text(85) = "01:00"
        End If
        ' ISI TEMPSTALL
         Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                     & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%') AND " _
                     & "tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%' ORDER BY noStall, Tanggal, JamMulai")
        
        gStall = "booking"
        gTanggalStall = Trim(text(23))
        
        If Log.isEmStall = 1 Then
            With frmStall
                .text(13).Enabled = True
                .text(7) = text(23)
                .text(14) = Trim(text(85))
                .text(4).Mask = ""
                .text(4) = noPol
                .text(5) = text(2)
                .text(6) = text(3)
                .text(12) = text(1)
                .text(9) = IIf(chEM = Checked, 0.5, IIf(chSBNP = Checked, 1, nRate))
                .text(8) = nMech
                .chkThs.value = chkThs.value
                .chEM.value = IIf(chEM = Checked, chEM.value, IIf(chSBNP = Checked, chSBNP.value, 0))
                .frmStal.Visible = False
                If isBookByMrs = False Then
                    NotChild frmStall.hWnd
                Else
                    TOChild frmBookingDMSNonDSS.hWnd, frmStall.hWnd
                    '.Show
                End If
                .WindowState = vbMaximized
                'frmStall.Show
            End With
        Else
            With frmStallNonEM
                '.Text(13).Enabled = True
                .text(7) = text(23)
                .text(4).Mask = ""
                .text(4) = noPol
                .text(5) = text(2)
                .text(6) = text(3)
                .text(12) = text(1)
                .text(9) = IIf(chEM = Checked, 0.5, IIf(chSBNP = Checked, 1, nRate))
                .text(8) = nMech
                .chkThs.value = chkThs.value
                .chEM.value = IIf(chEM = Checked, chEM.value, IIf(chSBNP = Checked, chSBNP.value, 0))
                .frmStal.Visible = False
                If isBookByMrs = False Then
                    NotChild frmStallNonEM.hWnd
                Else
                    TOChild frmBookingDMSNonDSS.hWnd, frmStallNonEM.hWnd
                    '.Show
                End If
                .WindowState = vbMaximized
                'frmStall.Show
            End With
        End If
        
        
    Else
        MsgBox "Lengkapi data Booking !!" & vbCr & pesan, vbCritical, "Isi Field Kosong"
        Exit Sub
    End If
End Sub

Private Sub cmdTutup_Click()
    CRViewer1.Visible = False
    cmdTutup.Visible = False
    cmdSettingPrinter.Visible = False
End Sub

Private Sub cmdWaktu_Click()
    If frmeWaktu.Visible Then
        frmeWaktu.Visible = False
    Else
        If text(23) <> "" And text(23) <> ">>->>->>>>" Then
            frmeWaktu.Visible = True
            frmeWaktu.Top = cmdWaktu.Top + cmdWaktu.Height + 100
            frmeWaktu.Left = cmdWaktu.Left - (frmeWaktu.Width - cmdWaktu.Width)
            Call TampilWaktu(Trim(text(23)))
        Else
            MsgBox "Tanggal booking harus diisi"
            text(23).SelLength = Len(text(23))
            text(23).SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdWarranty_Click()
    If text(23) = ">>->>->>>>" And text(24) = ">>:>>" Then
        MsgBox "Isi tanggal & jam Booking !"
        text(23).SetFocus
        Exit Sub
    Else
        gTag = "BOOKING"
        With frmWarranty
            .text(0) = gNoRangka
            .text(1) = text(3)
            .text(2) = noPol
            .Brs.Caption = baris
            .BrsPart.Caption = barispart
            .text(3) = Trim(text(23))
            .text(4) = Trim(text(24))
            .Show (vbModal)
        End With
    End If
End Sub


Private Sub Command3_Click()

End Sub

Private Sub DGrid_DblClick()
    If baris > 1 And DGrid.TextMatrix(DGrid.row, 1) <> "" Then
        a = DGrid.row
        
        If Trim(text(1)) <> "" Then
            Dim nCexPP As String
            Dim rsPP As ADODB.Recordset
            'Set rsPP = objO.FetchData("Select Item_Part from HO_PartPack Where MDL = '" & Text(10) & "' and FS_Item = '" & dGrid.TextMatrix(a, 1) & "' and JOB = '" & dGrid.TextMatrix(a, 4) & "'")
            Set rsPP = objO.FetchData("Select Nama_Part from HO_PartPack WITH (NOLOCK) Where MDL = '" & text(10) & "' and FS_Item = '" & DGrid.TextMatrix(a, 1) & "' and JOB = '" & DGrid.TextMatrix(a, 4) & "'")
            If rsPP.RecordCount > 0 Then
                nCexPP = "Select NOSO FROM SOORD WITH (NOLOCK) Where NoSPK = '" & Trim(text(1)) & "' and ("
                rsPP.MoveFirst
                Do Until rsPP.EOF
                    'nCexPP = nCexPP & " KodeItem like '%" & rsPP.Fields(0).value & "' OR "
                    nCexPP = nCexPP & " NamaBySA = '" & rsPP.Fields(0).value & "' OR "
                    rsPP.MoveNext
                Loop
                nCexPP = Left(nCexPP, Len(nCexPP) - 3) & ")"
                nCexPP = nCexPP & " and ((IsNull(QtyKirimBL,0) + IsNull(QtyKirimBI,0) + IsNull(QtyKeluar,0)) <> 0 Or IsNull(QtyReserve,0) <> 0 ) AND ([TO] <> 0 Or BO not in ('','CANCEL') )"
                Dim rsPSO As ADODB.Recordset
                Set rsPSO = objO.FetchDataACC(nCexPP)
                If rsPSO.RecordCount > 0 Then
                    MsgBox "Paket Part sudah di proses oleh partman !," & Chr(13) & "Job Order tidak dapat di hapus" & Chr(13) & Chr(13) & "Lakukan pembatalan proses part oleh Partman, untuk menghapus Job Order"
                    Exit Sub
                End If
            End If

        End If
        
        'AMBIL
        With DGrid
            text(34) = .TextMatrix(a, 1)
            Jrate = .TextMatrix(a, 2)
            text(32) = .TextMatrix(a, 4)
'            Text(50) = .TextMatrix(a, 5)
            If .TextMatrix(a, 3) <> "" Then
                If .TextMatrix(a, 3) Then
                    fOther = True
                    text(42).Visible = True
                    text(42).Mask = ""
                    text(42) = Jrate
                Else
                    fOther = False
                    text(42).Visible = False
                End If
            Else
                fOther = False
                text(42).Visible = False
            End If
        End With
        myChassis = Trim(text(41))
        Call cekPartPack(DGrid.TextMatrix(a, 4), DGrid.TextMatrix(a, 1), text(10), 2, 1, barispart, frmBookingDMSNonDSS)
        'hrgLabour = hitungLabour(DGrid.TextMatrix(baris - 1, 2), Text(11))
        hrgLabour = hitungLabour(DGrid.TextMatrix(a, 2), text(11))
        If isJualNRM = True Then
            Me.JualNRM.Caption = nJualNRM
        Else
            Me.JualNRM.Caption = ""
        End If
        nJRate = nJRate - Jrate
        If text(58) <> ">>:>>" And text(23) <> ">>->>->>>>" Then
            nWKT = hitungWaktu(CDbl(nJRate), text(58), text(23))
            text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
            text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
            text(84) = Trim(text(59))
            text(83) = Trim(text(45))
            
            Call hitungEstimasiSelesai
        End If
        
        'HAPUS
        For i = a To baris
            With DGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
            End With
        Next
        
        text(39) = Format(CDbl(text(39) - (CDbl(hrgLabour))), "#0,#")
        text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#0,#")
        baris = baris - 1
    End If
End Sub

Private Sub dGrid_GotFocus()
    Dim bolehTambah As Boolean

    If Trim(text(34)) <> "" And text(23) <> ">>->>->>>>" And text(24) <> ">>->>->>>>" And text(64) <> "" Then
        For i = 0 To baris
            If DGrid.TextMatrix(i, 1) = Trim(text(34)) Then
                bolehTambah = False
                text(34) = ""
                Exit Sub
            Else
                bolehTambah = True
            End If
        Next
        
        nMech = Trim(text(64))
        If bolehTambah Then
            ' JIKA PEKERJAAN INPUT SENDIRI
            If fOther Then
                Jrate = CDbl(Replace(Replace(text(42), " ", ""), ">", ""))
            End If
            
            If MaxBrs >= baris Then
                With DGrid
                    .TextMatrix(baris, 1) = text(34)
                    .TextMatrix(baris, 2) = Jrate
                    .TextMatrix(baris, 3) = fOther
                    .TextMatrix(baris, 4) = text(32)
                End With
                baris = baris + 1
                
                ' ISI PART PACK
                Call cekPartPack(text(32), text(34), text(10), 1, 1, barispart, frmBookingDMSNonDSS)
                myChassis = Trim(text(41))
                text(39) = Format(CDbl(text(39)) + Val(hitungLabour(CDbl(Jrate), Trim(text(11)))), "#0,0")
                nJRate = nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(DGrid.TextMatrix(baris - 1, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                If text(58) <> ">>:>>" And text(23) <> ">>->>->>>>" Then
                    nWKT = hitungWaktu(CDbl(nJRate), text(58), text(23))
                    text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
                    text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
                    text(84) = Trim(text(59))
                    text(83) = Trim(text(45))
                    Call hitungEstimasiSelesai
                End If
                If isJualNRM = True Then
                    Me.JualNRM.Caption = nJualNRM
                Else
                    Me.JualNRM.Caption = ""
                End If
                text(34) = ""
                text(32) = ""
                text(42) = ">>,>>"
                text(42).Visible = False
                text(32).SetFocus
                hrgLabour = Val(hitungLabour(CDbl(Jrate), Trim(text(11)))) / IIf(CDbl(Jrate) = 0, 1, CDbl(Jrate))
                If isJualNRM = True Then
                    Me.JualNRM.Caption = nJualNRM
                Else
                    Me.JualNRM.Caption = ""
                End If
                text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#0,#")
            Else
                MsgBox "Maksimal " & str(MaxBrs) & " !"
                text(34).SetFocus
            End If
        End If
    End If
End Sub

Private Sub DGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call DGrid_DblClick
    End If
    If KeyCode = vbKeyReturn Then
        Call DGrid_DblClick
    End If
End Sub

Private Sub Form_Activate()
Dim nDpn As String
    Dim nTgh As String
    Dim nBlg As String
    If isBookingOnline Then
        Call pecahNoPol(noPolBookingOnline, nDpn, nTgh, nBlg)
        
        text(6) = noTelpBookingOnline
        text(3) = namaCustomerBookingOnline
        text(5) = kotaBookingOnline
        
        text(23) = TglDDMMYYYY(CDate(tanggalBookingOnline)) 'Format(tanggalBookingOnline, "")
        text(24) = Left(jamBookingOnline, 5)
        
        text(0) = nDpn
        text(56) = nTgh
        text(57) = nBlg
        SendKeys ("{TAB}")
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        text(0).SetFocus
    End If
End Sub

Private Sub Form_Load()
    isBookDiag = False
    bookAktif = True
    baris = 1
    DGrid.Rows = MaxBrs
    DGrid.cols = 5
    DGrid.ColWidth(0) = 0
    DGrid.ColWidth(1) = 3500
    DGrid.ColWidth(2) = 800
    DGrid.ColWidth(3) = 0
    DGrid.ColWidth(4) = 0
    DGrid.ColAlignment(1) = flexAlignLeftCenter
    DGrid.TextMatrix(0, 1) = "Order Booking"
    
    barispart = 1
    PartGrid.Rows = MaxBrsPart
    PartGrid.cols = 9
    PartGrid.ColWidth(0) = 0
    PartGrid.ColWidth(1) = 1500
    PartGrid.ColWidth(2) = 2500
    PartGrid.ColWidth(3) = 0
    PartGrid.ColWidth(4) = 800
    PartGrid.ColWidth(5) = 0
    PartGrid.ColWidth(6) = 0
    PartGrid.ColWidth(7) = 2500
    PartGrid.ColWidth(8) = 0
    PartGrid.ColAlignment(1) = flexAlignLeftCenter
    PartGrid.TextMatrix(0, 1) = "Kode Part"
    PartGrid.TextMatrix(0, 2) = "Nama Part"
    PartGrid.TextMatrix(0, 4) = "Qty"
    PartGrid.TextMatrix(0, 7) = "Nama By PIC"
    
    ' ATUR GRID SUBLET
    barisSublet = 1
    subletGrid.Rows = 31
    subletGrid.cols = 6
    subletGrid.Clear
    subletGrid.FormatString = " | Sublet | Qty | Satuan | Harga | Part | Est Waktu"
    subletGrid.ColWidth(0) = 0
    subletGrid.ColWidth(1) = 2500
    subletGrid.ColWidth(2) = 600
    subletGrid.ColWidth(3) = 1000
    subletGrid.ColWidth(4) = 1200
    subletGrid.ColWidth(5) = 0
    subletGrid.ColWidth(6) = 0
    subletGrid.TextMatrix(0, 1) = "Order Sublet"
    
    Call hapusForm
    If Layar.Tag = "estBook" Then
        baris = gBantuan
        barispart = gBantuan2
        barisSublet = gBantuan4
        noPol = gBantuan3
    End If
    Layar.Tag = ""
    
    text(0).Mask = "??"
    text(56).Mask = "9999"
    text(57).Mask = "CCC"
    text(26) = gSA
    
    ' ISI COMBO KARAKTER
    cboPelanggan.Clear
    Set rsTemp = objO.FetchData("SELECT * FROM T_Karakter WITH (NOLOCK) ORDER BY kode")
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            cboPelanggan.AddItem rsTemp("KODE")
            rsTemp.MoveNext
        Loop
    End If
    cboPelanggan.ListIndex = 0
    If cmdNo.Caption = "No &Polisi" Then cmdBersih.Visible = False
    
    simpan = True
    nX = False
    
    If Log.isEmStall = 0 Then
        Label60.Visible = False
        text(85).Visible = False
    End If
    NewId = (objO.FetchData_DoubleACC("select top 1 Nomor from SOORD WITH (NOLOCK) Where NOSO = '' Order By Nomor DESC")) + 1
    
    
End Sub

Private Sub Form_Resize()
    If Me.Height > 500 Then
        asxPowerBanner1.Height = Me.Height - 500
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    bookAktif = False
End Sub

Private Sub opSvc_KeyPress(Index As Integer, KeyAscii As Integer)
    Call CekTombol(KeyAscii)
End Sub

Private Sub partGrid_DblClick()
    If barispart > 1 And (PartGrid.TextMatrix(PartGrid.row, 1) <> "" Or PartGrid.TextMatrix(PartGrid.row, 7) <> "") Then

    If Trim(text(1)) <> "" Then
        Dim rsSO As ADODB.Recordset
        Set rsSO = objO.FetchDataACC("Select NOSPK from SOORD WITH (NOLOCK) Where NOSPK = '" & Trim(text(1)) & "' and KodeItem = '" & (PartGrid.TextMatrix(PartGrid.row, 1)) & "' and isNull(NamaBYSA,'') = '" & (PartGrid.TextMatrix(PartGrid.row, 7)) & "' and Nomor = " & PartGrid.TextMatrix(PartGrid.row, 8) & "  AND ((IsNull(qtyreserve,0)) <> 0 ) AND ([TO] <> 0 Or BO not in ('','CANCEL'))")
        If rsSO.RecordCount > 0 Then
            MsgBox "Spare Part : " & (PartGrid.TextMatrix(PartGrid.row, 1)) & " " & (PartGrid.TextMatrix(PartGrid.row, 7)) & Chr(13) & "Sudah disupply !, tidak dapat dihapus/dirubah !", vbInformation + vbOKOnly
            Exit Sub
        End If
    End If
    
        a = PartGrid.row
        'AMBIL
        With PartGrid
            If .TextMatrix(PartGrid.row, 1) <> "" Then
                text(35) = .TextMatrix(a, 1)
                namaPart = .TextMatrix(a, 2)
                satPart = .TextMatrix(a, 3)
                text(36) = CDbl(.TextMatrix(a, 4))
                HrgPart = CDbl(.TextMatrix(a, 5))
                oliPart = NulA(.TextMatrix(a, 6))
                namabySA = .TextMatrix(a, 7)
                
                If oliPart Then
                    text(55) = Format(CDbl(text(55)) - hitungPart(PartGrid.TextMatrix(a, 4), PartGrid.TextMatrix(a, 5)), "#0,#")
                Else
                    text(40) = Format(CDbl(text(40)) - hitungPart(PartGrid.TextMatrix(a, 4), PartGrid.TextMatrix(a, 5)), "#0,#")
                End If
        
            Else
                text(35) = .TextMatrix(a, 7)
                namaPart = ""
                satPart = .TextMatrix(a, 3)
                text(36) = CDbl(.TextMatrix(a, 4))
                HrgPart = 0
                oliPart = .TextMatrix(a, 6)
                namabySA = .TextMatrix(a, 7)
                
                text(55) = Format(CDbl(text(55)) - hitungPart(PartGrid.TextMatrix(a, 4), 0), "#0,#")

            End If
        End With
        
        'If oliPart Then
        '    Text(55) = Format(CDbl(Text(55)) - hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5)), "#0,#")
        'Else
        '    Text(40) = Format(CDbl(Text(40)) - hitungPart(partGrid.TextMatrix(a, 4), partGrid.TextMatrix(a, 5)), "#0,#")
        'End If
        text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#0,#")
        
        'HAPUS
        For i = a To barispart
            With PartGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
                .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)
                .TextMatrix(i, 6) = .TextMatrix(i + 1, 6)
                .TextMatrix(i, 7) = .TextMatrix(i + 1, 7)
                .TextMatrix(i, 8) = .TextMatrix(i + 1, 8)
            End With
        Next

        barispart = barispart - 1
        text(35).SetFocus
    End If
End Sub

Private Sub PartGrid_GotFocus()
Dim bolehTambah As Boolean

    If text(36) = "" Then
        text(36) = 0
    End If
    If Trim(text(35)) <> "" And CDbl(text(36)) > 0 Then
        For i = 1 To barispart
            If PartGrid.TextMatrix(i, 1) = Trim(text(35)) Or PartGrid.TextMatrix(i, 7) = Trim(text(35)) Then
                If MsgBox("Part sudah pernah diinput, lanjutkan ?", vbQuestion + vbYesNo, "Part Booking") = vbYes Then
                    bolehTambah = True
                Else
                    text(35) = ""
                    Exit Sub
                End If
            Else
                bolehTambah = True
                If PartGrid.TextMatrix(i, 8) = "" Then
                    NewId = NewId + 1
                Else
                    NewId = PartGrid.TextMatrix(i, 8)
                End If
            End If
        Next
        
        If bolehTambah Then
            If MaxBrsPart >= barispart Then
                With PartGrid
                    If namaPart <> "" Then
                    .TextMatrix(barispart, 1) = text(35)
                    .TextMatrix(barispart, 2) = namaPart
                    .TextMatrix(barispart, 3) = satPart
                    .TextMatrix(barispart, 4) = text(36)
                    .TextMatrix(barispart, 5) = HrgPart
                    .TextMatrix(barispart, 6) = oliPart
                    .TextMatrix(barispart, 7) = ""
                    .TextMatrix(barispart, 8) = NewId
                    Else
                    .TextMatrix(barispart, 1) = ""
                    .TextMatrix(barispart, 2) = ""
                    .TextMatrix(barispart, 3) = satPart
                    .TextMatrix(barispart, 4) = text(36)
                    .TextMatrix(barispart, 5) = HrgPart
                    .TextMatrix(barispart, 6) = oliPart
                    .TextMatrix(barispart, 7) = text(35)
                    .TextMatrix(barispart, 8) = NewId
                    End If
                End With
                barispart = barispart + 1
                If oliPart Then
                    text(55) = Format(CDbl(text(55)) + hitungPart(PartGrid.TextMatrix(barispart - 1, 4), PartGrid.TextMatrix(barispart - 1, 5)), "#0,0")
                Else
                    text(40) = Format(CDbl(text(40)) + hitungPart(PartGrid.TextMatrix(barispart - 1, 4), PartGrid.TextMatrix(barispart - 1, 5)), "#0,0")
                End If
                text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#0,#")
                
                text(35) = ""
                text(36) = 0
                namaPart = ""
                text(35).SetFocus
                NewId = NewId + 1
            Else
                MsgBox "Maksimal Part " & str(MaxBrsPart) & " !"
                text(35).SetFocus
            End If
        End If
    End If
End Sub

Private Sub PartGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call partGrid_DblClick
    End If
    If KeyCode = vbKeyReturn Then
        Call partGrid_DblClick
    End If
End Sub

Private Sub tblBatal_Click()
    If MsgBox("Yakin Membatalkan booking ?", vbYesNo, "Konfirmasi") = vbYes Then
        Call objO.UpdateData("UPDATE BOOK set DELIV='BTL',PAID='D' where NO_POLISI ='" & noPol & "' and DELIV='BOK'")
        Call hapusForm
        text(0) = ""
        text(56) = ""
        text(57) = ""
    End If
End Sub

Private Sub TblKeluar_Click()
    Call cekTempStal
    If cmdWarranty.Visible Then
        Call cekWarranty(Trim(text(41)), Trim(text(71)))
    End If
    If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then Call objO.UpdateDataACC("Update SOORD Set Status = '' Where Status = '" & gSA & "'")
    Unload Me
End Sub

Private Sub cekWarranty(nRgk As String, kdWar As String)
    Dim rsHlp As ADODB.Recordset
    Set rsHlp = objO.FetchData("SELECT NO_CHASSIS FROM BOOK WITH (NOLOCK) WHERE NO_CHASSIS = '" & nRgk & "' AND KODETIPE = '" & kdWar & "'")
    If rsHlp.EOF Then
       Call objO.DeleteData("DELETE TEMPWAR WHERE NO_RANGKA = '" & nRgk & "' AND ID_WARRANTY = '" & kdWar & "'")
    End If
End Sub

Private Sub cekTempStal()
    If Trim(noPol) = "" Then
        text(0) = ""
        text(56) = ""
        text(57) = ""
    ElseIf noPol <> "--" And Tanggal <> ">>->>->>>>" Then
        If text(23) <> ">>->>->>>>" Then
            Set rsTemp = objO.FetchData("SELECT tanggal FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%'")
            If Not rsTemp.EOF Then
                Call objO.DeleteData("DELETE tempStall WHERE tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '" & Trim(noPol) & " |%'")
            End If
        End If
    End If
End Sub

Private Sub TblKeluar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        text(0).SetFocus
    End If
End Sub

Private Sub tblSimpan_Click()
    Dim boleh As Boolean
    Dim pesan As String
    boleh = True
    
    'On Error GoTo vbError
    ' MEMERIKSA FIELD MANDATORY YANG KOSONG, DENGAN FLAG
    If Trim(noPol) <> "" Then
        If CekTgl(text(23)) = False Then
            pesan = pesan + "Tgl Booking. "
            boleh = False
        End If
        If text(24) = ">>:>>" Then
            pesan = pesan + "Jam Booking. "
            boleh = False
        End If
        If Trim(text(10)) = "" Then
            pesan = pesan + "Model. "
            boleh = False
        End If
        If Trim(text(3)) = "" Then
            pesan = pesan + "Nama. "
            boleh = False
        End If
        If Trim(text(4)) = "" Then
            pesan = pesan + "Alamat. "
            boleh = False
        End If
        If Trim(text(8)) = "" And Trim(text(6)) = "" And Trim(text(7)) = "" Then
            pesan = pesan + "Telepon. "
            boleh = False
        End If
        If CekTgl(text(23)) Then
            If baris > 1 Then
                If cmdStall.Tag = "" Or cmdStall.Tag = "belum" Then
                    Set rsTemp = objO.FetchData("SELECT tanggal FROM tempstall WITH (NOLOCK) WHERE tanggal = '" & Format(text(23), "mm/dd/yyyy") & "' AND keterangan like '%" & Trim(noPol) & " |%'")
                Else
                    Set rsTemp = objO.FetchData("SELECT tanggal FROM stall WITH (NOLOCK) WHERE tanggal = '" & Format(text(23), "mm/dd/yyyy") & "' AND keterangan like '%" & Trim(noPol) & " |%'")
                End If
                If rsTemp.EOF Then
                    pesan = pesan + "Stall. "
                    boleh = False
                    cmdStall.Tag = "belum"
                    cmdStall.BackColor = &HFFFF&
                Else
                    cmdStall.Tag = ""
                    cmdStall.BackColor = &H8000000F
                End If
            Else
                ' non job order only sublet
                If barisSublet > 1 Then
                Else
                    boleh = False
                    pesan = pesan + "Job / Sublet Order. "
                    cmdStall.Tag = "belum"
                    cmdStall.BackColor = &HFFFF&
                End If
            End If
        End If
        If text(27) = "" And text(28) = "" And text(29) = "" And text(30) = "" And text(31) = "" Then
            boleh = False
            pesan = pesan + "Permintaan Customer. "
        End If
        If text(64) <> "" Then
            If text(1) = "" Then
                nTgl = text(23)
                gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
                Set rsTemp = objO.FetchData("select A.MECH from mekanik A WITH (NOLOCK) LEFT JOIN T_MEKANIK B WITH (NOLOCK) ON A.MECH = B.MECH AND hari = '" & gBantuan2 & "' AND TANGGAL = '" & gBantuan3 & "' LEFT JOIN T_STALL C WITH (NOLOCK) ON C.HARI = '" & gBantuan2 & "' WHERE DATEDIFF(n,WAKTUMULAI,WAKTUSELESAI) - ISNULL(ALOKASI,0) > 0 AND A.MECH = '" & text(64) & "'")
                If rsTemp.EOF Then
                    boleh = False
                    pesan = pesan + "Waktu Mekanik tidak mencukupi"
                End If
            End If
        End If
        
        If Log.DmsNonDss = 1 Then
            If ((Log.persentaseDP / 100) * CDbl(Trim(text(40)))) >= Log.plafonDP Then
                If nilaiDPPart < ((Log.persentaseDP / 100) * CDbl(Trim(text(0)))) Then
                    MsgBox "Nilai DP Part minimal harus Rp. " & ((Log.persentaseDP / 100) * CDbl(Trim(text(0)))), vbCritical
                                        
                Else
                
                End If
            End If
            
        End If
        '20/12/2012  Cust Non Axtif
        If Trim(text(2).text) <> "" Then
            Dim rsNon As ADODB.Recordset
            Set rsNon = objO.FetchDataACC("SELECT kodeLgn FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & text(2) & "' and nonaktifflag = 0")
            If rsNon.RecordCount = 0 Then
                boleh = False
                pesan = "Kode Langgan : " & text(2).text & ", Adalah Cust Non Aktif !" & Chr(13) & "Silahkan Rubah Kode Langgan"
            End If
        End If
        
        If boleh = True Then
            Call saveLogMenu(gSA, ComptName, gVersi, "frmBookingDMSNonDSS", "tblSimpan", 0)
            
            On Error GoTo vbError
            Call objHT.CloseTransactionAll_Err
            Call objHT.OpenTransactionAll
            erTranAll = False
            
            'Call SaveDataCust
            Call SaveDataBooking(Trim(noPol))
            Call savePart(text(1))
            If barisSublet > 1 Then
                Call saveSublet
            End If
            Call saveStall
            If chkRtn Then
                Call rtjFU(text(41))
            End If
            If isBookingOnline Then
                objHT.InsertData_Bengkel ("UPDATE BOOKINGONLINE SET FPROSES = 1,NOFAK = '" & Trim(text(1)) & "' WHERE ID = '" & IDBookingOnline & "'")
                
            End If
            'Call CetakForm

            If erTranAll = False Then
                Call objHT.CloseTransactionAll
                Call saveLogMenu(gSA, ComptName, gVersi, "frmBookingDMSNonDSS", "tblSimpan", 1)
                If gPreviewCetak = True Then text(0).SetFocus
                nJRate = 0
                'pengecekan apakah dari diag
                If isBookDiag Then
                    'jika booking dari diagnose
                    'lakukan proses pengupdatetan no booking ke dalam diagdta dan diagapproval
                    objO.InsertData ("UPDATE DIAGDTA SET NOFAK = '" & Trim(text(1)) & "' WHERE NO_DIAG = '" & no_approval & "'")
                    objO.InsertData ("UPDATE DIAGAPPROVAL SET NOFAK = '" & Trim(text(1)) & "' WHERE NO_DIAG = '" & no_approval & "'")
                End If
                Call CetakForm1
            Else
                Call objHT.CloseTransactionAll_Err
                Call saveLogMenu(gSA, ComptName, gVersi, "frmBookingDMSNonDSS", "tblSimpan", 2)
                MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
            End If
            simpan = True
        Else
            MsgBox "Masukkan " + pesan + "!"
        End If
    Else
        MsgBox "Masukkan No. Polisi !", vbOKOnly, "Perhatian"
    End If
    Exit Sub
    
vbError:
    If erTranAll = True Then
        err.Raise err.Description
        Call objHT.CloseTransactionAll_Err
        Call saveLogMenu(gSA, ComptName, gVersi, "frmBookingDMSNonDSS", "tblSimpan", 2)
        MsgBox "Proses gagal disimpan." & vbCr & "Ulangi penyimpanan ", vbInformation, "Gagal Simpan"
    End If
    err.Raise err.Number, err.Source, err.Description
    Exit Sub
End Sub

Private Sub CetakForm1()
Set rpt = New RptBooking2
    
    If MsgBox("Mencetak Booking ?", vbYesNo, "Konfirmasi") = vbYes Then
        Dim sHlp As Object
        Dim sHlp1 As Object
        Dim sHlp2 As Object
        Dim sHlp3 As Object
        Dim sHlp4 As Object
        Dim rRs As New ADODB.Recordset
        Dim crxSubReportPart As CRAXDRT.Report
        Dim crxSubReportKeluhan As CRAXDRT.Report
        Dim crxSubReportJob As CRAXDRT.Report
        Dim crxSubReportCat As CRAXDRT.Report
        Dim crxSubReportReq As CRAXDDRT.Report
            
            Set rRs = objO.FetchData("SELECT CUSTOMER, CUST_ADRES, KOTA, TELPON, NO_POLISI, NO_ENGINE, VDS, NO_CHASSIS, MDL, MODEL, FMDL, COLOR, " _
                                  & "FS_KD_DEALER, FS_NM_DEALER, DDATE, CODE, DOTAM, NPWP, PPN, DISC, DISCT, DATE, FRONT, TIME, TIME_D, MECH, REQ1, REQ2, REQ3, REQ4, " _
                                  & "REQ5, REP1, REP2, REP3, REP4, REP5, REP6, REP7, REP8, RTE1, RTE2, RTE3, RTE4, RTE5, RTE6, RTE7, RTE8, TOT_RATE, PAID, NOFAK, " _
                                  & "REMARK1, SUGGEST, ESTDATE, ESTTIME, DELIV, FS_TH_MOBIL, FS_LOG_ADD, FS_LOG_DEL, FD_TGL_DEL, FS_RC, FS_NM_HOBI, DATE_BOOK, " _
                                  & "FS_KD_PLG, FS_KD_WIL, GSM, CDMA, FLAGSVC, TOTPART, TOTLABOUR, TOTSUBLET, TOTOLI, REKOM, COND1, COND2, COND3, COND4, COND5, " _
                                  & "COND6, THS, FN_RETURN, KARPLG, ALASAN, CATATAN1, CATATAN2, CATATAN3, CATATAN4, KODETIPE, CMPL1, CMPL2, CMPL3, CMPL4, CMPL5, AMBIL, DATESERAH, TIMESERAH FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & Trim(text(1)) & "'")
            rpt.Database.SetDataSource rRs
            
            sSQL = "select NOFAK, KODEITEM, NAMABARANG, QTY, NAMABYSA FROM BOOK_PART WITH (NOLOCK) WHERE NOFAK = '" & Trim(text(1)) & "'"
            Set sHlp = objO.FetchData(sSQL)
            Set crxSubReportPart = rpt.Subreport1.OpenSubreport
            crxSubReportPart.Database.SetDataSource sHlp
            
            sSQL = "select KELUHAN FROM repoBook('" & Trim(text(1)) & "') WHERE KELUHAN <> '' ORDER BY IDB"
            Set sHlp1 = objO.FetchData(sSQL)
            Set crxSubReportKeluhan = rpt.Subreport4.OpenSubreport
            crxSubReportKeluhan.Database.SetDataSource sHlp1
            
            sSQL = "select JOBORDER FROM repoBook('" & Trim(text(1)) & "') WHERE JOBORDER <> '' UNION SELECT NAMABARANG FROM BOOK_SUBLET WHERE NOFAK = '" & Trim(text(1)) & "'"
            Set sHlp2 = objO.FetchData(sSQL)
            Set crxSubReportJob = rpt.Subreport5.OpenSubreport
            crxSubReportJob.Database.SetDataSource sHlp2
            
            sSQL = "select CATATAN FROM repoBook('" & Trim(text(1)) & "') WHERE CATATAN <> '' ORDER BY IDB"
            Set sHlp3 = objO.FetchData(sSQL)
            Set crxSubReportCat = rpt.Subreport6.OpenSubreport
            crxSubReportCat.Database.SetDataSource sHlp3
            
            sSQL = "select REQUEST FROM repoBook('" & Trim(text(1)) & "') WHERE REQUEST <> '' ORDER BY IDB"
            Set sHlp4 = objO.FetchData(sSQL)
            Set crxSubReportReq = rpt.Subreport2.OpenSubreport
            crxSubReportReq.Database.SetDataSource sHlp4
            
            rpt.txtPT.SetText Trim(gPerush)
            rpt.txtPT2.SetText Trim(gPerush)
            
            Dim rsSearch As ADODB.Recordset
            Set rsSearch = objO.FetchData("SELECT NOSTALL FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & Trim(noPol) & " |%' AND FLAG = ''")
            If Not rsSearch.EOF Then
                rpt.txtStall.SetText Trim(rsSearch("NOSTALL"))
            End If
             
            
            With CRViewer1
                .ReportSource = rpt
                .ViewReport
                .Zoom 100
                .Visible = True
                .Top = 0
                .Left = asxPowerBanner1.Left + asxPowerBanner1.Width
                .Height = asxPowerBanner1.Height + 200
                .Width = Me.Width - .Left
            End With
            cmdTutup.Visible = True
            cmdTutup.Left = CRViewer1.Left + 20
            cmdTutup.Top = 435
            
            cmdSettingPrinter.Top = cmdTutup.Top
            cmdSettingPrinter.Left = cmdTutup.Left + cmdTutup.Width + 50
            cmdSettingPrinter.Visible = True
            Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub tblSimpan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        text(0).SetFocus
    End If
End Sub

Private Sub Text_Change(Index As Integer)
    If Index = 52 Or Index = 58 Or Index = 70 Or Index = 46 Or Index = 47 Or Index = 48 Or Index = 49 Or Index = 50 Or Index = 51 Or Index = 57 Or Index = 1 Then
        If text(0) <> "" Then
            Select Case Index
                Case 52
                    If text(Index) = "" Then
                        text(Index) = 0
                    ElseIf CDbl(text(Index)) >= 0 Then
                        text(Index) = Format(text(Index), "#0,#")
                        text(Index).SelStart = Len(text(Index))
                        text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#0,#")
                    End If
                Case 58
                    If text(58) <> ">>:>>" And text(23) <> ">>->>->>>>" Then
                        nWKT = hitungWaktu(CDbl(nJRate), text(58), text(23))
                        text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
                        text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
                        text(84) = Trim(text(59))
                        text(83) = Trim(text(45))
                        Call hitungEstimasiSelesai
                    End If
                Case 70
                    baris = CDbl(gBantuan)
                    barispart = CDbl(gBantuan2)
                    
                Case 46
                    Me.text(54).Visible = IIf(Me.text(46) = "LAINNYA", True, False)
                Case 47
                    Me.text(65).Visible = IIf(Me.text(47) = "LAINNYA", True, False)
                Case 48
                    Me.text(66).Visible = IIf(Me.text(48) = "LAINNYA", True, False)
                Case 49
                    Me.text(67).Visible = IIf(Me.text(49) = "LAINNYA", True, False)
                Case 50
                    Me.text(68).Visible = IIf(Me.text(50) = "LAINNYA", True, False)
                Case 51
                    Me.text(69).Visible = IIf(Me.text(51) = "LAINNYA", True, False)
                Case 57
                    If nX Then
                        Call text_Validate(Index, False)
                    End If
                Case 1
                    If Trim(text(1)) <> "" Then
                        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then
                            Dim rsPM As ADODB.Recordset
                            Set rsPM = objO.FetchDataACC("select Top 1 Status from SOORD WITH (NOLOCK) Where NoSPK = '" & Trim(text(1)) & "'")
                            If rsPM.RecordCount > 0 Then
                                If rsPM.Fields(0).value <> "" And rsPM.Fields(0).value <> gSA Then
                                    MsgBox "Sparepart sedang diproses oleh Partman (" & rsPM.Fields(0).value & ") !", vbInformation + vbOKOnly, "Reserve Sparepart"
                                    Call cmdBersih_Click
                                    Exit Sub
                                Else
                                    Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NoSPK = '" & Trim(text(1)) & "'")
                                End If
                            Else
                                Call objO.UpdateDataACC("Update SOORD Set Status = '" & gSA & "' Where NoSPK = '" & Trim(text(1)) & "'")
                            End If
                        End If
                    Else
                        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then Call objO.UpdateDataACC("Update SOORD Set Status = '' Where Status ='" & gSA & "'")
                    End If
            End Select
        End If
        
        'lakukan pengecekan booking
        
    End If
End Sub

Private Sub text_GotFocus(Index As Integer)
    Dim DiTab As Boolean
    DiTab = False
    
    If text(Index).BackColor = &HC0E0FF Then
        DiTab = True
    Else
        If gMainKeyCari <> "" Then
            text(Index).Mask = ""
            text(Index) = gMainKeyCari
            If Index = 0 Then
                Dim nDpn1 As String
                Dim nTgh1 As String
                Dim nBlk1 As String
                nDpn1 = ""
                nTgh1 = ""
                nBlk1 = ""
                Call pecahNoPol(gMainKeyCari, nDpn1, nTgh1, nBlk1)
                text(0) = Trim(nDpn1)
                text(56) = Trim(nTgh1)
                text(57) = Trim(nBlk1)
            End If
            gMainKeyCari = ""
            DiTab = True
        End If
    End If
    Select Case Index
        Case 23, 46
            VScroll1.value = 0
        Case 32
            VScroll1.value = 4900
        Case 60
            VScroll1.value = 1000
        Case 63, 37
            VScroll1.value = 2000
    End Select
    If DiTab Then
        SendKeys ("{TAB}")
    Else
        SendKeys ("{HOME}+{END}")
    End If
End Sub

Private Sub SaveDataBooking(NoPolisi As String)
    
    ' *** MENYIMPAN DATA BOOKING ***
    If text(1) = "" Then
        Call isiNoBook
    End If
    
    Call ClearsSql
    gsSql1(0) = "NO_POLISI"
    gsSql2(0) = noPol
    gsSql1(1) = "NOFAK"
    gsSql2(1) = text(1)
    gsSql1(2) = "CUSTOMER"
    gsSql2(2) = text(3)
    gsSql1(3) = "CUST_ADRES"
    gsSql2(3) = text(4)
    gsSql1(4) = "KOTA"
    gsSql2(4) = text(5)
    gsSql1(5) = "GSM"
    gsSql2(5) = text(6)
    gsSql1(6) = "CDMA"
    gsSql2(6) = text(7)
    gsSql1(7) = "TELPON"
    gsSql2(7) = text(8)
    gsSql1(8) = "FS_TH_MOBIL"
    gsSql2(8) = text(9)
    gsSql1(9) = "MDL"
    gsSql2(9) = text(10)
    gsSql1(10) = "MODEL"
    gsSql2(10) = text(11)
    gsSql1(11) = "FMDL"
    gsSql2(11) = text(12)
    gsSql1(12) = "COLOR"
    gsSql2(12) = text(13)
    gsSql1(13) = "FS_KD_DEALER"
    gsSql2(13) = text(14)
    gsSql1(14) = "FS_NM_DEALER"
    gsSql2(14) = text(15)
    gsSql1(15) = "DOTAM"
    gsSql2(15) = Replace(text(16), ">", "")
    gsSql1(16) = "DDATE"
    gsSql2(16) = Replace(text(17), ">", "")
    
    gsSql1(17) = "[DATE]"
    gsSql2(17) = Format(TglYYYYMMDD(text(23)), "mm/dd/yyyy")
    gsSql1(18) = "TIME"
    gsSql2(18) = text(24)
    gsSql1(19) = "PAID"
    gsSql2(19) = text(25)
    
    gsSql1(20) = "REQ1"
    gsSql2(20) = text(27)
    gsSql1(21) = "REQ2"
    gsSql2(21) = text(28)
    gsSql1(22) = "REQ3"
    gsSql2(22) = text(29)
    gsSql1(23) = "REQ4"
    gsSql2(23) = text(30)
    gsSql1(24) = "REQ5"
    gsSql2(24) = text(31)

    Dim hlp As String
    Dim hlp2 As String
    Dim nTotRt As Double
    Dim nTotAct As Integer
    Dim rsLama As ADODB.Recordset
    Dim rsAct As ADODB.Recordset
    
    ' HAPUS JAM MEKANIK LAMA
    Set rsLama = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & Trim(text(1)) & "'")
    If Not rsLama.EOF Then
        For i = 1 To 7
            hlp = "REP" & i
            If rsLama(hlp) <> "" Then
                Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(rsLama(hlp)) & "' AND MDL = '" & Trim(rsLama("MDL")) & "'")
                If Not rsAct.EOF Then
                    nTotAct = nTotAct + CDbl(rsAct("RATE"))
                End If
                rsAct.Close
            End If
        Next
        Call createMekanik
        'Call decresRestTime(rsLama("MECH"), CDbl(nTotAct), 1, Format(rsLama("DATE"), "MM/DD/YYYY"))
    End If
    rsLama.Close
    
    nTotRt = 0
    nTotAct = 0
    For i = 1 To 7
        hlp = "REP" & i
        hlp2 = "RTE" & i
        gsSql1(24 + i) = hlp
        gsSql2(24 + i) = DGrid.TextMatrix(i, 1)
        gsSql1(31 + i) = hlp2
        gsSql2(31 + i) = Replace(IIf(DGrid.TextMatrix(i, 2) <> "", DGrid.TextMatrix(i, 2), 0), ",", ".")
        nTotRt = nTotRt + IIf(DGrid.TextMatrix(i, 2) <> "", DGrid.TextMatrix(i, 2), 0)
        If DGrid.TextMatrix(i, 1) <> "" Then
            Set rsAct = objO.FetchData("SELECT ISNULL(RATE,0)*60 AS RATE FROM ACTRATE WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(DGrid.TextMatrix(i, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
            If Not rsAct.EOF Then
               nTotAct = nTotAct + CDbl(rsAct("RATE"))
            End If
            rsAct.Close
        End If
    Next
    Call countRestTime(Trim(text(64)), CDbl(nTotAct), 1, Format(text(23), "MM/DD/YYYY"))
    
    If text(18) = ">>->>->>>>" Then
        text(18) = TglDDMMYYYY(Format(isnow, "DD/MM/YYYY"))
        text(19) = Format(isnow, "HH:NN")
    End If
    
    gsSql1(39) = "DATE_BOOK"
    gsSql2(39) = Format(TglYYYYMMDD(text(18)), "mm/dd/yyyy")
    gsSql1(40) = "TIME_D"
    gsSql2(40) = text(19)
    
    gsSql1(41) = "REMARK1"
    gsSql2(41) = text(37)
    gsSql1(42) = "SUGGEST"
    gsSql2(42) = text(38)
    gsSql1(43) = "DELIV"
    gsSql2(43) = "BOK"
    
    gsSql1(44) = "FS_LOG_ADD"
    gsSql2(44) = gSA
    
    gsSql1(45) = "FLAGSVC"
    If opSvc(0).value Then
        gsSql2(45) = "TUNGGU"
    ElseIf opSvc(1).value Then
        gsSql2(45) = "TINGGAL"
    End If
    
    gsSql1(46) = "TOTPART"
    gsSql2(46) = CDbl(text(40))
    gsSql1(47) = "TOTLABOUR"
    gsSql2(47) = CDbl(text(39))
    gsSql1(48) = "TOTSUBLET"
    gsSql2(48) = CDbl(text(52))
    gsSql1(49) = "TOTOLI"
    gsSql2(49) = CDbl(text(55))
    gsSql1(50) = "FS_KD_PLG"
    gsSql2(50) = text(2)
    gsSql1(51) = "NO_CHASSIS"
    gsSql2(51) = gNoRangka
    
    gsSql1(52) = "ESTTIME"
    gsSql2(52) = text(45)
    gsSql1(53) = "TOT_RATE"
    gsSql2(53) = Replace(CDbl(nTotRt), ",", ".")
    
    gsSql1(54) = "REKOM"
    gsSql2(54) = IIf(cmdNo.Caption = "No &Rekom", "Y", "N")
    
    For i = 1 To 6
        hlp = "COND" & i
        gsSql1(54 + i) = hlp
        gsSql2(54 + i) = IIf(text(45 + i) = "LAINNYA", "L| " & Trim(bKondisi(i - 1)), text(45 + i))
    Next
    
    gsSql1(61) = "THS"
    gsSql2(61) = IIf(chkThs = Checked, "1", "0")
    gsSql1(62) = "FN_RETURN"
    gsSql2(62) = IIf(chkRtn = Checked, "1", "0")
    gsSql1(63) = "CATATAN1"
    gsSql2(63) = text(60)
    gsSql1(64) = "CATATAN2"
    gsSql2(64) = text(61)
    gsSql1(65) = "CATATAN3"
    gsSql2(65) = text(62)
    gsSql1(66) = "CATATAN4"
    gsSql2(66) = text(63)
    gsSql1(67) = "ESTDATE"
    gsSql2(67) = Format(text(59), "MM/DD/YYYY")
    gsSql1(68) = "MECH"
    gsSql2(68) = nMech
    gsSql1(69) = "KODETIPE"
    gsSql2(69) = Trim(text(71))
    gsSql1(70) = "AMBIL"
    gsSql2(70) = IIf(chkAmbil = Checked, "1", "0")
    gsSql1(71) = "CODE"
    gsSql2(71) = IIf(chEM = Checked, "1", "0")
    gsSql1(72) = "KD_K"
    gsSql2(72) = IIf(chSBNP = Checked, "1", "0")
    gsSql1(73) = "KIRIM"
    gsSql2(73) = IIf(chkKirim = Checked, "1", "0")
    gsSql1(74) = "FS_KD_WIL"
    gsSql2(74) = Trim(text(44))
    gsSql1(75) = "DATESERAH"
    gsSql2(75) = Format(text(84), "MM/DD/YYYY")
    gsSql1(76) = "TIMESERAH"
    gsSql2(76) = Trim(text(83))
    
    For i = 1 To 5
        gsSql1(76 + i) = "CMPL" & i
        'gsSql2(76 + i) = Trim(text(71 + i))
        gsSql2(76 + i) = Trim(Text1(-1 + i))
    Next
    
    If Trim(text(1)) = "" Then
        gsSql1(82) = "FRONT"
        gsSql2(82) = text(26)
    End If
    
    Set oRs = objO.FetchData("Select * from BOOK WITH (NOLOCK) where nofak = '" & text(1) & "'")
    If Not oRs.EOF Then
        If oRs("DELIV") <> "" Then
            gsSql1(43) = "DELIV"
            gsSql2(43) = Trim(oRs("DELIV"))
        End If
        
        If DateDiff("d", Format(oRs("DATE"), "MM/DD/YYYY"), Format(TglYYYYMMDD(text(23)), "mm/dd/yyyy")) <> 0 Then
            'If oRs("TIME") <> Trim(text(24)) Then
                If Left(oRs("DELIV"), 3) = "RES" Then
                    gsSql1(43) = "DELIV"
                    gsSql2(43) = "RES" & CDbl(Right(Trim(oRs("DELIV")), 1)) + 1
                Else
                    gsSql1(43) = "DELIV"
                    gsSql2(43) = "RES1"
                End If
            'End If
            Call objHT.DeleteData_Bengkel("DELETE STALL WHERE KETERANGAN LIKE '%|" & Trim(text(1)) & "|%' AND TANGGAL = '" & Format(oRs("DATE"), "MM/DD/YYYY") & "'")
        Else
            If oRs("TIME") <> Trim(text(24)) Then
                If Left(oRs("DELIV"), 3) = "RES" Then
                    gsSql1(43) = "DELIV"
                    gsSql2(43) = "RES" & CDbl(Right(Trim(oRs("DELIV")), 1)) + 1
                Else
                    gsSql1(43) = "DELIV"
                    gsSql2(43) = "RES1"
                End If
            End If
            Call objHT.DeleteData_Bengkel("DELETE STALL WHERE KETERANGAN LIKE '%|" & Trim(text(1)) & "|%' AND TANGGAL = '" & Format(oRs("DATE"), "MM/DD/YYYY") & "'")
        End If
        sSQL = GensSql("BOOK", True) + " where NOFAK = '" & text(1) & "'"
        Call updateAntrian(noPol, text(12), text(3), Format(text(23), "DD/MM/YYYY"), text(24), 2, Trim(text(1)))
    Else
        sSQL = GensSql("BOOK", False)
        noPol = Trim(text(0)) & IIf(Trim(text(0)) <> "", "-", "") & Trim(text(56)) & IIf(Trim(text(57)) <> "", "-", "") & Trim(text(57))
        Call updateAntrian(noPol, text(12), text(3), Format(text(23), "DD/MM/YYYY"), text(24), 1, Trim(text(1)))
    End If
    Call objHT.UpdateData_Bengkel(sSQL)
    Call objHT.UpdateData_Bengkel("UPDATE BOOK SET FS_RC ='' where NO_POLISI ='" & NoPolisi & "' and DELIV='BOK'")
    
    Set oRs = objO.FetchData("SELECT * FROM T_REGULAR WITH (NOLOCK)")
    If Not oRs.BOF Then
        Do While Not oRs.EOF
            sSQL = "UPDATE BOOK SET FS_RC ='" & Trim(oRs("FS_RC")) & "' where NO_POLISI ='" & NoPolisi & "' and DELIV='BOK'" _
                 & " AND REP1+ REP2+ REP3+ REP4+ REP5+ REP6+ REP7 like '% " & Trim(oRs("FS_KETERANGAN")) & "[^0]%'"
            Call objHT.UpdateData_Bengkel(sSQL)
            oRs.MoveNext
        Loop
    End If
    If oRs.State <> adStateClosed Then oRs.Close
End Sub

Private Sub PanggilDataBooking(NoPolisi As String, sDeliv As String)
    Dim j As Integer
    Dim rsHlp As ADODB.Recordset
    Dim rsHlp2 As ADODB.Recordset
    
    sSQL = "Select * from BOOK WITH (NOLOCK) where NO_POLISI = '" & NoPolisi & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE' ORDER BY DELIV, [date] desc"
    Set rsHlp2 = objO.FetchData(sSQL)
    If Not rsHlp2.BOF Then
        Set rsHlp = objO.FetchData("SELECT no_chassis FROM jobDta WITH (NOLOCK) where no_chassis = '" & rsHlp2("NO_CHASSIS") & "' AND DELIV <> 'BTL' AND RPR_COND <> 'K' AND TPSVC <> 'W' ORDER BY no_WO desc")
        If Not rsHlp.EOF Then
            If MsgBox("Transaksi Lama belum closing !!!" & vbCr & "Tetap akan melanjutkan ?", vbYesNo, "Closing") = vbYes Then
                GoTo lanjut
            Else
                GoTo LEWAT
            End If
        Else
            GoTo lanjut
        End If
lanjut:
                
                text(2) = Nul(rsHlp2("FS_KD_PLG"))
                text(44) = Nul(rsHlp2("FS_KD_WIL"))
                text(3) = Nul(rsHlp2("CUSTOMER"))
                text(4) = Nul(rsHlp2("CUST_ADRES"))
                text(5) = Nul(rsHlp2("KOTA"))
                text(6) = Nul(rsHlp2("GSM"))
                text(7) = Nul(rsHlp2("CDMA"))
                text(8) = Nul(rsHlp2("TELPON"))
                text(9) = Nul(rsHlp2("FS_TH_MOBIL"))
                text(10) = rsHlp2("MDL")
                text(11) = rsHlp2("MODEL")
                    Call CexChassis_Discon(text(11))
                    If isOnCarClass = False Then
                        MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
                    End If
                text(12) = rsHlp2("FMDL")
                text(13) = rsHlp2("COLOR")
                text(14) = rsHlp2("FS_KD_DEALER")
                text(15) = rsHlp2("FS_NM_DEALER")
                text(16).Mask = ""
                text(16) = TglDDMMYYYY(IIf(rsHlp2("DOTAM") = "", "01/01/1900", rsHlp2("DOTAM")))
                text(17).Mask = ""
                text(17) = TglDDMMYYYY(IIf(rsHlp2("DDATE") = "", "01/01/1900", rsHlp2("DDATE")))
                text(41) = rsHlp2("NO_CHASSIS")
                gNoRangka = rsHlp2("NO_CHASSIS")
                cmdStall.Tag = ""
                cmdStall.BackColor = &H8000000F
                
                If Trim(rsHlp2("DELIV")) = "BOK" Or Left(Trim(rsHlp2("DELIV")), 3) = "RES" Then
                    text(1) = rsHlp2("NOFAK")
                    text(18) = TglDDMMYYYY(rsHlp2("DATE_BOOK"))
                    text(19) = Format(rsHlp2("TIME_D"), "HH:NN")
                    text(23) = TglDDMMYYYY(rsHlp2("DATE"))
                    text(24) = Format(rsHlp2("TIME"), "HH:NN")
                    
                    text(23).Enabled = True
                    text(24).Enabled = True
                    If Trim(rsHlp2("DELIV")) = "RES3" Then
                        MsgBox "Sudah di Reschedule 3x, tidak dapat di reschedule kembali." & vbCr & "Batalkan Booking, buat booking baru ."
                        text(23).Enabled = False
                        text(24).Enabled = False
                    End If
                    text(25) = rsHlp2("PAID")
                    text(26) = gSA
                    text(27) = rsHlp2("REQ1")
                    text(28) = rsHlp2("REQ2")
                    text(29) = rsHlp2("REQ3")
                    text(30) = rsHlp2("REQ4")
                    text(31) = rsHlp2("REQ5")
                    
                    ' ISI KELUHAN
                    Dim hlp As String
                    For i = 1 To 5
                        hlp = "CMPL" & i
                        Text1(-1 + i) = rsHlp2(hlp)
                    Next
                    
                    text(39) = NulA(Format(rsHlp2("TOTLABOUR"), "#0,#"))
                    text(40) = NulA(Format(rsHlp2("TOTPART"), "#0,#"))
                    text(52) = NulA(Format(rsHlp2("TOTSUBLET"), "#0,#"))
                    text(55) = NulA(Format(rsHlp2("TOTOLI"), "#0,#"))
                    text(53) = Format(CDbl(text(39) + CDbl(text(40)) + CDbl(text(52))) + CDbl(text(55)), "#0,#")
                    
                    text(37) = Trim(rsHlp2("REMARK1"))
                    text(38) = Trim(rsHlp2("SUGGEST"))
                    text(75) = Trim(rsHlp2("KARPLG"))
                    text(60) = Trim(rsHlp2("CATATAN1"))
                    text(61) = Trim(rsHlp2("CATATAN2"))
                    text(62) = Trim(rsHlp2("CATATAN3"))
                    text(63) = Trim(rsHlp2("CATATAN4"))
                    
                    ' TAMPIL CONDITION
                    Dim kCond As String
                    Dim kBant As String
                    Dim rsTmp As ADODB.Recordset
                    Set rsTmp = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & rsHlp2("NOFAK") & "' ")
                    If Not rsTmp.EOF Then
                        For i = 1 To 6
                            kBant = ""
                            kCond = ""
                            hlp = "COND" & i
                            If rsTmp(hlp) <> "" Then
                                If Left(rsTmp(hlp), 3) = "L| " Then
                                    kBant = Mid(rsTmp(hlp), 4, Len(rsTmp(hlp)) - 3)
                                    kCond = "LAINNYA"
                                Else
                                    kCond = Trim(rsTmp(hlp))
                                    kBant = ""
                                End If
                            End If
                            text(45 + i) = kCond
                            bKondisi(i - 1) = kBant
                        Next
                    End If
                    text(54).Visible = False
                    rsTmp.Close
                        
                    chkRtn = IIf(rsHlp2("FN_RETURN") = "1", Checked, Unchecked)
                    chkThs = IIf(rsHlp2("THS") = "1", Checked, Unchecked)
                    chkAmbil = IIf(rsHlp2("AMBIL") = "1", Checked, Unchecked)
                    chkKirim = IIf(rsHlp2("KIRIM") = "1", Checked, Unchecked)
                    chEM = IIf(rsHlp2("CODE") = "1", Checked, Unchecked)
                    chSBNP = IIf(rsHlp2("KD_K") = "1", Checked, Unchecked)
                    
                    DGrid.TextMatrix(1, 1) = rsHlp2("REP1")
                    DGrid.TextMatrix(2, 1) = rsHlp2("REP2")
                    DGrid.TextMatrix(3, 1) = rsHlp2("REP3")
                    DGrid.TextMatrix(4, 1) = rsHlp2("REP4")
                    DGrid.TextMatrix(5, 1) = rsHlp2("REP5")
                    DGrid.TextMatrix(6, 1) = rsHlp2("REP6")
                    DGrid.TextMatrix(7, 1) = rsHlp2("REP7")
                    
                    DGrid.TextMatrix(1, 2) = IIf(DGrid.TextMatrix(1, 1) <> "", rsHlp2("RTE1"), "")
                    DGrid.TextMatrix(2, 2) = IIf(DGrid.TextMatrix(2, 1) <> "", rsHlp2("RTE2"), "")
                    DGrid.TextMatrix(3, 2) = IIf(DGrid.TextMatrix(3, 1) <> "", rsHlp2("RTE3"), "")
                    DGrid.TextMatrix(4, 2) = IIf(DGrid.TextMatrix(4, 1) <> "", rsHlp2("RTE4"), "")
                    DGrid.TextMatrix(5, 2) = IIf(DGrid.TextMatrix(5, 1) <> "", rsHlp2("RTE5"), "")
                    DGrid.TextMatrix(6, 2) = IIf(DGrid.TextMatrix(6, 1) <> "", rsHlp2("RTE6"), "")
                    DGrid.TextMatrix(7, 2) = IIf(DGrid.TextMatrix(7, 1) <> "", rsHlp2("RTE7"), "")
                    
                    baris = 1
                    nJRate = 0
                    For j = 1 To 7
                        If Trim(DGrid.TextMatrix(j, 1)) <> "" Then
                            Set rsTemp = objO.FetchData("SELECT rate,job FROM HO_flate WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(DGrid.TextMatrix(j, 1)) & "' AND MDL = '" & rsHlp2("MDL") & "' and FS_LOG_DEL = ''")
                            If Not rsTemp.EOF Then
                                DGrid.TextMatrix(j, 2) = rsTemp("RATE")
                                DGrid.TextMatrix(j, 3) = ""
                                DGrid.TextMatrix(j, 4) = rsTemp("JOB")
                            Else
                                DGrid.TextMatrix(j, 3) = True
                            End If
                            nJRate = nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(DGrid.TextMatrix(j, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                            baris = baris + 1
                        End If
                    Next
                    
                    Set rsTemp = objO.FetchData("SELECT * FROM BOOK_SUBLET WITH (NOLOCK) WHERE NOFAK = '" & rsHlp2("NOFAK") & "'")
                    If Not rsTemp.EOF Then
                        i = 1
                        barisSublet = 1
                        Do While Not rsTemp.EOF
                            subletGrid.TextMatrix(i, 1) = rsTemp("NAMABARANG")
                            subletGrid.TextMatrix(i, 2) = CDbl(rsTemp("QTY"))
                            subletGrid.TextMatrix(i, 3) = rsTemp("KODESATUAN")
                            subletGrid.TextMatrix(i, 4) = Format(CDbl(rsTemp("HARGA")), "#0,#")
                            subletGrid.TextMatrix(i, 5) = rsTemp("FLAG")
                            subletGrid.TextMatrix(i, 6) = rsTemp("ESTWAKTU")
                            i = i + 1
                            rsTemp.MoveNext
                        Loop
                    End If
                    If rsTemp.State <> adStateClosed Then rsTemp.Close
                    'CEX 280514
                     If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                     Set rsTemp = objO.FetchDataACC("select NOSPK as NOFAK, KodeItem, Keterangan as namaBarang, SatuanNS as KodeSatuan, Qty, '' as FLAG, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID, '' as KDOLI from SOORD WITH (NOLOCK) Where NoSPK = '" & rsHlp2("NOFAK") & "' and (isnull(BO,'null') = 'null' or BO <> 'CANCEL')")
                     Else
                     Set rsTemp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM book_part A WITH (NOLOCK) LEFT JOIN T_OLI B WITH (NOLOCK) ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & rsHlp2("NOFAK") & "'")
                     End If
                   ' Set rsTemp = objO.FetchDataACC("select NOSPK as NOFAK, KodeItem, Keterangan as namaBarang, SatuanNS as KodeSatuan, Qty, '' as FLAG, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID, '' as KDOLI from SOORD Where NoSPK = '" & rsHlp2("NOFAK") & "' and BO <> 'CANCEL'")
                   ' If rsTemp.RecordCount = 0 Then
                   '     Set rsTemp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM book_part A LEFT JOIN T_OLI B ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & rsHlp2("NOFAK") & "'")
                   ' End If
                    If Not rsTemp.EOF Then
                        i = 1
                        barispart = 1
                        Do While Not rsTemp.EOF
                            PartGrid.TextMatrix(i, 1) = rsTemp("KODEITEM")
                            PartGrid.TextMatrix(i, 2) = rsTemp("NAMABARANG")
                            PartGrid.TextMatrix(i, 3) = rsTemp("KODESATUAN")
                            PartGrid.TextMatrix(i, 4) = rsTemp("QTY")
                            PartGrid.TextMatrix(i, 6) = IIf(rsTemp("KDOLI") <> "", True, False)
                            PartGrid.TextMatrix(i, 7) = rsTemp("NAMABYSA")
                            PartGrid.TextMatrix(i, 8) = Val(NulA(rsTemp("nID")))
                            Dim rshrgPart As ADODB.Recordset
                                Set rshrgPart = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM LIKE '%" & Trim(rsTemp("KODEITEM")) & "%'")
                                If Not rshrgPart.EOF Then
                                    PartGrid.TextMatrix(i, 5) = rshrgPart("HARGAJUALA")
                                End If
                                rshrgPart.Close
                                Set rshrgPart = Nothing
                            barispart = barispart + 1
                            i = i + 1
                            rsTemp.MoveNext
                        Loop
                        NewId = (objO.FetchData_DoubleACC("select top 1 Nomor from SOORD WITH (NOLOCK) Where NOSPK = '" & rsHlp2("NOFAK") & "' Order By Nomor DESC")) + 1
                    End If
                    
                    ' ISI STALL
                    Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                                 & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%') AND " _
                                 & "tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%' ORDER BY noStall, Tanggal, JamMulai")
                    cmdStall.Tag = ""
                    Set rsHlp = objO.FetchData("SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%'")
                    If Not rsHlp.EOF Then
                        text(58) = Trim(rsHlp("JAMMULAI"))
                        nMech = Trim(rsHlp("MECH"))
                    Else
                        text(58) = Trim(text(24))
                        nMech = Trim(rsHlp2("MECH"))
                    End If
                    rsHlp.Close
                    text(64) = nMech
                    
                    If text(58) <> ">>:>>" And text(23) <> ">>->>->>>>" Then
                        nWKT = hitungWaktu(CDbl(nJRate), text(58), text(23))
                        text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
                        text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
                        text(84) = TglDDMMYYYY(rsHlp2("DATESERAH"))
                        text(83) = Trim(rsHlp2("TIMESERAH"))
                        Call hitungEstimasiSelesai
                    End If
                Else
                    If PanggilDataCust(NoPolisi) Then
                        cmdStall.Tag = "belum"
                    End If
                End If
                lblEdit.Visible = True
    If text(1) = "" Then Call cmdBersih_Click
            'End If
LEWAT:
    
    End If
    sSQL = "select * from MODEL WITH (NOLOCK) where MDL ='" & text(10) & "'"
    Set rsHlp2 = objO.FetchData(sSQL)
    If rsHlp2.BOF Then
        gGroupModel = ""
    Else
        gGroupModel = rsHlp2("FS_GRP_MDL")
    End If
    If rsHlp2.State <> adStateClosed Then rsHlp2.Close
End Sub

Private Sub PanggilDataBooking2(NoPolisi As String, NoFak As String, sDeliv As String)
    Dim j As Integer
    Dim rsHlp As ADODB.Recordset
    Dim rsHlp2 As ADODB.Recordset
    
    sSQL = "Select * from BOOK WITH (NOLOCK) where NO_POLISI = '" & NoPolisi & "' AND nofak = '" & NoFak & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE' ORDER BY DELIV, [date] desc"
    Set rsHlp2 = objO.FetchData(sSQL)
    If Not rsHlp2.BOF Then
        Set rsHlp = objO.FetchData("SELECT no_chassis FROM jobDta WITH (NOLOCK) where no_chassis = '" & rsHlp2("NO_CHASSIS") & "' AND DELIV <> 'BTL' AND RPR_COND <> 'K' AND TPSVC <> 'W' ORDER BY no_WO desc")
        If Not rsHlp.EOF Then
            If MsgBox("Transaksi Lama belum closing !!!" & vbCr & "Tetap akan melanjutkan ?", vbYesNo, "Closing") = vbYes Then
                GoTo lanjut
            Else
                GoTo LEWAT
            End If
        Else
            GoTo lanjut
        End If
lanjut:
                text(2) = Nul(rsHlp2("FS_KD_PLG"))
                text(44) = Nul(rsHlp2("FS_KD_WIL"))
                text(3) = Nul(rsHlp2("CUSTOMER"))
                text(4) = Nul(rsHlp2("CUST_ADRES"))
                text(5) = Nul(rsHlp2("KOTA"))
                text(6) = Nul(rsHlp2("GSM"))
                text(7) = Nul(rsHlp2("CDMA"))
                text(8) = Nul(rsHlp2("TELPON"))
                text(9) = Nul(rsHlp2("FS_TH_MOBIL"))
                text(10) = rsHlp2("MDL")
                text(11) = rsHlp2("MODEL")
                    Call CexChassis_Discon(text(11))
                    If isOnCarClass = False Then
                        MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
                    End If
                text(12) = rsHlp2("FMDL")
                text(13) = rsHlp2("COLOR")
                text(14) = rsHlp2("FS_KD_DEALER")
                text(15) = rsHlp2("FS_NM_DEALER")
                text(16).Mask = ""
                text(16) = TglDDMMYYYY(IIf(rsHlp2("DOTAM") = "", "01/01/1900", rsHlp2("DOTAM")))
                text(17).Mask = ""
                text(17) = TglDDMMYYYY(IIf(rsHlp2("DDATE") = "", "01/01/1900", rsHlp2("DDATE")))
                text(41) = rsHlp2("NO_CHASSIS")
                gNoRangka = rsHlp2("NO_CHASSIS")
                cmdStall.Tag = ""
                cmdStall.BackColor = &H8000000F
                
                If Trim(rsHlp2("DELIV")) = "BOK" Or Left(Trim(rsHlp2("DELIV")), 3) = "RES" Then
                    text(1) = rsHlp2("NOFAK")
                    text(18) = TglDDMMYYYY(rsHlp2("DATE_BOOK"))
                    text(19) = Format(rsHlp2("TIME_D"), "HH:NN")
                    text(23) = TglDDMMYYYY(rsHlp2("DATE"))
                    text(24) = Format(rsHlp2("TIME"), "HH:NN")
                    
                    text(23).Enabled = True
                    text(24).Enabled = True
                    If Trim(rsHlp2("DELIV")) = "RES3" Then
                        MsgBox "Sudah di Reschedule 3x, tidak dapat di reschedule kembali." & vbCr & "Batalkan Booking, buat booking baru ."
                        text(23).Enabled = False
                        text(24).Enabled = False
                    End If
                    text(25) = rsHlp2("PAID")
                    text(26) = gSA
                    text(27) = rsHlp2("REQ1")
                    text(28) = rsHlp2("REQ2")
                    text(29) = rsHlp2("REQ3")
                    text(30) = rsHlp2("REQ4")
                    text(31) = rsHlp2("REQ5")
                    
                    ' ISI KELUHAN
                    Dim hlp As String
                    For i = 1 To 5
                        hlp = "CMPL" & i
                        Text1(-1 + i) = rsHlp2(hlp)
                    Next
                    
                    text(39) = NulA(Format(rsHlp2("TOTLABOUR"), "#0,#"))
                    text(40) = NulA(Format(rsHlp2("TOTPART"), "#0,#"))
                    text(52) = NulA(Format(rsHlp2("TOTSUBLET"), "#0,#"))
                    text(55) = NulA(Format(rsHlp2("TOTOLI"), "#0,#"))
                    text(53) = Format(CDbl(text(39) + CDbl(text(40)) + CDbl(text(52))) + CDbl(text(55)), "#0,#")
                    
                    text(37) = Trim(rsHlp2("REMARK1"))
                    text(38) = Trim(rsHlp2("SUGGEST"))
                    text(75) = Trim(rsHlp2("KARPLG"))
                    text(60) = Trim(rsHlp2("CATATAN1"))
                    text(61) = Trim(rsHlp2("CATATAN2"))
                    text(62) = Trim(rsHlp2("CATATAN3"))
                    text(63) = Trim(rsHlp2("CATATAN4"))
                    
                    ' TAMPIL CONDITION
                    Dim kCond As String
                    Dim kBant As String
                    Dim rsTmp As ADODB.Recordset
                    Set rsTmp = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NOFAK = '" & rsHlp2("NOFAK") & "' ")
                    If Not rsTmp.EOF Then
                        For i = 1 To 6
                            kBant = ""
                            kCond = ""
                            hlp = "COND" & i
                            If rsTmp(hlp) <> "" Then
                                If Left(rsTmp(hlp), 3) = "L| " Then
                                    kBant = Mid(rsTmp(hlp), 4, Len(rsTmp(hlp)) - 3)
                                    kCond = "LAINNYA"
                                Else
                                    kCond = Trim(rsTmp(hlp))
                                    kBant = ""
                                End If
                            End If
                            text(45 + i) = kCond
                            bKondisi(i - 1) = kBant
                        Next
                    End If
                    text(54).Visible = False
                    rsTmp.Close
                        
                    chkRtn = IIf(rsHlp2("FN_RETURN") = "1", Checked, Unchecked)
                    chkThs = IIf(rsHlp2("THS") = "1", Checked, Unchecked)
                    chkAmbil = IIf(rsHlp2("AMBIL") = "1", Checked, Unchecked)
                    chkKirim = IIf(rsHlp2("KIRIM") = "1", Checked, Unchecked)
                    
                    
                    DGrid.TextMatrix(1, 1) = rsHlp2("REP1")
                    DGrid.TextMatrix(2, 1) = rsHlp2("REP2")
                    DGrid.TextMatrix(3, 1) = rsHlp2("REP3")
                    DGrid.TextMatrix(4, 1) = rsHlp2("REP4")
                    DGrid.TextMatrix(5, 1) = rsHlp2("REP5")
                    DGrid.TextMatrix(6, 1) = rsHlp2("REP6")
                    DGrid.TextMatrix(7, 1) = rsHlp2("REP7")
                    
                    DGrid.TextMatrix(1, 2) = IIf(DGrid.TextMatrix(1, 1) <> "", rsHlp2("RTE1"), "")
                    DGrid.TextMatrix(2, 2) = IIf(DGrid.TextMatrix(2, 1) <> "", rsHlp2("RTE2"), "")
                    DGrid.TextMatrix(3, 2) = IIf(DGrid.TextMatrix(3, 1) <> "", rsHlp2("RTE3"), "")
                    DGrid.TextMatrix(4, 2) = IIf(DGrid.TextMatrix(4, 1) <> "", rsHlp2("RTE4"), "")
                    DGrid.TextMatrix(5, 2) = IIf(DGrid.TextMatrix(5, 1) <> "", rsHlp2("RTE5"), "")
                    DGrid.TextMatrix(6, 2) = IIf(DGrid.TextMatrix(6, 1) <> "", rsHlp2("RTE6"), "")
                    DGrid.TextMatrix(7, 2) = IIf(DGrid.TextMatrix(7, 1) <> "", rsHlp2("RTE7"), "")
                    
                    baris = 1
                    nJRate = 0
                    For j = 1 To 7
                        If Trim(DGrid.TextMatrix(j, 1)) <> "" Then
                            Set rsTemp = objO.FetchData("SELECT rate,job FROM HO_flate WITH (NOLOCK) WHERE FS_ITEM = '" & Trim(DGrid.TextMatrix(j, 1)) & "' AND MDL = '" & rsHlp2("MDL") & "' and FS_LOG_DEL = ''")
                            If Not rsTemp.EOF Then
                                DGrid.TextMatrix(j, 2) = rsTemp("RATE")
                                DGrid.TextMatrix(j, 3) = ""
                                DGrid.TextMatrix(j, 4) = rsTemp("JOB")
                            Else
                                DGrid.TextMatrix(j, 3) = True
                            End If
                            nJRate = nJRate + objO.FetchData_Double("select Rate from ACTRATE WITH (NOLOCK) Where FS_ITEM = '" & Trim(DGrid.TextMatrix(j, 1)) & "' AND MDL = '" & Trim(text(10)) & "'")
                            baris = baris + 1
                        End If
                    Next
                    
                    Set rsTemp = objO.FetchData("SELECT * FROM BOOK_SUBLET WITH (NOLOCK) WHERE NOFAK = '" & rsHlp2("NOFAK") & "'")
                    If Not rsTemp.EOF Then
                        i = 1
                        barisSublet = 1
                        Do While Not rsTemp.EOF
                            subletGrid.TextMatrix(i, 1) = rsTemp("NAMABARANG")
                            subletGrid.TextMatrix(i, 2) = CDbl(rsTemp("QTY"))
                            subletGrid.TextMatrix(i, 3) = rsTemp("KODESATUAN")
                            subletGrid.TextMatrix(i, 4) = Format(CDbl(rsTemp("HARGA")), "#0,#")
                            subletGrid.TextMatrix(i, 5) = rsTemp("FLAG")
                            subletGrid.TextMatrix(i, 6) = rsTemp("ESTWAKTU")
                            i = i + 1
                            barisSublet = barisSublet + 1
                            rsTemp.MoveNext
                        Loop
                    End If
                    If rsTemp.State <> adStateClosed Then rsTemp.Close
                    'CEX 280514
                     If Log.isFill = 1 And Log.startAppRate <= Format(text(65), "dd/mm/yyyy") Then
                     Set rsTemp = objO.FetchDataACC("select NOSPK as NOFAK, KodeItem, Keterangan as namaBarang, SatuanNS as KodeSatuan, Qty, '' as FLAG, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID, '' as KDOLI from SOORD WITH (NOLOCK) Where NoSPK = '" & rsHlp2("NOFAK") & "' and (isnull(BO,'null') = 'null' or BO <> 'CANCEL')")
                     Else
                     Set rsTemp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM book_part A WITH (NOLOCK) LEFT JOIN T_OLI B WITH (NOLOCK) ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & rsHlp2("NOFAK") & "'")
                     End If
                   ' Set rsTemp = objO.FetchDataACC("select NOSPK as NOFAK, KodeItem, Keterangan as namaBarang, SatuanNS as KodeSatuan, Qty, '' as FLAG, TglPerlu as ETA, '' as Keterangan, NamaBYSA, Nomor as nID, '' as KDOLI from SOORD Where NoSPK = '" & rsHlp2("NOFAK") & "' and BO <> 'CANCEL'")
                   ' If rsTemp.RecordCount = 0 Then
                   '     Set rsTemp = objO.FetchData("SELECT A.*, ISNULL(FS_KD_OLI,'') AS KDOLI FROM book_part A LEFT JOIN T_OLI B ON LEFT(A.KODEITEM,1) = B.FS_KD_OLI AND B.TIPE = 'O' WHERE nofak = '" & rsHlp2("NOFAK") & "'")
                   ' End If
                    If Not rsTemp.EOF Then
                        i = 1
                        barispart = 1
                        Do While Not rsTemp.EOF
                            PartGrid.TextMatrix(i, 1) = rsTemp("KODEITEM")
                            PartGrid.TextMatrix(i, 2) = rsTemp("NAMABARANG")
                            PartGrid.TextMatrix(i, 3) = rsTemp("KODESATUAN")
                            PartGrid.TextMatrix(i, 4) = rsTemp("QTY")
                            PartGrid.TextMatrix(i, 6) = IIf(rsTemp("KDOLI") <> "", True, False)
                            PartGrid.TextMatrix(i, 7) = rsTemp("NAMABYSA")
                            PartGrid.TextMatrix(i, 8) = Val(NulA(rsTemp("nID")))
                            Dim rshrgPart As ADODB.Recordset
                                Set rshrgPart = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM LIKE '%" & Trim(rsTemp("KODEITEM")) & "%'")
                                If Not rshrgPart.EOF Then
                                    PartGrid.TextMatrix(i, 5) = rshrgPart("HARGAJUALA")
                                End If
                                rshrgPart.Close
                                Set rshrgPart = Nothing
                            barispart = barispart + 1
                            i = i + 1
                            rsTemp.MoveNext
                        Loop
                        NewId = (objO.FetchData_DoubleACC("select top 1 Nomor from SOORD WITH (NOLOCK) Where NOSPK = '" & rsHlp2("NOFAK") & "' Order By Nomor DESC")) + 1
                    End If
                    
                    ' ISI STALL
                    Call objO.InsertData("INSERT INTO tempStall SELECT noStall, Tanggal, JamMulai, JamSelesai, rate, [group], Mech, flag, Keterangan,adjust FROM stall WITH (NOLOCK) " _
                                 & "WHERE NOT EXISTS (SELECT * FROM tempStall WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%') AND " _
                                 & "tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%' ORDER BY noStall, Tanggal, JamMulai")
                    cmdStall.Tag = ""
                    Set rsHlp = objO.FetchData("SELECT * FROM TEMPSTALL WITH (NOLOCK) WHERE tanggal = '" & Format(Trim(text(23)), "mm/dd/yyyy") & "' AND keterangan like '%" & noPol & " |%'")
                    If Not rsHlp.EOF Then
                        text(58) = Trim(rsHlp("JAMMULAI"))
                        nMech = Trim(rsHlp("MECH"))
                    Else
                        text(58) = Trim(text(24))
                        nMech = Trim(rsHlp2("MECH"))
                    End If
                    rsHlp.Close
                    text(64) = nMech
                    
                    chEM = IIf(rsHlp2("CODE") = "1", Checked, Unchecked)
                    chSBNP = IIf(rsHlp2("KD_K") = "1", Checked, Unchecked)
                    
                    If text(58) <> ">>:>>" And text(23) <> ">>->>->>>>" Then
                        nWKT = hitungWaktu(CDbl(nJRate), text(58), text(23))
                        text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
                        text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
                        text(84) = TglDDMMYYYY(rsHlp2("DATESERAH"))
                        text(83) = Trim(rsHlp2("TIMESERAH"))
                        Call hitungEstimasiSelesai
                    End If
                Else
                    If PanggilDataCust(NoPolisi) Then
                        cmdStall.Tag = "belum"
                    End If
                End If
                lblEdit.Visible = True
                cmdHistory.Enabled = True
                Command1.Enabled = True
    If text(1) = "" Then Call cmdBersih_Click
            'End If
LEWAT:
    
    End If
    sSQL = "select * from MODEL WITH (NOLOCK) where MDL ='" & text(10) & "'"
    Set rsHlp2 = objO.FetchData(sSQL)
    If rsHlp2.BOF Then
        gGroupModel = ""
    Else
        gGroupModel = rsHlp2("FS_GRP_MDL")
    End If
    If rsHlp2.State <> adStateClosed Then rsHlp2.Close
End Sub

Private Sub Text_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
    If KeyCode = vbKeyF12 Then
        Select Case Index
            Case 0
                gMainKeyCari = text(Index)
                gSearchOption = "BOOK"
                frmCari.Show (vbModal)
            Case 2
                gMainKeyCari = text(Index)
                gSearchOption = "KODEPLG"
                frmCariAcess.Show (vbModal)
            Case 10
                If noPol = "" Then
                    MsgBox "Isi No Polisi kendaraan !!!"
                    text(0).SelLength = Len(text(0))
                    text(0).SetFocus
                    Exit Sub
                Else
                    gMainKeyCari = text(Index)
                    gSearchOption = "MDL"
                    frmCari.Show (vbModal)
                End If
            Case 13
                If text(10) <> "" Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "WARNA"
                    gBantuan = Trim(text(10))
                    frmCariSPK.Show (vbModal)
                Else
                    MsgBox "Isi model kendaraan !!!"
                    text(10).SelLength = 10
                    text(10).SetFocus
                    Exit Sub
                End If
            Case 32
                If text(10) = "" Then
                    MsgBox "Isi model kendaraan !!!"
                    text(10).SelLength = 10
                    text(10).SetFocus
                    Exit Sub
                Else
                    gMainKeyCari = text(Index)
                    gSearchOption = "JENIS_KERJA"
                    frmCari.Show (vbModal)
                End If
            Case 34
                If text(32) = "" Then
                    MsgBox "Isi Tipe Job Order !!!"
                    text(32).SelLength = 3
                    text(32).SetFocus
                    Exit Sub
                Else
                    If text(24) = ">>:>>" Or text(23) = ">>->>->>>>" Then
                        MsgBox "Tanggal dan Jam Booking !!!"
                        text(23).SelLength = Len(text(23))
                        text(23).SetFocus
                        Exit Sub
                    Else
                        gMainKeyCari = text(Index)
                        gSearchOption = "JOBLIST2"
                        gBantuan = text(10)
                        gBantuan2 = text(32)
                        frmCari.Show (vbModal)
                    End If
                End If
            Case 35
                If text(10) = "" Then
                    MsgBox "Isi model kendaraan !!!"
                    text(10).SelLength = 10
                    text(10).SetFocus
                    Exit Sub
                Else
                    gMainKeyCari = text(Index)
                    gSearchOption = "PART"
                    frmCariAcess.Show (vbModal)
                End If
            Case 46, 47, 48, 49, 50, 51
                gMainKeyCari = text(Index)
                gSearchOption = "KONDISI"
                gBantuan = Index - 45
                frmCari.Show (vbModal)
            Case 56, 57
                text(0).SetFocus
                Exit Sub
            Case 64
                If text(23) <> ">>->>->>>>" Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "MEKANIK"
                    If DateDiff("d", Format(text(23), "MM/DD/YYYY"), Format(isnow, "MM/DD/YYYY")) = 0 Then
                        gBantuan = "ORDER"
                        gBantuan2 = IIf(Weekday(isnow) = 1, 6, Weekday(isnow) - 2) & WeekdayName(IIf(Weekday(isnow) = 1, 8, Weekday(isnow)) - 1)
                        gBantuan3 = Format(isnow, "DD/MM/YYYY")
                        gBantuan4 = ""
                    Else
                        gBantuan = "BOOKING"
                        Dim nTgl As String
                        nTgl = Format(text(23), "MM/DD/YYYY")
                        gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
                        gBantuan3 = Format(nTgl, "DD/MM/YYYY")
                        gBantuan4 = ""
                    End If
                    frmCari.Show (vbModal)
                End If
            Case 79
                gMainKeyCari = text(Index)
                gSearchOption = "SUBLET"
                frmCari.Show (vbModal)
        End Select
    End If
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        cmdNo.Caption = "No &Polisi"
        text(0).SetFocus
    End If
End Sub

Private Sub text_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        SendKeys ("´")
    End If
    If KeyAscii = 34 Then
        KeyAscii = 0
        SendKeys ("”")
    End If
    KeyAscii = UpCase(KeyAscii)
    If Index = 27 And KeyAscii = 13 Then
        Dim nCase As Boolean
        nCase = False
        Dim xSB As Integer
        For xSB = 5 To 500 Step 5
            If Val(text(27).text) = xSB Then
                nCase = True
                Exit For
            End If
        Next xSB
        If nCase = True Then text(27).text = "SERVIS BERKALA " & Trim(text(27)) & ".000 KM"
    End If
    Select Case Index
        Case 52, 23
            KeyAscii = CekNumerik(KeyAscii)
            simpan = False
    End Select
End Sub

Private Sub PanggilDataCustBook(nPol As String)
Dim rsHlp As ADODB.Recordset
    Set rsHlp = objO.FetchData("SELECT * FROM BOOK WITH (NOLOCK) WHERE NO_POLISI = '" & nPol & "' ORDER BY [DATE] DESC")
    If Not rsHlp.EOF Then
        text(3) = rsHlp("CUSTOMER")
        text(4) = rsHlp("CUST_ADRES")
        text(5) = rsHlp("KOTA")
        text(6) = rsHlp("GSM")
        text(7) = rsHlp("CDMA")
        text(8) = rsHlp("TELPON")
        text(9) = rsHlp("FS_TH_MOBIL")
        text(10) = rsHlp("MDL")
        text(11) = rsHlp("MODEL")
                    Call CexChassis_Discon(text(11))
                    If isOnCarClass = False Then
                        MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
                    End If
        text(12) = rsHlp("FMDL")
        text(13) = rsHlp("COLOR")
    End If
End Sub

Private Sub PanggilDataCustDiagnose(noPol As String, noPrediag As String)
    Dim rsDiagapproval As New ADODB.Recordset
    Set rsDiagapproval = objO.FetchData("select da.*,d.no_polisi,d.* from diagdta d WITH (NOLOCK) inner join diagapproval da WITH (NOLOCK) on da.no_diag = d.no_diag where da.no_wo = '' and isnull(da.alasan_tdksetuju,'') = '' and  d.no_polisi = '" & noPol & "' and da.NO_DIAG = '" & noPrediag & "'")
    
    If Not rsDiagapproval.EOF Then
        'jika ada data
        no_approval = Trim(rsDiagapproval.Fields(0))
        Dim rsDiagdta As New ADODB.Recordset
        Dim rsDiagPart As New ADODB.Recordset
        Dim rsDiagDtl As New ADODB.Recordset
        
        Set rsDiagdta = objO.FetchData("select * from diagdta WITH (NOLOCK) where no_polisi = '" & Trim(noPol) & "' and no_diag = '" & no_approval & "' order by no_diag desc")
        Set rsDiagPart = objO.FetchData("select * from diagpart WITH (NOLOCK) where no_diag = '" & Trim(rsDiagapproval.Fields(0)) & "'")
        Set rsDiagDtl = objO.FetchData("select cmpl1,cmpl2,cmpl3,cmpl4,cmpl5 from diagdtl WITH (NOLOCK) where no_diag = '" & Trim(rsDiagapproval.Fields(0)) & "'")
        
        If Not rsDiagdta.EOF Then
            
            isBookDiag = True
            
            'jika ada data berdasarkan no polisi
            text(2) = Nul(rsDiagdta("FS_KD_PLG"))
            text(44) = Nul(rsDiagdta("FS_KD_WIL"))
            text(3) = Nul(rsDiagdta("CUSTOMER"))
            text(4) = Nul(rsDiagdta("CUST_ADRES"))
            text(5) = Nul(rsDiagdta("KOTA"))
            'text(6) = Nul(rsDiagdta("GSM"))
            'text(7) = Nul(rsDiagdta("CDMA"))
            'text(8) = Nul(rsDiagdta("TELPON"))
            text(9) = Nul(rsDiagdta("FS_TH_MOBIL"))
            text(10) = rsDiagdta("MDL")
            text(11) = rsDiagdta("MODEL")
                Call CexChassis_Discon(text(11))
                If isOnCarClass = False Then
                    MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
                End If
            text(12) = rsDiagdta("FMDL")
            text(13) = rsDiagdta("COLOR")
            text(14) = rsDiagdta("FS_KD_DEALER")
            text(15) = rsDiagdta("FS_NM_DEALER")
            text(16).Mask = ""
            text(16) = TglDDMMYYYY(IIf(rsDiagdta("DOTAM") = "", "01/01/1900", rsDiagdta("DOTAM")))
            text(17).Mask = ""
            text(17) = TglDDMMYYYY(IIf(rsDiagdta("DDATE") = "", "01/01/1900", rsDiagdta("DDATE")))
            text(41) = rsDiagdta("NO_CHASSIS")
            gNoRangka = rsDiagdta("NO_CHASSIS")
            cmdStall.Tag = ""
            cmdStall.BackColor = &H8000000F
            
            text(1) = rsDiagdta("NOFAK")
            'text(18) = TglDDMMYYYY(rsDiagdta("DATE_BOOK"))
            'text(19) = Format(rsDiagdta("TIME_D"), "HH:NN")
            text(23) = TglDDMMYYYY(rsDiagdta("DATE"))
            text(24) = Format(rsDiagdta("TIME"), "HH:NN")
            
            text(23).Enabled = True
            text(24).Enabled = True
'                        If Trim(rsHlp2("DELIV")) = "RES3" Then
'                            MsgBox "Sudah di Reschedule 3x, tidak dapat di reschedule kembali." & vbCr & "Batalkan Booking, buat booking baru ."
'                            text(23).Enabled = False
'                            text(24).Enabled = False
'                        End If

            text(26) = gSA
            text(27) = rsDiagdta("REQ1")
            text(28) = rsDiagdta("REQ2")
            text(29) = rsDiagdta("REQ3")
            text(30) = rsDiagdta("REQ4")
            text(31) = rsDiagdta("REQ5")
            
            ' ISI KELUHAN
            Dim hlp As String
            If Not rsDiagDtl.EOF Then
                For i = 1 To 5
                    hlp = "CMPL" & i
                    Text1(-1 + i) = rsDiagDtl(hlp)
                Next
            End If
        End If
        
        If Not rsDiagPart.EOF Then
            i = 1
            barispart = 1
            Do While Not rsDiagPart.EOF
                PartGrid.TextMatrix(i, 1) = rsDiagPart.Fields(4) '("KODEITEM")
                PartGrid.TextMatrix(i, 2) = rsDiagPart.Fields(5) '("NAMAITEM")
                PartGrid.TextMatrix(i, 3) = rsDiagPart.Fields(6) '("SAT_ITEM")
                PartGrid.TextMatrix(i, 4) = rsDiagPart.Fields(7) '("QTY")
                'PartGrid.TextMatrix(i, 6) = IIf(rsDiagPart.Fields() <> "", True, False) '("KDOLI") <> "", True, False)
                PartGrid.TextMatrix(i, 7) = "" 'rsDiagPart("NAMABYSA")
                PartGrid.TextMatrix(i, 8) = barispart 'Val(NulA(rsDiagPart("nID")))
                Dim rshrgPart As ADODB.Recordset
                If Trim(rsDiagPart.Fields(4)) <> "" Then
                    Set rshrgPart = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM LIKE '%" & Trim(rsDiagPart.Fields(4)) & "%'") '("KODEITEM")) & "%'")
                Else
                    Set rshrgPart = objO.FetchDataACC("SELECT HARGAJUALA FROM INMST WITH (NOLOCK) WHERE KODEITEM LIKE '%zzzzzzzzzzzzzzzzzzzzzzzzzzzzz%'") '("KODEITEM")) & "%'")
                End If
                
                If Not rshrgPart.EOF Then
                    PartGrid.TextMatrix(i, 5) = rshrgPart("HARGAJUALA")
                Else
                    PartGrid.TextMatrix(i, 5) = 0
                End If
                rshrgPart.Close
                Set rshrgPart = Nothing
                
                'If oliPart Then
                '    Text(55) = Format(CDbl(Text(55)) + hitungPart(partGrid.TextMatrix(barispart - 1, 4), partGrid.TextMatrix(barispart - 1, 5)), "#0,0")
                'Else
                    text(40) = Format(CDbl(text(40)) + hitungPart(PartGrid.TextMatrix(i, 4), PartGrid.TextMatrix(i, 5)), "#,0")
                'End If
                text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#,0#")
                barispart = barispart + 1
                i = i + 1
                rsDiagPart.MoveNext
            Loop
            NewId = barispart '(objO.FetchData_DoubleACC("select top 1 Nomor from SOORD Where NOSPK = '" & rsHlp2("NOFAK") & "' Order By Nomor DESC")) + 1
        End If
    End If
End Sub


Private Function PanggilDataCust(nPol As String) As Boolean
    Dim rsHlp3 As ADODB.Recordset
    On Error Resume Next
    Set rsHlp3 = objO.FetchData("Select * from CUSTDAT WITH (NOLOCK) where NO_POLISI = '" & Trim(Replace(Replace(nPol, ">", ""), " ", "")) & "'")
    If Not rsHlp3.EOF Then
        rsHlp3.MoveLast
        If rsTemp.State Then rsTemp.Close
        
        Set rsTemp = objO.FetchData("SELECT * FROM jobdta WITH (NOLOCK) WHERE no_polisi = '" & Trim(Replace(Replace(nPol, ">", ""), " ", "")) & "' AND RPR_COND <> 'K' AND DELIV <> 'BTL' AND TPSVC <> 'W' ORDER BY no_WO desc")
        If rsTemp.EOF Then
balik:
            text(2) = rsHlp3("FS_KD_PLG")
            text(44) = rsHlp3("FS_KD_WIL")
            text(3) = rsHlp3("CUSTOMER")
            text(4) = rsHlp3("CUST_ADRES")
            text(5) = rsHlp3("KOTA")
            text(6) = rsHlp3("GSM")
            text(7) = rsHlp3("CDMA")
            text(8) = rsHlp3("TELPON")
            text(9) = rsHlp3("FS_TH_MOBIL")
            text(10) = rsHlp3("MDL")
            text(11) = rsHlp3("MODEL")
                    Call CexChassis_Discon(text(11))
                    If isOnCarClass = False Then
                        MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
                    End If
            text(12) = rsHlp3("FMDL")
            text(13) = rsHlp3("COLOR")
            text(14) = rsHlp3("FS_KD_DEALER")
            text(15) = rsHlp3("FS_NM_DEALER")
            text(16) = TglDDMMYYYY(rsHlp3("DOTAM"))
            text(17) = TglDDMMYYYY(rsHlp3("DDATE"))
            gNoRangka = rsHlp3("NO_CHASSIS")
            text(41) = gNoRangka
            text(26) = gSA
            cmdHistory.Enabled = True
            Command1.Enabled = True
        Else
            If cmdNo.Caption = "No &Rekom" Then
                text(3).SetFocus
            Else
                'proses ada data no polisi dari prediagnos
                '- lakukan pengecekan no polisi di diagapproval
                
                    If rsTemp("RPR_COND") <> "K" Then
                        If Trim(cekMOU(rsHlp3("MDL"), rsHlp3("NO_CHASSIS"))) = "" Then
                            If cekWoInvoice(Trim(rsHlp3("NO_CHASSIS"))) Then
                                MsgBox "WO terakhir harus ditutup !", vbCritical, "CLOSING"
                                Call hapusForm
                                text(57).SetFocus
                            Else
                                If Not cekLastIsSBIDealer(Trim(rsHlp3("NO_CHASSIS"))) Then
                                    If Not cekLastIsPDS(Trim(rsHlp3("NO_CHASSIS"))) Then
                                        MsgBox "WO terakhir harus ditutup !", vbCritical, "CLOSING"
                                        Call hapusForm
                                        text(57).SetFocus
                                    Else
                                        If Not cekSBIDealerInvoice(Trim(rsHlp3("NO_CHASSIS"))) Then
                                            MsgBox "WO terakhir harus ditutup !", vbCritical, "CLOSING"
                                            Call hapusForm
                                            text(57).SetFocus
                                        Else
                                            GoTo balik
                                        End If
                                    End If
                                Else
                                    GoTo balik
                                End If
                            End If
                        Else
                            GoTo balik
                        End If
                End If
            End If
        End If
        PanggilDataCust = True
    Else
        cmdHistory.Enabled = False
        Command1.Enabled = False
        PanggilDataCust = False
    End If
    If rsHlp3.State <> adStateClosed Then rsHlp3.Close
    gGroupModel = objO.FetchData_String("select FS_GRP_MDL from MODEL WITH (NOLOCK) where MDL ='" & text(10) & "'")
End Function

Private Sub Text_LostFocus(Index As Integer)
    Select Case Index
        Case 9
            If text(Index) = 0 Then
                text(Index).SelLength = Len(text(Index))
                text(Index).SetFocus
                Exit Sub
            End If
        Case 23
            If text(Index) <> "" Or text(Index) <> ">>->>->>>>" Then
                cmdWaktu.Enabled = True
            Else
                cmdWaktu.Enabled = False
            End If
    End Select
End Sub

Private Sub isCurrentAR()
    If Trim(text(2).text) = "" Then Exit Sub
    Dim aBatasKr As Double
    Dim aCurrent As Double
    aBatasKr = objO.Fetch_BatasKredit_Ar(Trim(Replace(text(2).text, ">", "")))

    txtRp(1) = aBatasKr
    aCurrent = objO.Fetch_Current_Ar(Trim(Replace(text(2).text, ">", "")))
    txtRp(0) = aCurrent
    If aCurrent <> 0 And aBatasKr <> 0 Then
        nPersenAR = Round((aCurrent / aBatasKr) * 100, 0)
    Else
        nPersenAR = 0
    End If
    nAR2.Caption = nPersenAR & "%"
    Call isOverdueAR
End Sub

Private Sub isOverdueAR()
    Dim nOAR As ADODB.Recordset
    Set nOAR = objO.FetchOverdueAR(Trim(Replace(text(2).text, ">", "")))
    If nOAR.RecordCount > 0 Then
        Set nOAR = objO.FetchOverdueAR_Data(Trim(Replace(text(2).text, ">", "")))
        If nOAR.RecordCount > 0 Then
            nAR2.Caption = nAR2.Caption & "  ( AR OVERDUE )"
        End If
    End If
End Sub

Private Sub text_Validate(Index As Integer, Cancel As Boolean)
    Dim boleh As Boolean
    Dim birthDay As String
    Dim rsSearch As ADODB.Recordset
    
    Select Case Index
        Case 57
            noPol = Trim(text(0)) & IIf(Trim(text(0)) <> "", "-", "") & Trim(text(56)) & IIf(Trim(text(57)) <> "", "-", "") & Trim(text(57))
            text(78) = Trim(noPol)
            If Trim(noPol) <> "" Then
                isBookDiag = False
                no_approval = ""
                boleh = True
                
                If simpan = False Then
                    If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                        boleh = False
                    End If
                End If
                If boleh Then
                    If Layar.Tag <> "estBook" Then
                        lblEdit.Visible = False
                        Call hapusForm
                        If frmBookingDMSNonDSS.Tag = "FU" Then
                            chkRtn = Checked
                        End If
                        Dim nDpn1 As String
                        Dim nTgh1 As String
                        Dim nBlk1 As String
                        nDpn1 = ""
                        nTgh1 = ""
                        nBlk1 = ""
                        Call pecahNoPol(noPol, nDpn1, nTgh1, nBlk1)
                        simpan = True
                        text(0) = Trim(nDpn1)
                        text(56) = Trim(nTgh1)
                        text(57) = Trim(nBlk1)
                        'lakukan pengecekan dari diagnose yang oke
                        Dim rsDiagapproval As New ADODB.Recordset
                        Dim pilihNoPrediagnose As String
                        Dim kumpulanPrediagnose As String
                        Dim jumlahPrediagnose As Integer
                        
                        Set rsDiagapproval = objO.FetchData("select da.*,d.no_polisi,d.* from diagdta d WITH (NOLOCK) inner join diagapproval da WITH (NOLOCK) on da.no_diag = d.no_diag where da.no_wo = '' and isnull(da.alasan_tdksetuju,'') = '' and  d.no_polisi = '" & noPol & "' and d.VALIDTWC = 'OKE' and isnull(da.NOFAK,'') = ''")
                        
                        If Not rsDiagapproval.EOF Then
                            'jika ada data
                            jumlahPrediagnose = 1
                            rsDiagapproval.MoveFirst
                            Do While Not rsDiagapproval.EOF
                                kumpulanPrediagnose = kumpulanPrediagnose & jumlahPrediagnose & ". " & rsDiagapproval.Fields(0) & vbCrLf
                                jumlahPrediagnose = jumlahPrediagnose + 1
                                rsDiagapproval.MoveNext
                            Loop
                            
                            If MsgBox("Sudah ada Prediagnosis dengan No Prediagnosis :" & vbCrLf & kumpulanPrediagnose & " Untuk No Polisi tersebut. " & vbCrLf & " Apakah ingin membuat booking atas prediagnosis tersebut?", vbYesNo, "Konfirmasi") = vbYes Then
ulang:
                                pilihNoPrediagnose = UCase(InputBox("Masukan Prediagnose yang ingin dibuat Booking " & vbCrLf & kumpulanPrediagnose & " : ", "Pilih Prediagnose "))
                                If pilihNoPrediagnose = "" Then
                                    If MsgBox("Apakah ingin membatalkan Booking dari Prediagnose?", vbYesNo, "Konfirmasi") = vbYes Then
                                        Exit Sub
                                    Else
                                        GoTo ulang
                                    End If
                                Else
                                    rsDiagapproval.filter = "NO_DIAG = '" & pilihNoPrediagnose & "'"
                                    If rsDiagapproval.EOF Then
                                        MsgBox "Tidak ada No Prediagnose untuk No Polisi tersebut!", vbOKOnly, ""
                                        GoTo ulang
                                    Else
                                        Call PanggilDataCustDiagnose(noPol, Trim(pilihNoPrediagnose))
                                        Exit Sub
                                    End If
                                End If
                            Else
                                'jika tidak membuat booking dari prediagnose
                                
                                'LAKUKAN PEMANGGILAN DATA BOOKING OPEN
                                'update tanggal 26-08-2014
                                Set rsSearch = objO.FetchData("Select DELIV,nofak from BOOK WITH (NOLOCK) where NO_POLISI ='" & Trim(noPol) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
                                Dim pilihNoBooking As String
                                Dim kumpulanBooking As String
                                Dim jumlahBooking As Integer
                                If Not rsSearch.EOF Then
                                    rsSearch.MoveFirst
                                    Do While Not rsSearch.EOF
                                        jumlahBooking = jumlahBooking + 1
                                        kumpulanBooking = kumpulanBooking & jumlahBooking & ". " & rsSearch.Fields(1) & vbCrLf
                                        rsSearch.MoveNext
                                    Loop
                                    If MsgBox("Sudah ada Booking dengan No Booking :" & vbCrLf & kumpulanBooking & " Untuk No Polisi tersebut. " & vbCrLf & " Apakah ingin mengubah booking tersebut?", vbYesNo, "Konfirmasi") = vbYes Then
balikLagi:
                                        pilihNoBooking = UCase(InputBox("Masukan No Booking yang ingin diubah " & vbCrLf & kumpulanBooking & " : ", "Pilih Booking "))
                                        If pilihNoBooking = "" Then
                                            If MsgBox("Apakah ingin membatalkan mengubah data Booking?", vbYesNo, "Konfirmasi") = vbYes Then
                                                Exit Sub
                                            Else
                                                GoTo balikLagi
                                            End If
                                        Else
                                            rsSearch.filter = "nofak = '" & pilihNoBooking & "'"
                                            If rsSearch.EOF Then
                                                MsgBox "Tidak ada No Booking untuk No Polisi tersebut!", vbOKOnly, ""
                                                GoTo balikLagi
                                            Else
                                                Call PanggilDataBooking2(noPol, Trim(pilihNoBooking), rsSearch("DELIV"))
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                    'jika tidak mengubah data booking
                                        Call hapusForm
                                        text(0).SetFocus
                                        Exit Sub
                                    End If
                                End If
                                Call PanggilDataCust(noPol)
                            End If
                        Else
                            'ambil data non diagnose
                            'update tanggal 26-08-2014
                            Set rsSearch = objO.FetchData("Select DELIV,nofak from BOOK WITH (NOLOCK) where NO_POLISI ='" & Trim(noPol) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
                                Dim pilihNoBooking2 As String
                                Dim kumpulanBooking2 As String
                                Dim jumlahBooking2 As Integer
                                If Not rsSearch.EOF Then
                                    rsSearch.MoveFirst
                                    Do While Not rsSearch.EOF
                                        jumlahBooking2 = jumlahBooking2 + 1
                                        kumpulanBooking2 = kumpulanBooking2 & jumlahBooking2 & ". " & rsSearch.Fields(1) & vbCrLf
                                        rsSearch.MoveNext
                                    Loop
                                    If MsgBox("Sudah ada Booking dengan No Booking :" & vbCrLf & kumpulanBooking2 & " Untuk No Polisi tersebut. " & vbCrLf & " Apakah ingin mengubah booking tersebut?", vbYesNo, "Konfirmasi") = vbYes Then
balikLagiUbahBook:
                                        pilihNoBooking2 = UCase(InputBox("Masukan No Booking yang ingin diubah " & vbCrLf & kumpulanBooking2 & " : ", "Pilih Booking "))
                                        If pilihNoBooking2 = "" Then
                                            If MsgBox("Apakah ingin membatalkan mengubah data Booking?", vbYesNo, "Konfirmasi") = vbYes Then
                                                Exit Sub
                                            Else
                                                GoTo balikLagiUbahBook
                                            End If
                                        Else
                                            rsSearch.filter = "nofak = '" & pilihNoBooking2 & "'"
                                            If rsSearch.EOF Then
                                                MsgBox "Tidak ada No Booking untuk No Polisi tersebut!", vbOKOnly, ""
                                                GoTo balikLagiUbahBook
                                            Else
                                                Call PanggilDataBooking2(noPol, Trim(pilihNoBooking2), rsSearch("DELIV"))
                                                Exit Sub
                                            End If
                                        End If
                                    Else
                                    'jika tidak mengubah data booking
                                        Call hapusForm
                                        text(0).SetFocus
                                        Exit Sub
                                    End If
                                End If
                                Call PanggilDataCust(noPol)
                            End If
                            
'                            Set rsSearch = objO.FetchData("Select DELIV from BOOK where NO_POLISI ='" & Trim(noPol) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
'                            If rsSearch.BOF Then
'                                baris = 1
'                                If Not PanggilDataCust(noPol) Then
'                                '    Call PanggilDataCustBook(noPol)
'                                End If
'                            Else
'                                Call PanggilDataBooking(Trim(noPol), rsSearch("DELIV"))
'                            End If
                        'End If
'                        Set rsSearch = objO.FetchData("Select DELIV from BOOK where NO_POLISI ='" & Trim(noPol) & "' AND DELIV <> 'BTL' AND DELIV <> 'OKE'")
'                        If rsSearch.BOF Then
'                            baris = 1
'                            If Not PanggilDataCust(noPol) Then
'                            '    Call PanggilDataCustBook(noPol)
'                            End If
'                        Else
'                            Call PanggilDataBooking(Trim(noPol), rsSearch("DELIV"))
'                        End If
                        If cekPiutang(Trim(text(2)), Trim(text(10)), gNoRangka) Then
                        End If
                        birthDay = ""
                        birthDay = birthDayReminder(Trim(text(41)))
                        If birthDay <> "" Then
                            If Abs(DateDiff("d", Format(birthDay, "mm/dd/yyyy"), isnow)) <= 5 Then
                                MsgBox "Selamat Ulang tahun Bp./Ibu./Sdr." & cariNamaBirth(Trim(text(15))) & vbCr & "Pada Tanggal " & birthDay
                            End If
                        End If
                        text(43) = cekMOU(Trim(text(10)), gNoRangka)
                        text(26) = gSA
                        simpan = False
                    End If
                End If
                If cmdNo.Caption = "No &Polisi" Then
                    cmdBersih.Visible = False
                    If Trim(gNoRangka) <> "" Then
                        Call panggilHistoryBok(gNoRangka, Trim(noPol))
                        Call panggilWarrantyBok(gNoRangka, Trim(text(10)))
                    End If
                ElseIf cmdNo.Caption = "No &Rekom" Then
                    cmdBersih.Visible = True
                End If
                
            Else
                Call hapusForm
            End If
        Case 2
            If text(Index) <> "" Then
                    Call isCurrentAR
                    'Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WHERE kodeLgn = '" & Text(Index) & "'")
                    '20/12/2012 Cust Non Axtif
                    Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WITH (NOLOCK) WHERE kodeLgn = '" & text(Index) & "' and nonaktifflag = 0")
                    If rsSearch.EOF Then
                        gMainKeyCari = text(Index)
                        gSearchOption = "KODEPLG"
                        frmCariAcess.Show (vbModal)
                    Else
                        text(44) = Nul(rsSearch("KODEWIL"))
                        text(3) = Nul(rsSearch("NAMALGN"))
                        text(4) = Nul(rsSearch("ALAMAT1"))
                        text(5) = Nul(rsSearch("ALAMAT2"))
                        
                        nDpn1 = ""
                        nTgh1 = ""
                        nBlk1 = ""
                        Call pecahNoTelp(Trim(Nul(rsSearch("TELEPON"))), nDpn1, nTgh1, nBlk1, ",")
                        text(8) = Trim(nDpn1)
                        text(6) = Trim(nTgh1)
                        text(7) = Trim(nBlk1)
                    End If
            End If
        Case 3
            If text(Index) <> "" Then
balik:
                    Set rsSearch = objO.FetchDataACC("SELECT * FROM armst WITH (NOLOCK) WHERE NAMALgn = '" & text(Index) & "'")
                    If rsSearch.EOF Then
                        text(2) = ""
                        text(44) = ""
                    Else
                        Dim rsKode As ADODB.Recordset
lanjut:
                        Dim telp As String
                        Set rsKode = objO.FetchDataACC("SELECT KODEWIL,ALAMAT1,ALAMAT2,TELEPON FROM ARMST WITH (NOLOCK) WHERE KODELGN = '" & Trim(text(2)) & "'")
                        If Not rsKode.EOF Then
                            text(44) = rsKode("KODEWIL")
                            text(4) = rsKode("ALAMAT1")
                            text(5) = rsKode("ALAMAT2")
                            telp = rsKode("TELEPON")
                        Else
                            text(44) = rsSearch("KODEWIL")
                            text(4) = rsSearch("ALAMAT1")
                            text(5) = rsSearch("ALAMAT2")
                            telp = rsSearch("TELEPON")
                        End If
                        If rsKode.State <> adStateClosed Then rsKode.Close
                        
                        nDpn1 = ""
                        nTgh1 = ""
                        nBlk1 = ""
                        Call pecahNoTelp(Trim(telp), nDpn1, nTgh1, nBlk1, ",")
                        text(8) = Trim(nDpn1)
                        text(6) = Trim(nTgh1)
                        text(7) = Trim(nBlk1)
                    End If
            End If
        Case 9
            If text(Index) = "" Then
                text(Index) = 0
                text(Index).SetFocus
            ElseIf CDbl(text(Index)) < (Year(isnow) - 50) Or CDbl(text(Index)) > Year(isnow) Then
                text(Index) = 0
                text(Index).SelLength = 4
                text(Index).SetFocus
                Exit Sub
            End If
        Case 10
            If text(Index) <> "" Then
                Set rsSearch = objO.FetchData("Select * from MODEL WITH (NOLOCK) where MDL ='" & text(Index) & "'")
                If rsSearch.BOF Then
                    Cancel = True
                    gMainKeyCari = text(Index)
                    gSearchOption = "MDL"
                    frmCari.Show (vbModal)
                    gGroupModel = ""
                Else
                    text(10) = rsSearch("MDL")
                    text(12) = rsSearch("Desc")
                    text(11) = rsSearch("MODEL")
                    Call CexChassis_Discon(text(11))
                    If isOnCarClass = False Then
                        MsgBox "Kelas Kendaraan Tidak Terdaftar, Hubungi HO Service !", vbCritical + vbOKOnly, "Car Class"
                    End If
                    gGroupModel = rsSearch("FS_GRP_MDL")
                End If
            Else
                text(12) = ""
                text(11) = ""
                gGroupModel = ""
            End If
        Case 23
            If CekTgl(text(Index)) Then
                If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(text(Index), "DD/MM/YYYY")) < 0 Then
                    text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
                    MsgBox "Tanggal Booking tidak boleh lebih kecil dari hari ini"
                    'Exit Sub
                ElseIf DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(text(Index), "DD/MM/YYYY")) > 30 Then
                    text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
                    MsgBox "Tanggal Booking tidak boleh lebih besar dari 1 bulan"
                    'Exit Sub
                End If
            ElseIf Not CekTgl(text(Index)) Then
                If text(Index) <> ">>->>->>>>" Then
                    MsgBox "Format tanggal salah"
                End If
                text(Index) = TglDDMMYYYY(Format(isnow, "dd/mm/yyyy"))
                'Exit Sub
            End If
            If text(58) <> ">>:>>" And text(23) <> ">>->>->>>>" Then
                nWKT = hitungWaktu(CDbl(nJRate), text(58), text(23))
                text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
                text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
            End If
        Case 24
            If Not CekJam(text(Index)) Then
                text(Index) = Format(isnow, "HH:NN")
            Else
                If DateDiff("d", Format(isnow, "dd/mm/yyyy"), Format(text(23), "DD/MM/YYYY")) = 0 Then
                    If DateDiff("s", Format(isnow, "HH:NN"), Format(text(Index), "HH:NN")) < 0 Then
                        text(Index) = Format(isnow, "HH:NN")
                    End If
                End If
            End If
            text(58) = Trim(text(24))
            If text(58) <> ">>:>>" And text(23) <> ">>->>->>>>" And Not CekJam(text(58)) Then
                nWKT = hitungWaktu(CDbl(nJRate), text(58), text(23))
                text(45) = Left(nWKT, InStr(nWKT, "|") - 1)
                text(59) = TglDDMMYYYY(Right(nWKT, Len(nWKT) - InStr(nWKT, "|")))
            End If
            
            If Trim(text(1)) <> "" Then
                noPolis = Trim(text(0)) & IIf(Trim(text(0)) <> "", "-", "") & Trim(text(56)) & IIf(Trim(text(57)) <> "", "-", "") & Trim(text(57))
                Set rsTemp = objO.FetchData("SELECT * FROM STALL WITH (NOLOCK) WHERE KETERANGAN LIKE '%" & noPolis & " |" & Trim(text(1)) & "%'")
                If Not rsTemp.EOF Then
                    If DateDiff("d", Format(rsTemp("TANGGAL"), "MM/DD/YYYY"), Format(Trim(text(23)), "MM/DD/YYYY")) = 0 Then
                        If Trim(rsTemp("JAMMULAI")) < Trim(text(24)) Then
                            Call objO.DeleteData("DELETE TEMPSTALL WHERE KETERANGAN LIKE '" & noPolis & " |" & Trim(text(1)) & "|%'")
                            MsgBox "Atur Ulang Posisi Stall"
                            cmdStall.Tag = "belum"
                        End If
                    Else
                        MsgBox "Atur Ulang Posisi Stall"
                        cmdStall.Tag = "belum"
                    End If
                End If
                If rsTemp.State <> adStateClosed Then rsTemp.Close
            End If
        Case 25
            If text(Index) = "" Then
                text(Index) = "T"
            ElseIf (Trim(text(Index)) <> "T" And Trim(text(Index)) <> "D" And Trim(text(Index)) <> "S") Then
                text(Index) = "T"
            End If
        Case 32
            If text(Index) <> "" Then
                If text(10) = "" Then
                    MsgBox "Isi model kendaraan !!!"
                    SendKeys ("+{TAB}")
                    text(10).SelLength = Len(text(10))
                    text(10).SetFocus
                    Exit Sub
                Else
                    Set rsSearch = objO.FetchData("Select DISTINCT JOB, REMARK from HO_FLATE WITH (NOLOCK) where job = '" & text(Index) & "' AND FS_LOG_DEL = ''")
                    If rsSearch.BOF Then
                        Cancel = True
                        gMainKeyCari = text(Index)
                        gSearchOption = "JENIS_KERJA"
                        frmCari.Show (vbModal)
                    Else
                        text(33) = rsSearch("REMARK")
                    End If
                End If
            End If
        Case 34
            If text(Index) <> "" Then
                If text(32) = "" Then
                    MsgBox "Isi Tipe Job Order !!!"
                    SendKeys ("+{TAB}")
                    text(32).SelLength = 3
                    text(32).SetFocus
                    Exit Sub
                Else
                    If text(23) <> ">>->>->>>>" Then
                        Set rsSearch = objO.FetchData("select DISTINCT FS_ITEM, RATE from HO_FLATE WITH (NOLOCK) where job = '" & text(32) & "' and fs_item = '" & text(Index) & "' AND mdl = '" & text(10) & "' and FS_LOG_DEL = ''")
                        If rsSearch.BOF Then
                            Cancel = True
                            gMainKeyCari = text(Index)
                            gSearchOption = "JOBLIST2"
                            gBantuan = text(10)
                            gBantuan2 = text(32)
                            frmCari.Show (vbModal)
                            If fOther Then
                                text(42).SelLength = Len(text(42))
                            End If
                        Else
                            If Trim(rsSearch("FS_ITEM")) = "OTHER" Then
                                fOther = True
                                text(34) = ""
                                text(34).Mask = ">CCCCCCCCCCCCCCCCCCCCCCCCC"
                                text(42).Visible = True
                                text(42).Enabled = True
                                text(34).SelLength = Len(text(34))
                                text(34).SetFocus
                            Else
                                fOther = False
                                text(42).Visible = False
                                Jrate = CDbl(rsSearch("Rate"))
                            End If
                        End If
                    Else
                        MsgBox "Isi Tanggal & Jam Booking", vbCritical, "KELENGKAPAN"
                        SendKeys ("+{TAB}")
                        text(23).SelLength = Len(text(23))
                        text(23).SetFocus
                        Exit Sub
                    End If
                End If
            End If
        Case 35
            'PART
            If text(Index) <> "" Then
                    Set rsSearch = objO.FetchDataACC("SELECT A.KODEITEM, NAMABARANG, KODESATUAN, SUM(ISNULL(B.QTY,0)) as QTYSTOCK, HARGAJUALA FROM inmst A WITH (NOLOCK) INNER JOIN intrn B WITH (NOLOCK) ON A.kodeItem = B.kodeItem WHERE A.kodegdg = '03' and A.kodeItem = '" & text(Index) & "' GROUP BY A.kodeitem, namabarang, kodeSatuan, hargajuala ORDER BY A.kodeitem")
                    If rsSearch.BOF Then
                        If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then
                            stkPart = 0
                            namaPart = ""
                            HrgPart = 0
                            satPart = "ITEM"
                            oliPart = False
                            namabySA = text(Index)
                            If rsTemp.State <> adStateClosed Then rsTemp.Close
                        Else
                            Cancel = True
                            gMainKeyCari = text(Index)
                            text(Index) = ""
                            gSearchOption = "PART"
                            frmCariAcess.Show (vbModal)
                        End If
                    Else
                        stkPart = CDbl(rsSearch("QTYSTOCK"))
                        namaPart = rsSearch("NAMABARANG")
                        HrgPart = Val(rsSearch("HARGAJUALA"))
                        satPart = rsSearch("KODESATUAN")
                        Set rsTemp = objO.FetchData("SELECT FS_KD_OLI FROM T_OLI WITH (NOLOCK) WHERE FS_KD_OLI = '" & Left(text(Index), 1) & "' AND TIPE = 'O'")
                        If Not rsTemp.EOF Then
                            oliPart = True
                        Else
                            oliPart = False
                        End If
                        namabySA = ""
                        If rsTemp.State <> adStateClosed Then rsTemp.Close
                    End If
                
            End If
        
        Case 36
            If text(35) <> "" Then
                If Log.isFill = 1 And Log.startAppRate <= Format(Trim(text(18)), "DD/MM/YYYY") Then
                    If Trim(text(Index)) = "" Then
                            text(Index) = 0
                        ElseIf CDbl(text(Index)) = 0 Then
                            SendKeys ("+{TAB}")
                            Exit Sub
                        Else
        ''''                    If CDbl(text(Index)) > CDbl(stkPart) Then
        ''''                        temp = MsgBox("Melebihi Stok yang ada !!!" & vbCr & "Tetap akan memproses ?", vbYesNo, "PESAN MELEBIHI STOCK")
        ''''                        If temp = vbYes Then
        ''''                            ' NOTE KE BAGIAN SPAREPART UNTUK PESAN
        ''''                        Else
        ''''                            text(Index) = stkPart
        ''''                        End If
        ''''                    End If
                        End If
                    
                    Else
                        For i = 1 To barispart
                            If UCase(PartGrid.TextMatrix(i, 1)) = UCase(Trim(text(35))) Or UCase(PartGrid.TextMatrix(i, 7)) = UCase(Trim(text(35))) Then
                                text(Index) = 0
                                Exit Sub
                            End If
                        Next
                    End If
    '
                
            End If
        Case 42
            If text(Index) = "" Then
                text(Index) = 0
            Else
                If text(32) = "" Then
                    MsgBox "Isi Tipe Job Order !!!"
                    text(32).SelLength = 3
                    text(32).SetFocus
                    Exit Sub
                End If
            End If
        Case 46, 47, 48, 49, 50, 51
            If text(Index) <> "" Then
                gBantuan = Index - 45
                Select Case gBantuan
                    Case 1
                        Set rsSearch = objO.FetchData("SELECT waktu FROM T_KONDISI WITH (NOLOCK) WHERE waktu = '" & text(Index) & "' ORDER BY id")
                    Case 2
                        Set rsSearch = objO.FetchData("SELECT frekuensi FROM T_KONDISI WITH (NOLOCK) WHERE frekuensi = '" & text(Index) & "' ORDER BY id")
                    Case 3
                        Set rsSearch = objO.FetchData("SELECT tempat FROM T_KONDISI WITH (NOLOCK) WHERE tempat = '" & text(Index) & "' ORDER BY id")
                    Case 4
                        Set rsSearch = objO.FetchData("SELECT warn FROM T_KONDISI WITH (NOLOCK) WHERE warn = '" & text(Index) & "' ORDER BY id")
                    Case 5
                        Set rsSearch = objO.FetchData("SELECT mulai FROM T_KONDISI WITH (NOLOCK) WHERE mulai = '" & text(Index) & "' ORDER BY id")
                    Case 6
                        Set rsSearch = objO.FetchData("SELECT selesai FROM T_KONDISI WITH (NOLOCK) WHERE selesai = '" & text(Index) & "' ORDER BY id")
                End Select
                If rsSearch.EOF Then
                    text(Index) = ""
                Else
                    Dim nT As Integer
                    Select Case gBantuan
                        Case 1
                            text(54) = ""
                            nT = 54
                        Case 2
                            text(65) = ""
                            nT = 65
                        Case 3
                            text(66) = ""
                            nT = 66
                        Case 4
                            text(67) = ""
                            nT = 67
                        Case 5
                            text(68) = ""
                            nT = 68
                        Case 6
                            text(69) = ""
                            nT = 69
                    End Select
                    If text(Index) = "LAINNYA" Then
                        Select Case gBantuan
                            Case 1
                                text(54) = bKondisi(0)
                            Case 2
                                text(65) = bKondisi(1)
                            Case 3
                                text(66) = bKondisi(2)
                            Case 4
                                text(67) = bKondisi(3)
                            Case 5
                                text(68) = bKondisi(4)
                            Case 6
                                text(69) = bKondisi(5)
                        End Select
                        text(nT).Visible = True
                        text(nT).SetFocus
                    Else
                        text(nT) = ""
                        Me.text(nT).Visible = IIf(Me.text(nT).Visible, False, False)
                    End If
                End If
            End If
        Case 54
            If text(Index) <> "" Then
                bKondisi(0) = text(Index)
            End If
        Case 65
            If text(Index) <> "" Then
                bKondisi(1) = text(Index)
            End If
        Case 66
            If text(Index) <> "" Then
                bKondisi(2) = text(Index)
            End If
        Case 67
            If text(Index) <> "" Then
                bKondisi(3) = text(Index)
            End If
        Case 68
            If text(Index) <> "" Then
                bKondisi(4) = text(Index)
            End If
        Case 69
            If text(Index) <> "" Then
                bKondisi(5) = text(Index)
            End If
        Case 64
            If text(Index) <> "" Then
                Set rsSearch = objO.FetchData("SELECT * FROM MEKANIK WITH (NOLOCK) WHERE MECH = '" & text(Index) & "'")
                If rsSearch.EOF Then
                    gMainKeyCari = text(Index)
                    gSearchOption = "MEKANIK"
                    gTanggal = isnow
                    If DateDiff("d", Format(text(23), "MM/DD/YYYY"), Format(gTanggal, "MM/DD/YYYY")) = 0 Then
                        gBantuan = "ORDER"
                        gBantuan2 = IIf(Weekday(gTanggal) = 1, 6, Weekday(gTanggal) - 2) & WeekdayName(IIf(Weekday(gTanggal) = 1, 8, Weekday(gTanggal)) - 1)
                        gBantuan3 = Format(gTanggal, "DD/MM/YYYY")
                        gBantuan4 = ""
                    Else
                        gBantuan = "BOOKING"
                        Dim nTgl As String
                        nTgl = Format(text(23), "MM/DD/YYYY")
                        gBantuan2 = IIf(Weekday(nTgl) = 1, 6, Weekday(nTgl) - 2) & WeekdayName(IIf(Weekday(nTgl) = 1, 8, Weekday(nTgl)) - 1)
                        gBantuan3 = Format(nTgl, "DD/MM/YYYY")
                        gBantuan4 = ""
                    End If
                    frmCari.Show (vbModal)
                End If
            End If
    End Select
     
End Sub


Public Sub panggilHistoryBok(noRgk As String, noPols As String)
    On Error Resume Next
    Dim rsHist As ADODB.Recordset
    
        ' AMBIL HISTORY DARI HO
    Dim rsHlp  As ADODB.Recordset
    Dim rsHlp2 As ADODB.Recordset
    Dim dBaseHO As String
    
    If Trim(noPol) = "" Then
        noPol = Trim(text(0)) & IIf(Trim(text(0)) <> "", "-", "") & Trim(text(56)) & IIf(Trim(text(57)) <> "", "-", "") & Trim(text(57))
    End If
    gNoRangka = noRgk

    If rsHist.State Then rsHist.Close
    If noRgk = "" Then
        sSQL = "SELECT * FROM History WITH (NOLOCK) WHERE no_polisi = '" & Trim(noPols) & "' AND FS_LOG_DEL = '' ORDER BY [date] DESC, NO_WO DESC"
    Else
        sSQL = "SELECT * FROM History WITH (NOLOCK) WHERE no_chassis = '" & Trim(noRgk) & "' AND FS_LOG_DEL = '' ORDER BY [date] DESC, NO_WO DESC"
    End If
    Set rsHist = objO.FetchData(sSQL)
    If Not rsHist.EOF Then
        text(20) = Format(Trim(rsHist("DATE")), "dd-mm-yyyy")
        text(21) = Nul(Trim(rsHist("FS_NM_DEALER")))
        text(22) = Format(CDbl(rsHist("KM")), "#,#")
        gNoRangka = Trim(rsHist("NO_CHASSIS"))
        If DateDiff("d", Format(rsHist("INVDATE"), "mm/dd/yyyy"), Format(isnow, "mm/dd/yyyy")) > 15 Then
            chkRtn.Enabled = False
        Else
            chkRtn.Enabled = True
        End If
        cmdHistory.Enabled = True
        Command1.Enabled = True
    Else
        chkRtn.Enabled = False
        text(22) = 0
        'cmdHistory.Enabled = False
    End If
    cmdHistory.Enabled = True
    Command1.Enabled = True
End Sub

Private Sub panggilHistoryHO(noRgk As String)
Dim rsHlp As ADODB.Recordset

            
                
    Set rsRec = objO.FetchDataHO("SELECT no_chassis FROM Historynew WITH (NOLOCK) WHERE no_chassis = '" & noRangka & "' AND kodeTipe = '" & rsTemp("ID_WARRANTY") & "' AND FS_LOG_DEL = ''")
    If Not rsRec.EOF Then
        terlibat = False
    Else
        terlibat = True
    End If


    If Log.LogHOConn = 1 Then
        If ConHO = True Then
        dBaseHO = HODb
         
        Set rsHlp = objO.FetchDataHO("SELECT * FROM HISTORYnew WITH (NOLOCK) WHERE NO_CHASSIS = '" & noRgk & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC, NO_WO DESC")
        If Not rsHlp.EOF Then
            Do While Not rsHlp.EOF
                If rsHlp("NO_CHASSIS") <> "" Then
                     
                    Set rsHlp2 = objO.FetchData("SELECT * FROM HISTORY WITH (NOLOCK) WHERE NO_CHASSIS = '" & rsHlp("NO_CHASSIS") & "' AND NO_WO = '" & rsHlp("NO_WO") & "' AND NOINV = '" & rsHlp("NOINV") & "' AND KM= '" & rsHlp("KM") & "' AND FS_LOG_DEL = '' ORDER BY [DATE] DESC")
                    If rsHlp2.EOF Then
                        Call objO.InsertData("INSERT INTO HISTORY VALUES ('" & rsHlp("CUSTOMER") & "', '" & rsHlp("CUST_ADRES") & "', '" & rsHlp("KOTA") & "','" & rsHlp("NO_POLISI") & "', '" & rsHlp("NO_ENGINE") & "','" & rsHlp("NO_CHASSIS") & "','" & rsHlp("FRONT") & "','" & rsHlp("MECH") & "','" & rsHlp("KM") & "','" & rsHlp("DATE") & "'," & rsHlp("PART_G") & ", " _
                                     & "" & rsHlp("PART_NG") & "," & rsHlp("OIL_G") & "," & rsHlp("OIL_NG") & "," & rsHlp("MTRL") & "," & rsHlp("LABOR") & "," & rsHlp("BODY") & "," & rsHlp("SBLT") & "," & rsHlp("MET") & "," & rsHlp("TOTAL") & ",'" & rsHlp("JOB") & "','" & rsHlp("REP1") & "','" & rsHlp("REP2") & "','" & rsHlp("REP3") & "','" & rsHlp("REP4") & "'," _
                                     & "'" & rsHlp("REP5") & "','" & rsHlp("REP6") & "','" & rsHlp("REP7") & "','" & rsHlp("MEKANIK1") & "','" & rsHlp("MEKANIK2") & "','" & rsHlp("MEKANIK3") & "'," _
                                     & "'" & rsHlp("MEKANIK4") & "','" & rsHlp("MEKANIK5") & "','" & rsHlp("MEKANIK6") & "','" & rsHlp("MEKANIK7") & "','" & rsHlp("MEKANIK8") & "','" & rsHlp("KLPKERJA1") & "','" & rsHlp("KLPKERJA2") & "','" & rsHlp("KLPKERJA3") & "','" & rsHlp("KLPKERJA4") & "','" & rsHlp("KLPKERJA5") & "','" & rsHlp("KLPKERJA6") & "'," _
                                     & "'" & rsHlp("KLPKERJA7") & "','" & rsHlp("KLPKERJA8") & "','" & rsHlp("NO_WO") & "','" & rsHlp("NOINV") & "','" & rsHlp("NOFAK") & "','" & rsHlp("STATUS") & "','" & rsHlp("PAID") & "','" & rsHlp("DISC") & "','" & rsHlp("DISCT") & "','" & rsHlp("REMARK1") & "','" & rsHlp("FN_RETURN") & "','" & rsHlp("FS_RC") & "'," _
                                     & "'" & rsHlp("FS_KD_DEALER") & "','" & rsHlp("FS_NM_DEALER") & "','" & rsHlp("TIPEWO") & "','" & rsHlp("KODETIPE") & "', '" & rsHlp("SBE") & "','" & rsHlp("FS_KD_PLG") & "','" & rsHlp("FS_KD_WIL") & "','" & rsHlp("THS") & "','" & rsHlp("MOU") & "','" & rsHlp("PDS") & "','" & rsHlp("FS_LOG_DEL") & "','" & rsHlp("FD_TGL_DEL") & "', " _
                                     & "'" & rsHlp("ALASANDEL") & "','" & Format(rsHlp("INVDATE"), "MM/DD/YYYY") & "')")
                    End If
                End If
                rsHlp.MoveNext
            Loop
        End If
        End If
         
    End If
End Sub

Private Sub saveStall()
    Dim rsExist As ADODB.Recordset
    Dim rsSTemp As ADODB.Recordset
    Dim nKet As String
    Dim nTD As Double
    Dim doup As Integer
    Dim filterStall As String
    
    If lblEdit.Visible = True Then
        filterStall = noPol & " |" & text(1) & "%'"
    Else
        filterStall = noPol & " |%'"
    End If
    
    
    Set rsSTemp = objO.FetchData("SELECT noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, keterangan, CHARINDEX('||',KETERANGAN) AS TNDA,adjust FROM TempStall WITH (NOLOCK) WHERE tanggal = '" & Format(text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
    If rsSTemp.RecordCount > 0 Then
        nKet = rsSTemp.Fields(7).value
        nKet = Left(nKet, InStr(nKet, "|")) & text(1) & Right(nKet, Len(nKet) - InStr(nKet, "| ") + 1)
        nTD = CDbl(rsSTemp("TNDA"))
    End If
    doup = 0
    If Not rsSTemp.EOF Then
        If nTD > 0 Then
            Set rsSTemp = objO.FetchData("SELECT noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, '" & nKet & "' as keterangan ,FLAG,adjust FROM TempStall WITH (NOLOCK) WHERE tanggal = '" & Format(text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
            doup = 1
        Else
            Set rsSTemp = objO.FetchData("SELECT noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, keterangan,FLAG,adjust FROM TempStall WITH (NOLOCK) WHERE tanggal = '" & Format(text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
        End If
    Else
        Set rsSTemp = objO.FetchData("SELECT noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, keterangan,FLAG,adjust FROM TempStall WITH (NOLOCK) WHERE tanggal = '" & Format(text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
    End If
    On Error GoTo errStall
    If Not rsSTemp.EOF Then
        Call objHT.DeleteData_Bengkel("DELETE FROM stall WHERE tanggal = '" & Format(text(23), "mm/dd/yyyy") & "' AND keterangan like '" & noPol & " |" & text(1) & "%'")
            Call objHT.InsertData_BengkelOnly("INSERT INTO stall (noStall, tanggal, jamMulai, jamSelesai, rate, [group], mech, keterangan, FLAG,adjust)  VALUES " _
                       & "('" & rsSTemp("NOSTALL") & "','" & Format(rsSTemp("TANGGAL"), "MM/DD/YYYY") & "', '" & rsSTemp("JAMMULAI") & "', " _
                       & "'" & rsSTemp("JAMSELESAI") & "', '" & Replace(rsSTemp("RATE"), ",", ".") & "', '" & rsSTemp("GROUP") & "', '" & rsSTemp("MECH") & "', '" & nKet & "', '" & rsSTemp("FLAG") & "','" & rsSTemp("adjust") & "')")
            Call objHT.DeleteData_Bengkel("DELETE tempStall WHERE keterangan like '" & filterStall & "")
        If doup = 1 Then Call objHT.UpdateData_Bengkel("UPDATE TEMPSTALL SET KETERANGAN = Left(KETERANGAN, CHARINDEX('|',KETERANGAN)) + '" & text(1) & "' + Right(KETERANGAN, Len(KETERANGAN) - CHARINDEX('||',KETERANGAN))  WHERE tanggal = '" & Format(text(23), "mm/dd/yyyy") & "' AND keterangan like '" & filterStall & "")
    End If
    Exit Sub
errStall:
    err.Raise err.Number, err.Source, err.Description
End Sub

Private Sub rtjFU(nChs As String)
    ' UPDATE STATUS BOOK DI DTLFU
    Call objHT.UpdateData_Bengkel("UPDATE DTLFU SET BOOK = '1' WHERE NO_CHASSIS = '" & Trim(nChs) & "' AND TANGGAL = '" & Format(isnow, "MM/DD/YYYY") & "'")
End Sub

Private Sub Text1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Call CekTombol(KeyCode)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpCase(KeyAscii)
End Sub

Private Sub VScroll1_Change()
    If VScroll1.value > 0 Then
        Frame2.Top = -VScroll1.value + 200
    ElseIf VScroll1.value = 0 Then
        Frame2.Top = 0
    End If
End Sub

Private Sub VScroll1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Not simpan Then
            If MsgBox("Data Belum disimpan, menampilkan data baru ?", vbYesNo, "Konfirmasi") = vbNo Then
                Exit Sub
            End If
        End If
        Call hapusForm
        text(0).SetFocus
    End If
End Sub

Private Sub VScroll1_Scroll()
    If VScroll1.value > 0 Then
        Frame2.Top = -VScroll1.value + 200
    ElseIf VScroll1.value = 0 Then
        Frame2.Top = 0
    End If
End Sub

Private Sub subletGrid_DblClick()
Dim a As Integer

    If barisSublet > 1 And subletGrid.TextMatrix(subletGrid.row, 1) <> "" Then
        a = subletGrid.row
        
        With subletGrid
            text(79) = Trim(.TextMatrix(a, 1))
            text(80) = CDbl(.TextMatrix(a, 2))
            text(81) = Trim(.TextMatrix(a, 3))
            text(82) = CDbl(.TextMatrix(a, 4)) '* (100 / (100 + gUpSublet))
            chSublet = IIf(CDbl(.TextMatrix(a, 5)) = 0, Unchecked, Checked)
        End With
        
        For i = a To barisSublet
            With subletGrid
                .TextMatrix(i, 0) = .TextMatrix(i + 1, 0)
                .TextMatrix(i, 1) = .TextMatrix(i + 1, 1)
                .TextMatrix(i, 2) = .TextMatrix(i + 1, 2)
                .TextMatrix(i, 3) = .TextMatrix(i + 1, 3)
                .TextMatrix(i, 4) = .TextMatrix(i + 1, 4)
                .TextMatrix(i, 5) = .TextMatrix(i + 1, 5)
                .TextMatrix(i, 6) = .TextMatrix(i + 1, 6)
            End With
        Next
        
        text(52) = Format(CDbl(text(52)) - (CDbl(text(80)) * CDbl(text(82))), "#0,#")
        text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#0,#")
                
        barisSublet = barisSublet - 1
        Call hitungEstimasiSelesai
        text(79).SetFocus
    End If
End Sub

Private Sub subletGrid_GotFocus()
Dim bolehTambah As Boolean
    
    If Trim(text(79)) <> "" And Trim(text(80)) <> "" And Trim(text(81)) <> "" And Trim(text(82)) <> "" Then
    
        For i = 0 To baris
            If subletGrid.TextMatrix(i, 1) = Trim(text(79)) Then
                bolehTambah = False
                text(79) = ""
                Exit Sub
            Else
                bolehTambah = True
            End If
        Next
        If bolehTambah Then
            If Log.isEmStall = 1 Then
                frmEstimasiSublet.Visible = True
                text(91) = "00:00"
            Else
                estimasiSublet = "00:00"
                If MaxBrs >= barisSublet Then
                    With subletGrid
                        .TextMatrix(barisSublet, 1) = Trim(text(79))                                ' NAMA SUBLET
                        .TextMatrix(barisSublet, 2) = CDbl(text(80))                                ' QTY
                        .TextMatrix(barisSublet, 3) = Trim(text(81))                                ' SATUAN
                        .TextMatrix(barisSublet, 4) = Format(CDbl(text(82)), "#0,#")                ' * ((100 + gUpSublet) / 100)    ' HARGA
                        .TextMatrix(barisSublet, 5) = IIf(chSublet = Checked, 1, 0)
                        .TextMatrix(barisSublet, 6) = estimasiSublet
                    End With
                    barisSublet = barisSublet + 1
                    text(52) = Format(CDbl(text(52)) + (CDbl(text(80)) * CDbl(text(82))), "#0,#")
                    text(53) = Format(CDbl(text(39)) + CDbl(text(40)) + CDbl(text(52)) + CDbl(text(55)), "#0,#")
                    
                    text(79) = ""
                    text(80) = 0
                    text(81) = ""
                    text(82) = 0
                    chSublet = Unchecked
                    Call hitungEstimasiSelesai
                    text(79).SetFocus
                End If
                '
            End If
        End If
    End If
End Sub


Public Function GetDelIndex(ModilID As String) As String
    Dim acak As Double
    Randomize Timer
    acak = Int(Rnd * 99999)
    GetDelIndex = ModilID & acak
End Function

Public Sub hitungEstimasiSelesai()
    Dim rsJob As ADODB.Recordset
    Dim rsSublet As ADODB.Recordset
    Dim totalJam As Integer
    Dim totalSublet As Integer
    Dim totalJob As Double
    Dim hasilJam As String
    Dim nRate As Double
    Dim keterangan As String
    Dim jamTemSkrg As Integer
    Dim menitTemSkrg As Integer
    Dim jamMulai As String
    Dim jamTempSublet As Integer
    Dim menitTempSublet As Integer
    Dim totalTempSublet As Integer
    If Trim(text(58)) <> ">>:>>" Then
        jamMulai = Trim(text(58))
    Else
        jamMulai = Trim(text(24))
    End If
    
    totalSublet = 0
    
    For i = 1 To barisSublet
        With subletGrid
            If .TextMatrix(i, 1) <> "" Then
                jamTempSublet = CInt(Left(Trim(.TextMatrix(i, 6)), InStr(.TextMatrix(i, 6), ":") - 1))
                menitTempSublet = Mid(Trim(.TextMatrix(i, 6)), InStr(Trim(.TextMatrix(i, 6)), ":") + 1, Len(Trim(.TextMatrix(i, 6))) - InStr(Trim(.TextMatrix(i, 6)), ":"))
                totalTempSublet = (jamTempSublet * 60) + menitTempSublet
                totalSublet = totalSublet + totalTempSublet
            End If
        End With
    Next
        
    totalJob = 0
    nRate = 0
    For i = 1 To baris - 1
        If Trim(DGrid.TextMatrix(i, 1)) <> "" Then
            Set rsTemp = objO.FetchData("SELECT RATE FROM actRAte WITH (NOLOCK) WHERE MDL = '" & text(10) & "' AND FS_ITEM = '" & Trim(DGrid.TextMatrix(i, 1)) & "' AND FS_LOG_DEL = ''")
            If Not rsTemp.EOF Then
                nRate = nRate + CDbl(rsTemp("RATE")) 'DGrid.TextMatrix(i, 2)
            End If
        End If
    Next
    
   ' totalJob = 0
    If Trim(text(1)) <> "" Then
        keterangan = Trim(text(0)) & "-" & Trim(text(56)) & "-" & Trim(text(57)) & " |" & Trim(text(1))
    Else
        keterangan = Trim(text(0)) & "-" & Trim(text(56)) & "-" & Trim(text(57)) & " |"
    End If
    If chEM = Checked Then
        totalJob = 0.5 * 60
    ElseIf chSBNP = Checked Then
        totalJob = 1 * 60
    Else
        
        Set rsJob = objO.FetchData("SELECT adjust,jammulai FROM tempstall WITH (NOLOCK) WHERE flag NOT IN ('k','x') and keterangan like '" & keterangan & "%'")
        If Not rsJob.EOF Then
            jamTemSkrg = CInt(Left(Trim(rsJob("adjust")), InStr(rsJob("adjust"), ":") - 1))
            menitTemSkrg = Mid(Trim(rsJob("adjust")), InStr(Trim(rsJob("adjust")), ":") + 1, Len(Trim(rsJob("adjust"))) - InStr(Trim(rsJob("adjust")), ":"))
            totalJob = (jamTemSkrg * 60) + menitTemSkrg
            jamMulai = Trim(rsJob("jammulai"))
        Else
            Set rsJob = objO.FetchData("SELECT adjust,jammulai FROM stall WITH (NOLOCK) WHERE flag NOT IN ('k','x') and keterangan like '" & keterangan & "%'")
            If Not rsJob.EOF Then
                jamTemSkrg = CInt(Left(Trim(rsJob("adjust")), InStr(rsJob("adjust"), ":") - 1))
                menitTemSkrg = Mid(Trim(rsJob("adjust")), InStr(Trim(rsJob("adjust")), ":") + 1, Len(Trim(rsJob("adjust"))) - InStr(Trim(rsJob("adjust")), ":"))
                totalJob = (jamTemSkrg * 60) + menitTemSkrg
                jamMulai = Trim(rsJob("jammulai"))
            End If
        End If
'        If Trim(text(114)) <> "" Then
'            totalJob = CDbl(CDbl(Trim(text(114))) * 60)
'        End If
    End If
    
    totalMenit = CInt(totalSublet) + CInt(totalJob) + CInt(Log.estInv)
    hasilJam = hitungTotalPenyerahan(CInt(totalMenit), jamMulai, text(23))
        
    text(45) = Left(hasilJam, InStr(hasilJam, "|") - 1)
    text(59) = TglDDMMYYYY(Right(hasilJam, Len(hasilJam) - InStr(hasilJam, "|")))
    
    text(83) = Left(hasilJam, InStr(hasilJam, "|") - 1)
    text(84) = TglDDMMYYYY(Right(hasilJam, Len(hasilJam) - InStr(hasilJam, "|")))
    
End Sub
